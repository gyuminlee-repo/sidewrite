/* global Word */
import {
  normalizeParaText,
  paragraphHash,
  computeAfterText,
  countRevisionMarks,
  trackedRuns,
  isToolComment,
  planBackupDelete,
} from "./review-view.js";
//
// Word tool implementations + Word-specific helpers.
//
// Each `tool*` function is invoked from the dispatcher in taskpane.js when a
// matching `office_*` tool_call arrives from the daemon. They all run inside
// Word.run(...) so Office.js manages the request context lifecycle. The
// helpers above (paragraph addressing, withTrackChanges, snapshot, heading
// detection) are private to this module — taskpane.js doesn't need them.

// ---------------------------------------------------------------------------
// Paragraph addressing
//
// We use Word's `paragraph.uniqueLocalId` (the OOXML w:paraId) — stable
// across insertions/deletions, which is what the agent needs when it does
// "read these paragraphs, then edit p7" across multiple tool calls. Available
// in WordApi 1.6+ (every recent Microsoft 365 Word).
//
// On older Word builds where uniqueLocalId is unsupported, we fall back to
// position-based IDs (`p${index}`). These are NOT stable across structural
// edits — when the fallback is in play, snapshotParagraphs() flags the mode
// and tools should re-read before each operation rather than caching IDs.
// ---------------------------------------------------------------------------
const fallbackId = (index) => `p${index}`;
const parseFallbackId = (id) => {
  const m = /^p(\d+)$/.exec(id);
  return m ? parseInt(m[1], 10) : null;
};

async function getParagraphsWithIds(context) {
  const paragraphs = context.document.body.paragraphs;
  // CRITICAL: load every per-item property we'll need in a SINGLE
  // paragraphs.load() before the first sync. A second load/sync on the same
  // paragraphs collection silently un-loads the previously-loaded scalars,
  // so every later `p.text` read throws "property 'text' is not available."
  // changeTrackingMode is a document-level scalar, not a paragraph property —
  // safe to piggyback on the same sync, no scalar-refresh issue.
  try {
    context.document.load("changeTrackingMode");
    paragraphs.load(
      "items/text, items/style, items/uniqueLocalId, items/listItemOrNullObject/listString",
    );
    await context.sync();
    if (paragraphs.items.length > 0 && !paragraphs.items[0].uniqueLocalId) {
      throw new Error("uniqueLocalId not populated");
    }
    const trackChangesOn = String(context.document.changeTrackingMode) !== "Off";
    return { paragraphs, idMode: "uniqueLocalId", trackChangesOn };
  } catch {
    context.document.load("changeTrackingMode");
    paragraphs.load("items/text, items/style");
    await context.sync();
    const trackChangesOn = String(context.document.changeTrackingMode) !== "Off";
    return { paragraphs, idMode: "index", trackChangesOn };
  }
}

// Pull a literally-typed bracketed paragraph number (e.g. "[0034]") off the
// front of a paragraph's text. Fallback for docs that number paragraphs by
// typing, not by Word list-numbering.
function extractBracketedNumber(text) {
  if (!text) return null;
  const m = /^\s*(\[[^\]]+\])/.exec(text);
  return m ? m[1] : null;
}

// The user-facing number for a paragraph. Preferred: Word's rendered list
// number (e.g. "[0034]") via Paragraph.listItemOrNullObject.listString —
// NOT present in paragraph.text, so a regex on text alone misses every
// auto-numbered doc. Fallback: a literally-typed [NNNN] prefix in the text.
// Used so the agent can cite paragraph numbers the user actually sees on
// the page, instead of opaque uniqueLocalId hex strings.
function listNumberOf(p, text) {
  try {
    const li = p.listItemOrNullObject;
    if (li && !li.isNullObject) {
      const s = li.listString;
      if (s && String(s).trim()) return String(s).trim();
    }
  } catch {
    /* listItemOrNullObject not loaded on this build / fallback path */
  }
  return extractBracketedNumber(text);
}

// Run a batch of getReviewedText(...) calls on ranges, working around
// office-js#3691 (getReviewedText scrolls the document to the change). Save
// the user's selection once, do the reads, restore the selection. Returns
// an array of reviewed-text strings aligned with `ranges`, or null on any
// failure (older builds without WordApi 1.4 — caller should fall back to
// raw .text). office-js#5188 makes getTrackedChanges() unusable as an
// alternative (TrackedChange.text returns empty for Deleted), so this is
// the only supported path.
async function getReviewedTextBatch(context, ranges) {
  if (!ranges || ranges.length === 0) return [];
  const saved = context.document.getSelection();
  try {
    const results = ranges.map((r) => r.getReviewedText(Word.ChangeTrackingVersion.current));
    await context.sync();
    try {
      saved.select();
      await context.sync();
    } catch {
      /* best effort — restore failure shouldn't mask the read */
    }
    return results.map((r) => r.value);
  } catch {
    try {
      saved.select();
      await context.sync();
    } catch {
      /* best effort */
    }
    return null;
  }
}

function getId(paragraph, index, idMode) {
  return idMode === "uniqueLocalId" ? paragraph.uniqueLocalId : fallbackId(index);
}

function findIndexById(paragraphs, id, idMode) {
  if (idMode === "uniqueLocalId") {
    return paragraphs.items.findIndex((p) => p.uniqueLocalId === id);
  }
  const i = parseFallbackId(id);
  if (i === null || i < 0 || i >= paragraphs.items.length) return -1;
  return i;
}

// Index in paragraphs.items of each paragraph in `firsts` (a Paragraph taken
// from a range, a comment or the selection; null gives -1). With
// uniqueLocalId the id decides. Without it the text does, and when several
// paragraphs have that text, the place: the paragraph's own range is
// compared with each of them (one more sync, only then). Looking a twin up
// by its text alone gave the first one (Codex review 14).
const SAME_PLACE = ["Equal", "Inside", "InsideStart", "InsideEnd"];
export async function paragraphIndexes(context, paragraphs, idMode, firsts) {
  const out = [];
  const open = [];
  firsts.forEach((first, n) => {
    if (!first) {
      out[n] = -1;
    } else if (idMode === "uniqueLocalId") {
      out[n] = paragraphs.items.findIndex((p) => p.uniqueLocalId === first.uniqueLocalId);
    } else {
      const same = [];
      paragraphs.items.forEach((p, i) => {
        if (p.text === first.text) same.push(i);
      });
      if (same.length <= 1) {
        out[n] = same.length ? same[0] : -1;
      } else {
        const own = first.getRange("Whole");
        const where = same.map((i) =>
          own.compareLocationWith(paragraphs.items[i].getRange("Whole")),
        );
        open.push({ n, same, where });
      }
    }
  });
  if (open.length) {
    await context.sync();
    for (const { n, same, where } of open) {
      const k = where.findIndex((w) => SAME_PLACE.includes(w.value));
      out[n] = k === -1 ? -1 : same[k];
    }
  }
  return out;
}

// Run `body` with the document's changeTrackingMode set to TrackAll if
// `track_changes` is truthy, and unconditionally restore it in `finally` so
// that exceptions thrown mid-edit don't leave Word stuck in Track All mode.
// The restore-sync is wrapped in a try/catch so a failed restore doesn't
// shadow the original error.
async function withTrackChanges(context, track_changes, body) {
  let prevTracking = null;
  if (track_changes) {
    context.document.load("changeTrackingMode");
    await context.sync();
    prevTracking = context.document.changeTrackingMode;
    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
  }
  try {
    return await body();
  } finally {
    if (track_changes && prevTracking !== null) {
      try {
        context.document.changeTrackingMode = prevTracking;
        await context.sync();
      } catch {
        /* swallow — don't shadow original error */
      }
    }
  }
}

// ---------------------------------------------------------------------------
// Heading / section helpers
// ---------------------------------------------------------------------------
function headingLevel(style) {
  // Returns the heading level if the style name *contains* "Heading <N>".
  // Many organizations use prefixed style names like "Acme-Heading 2" or
  // "Firm-Heading-1" — we want to recognize those as headings too, not just
  // Word's built-in "Heading 1" / "Heading 2".
  if (!style) return null;
  const m = /heading[\s_-]*(\d)/i.exec(style);
  return m ? parseInt(m[1], 10) : null;
}

// Build the "view" the agent sees: list of {id, style, text, [number]}.
// `number` is the user-facing paragraph number (Word's rendered list string,
// or a literally-typed [NNNN]) when present — omitted otherwise, so the
// agent's `if (p.number)` checks stay clean.
function snapshotParagraphs(paragraphs, idMode) {
  return paragraphs.items.map((p, i) => {
    const text = p.text;
    const number = listNumberOf(p, text);
    return {
      index: i,
      id: getId(p, i, idMode),
      style: p.style,
      text,
      ...(number ? { number } : {}),
    };
  });
}

// Find the index of a heading paragraph by case-insensitive exact text match.
function findHeadingIndex(snapshot, headingText) {
  const target = headingText.trim().toLowerCase();
  return snapshot.findIndex(
    (p) => headingLevel(p.style) !== null && p.text.trim().toLowerCase() === target,
  );
}

// Find the canonical body paragraph for the section that contains the
// insert position `fromIdx`.
//
//   1. Walk BACKWARD from fromIdx-1 to the nearest heading — the top of
//      the section the insert sits in.
//   2. From that heading, walk FORWARD for the first body paragraph in the
//      section, stopping at the next heading.
//   3. If there's no heading above the insert, or the section has no body
//      paragraph above the next heading, fall back to the first non-heading
//      paragraph anywhere in the doc.
//
// Returns `{ idx, inSection }`:
//   - `idx`       — index into `snapshot` of the reference paragraph (or -1).
//   - `inSection` — true when `idx` was found *inside* the insert's own
//     section (steps 1–2); false when it came from the doc-wide fallback
//     (step 3).
//
// Why `inSection` matters: the style NAME is a safe thing to copy even from
// the fallback (low blast radius — old behavior). But the reference
// paragraph's *direct* formatting (indent / alignment / spacing) is only
// representative if it's from the same section. The doc-wide fallback might
// be a centered title block or an indented address line, and cloning its
// direct formatting onto a body insert would be visibly wrong. Callers use
// `inSection` to gate format-cloning while still always taking the style
// name.
function findBodyReferenceIndex(snapshot, fromIdx) {
  // Step 1: find the section's heading by walking backward.
  let sectionStart = -1;
  for (let i = fromIdx - 1; i >= 0; i--) {
    if (snapshot[i] && headingLevel(snapshot[i].style) !== null) {
      sectionStart = i;
      break;
    }
  }

  // Step 2: from the heading, the first body paragraph in THIS section.
  // Stop at the next heading.
  if (sectionStart !== -1) {
    for (let i = sectionStart + 1; i < snapshot.length; i++) {
      if (!snapshot[i]) continue;
      if (headingLevel(snapshot[i].style) !== null) break;
      return { idx: i, inSection: true };
    }
  }

  // Step 3: doc-wide fallback — style name only, NOT a format-clone source.
  for (let i = 0; i < snapshot.length; i++) {
    if (snapshot[i] && headingLevel(snapshot[i].style) === null) {
      return { idx: i, inSection: false };
    }
  }
  return { idx: -1, inSection: false };
}

function inferBodyStyle(snapshot, fromIdx) {
  const { idx } = findBodyReferenceIndex(snapshot, fromIdx);
  return idx === -1 ? null : snapshot[idx].style || null;
}

// Paragraph-format properties cloned from a reference body paragraph onto
// inserted ones. Many templates set these directly (not via the named
// style), so a style-name copy alone leaves inserts left-aligned with no
// first-line indent and no inter-paragraph spacing.
const CLONED_PARA_FORMAT_PROPS = [
  "alignment",
  "firstLineIndent",
  "leftIndent",
  "rightIndent",
  "lineSpacing",
  "spaceBefore",
  "spaceAfter",
  "lineUnitBefore",
  "lineUnitAfter",
];

// Queue a load of the format props off a reference paragraph. Call before
// the next context.sync().
function loadParagraphFormat(refPara) {
  refPara.load(CLONED_PARA_FORMAT_PROPS.join(", "));
}

// Snapshot a synced reference paragraph's format values into a plain
// object. Decouples the captured formatting from the live proxy —
// necessary for office_replace_section, where the reference paragraph is
// deleted before the new paragraphs are inserted.
function snapshotParagraphFormat(refPara) {
  const out = {};
  for (const prop of CLONED_PARA_FORMAT_PROPS) {
    try {
      const v = refPara[prop];
      if (v !== undefined && v !== null) out[prop] = v;
    } catch {
      /* prop unsupported on this build */
    }
  }
  return out;
}

// Apply a snapshotted format object onto a target paragraph.
function applyParagraphFormat(targetPara, fmt) {
  if (!fmt) return;
  for (const prop of CLONED_PARA_FORMAT_PROPS) {
    const v = fmt[prop];
    if (v !== undefined && v !== null) {
      try {
        targetPara[prop] = v;
      } catch {
        /* prop unsupported on this build */
      }
    }
  }
}

// Find the end of a section starting at headingIdx (exclusive): next paragraph
// with a heading style of same-or-higher rank (i.e. lower or equal level).
function findSectionEnd(snapshot, headingIdx) {
  const startLevel = headingLevel(snapshot[headingIdx].style);
  for (let i = headingIdx + 1; i < snapshot.length; i++) {
    const lvl = headingLevel(snapshot[i].style);
    if (lvl !== null && lvl <= startLevel) return i;
  }
  return snapshot.length;
}

// ---------------------------------------------------------------------------
// Tool: office_get_selection
// ---------------------------------------------------------------------------
export async function toolGetSelection() {
  return await Word.run(async (context) => {
    const sel = context.document.getSelection();
    sel.load("text");
    const selParas = sel.paragraphs;
    context.document.load("changeTrackingMode");

    // Try to load uniqueLocalId directly on the selected paragraphs — that
    // gives us stable IDs without any text-matching against the full doc
    // (which is unsafe whenever the doc has repeated/boilerplate paragraphs).
    // Pack ALL per-item properties (incl. listItemOrNullObject/listString
    // for rendered numbers) into one load before the first sync — a second
    // load/sync on the same paragraphs collection would silently un-load
    // the previously-loaded scalars and make every subsequent .text read
    // throw.
    let idMode = "uniqueLocalId";
    try {
      selParas.load(
        "items/text, items/style, items/uniqueLocalId, items/listItemOrNullObject/listString",
      );
      await context.sync();
      if (selParas.items.length > 0 && !selParas.items[0].uniqueLocalId) {
        throw new Error("uniqueLocalId not populated");
      }
    } catch {
      selParas.load("items/text, items/style");
      await context.sync();
      idMode = "index";
    }
    const trackChangesOn = String(context.document.changeTrackingMode) !== "Off";

    // When track changes is on, .text concatenates tracked deletions with
    // their inserted replacements — the agent then sees "clean + stale"
    // and tries to "fix" the duplicate by re-replacing it, stacking more
    // tracked changes every turn. Use getReviewedText(current) on the
    // selection AND each selected paragraph so the agent sees what the
    // user sees with track changes accepted. Skipped when track changes
    // is off (raw .text is correct then — zero behavior change).
    let reviewedSelText = null;
    let reviewedPerPara = null;
    if (trackChangesOn) {
      const ranges = [sel, ...selParas.items.map((p) => p.getRange())];
      const reviewed = await getReviewedTextBatch(context, ranges);
      if (reviewed) {
        reviewedSelText = reviewed[0];
        reviewedPerPara = reviewed.slice(1);
      }
    }

    const selSnapshot = selParas.items.map((p, i) => {
      const text = reviewedPerPara ? reviewedPerPara[i] : p.text;
      const number = listNumberOf(p, p.text);
      return {
        // When uniqueLocalId is unavailable, return null rather than guessing
        // an index by text-matching — duplicate paragraphs in any doc make
        // text-match unreliable. The agent should fall back to text references.
        id: idMode === "uniqueLocalId" ? p.uniqueLocalId : null,
        style: p.style,
        text,
        ...(number ? { number } : {}),
      };
    });

    const finalText = reviewedSelText != null ? reviewedSelText : sel.text;
    return {
      text: finalText,
      is_empty: !finalText || finalText.length === 0,
      paragraphs: selSnapshot,
      addressing: idMode,
    };
  });
}

// ---------------------------------------------------------------------------
// Tool: office_read_paragraphs
// ---------------------------------------------------------------------------
export async function toolReadParagraphs({ ids, heading_section, range, preview }) {
  return await Word.run(async (context) => {
    const { paragraphs, idMode, trackChangesOn } = await getParagraphsWithIds(context);
    const snapshot = snapshotParagraphs(paragraphs, idMode);

    let picked;
    // Truncation is on by default in the no-args ("dump the whole doc")
    // mode — that's an orientation pass, not a content read. Targeted forms
    // (`ids`, `heading_section`, `range`) always return full text. The agent
    // can also explicitly force `preview: false` in the no-args mode for a
    // full dump (heavy, but escape-hatch valid).
    let truncate = false;
    // Track-changes-safe reads are only worth doing in targeted modes — the
    // no-args preview is orientation only, calling getReviewedText on every
    // paragraph would thrash the document view (office-js#3691) for no gain.
    let isTargeted = false;
    if (ids && ids.length > 0) {
      const set = new Set(ids);
      picked = snapshot.filter((p) => set.has(p.id));
      isTargeted = true;
    } else if (heading_section) {
      const startIdx = findHeadingIndex(snapshot, heading_section);
      if (startIdx === -1) {
        throw new Error(`Heading not found: "${heading_section}"`);
      }
      const endIdx = findSectionEnd(snapshot, startIdx);
      picked = snapshot.slice(startIdx, endIdx);
      isTargeted = true;
    } else if (range) {
      const [s, e] = range;
      picked = snapshot.slice(s, e);
      isTargeted = true;
    } else {
      // Default: outline-ish view. Word documents in the wild use varied
      // heading conventions (built-in Heading N, custom style names, all-caps
      // direct formatting), so we don't pre-filter — return every paragraph
      // with its style and a truncated text preview. The agent picks out the
      // headings.
      picked = snapshot;
      truncate = preview !== false;
    }

    // With track changes on, raw .text concatenates a tracked deletion's
    // original run with its inserted replacement — the agent reads back
    // "clean + stale" and tries to "fix" the duplicate, stacking more
    // tracked changes every turn. getReviewedText(current) returns what
    // the user sees with the deletions accepted, which is what the agent
    // should reason about. Skip in the preview pass (see above).
    let reviewedByPicked = null;
    if (isTargeted && trackChangesOn && picked.length > 0) {
      const ranges = picked.map((p) => paragraphs.items[p.index].getRange());
      reviewedByPicked = await getReviewedTextBatch(context, ranges);
    }

    // Preview length: long enough to surface most in-paragraph references
    // (numbers, defined terms, citations) while still bounding the response
    // size on large docs. Was 117 originally — too aggressive; the agent
    // treated previews as full reads and missed mid-paragraph content.
    const PREVIEW_LEN = 500;

    return {
      paragraphs: picked.map((p, k) => {
        const full = reviewedByPicked && reviewedByPicked[k] != null ? reviewedByPicked[k] : p.text;
        const isTruncated = truncate && full.length > PREVIEW_LEN;
        return {
          id: p.id,
          style: p.style,
          ...(p.number ? { number: p.number } : {}),
          text: isTruncated ? full.slice(0, PREVIEW_LEN - 1) + "…" : full,
          // Structured truncation flag so the agent doesn't have to detect
          // "…" suffixes. When `truncated` is true, the paragraph has more
          // content; re-read it via `ids: [<id>]` (always full text) to get
          // it. `full_length` lets the agent decide if a re-read is worth
          // the round-trip.
          ...(isTruncated ? { truncated: true, full_length: full.length } : {}),
        };
      }),
      total_in_doc: snapshot.length,
      addressing: idMode,
      // When truncation is active across the response, surface it at the
      // top level too. Cheap signal for the agent's next-action decision.
      ...(truncate ? { preview_mode: true, preview_limit: PREVIEW_LEN } : {}),
    };
  });
}

// ---------------------------------------------------------------------------
// Tool: office_insert_paragraphs
// ---------------------------------------------------------------------------
export async function toolInsertParagraphs({
  after,
  content,
  track_changes,
  style_per_para,
  provenance_comment,
}) {
  if (!content || content.length === 0) throw new Error("content must be non-empty");

  return await Word.run(async (context) => {
    return await withTrackChanges(context, track_changes, async () => {
      const { paragraphs, idMode } = await getParagraphsWithIds(context);
      const snapshot = snapshotParagraphs(paragraphs, idMode);

      let anchorIdx;
      if (after?.id) {
        anchorIdx = findIndexById(paragraphs, after.id, idMode);
        if (anchorIdx === -1) {
          throw new Error(`Anchor paragraph not found: ${after.id}`);
        }
      } else if (after?.heading) {
        anchorIdx = findHeadingIndex(snapshot, after.heading);
        if (anchorIdx === -1) throw new Error(`Heading not found: "${after.heading}"`);
      } else {
        throw new Error("after must specify either id or heading");
      }

      const anchor = paragraphs.items[anchorIdx];
      // If no explicit style is given, infer the section's body style so we
      // don't inherit a heading's formatting. Start one past the anchor (the
      // first paragraph the new content will sit alongside).
      const useInferred = !(style_per_para && style_per_para.length === content.length);
      const ref = useInferred
        ? findBodyReferenceIndex(snapshot, anchorIdx + 1)
        : { idx: -1, inSection: false };
      const refIdx = ref.idx;
      const inferredStyle = refIdx === -1 ? null : snapshot[refIdx].style || null;

      // Clone the reference body paragraph's *direct* formatting (indent,
      // alignment, spacing) — templates apply these manually, so the style
      // name alone isn't enough. ONLY when the reference is from the
      // insert's own section: the doc-wide fallback could be a title block
      // or address line, and cloning its direct formatting onto a body
      // insert would be visibly wrong. The style name is still taken from
      // the fallback (low blast radius); only format-cloning is gated.
      let refFmt = null;
      if (refIdx !== -1 && ref.inSection) {
        const refPara = paragraphs.items[refIdx];
        loadParagraphFormat(refPara);
        await context.sync();
        refFmt = snapshotParagraphFormat(refPara);
      }

      let cursor = anchor;
      const inserted = [];
      for (let i = 0; i < content.length; i++) {
        const p = cursor.insertParagraph(content[i], Word.InsertLocation.after);
        const explicit = style_per_para && style_per_para[i];
        if (explicit) p.style = explicit;
        else if (inferredStyle) p.style = inferredStyle;
        if (refFmt && !explicit) applyParagraphFormat(p, refFmt);
        inserted.push(p);
        cursor = p;
      }

      if (provenance_comment && inserted.length > 0) {
        inserted[0].getRange().insertComment(provenance_comment);
      }

      await context.sync();

      let newIds;
      if (idMode === "uniqueLocalId") {
        for (const p of inserted) p.load("uniqueLocalId");
        await context.sync();
        newIds = inserted.map((p) => p.uniqueLocalId);
      } else {
        newIds = inserted.map((_, k) => fallbackId(anchorIdx + 1 + k));
      }

      return {
        inserted_count: inserted.length,
        new_para_ids: newIds,
        addressing: idMode,
      };
    });
  });
}

// ---------------------------------------------------------------------------
// Tool: office_replace_paragraphs
// ---------------------------------------------------------------------------
export async function toolReplaceParagraphs({
  ids,
  content,
  track_changes,
  style_per_para,
  provenance_comment,
}) {
  if (!ids || ids.length === 0) throw new Error("ids must be non-empty");
  if (!content || content.length === 0) throw new Error("content must be non-empty");
  if (ids.length !== content.length) {
    throw new Error(
      `ids.length (${ids.length}) must equal content.length (${content.length}). To grow or shrink a section, use office_insert_paragraphs or office_replace_section.`,
    );
  }

  return await Word.run(async (context) => {
    return await withTrackChanges(context, track_changes, async () => {
      const { paragraphs, idMode } = await getParagraphsWithIds(context);

      const targets = ids.map((id) => {
        const idx = findIndexById(paragraphs, id, idMode);
        if (idx === -1) {
          throw new Error(`Paragraph not found: ${id}`);
        }
        return { id, idx, paragraph: paragraphs.items[idx] };
      });

      // `insertText(..., "Replace")` swaps the text content while keeping the
      // paragraph element (so style is preserved unless we override below).
      for (let i = 0; i < targets.length; i++) {
        const t = targets[i];
        t.paragraph.insertText(content[i], Word.InsertLocation.replace);
        if (style_per_para && style_per_para[i]) {
          t.paragraph.style = style_per_para[i];
        }
      }

      if (provenance_comment && targets.length > 0) {
        targets[0].paragraph.getRange().insertComment(provenance_comment);
      }

      await context.sync();

      return {
        replaced_count: targets.length,
        paragraph_ids: targets.map((t) => t.id),
      };
    });
  });
}

// ---------------------------------------------------------------------------
// Tool: office_replace_text — surgical sub-paragraph search/replace
// ---------------------------------------------------------------------------
export async function toolReplaceText({
  paragraph_ids,
  find,
  replace,
  match_case,
  whole_word,
  track_changes,
}) {
  if (!Array.isArray(paragraph_ids) || paragraph_ids.length === 0) {
    throw new Error("paragraph_ids must be a non-empty array");
  }
  if (typeof find !== "string" || find.length === 0) {
    throw new Error("find must be a non-empty string");
  }
  if (typeof replace !== "string") {
    throw new Error('replace must be a string (use "" to delete)');
  }

  return await Word.run(async (context) => {
    return await withTrackChanges(context, track_changes, async () => {
      const { paragraphs, idMode } = await getParagraphsWithIds(context);

      const targets = [];
      for (const id of paragraph_ids) {
        const idx = findIndexById(paragraphs, id, idMode);
        if (idx === -1) {
          throw new Error(`Paragraph not found: ${id}`);
        }
        targets.push({ id, paragraph: paragraphs.items[idx] });
      }

      const probes = targets.map((t) => {
        const results = t.paragraph.search(find, {
          matchCase: !!match_case,
          matchWholeWord: !!whole_word,
        });
        results.load("items");
        return { id: t.id, results };
      });
      await context.sync();

      let totalReplaced = 0;
      const perParagraph = [];
      for (const p of probes) {
        const matches = p.results.items.length;
        for (const r of p.results.items) {
          r.insertText(replace, Word.InsertLocation.replace);
        }
        perParagraph.push({ paragraph_id: p.id, replacements: matches });
        totalReplaced += matches;
      }
      await context.sync();

      return {
        total_replacements: totalReplaced,
        per_paragraph: perParagraph,
      };
    });
  });
}

// ---------------------------------------------------------------------------
// Tool: office_delete_paragraphs — remove paragraphs entirely (including
// their paragraph marks). office_replace_text with replace:"" only blanks
// text — the paragraph mark survives, so a stale/duplicated paragraph stays
// as an empty line. This is the only path that actually collapses the doc.
// ---------------------------------------------------------------------------
export async function toolDeleteParagraphs({ paragraph_ids, track_changes }) {
  if (!Array.isArray(paragraph_ids) || paragraph_ids.length === 0) {
    throw new Error("paragraph_ids must be a non-empty array");
  }
  return await Word.run(async (context) =>
    withTrackChanges(context, track_changes, async () => {
      const { paragraphs, idMode } = await getParagraphsWithIds(context);
      // Resolve every target to its paragraph proxy BEFORE deleting any.
      // Each delete() shifts subsequent indices, but the already-resolved
      // proxy objects stay valid — so collect first, delete second.
      const targets = [];
      for (const id of paragraph_ids) {
        const idx = findIndexById(paragraphs, id, idMode);
        if (idx === -1) throw new Error(`Paragraph not found: ${id}`);
        targets.push(paragraphs.items[idx]);
      }
      for (const p of targets) p.delete();
      await context.sync();
      return { deleted: targets.length };
    }),
  );
}

// ---------------------------------------------------------------------------
// Tool: office_replace_section
// ---------------------------------------------------------------------------
export async function toolReplaceSection({
  heading,
  content,
  track_changes,
  style_per_para,
  provenance_comment,
}) {
  if (!content || content.length === 0) throw new Error("content must be non-empty");

  return await Word.run(async (context) => {
    return await withTrackChanges(context, track_changes, async () => {
      let para1 = await getParagraphsWithIds(context);
      let paragraphs = para1.paragraphs;
      let idMode = para1.idMode;
      let snapshot = snapshotParagraphs(paragraphs, idMode);

      const startIdx = findHeadingIndex(snapshot, heading);
      if (startIdx === -1) throw new Error(`Heading not found: "${heading}"`);
      const endIdx = findSectionEnd(snapshot, startIdx);

      // Capture the existing section's body style AND direct formatting
      // BEFORE deleting, so inserted paragraphs match. Without this, new
      // paragraphs inherit the heading's style (centered/bold/large) and
      // lose any manually-applied indent/justify/spacing.
      const ref =
        style_per_para && style_per_para.length === content.length
          ? { idx: -1, inSection: false }
          : findBodyReferenceIndex(snapshot, startIdx + 1);
      const refIdx = ref.idx;
      const bodyStyle = refIdx === -1 ? null : snapshot[refIdx].style || null;
      // Format-clone only when the reference is from this section (see
      // findBodyReferenceIndex). Style name is still taken from the
      // doc-wide fallback; direct formatting is not.
      let refFmt = null;
      if (refIdx !== -1 && ref.inSection) {
        const refPara = paragraphs.items[refIdx];
        loadParagraphFormat(refPara);
        await context.sync();
        refFmt = snapshotParagraphFormat(refPara);
      }

      const toDelete = paragraphs.items.slice(startIdx + 1, endIdx);
      const deletedCount = toDelete.length;
      for (const p of toDelete) p.delete();
      await context.sync();

      // Re-fetch. Heading index may have shifted by 0 (we only deleted *after* it)
      // but be safe and re-locate.
      const para2 = await getParagraphsWithIds(context);
      paragraphs = para2.paragraphs;
      idMode = para2.idMode;
      snapshot = snapshotParagraphs(paragraphs, idMode);
      const newHeadingIdx = findHeadingIndex(snapshot, heading);
      if (newHeadingIdx === -1) throw new Error("Lost the heading after delete; aborting");

      const headingPara = paragraphs.items[newHeadingIdx];
      let cursor = headingPara;
      const inserted = [];
      for (let i = 0; i < content.length; i++) {
        const p = cursor.insertParagraph(content[i], Word.InsertLocation.after);
        const explicit = style_per_para && style_per_para[i];
        if (explicit) p.style = explicit;
        else if (bodyStyle) p.style = bodyStyle;
        if (refFmt && !explicit) applyParagraphFormat(p, refFmt);
        inserted.push(p);
        cursor = p;
      }

      if (provenance_comment && inserted.length > 0) {
        inserted[0].getRange().insertComment(provenance_comment);
      }

      await context.sync();

      let newIds;
      if (idMode === "uniqueLocalId") {
        for (const p of inserted) p.load("uniqueLocalId");
        await context.sync();
        newIds = inserted.map((p) => p.uniqueLocalId);
      } else {
        newIds = inserted.map((_, k) => fallbackId(newHeadingIdx + 1 + k));
      }

      return {
        deleted_paragraphs: deletedCount,
        inserted_count: inserted.length,
        new_para_ids: newIds,
        addressing: idMode,
      };
    });
  });
}

// ---------------------------------------------------------------------------
// Tool: office_highlight
// ---------------------------------------------------------------------------
const SEVERITY_COLOR = {
  error: "Red",
  warning: "Yellow",
  info: "Turquoise",
  uncertain: "Pink",
};

export async function toolHighlight({ targets }) {
  if (!Array.isArray(targets) || targets.length === 0) {
    throw new Error("targets must be a non-empty array");
  }

  return await Word.run(async (context) => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);

    // Phase 1: queue all the searches. Office.js batches operations until
    // sync(), so doing this in one pass is faster than one-sync-per-target.
    const perTarget = [];
    const pendingSearches = []; // { idx, target, searchResults }
    for (let i = 0; i < targets.length; i++) {
      const t = targets[i];
      const idx = findIndexById(paragraphs, t.paragraph_id, idMode);
      if (idx === -1) {
        perTarget[i] = { paragraph_id: t.paragraph_id, ok: false, error: "paragraph_not_found" };
        continue;
      }
      const paragraph = paragraphs.items[idx];
      if (t.query) {
        const searchResults = paragraph.search(t.query, { matchCase: false });
        searchResults.load("items");
        pendingSearches.push({ i, t, paragraph, searchResults });
      } else {
        pendingSearches.push({ i, t, paragraph, wholeParagraph: true });
      }
    }
    await context.sync();

    // Phase 2: apply highlights.
    let totalHighlights = 0;
    for (const p of pendingSearches) {
      const color = SEVERITY_COLOR[p.t.severity || "warning"] || "Yellow";
      try {
        if (p.wholeParagraph) {
          p.paragraph.getRange().font.highlightColor = color;
          perTarget[p.i] = {
            paragraph_id: p.t.paragraph_id,
            ok: true,
            matches: 1,
            severity: p.t.severity || "warning",
          };
          totalHighlights += 1;
        } else {
          const items = p.searchResults.items;
          for (const r of items) {
            r.font.highlightColor = color;
          }
          perTarget[p.i] = {
            paragraph_id: p.t.paragraph_id,
            ok: items.length > 0,
            matches: items.length,
            severity: p.t.severity || "warning",
            ...(items.length === 0 ? { error: "query_not_found", query: p.t.query } : {}),
          };
          totalHighlights += items.length;
        }
      } catch (e) {
        perTarget[p.i] = { paragraph_id: p.t.paragraph_id, ok: false, error: e.message };
      }
    }
    await context.sync();

    return {
      total_highlights: totalHighlights,
      per_target: perTarget,
    };
  });
}

// ---------------------------------------------------------------------------
// Tool: office_clear_highlights
// ---------------------------------------------------------------------------
export async function toolClearHighlights(args) {
  const { paragraph_ids, heading_section, all } = args || {};

  return await Word.run(async (context) => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);
    const snapshot = snapshotParagraphs(paragraphs, idMode);

    let targets = [];
    if (all === true) {
      targets = paragraphs.items;
    } else if (Array.isArray(paragraph_ids) && paragraph_ids.length > 0) {
      for (const id of paragraph_ids) {
        const idx = findIndexById(paragraphs, id, idMode);
        if (idx !== -1) targets.push(paragraphs.items[idx]);
      }
    } else if (typeof heading_section === "string") {
      const startIdx = findHeadingIndex(snapshot, heading_section);
      if (startIdx === -1) throw new Error(`Heading not found: "${heading_section}"`);
      const endIdx = findSectionEnd(snapshot, startIdx);
      targets = paragraphs.items.slice(startIdx, endIdx);
    } else {
      throw new Error("Specify one of paragraph_ids, heading_section, or all: true");
    }

    // Office.js batches all property assignments and applies them at
    // context.sync(). A single bad paragraph (empty paragraph, paragraph
    // in a content control, paragraph with only tracked deletions, etc.)
    // fails the whole batch with a generic InvalidArgument. So:
    //   Fast path — try the whole batch in one sync.
    //   Slow path — if the batch fails, sync one paragraph at a time and
    //               record which paragraphs rejected the change.
    // null is the Office.js named value for clearing a highlight.
    try {
      for (const p of targets) {
        p.getRange().font.highlightColor = null;
      }
      await context.sync();
      return { cleared_paragraphs: targets.length, failed_paragraphs: 0 };
    } catch (batchErr) {
      let cleared = 0;
      const failed = [];
      for (let i = 0; i < targets.length; i++) {
        try {
          targets[i].getRange().font.highlightColor = null;
          await context.sync();
          cleared++;
        } catch (e) {
          failed.push({ index: i, error: e?.message ?? String(e) });
        }
      }
      return { cleared_paragraphs: cleared, failed_paragraphs: failed.length, failures: failed };
    }
  });
}

// ---------------------------------------------------------------------------
// Tool: office_clear_comments
// ---------------------------------------------------------------------------
export async function toolClearComments(args) {
  const { paragraph_ids, heading_section, all } = args || {};

  return await Word.run(async (context) => {
    const comments = context.document.body.getComments();
    comments.load("items");
    await context.sync();

    if (all === true) {
      const total = comments.items.length;
      for (const c of comments.items) c.delete();
      await context.sync();
      return { cleared_comments: total };
    }

    // Scoped path: determine which paragraph each comment is anchored on.
    // For each comment, load its contentRange.paragraphs (the paragraph[s]
    // the comment is anchored in) and find the index of the first one
    // (paragraphIndexes).
    const { paragraphs, idMode } = await getParagraphsWithIds(context);
    const snapshot = snapshotParagraphs(paragraphs, idMode);

    let allowedIdxs;
    if (Array.isArray(paragraph_ids) && paragraph_ids.length > 0) {
      allowedIdxs = new Set();
      for (const id of paragraph_ids) {
        const idx = findIndexById(paragraphs, id, idMode);
        if (idx !== -1) allowedIdxs.add(idx);
      }
    } else if (typeof heading_section === "string") {
      const startIdx = findHeadingIndex(snapshot, heading_section);
      if (startIdx === -1) throw new Error(`Heading not found: "${heading_section}"`);
      const endIdx = findSectionEnd(snapshot, startIdx);
      allowedIdxs = new Set();
      for (let i = startIdx; i < endIdx; i++) allowedIdxs.add(i);
    } else {
      throw new Error("Specify one of paragraph_ids, heading_section, or all: true");
    }

    const probes = comments.items.map((c) => {
      const paras = c.contentRange.paragraphs;
      paras.load(idMode === "uniqueLocalId" ? "items/uniqueLocalId, items/text" : "items/text");
      return { comment: c, paras };
    });
    await context.sync();

    // The paragraph of each comment's anchor: by id, or by its place when
    // paragraphs share a text (paragraphIndexes), so a comment on a twin
    // outside the scope is not cleared for the twin inside it.
    const anchors = await paragraphIndexes(
      context,
      paragraphs,
      idMode,
      probes.map(({ paras }) => paras.items[0] ?? null),
    );
    let deletedCount = 0;
    for (const [n, { comment }] of probes.entries()) {
      const idx = anchors[n];
      if (idx !== -1 && allowedIdxs.has(idx)) {
        comment.delete();
        deletedCount += 1;
      }
    }
    await context.sync();

    return { cleared_comments: deletedCount };
  });
}

// ---------------------------------------------------------------------------
// Tool: office_add_comment
// ---------------------------------------------------------------------------
export async function toolAddComment({ paragraph_id, query, text }) {
  if (!text || !text.trim()) throw new Error("text must be non-empty");

  return await Word.run(async (context) => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);
    const idx = findIndexById(paragraphs, paragraph_id, idMode);
    if (idx === -1) {
      throw new Error(`Paragraph not found: ${paragraph_id}`);
    }
    const paragraph = paragraphs.items[idx];

    let target;
    if (query) {
      const searchResults = paragraph.search(query, { matchCase: false });
      searchResults.load("items");
      await context.sync();
      if (searchResults.items.length === 0) {
        throw new Error(`Query "${query}" not found in paragraph ${paragraph_id}`);
      }
      target = searchResults.items[0];
    } else {
      target = paragraph.getRange();
    }

    target.insertComment(text);
    await context.sync();

    return { paragraph_id, anchored_on: query || "whole_paragraph", comment_added: true };
  });
}

// ---------------------------------------------------------------------------
// Tier 1 editing tools — formatting / structure / whole-doc read
// ---------------------------------------------------------------------------

// Resolve a list of paragraph IDs to live Paragraph objects (shared by the
// formatting tools below). Throws on the first unknown ID.
function resolveTargets(paragraphs, idMode, ids) {
  return ids.map((id) => {
    const idx = findIndexById(paragraphs, id, idMode);
    if (idx === -1) throw new Error(`Paragraph not found: ${id}`);
    return { id, paragraph: paragraphs.items[idx] };
  });
}

// Tool: office_apply_style — restyle existing paragraphs in place (without
// rewriting their text, unlike insert/replace which take a style alongside
// new content).
export async function toolApplyStyle({ ids, style, track_changes }) {
  if (!Array.isArray(ids) || ids.length === 0) throw new Error("ids must be non-empty");
  if (typeof style !== "string" || !style.trim())
    throw new Error("style must be a non-empty string");
  return await Word.run(async (context) => {
    return await withTrackChanges(context, track_changes, async () => {
      const { paragraphs, idMode } = await getParagraphsWithIds(context);
      const targets = resolveTargets(paragraphs, idMode, ids);
      for (const t of targets) t.paragraph.style = style;
      await context.sync();
      return { styled_count: targets.length, paragraph_ids: targets.map((t) => t.id), style };
    });
  });
}

// Tool: office_set_font — character formatting on whole paragraphs.
const UNDERLINE = (v) =>
  v === true ? "Single" : v === false ? "None" : typeof v === "string" ? v : undefined;
export async function toolSetFont({
  ids,
  bold,
  italic,
  underline,
  size,
  color,
  name,
  track_changes,
}) {
  if (!Array.isArray(ids) || ids.length === 0) throw new Error("ids must be non-empty");
  return await Word.run(async (context) => {
    return await withTrackChanges(context, track_changes, async () => {
      const { paragraphs, idMode } = await getParagraphsWithIds(context);
      const targets = resolveTargets(paragraphs, idMode, ids);
      for (const t of targets) {
        const f = t.paragraph.getRange().font;
        if (bold !== undefined) f.bold = bold;
        if (italic !== undefined) f.italic = italic;
        if (underline !== undefined) f.underline = UNDERLINE(underline);
        if (size !== undefined) f.size = size;
        if (color !== undefined) f.color = color;
        if (name !== undefined) f.name = name;
      }
      await context.sync();
      return { formatted_count: targets.length, paragraph_ids: targets.map((t) => t.id) };
    });
  });
}

// Tool: office_set_paragraph_formatting — alignment / indent / spacing.
const ALIGNMENT = { left: "Left", center: "Centered", right: "Right", justify: "Justified" };
export async function toolSetParagraphFormatting({
  ids,
  alignment,
  left_indent,
  space_before,
  space_after,
  line_spacing,
  track_changes,
}) {
  if (!Array.isArray(ids) || ids.length === 0) throw new Error("ids must be non-empty");
  if (alignment !== undefined && !ALIGNMENT[alignment]) {
    throw new Error(`alignment must be one of: ${Object.keys(ALIGNMENT).join(", ")}`);
  }
  return await Word.run(async (context) => {
    return await withTrackChanges(context, track_changes, async () => {
      const { paragraphs, idMode } = await getParagraphsWithIds(context);
      const targets = resolveTargets(paragraphs, idMode, ids);
      for (const t of targets) {
        const p = t.paragraph;
        if (alignment !== undefined) p.alignment = ALIGNMENT[alignment];
        if (left_indent !== undefined) p.leftIndent = left_indent;
        if (space_before !== undefined) p.spaceBefore = space_before;
        if (space_after !== undefined) p.spaceAfter = space_after;
        if (line_spacing !== undefined) p.lineSpacing = line_spacing;
      }
      await context.sync();
      return { formatted_count: targets.length, paragraph_ids: targets.map((t) => t.id) };
    });
  });
}

// Tool: office_insert_table — insert a table, optionally after a given
// paragraph (default: appended at the end of the document body).
export async function toolInsertTable({ rows, after_paragraph_id, header, track_changes }) {
  if (!Array.isArray(rows) || rows.length === 0 || !Array.isArray(rows[0])) {
    throw new Error("`rows` must be a non-empty 2D array of cell strings.");
  }
  const cols = rows[0].length;
  for (let i = 1; i < rows.length; i++) {
    if (!Array.isArray(rows[i]) || rows[i].length !== cols) {
      throw new Error(`Ragged rows: row ${i} has ${rows[i]?.length} cells but row 0 has ${cols}.`);
    }
  }
  return await Word.run(async (context) => {
    return await withTrackChanges(context, track_changes, async () => {
      let table;
      if (after_paragraph_id) {
        const { paragraphs, idMode } = await getParagraphsWithIds(context);
        const idx = findIndexById(paragraphs, after_paragraph_id, idMode);
        if (idx === -1) throw new Error(`Paragraph not found: ${after_paragraph_id}`);
        const anchor = paragraphs.items[idx].getRange(Word.RangeLocation.after);
        table = anchor.insertTable(rows.length, cols, Word.InsertLocation.after, rows);
      } else {
        table = context.document.body.insertTable(rows.length, cols, Word.InsertLocation.end, rows);
      }
      if (header) table.headerRowCount = 1;
      await context.sync();
      return { rows: rows.length, columns: cols, header: !!header };
    });
  });
}

// Tool: office_set_table_cell — overwrite one cell in an existing table.
// Tables are addressed by 0-based document order (body.tables).
export async function toolSetTableCell({ table_index, row, column, text, track_changes }) {
  if (!Number.isInteger(table_index) || table_index < 0)
    throw new Error("`table_index` must be a non-negative integer.");
  if (!Number.isInteger(row) || row < 0) throw new Error("`row` must be a non-negative integer.");
  if (!Number.isInteger(column) || column < 0)
    throw new Error("`column` must be a non-negative integer.");
  return await Word.run(async (context) => {
    return await withTrackChanges(context, track_changes, async () => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();
      if (table_index >= tables.items.length) {
        throw new Error(`Table ${table_index} not found (document has ${tables.items.length}).`);
      }
      const cell = tables.items[table_index].getCell(row, column);
      cell.body.clear();
      cell.body.insertText(text ?? "", Word.InsertLocation.start);
      await context.sync();
      return { table_index, row, column, set: true };
    });
  });
}

// Tool: office_get_document_text — the full body text in one call (the
// paragraph-range read tools are for targeted reads; this is the whole doc).
// When track changes is on, use getReviewedText(current) so the agent sees
// the doc as the user sees it with deletions accepted — raw body.text
// concatenates tracked deletions with their replacements (see
// toolReadParagraphs for the full rationale).
export async function toolGetDocumentText() {
  return await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    context.document.load("changeTrackingMode");
    await context.sync();
    const trackChangesOn = String(context.document.changeTrackingMode) !== "Off";

    let text = body.text || "";
    if (trackChangesOn) {
      const reviewed = await getReviewedTextBatch(context, [body]);
      if (reviewed && reviewed[0] != null) text = reviewed[0];
    }

    const words = text.trim() ? text.trim().split(/\s+/).length : 0;
    return { text, character_count: text.length, word_count: words };
  });
}

// Tool: office_get_outline — the heading tree (id, level, text) for
// orientation without dumping the whole document.
export async function toolGetOutline() {
  return await Word.run(async (context) => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);
    const outline = [];
    for (let i = 0; i < paragraphs.items.length; i++) {
      const p = paragraphs.items[i];
      const level = headingLevel(p.style);
      if (level > 0) {
        outline.push({ id: getId(p, i, idMode), level, text: p.text.trim() });
      }
    }
    return { heading_count: outline.length, outline };
  });
}

// ---------------------------------------------------------------------------
// Tier 2 editing tools — lists / media / links / find / comments / chrome
// ---------------------------------------------------------------------------

// Tool: office_set_list — turn existing paragraphs into a bulleted or
// numbered list (the first becomes a new list; the rest attach to it).
export async function toolSetList({ ids, ordered, track_changes }) {
  if (!Array.isArray(ids) || ids.length === 0) throw new Error("ids must be non-empty");
  return await Word.run(async (context) => {
    return await withTrackChanges(context, track_changes, async () => {
      const { paragraphs, idMode } = await getParagraphsWithIds(context);
      const targets = resolveTargets(paragraphs, idMode, ids);
      const list = targets[0].paragraph.startNewList();
      list.load("id");
      await context.sync();
      for (let i = 1; i < targets.length; i++) {
        targets[i].paragraph.attachToList(list.id, 0);
      }
      if (ordered) {
        list.setLevelNumbering(0, Word.ListNumbering.arabic, ["%1."]);
      } else {
        list.setLevelBullet(0, Word.ListBullet.solid);
      }
      await context.sync();
      return { list_count: targets.length, ordered: !!ordered, paragraph_ids: ids };
    });
  });
}

// Tool: office_insert_image — insert an inline picture from base64 (the
// agent obtains it with e.g. `base64 -i path`). Appended at the end of the
// body, or after a given paragraph.
export async function toolInsertImage({ base64, after_paragraph_id, alt_text, track_changes }) {
  if (typeof base64 !== "string" || !base64.trim()) {
    throw new Error("`base64` (a base64-encoded image) is required.");
  }
  const clean = base64.replace(/^data:image\/[a-z]+;base64,/i, "").trim();
  return await Word.run(async (context) => {
    return await withTrackChanges(context, track_changes, async () => {
      let pic;
      if (after_paragraph_id) {
        const { paragraphs, idMode } = await getParagraphsWithIds(context);
        const idx = findIndexById(paragraphs, after_paragraph_id, idMode);
        if (idx === -1) throw new Error(`Paragraph not found: ${after_paragraph_id}`);
        pic = paragraphs.items[idx]
          .getRange(Word.RangeLocation.after)
          .insertInlinePictureFromBase64(clean, Word.InsertLocation.after);
      } else {
        pic = context.document.body.insertInlinePictureFromBase64(clean, Word.InsertLocation.end);
      }
      if (alt_text) pic.altTextDescription = alt_text;
      await context.sync();
      return { inserted: true, anchored: after_paragraph_id || "end_of_document" };
    });
  });
}

// Resolve a paragraph + optional sub-string query to a Range (shared by
// hyperlink/bookmark). Returns the whole paragraph range when no query.
async function rangeForParagraphQuery(context, paragraph_id, query) {
  const { paragraphs, idMode } = await getParagraphsWithIds(context);
  const idx = findIndexById(paragraphs, paragraph_id, idMode);
  if (idx === -1) throw new Error(`Paragraph not found: ${paragraph_id}`);
  const paragraph = paragraphs.items[idx];
  if (!query) return paragraph.getRange();
  const found = paragraph.search(query, { matchCase: false });
  found.load("items");
  await context.sync();
  if (found.items.length === 0) {
    throw new Error(`Query "${query}" not found in paragraph ${paragraph_id}`);
  }
  return found.items[0];
}

// Tool: office_insert_hyperlink — make text (a query within a paragraph, or
// the whole paragraph) a hyperlink.
export async function toolInsertHyperlink({ paragraph_id, query, url, track_changes }) {
  if (!paragraph_id) throw new Error("`paragraph_id` is required.");
  if (typeof url !== "string" || !url.trim()) throw new Error("`url` is required.");
  return await Word.run(async (context) => {
    return await withTrackChanges(context, track_changes, async () => {
      const range = await rangeForParagraphQuery(context, paragraph_id, query);
      range.hyperlink = url;
      await context.sync();
      return { paragraph_id, anchored_on: query || "whole_paragraph", url };
    });
  });
}

// Tool: office_insert_bookmark — drop a named bookmark on a paragraph or a
// text query within it.
export async function toolInsertBookmark({ paragraph_id, query, name }) {
  if (!paragraph_id) throw new Error("`paragraph_id` is required.");
  if (typeof name !== "string" || !name.trim()) throw new Error("`name` is required.");
  return await Word.run(async (context) => {
    const range = await rangeForParagraphQuery(context, paragraph_id, query);
    try {
      range.insertBookmark(name);
      await context.sync();
    } catch (e) {
      throw new Error(`Bookmarks are not supported on this Word build (${e?.message ?? e}).`);
    }
    return { paragraph_id, anchored_on: query || "whole_paragraph", bookmark: name };
  });
}

// Tool: office_find — search the document; reports matches with the
// containing paragraph ID so the agent can follow up with an edit tool.
export async function toolFind({ query, match_case, whole_word, wildcards }) {
  if (typeof query !== "string" || !query.length) throw new Error("`query` is required.");
  return await Word.run(async (context) => {
    const results = context.document.body.search(query, {
      matchCase: !!match_case,
      matchWholeWord: !!whole_word,
      matchWildcards: !!wildcards,
    });
    results.load("items/text");
    await context.sync();
    const items = results.items.slice(0, 200);
    for (const r of items) r.paragraphs.load("items/uniqueLocalId");
    await context.sync();
    const matches = items.map((r) => {
      const p = r.paragraphs.items[0];
      return { text: r.text, paragraph_id: p ? p.uniqueLocalId : null };
    });
    return { query, match_count: results.items.length, matches };
  });
}

async function getCommentsLoaded(context, fields) {
  const comments = context.document.body.getComments();
  comments.load(fields);
  await context.sync();
  return comments;
}

// Tool: office_list_comments — every comment with id/author/text/resolved.
export async function toolListComments() {
  return await Word.run(async (context) => {
    const comments = await getCommentsLoaded(
      context,
      "items/id, items/authorName, items/content, items/resolved",
    );
    return {
      comment_count: comments.items.length,
      comments: comments.items.map((c) => ({
        id: c.id,
        author: c.authorName,
        text: c.content,
        resolved: c.resolved,
      })),
    };
  });
}

// Tool: office_reply_to_comment — add a reply to a comment thread by id.
export async function toolReplyToComment({ comment_id, text }) {
  if (!comment_id) throw new Error("`comment_id` is required.");
  if (typeof text !== "string" || !text.trim()) throw new Error("`text` is required.");
  return await Word.run(async (context) => {
    const comments = await getCommentsLoaded(context, "items/id");
    const c = comments.items.find((x) => x.id === comment_id);
    if (!c) throw new Error(`Comment not found: ${comment_id}`);
    c.reply(text);
    await context.sync();
    return { comment_id, replied: true };
  });
}

// Tool: office_resolve_comment — mark a comment resolved (or reopen it).
export async function toolResolveComment({ comment_id, resolved = true }) {
  if (!comment_id) throw new Error("`comment_id` is required.");
  return await Word.run(async (context) => {
    const comments = await getCommentsLoaded(context, "items/id, items/resolved");
    const c = comments.items.find((x) => x.id === comment_id);
    if (!c) throw new Error(`Comment not found: ${comment_id}`);
    c.resolved = !!resolved;
    await context.sync();
    return { comment_id, resolved: !!resolved };
  });
}

// Tool: office_header_footer — set the primary header or footer text on
// every section (replaces existing content).
export async function toolHeaderFooter({ which, text }) {
  if (which !== "header" && which !== "footer") {
    throw new Error("`which` must be 'header' or 'footer'.");
  }
  return await Word.run(async (context) => {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();
    for (const section of sections.items) {
      const hf =
        which === "header"
          ? section.getHeader(Word.HeaderFooterType.primary)
          : section.getFooter(Word.HeaderFooterType.primary);
      hf.clear();
      hf.insertText(text ?? "", Word.InsertLocation.start);
    }
    await context.sync();
    return { which, sections: sections.items.length, set: true };
  });
}

// ---------------------------------------------------------------------------
// Review tab (C1). office_tracked_edit is the only writer in a review turn;
// the other functions here are called by the Review tab directly, without
// the agent. Word evidence for the comment anchors: 2026-09-30 spike
// (comment on the deleted range survives accept and reject; comment on an
// inserted range is removed by reject).
// ---------------------------------------------------------------------------

// Capabilities this task pane reports in its hello (taskpane.js sendHello).
// target_hash: toolTrackedEdit compares the paragraph with
// expected_before_hash before it edits. The daemon sends office_tracked_edit
// only to a pane that names it (review-tools.mjs TARGET_HASH_CAP): a pane
// loaded before the daemon was updated would apply the edit unchecked.
export const PANE_CAPS = ["target_hash"];

// Tool: office_tracked_edit. Always tracked, whatever the
// track-changes mode setting (taskpane.js dispatches it without writeArgs()).
// narrow ({ find, replace }, from daemon/edit-facts.mjs narrowEdit, L3): the
// changed words only. Used when narrow.find occurs exactly once inside the
// located hit; otherwise the full find/replace. The result says which.
// expected_before_hash: paragraphHash of the paragraph the daemon judged the
// edit on (forbid list, lost facts). The paragraph this function is about to
// edit is hashed the same way right before the edit, and when the two differ
// nothing is changed (StaleTarget): the author changed the paragraph after
// the daemon read it, so the edit was judged on a text that is gone.
// What is hashed: the REVIEWED text of the paragraph (getReviewedText
// current: pending tracked changes counted as accepted), normalized. That is
// the text the daemon has: its snapshot is collectReviewSnapshot's
// reviewed_text, and it puts its own recorded changes in as accepted
// (review-tools.mjs currentTexts). The raw paragraph.text still holds
// tracked-deleted words, so it differs from the daemon's text after the
// first edit of a paragraph, and it does not show a tracked deletion the
// author made next to the edit. A Word without getReviewedText gives the raw
// text, as it does in the snapshot. The paragraph compared is the one the
// hit is in, wherever it now stands: one that moved with its text unchanged
// passes, and so does a twin with the same text.
// No hash (the daemon could not tell the paragraph, so it judged none) is
// refused as well. expected_paragraph ({ index, id }: where the daemon saw
// the paragraph) only goes into the message.
export async function toolTrackedEdit({
  find,
  replace,
  insert_after,
  paragraph_hint,
  narrow,
  comment_text,
  expected_before_hash,
  expected_paragraph,
}) {
  if (typeof find !== "string" || !find) throw new Error("`find` is required");
  if (find.length > 255) throw new Error("`find` exceeds the 255-character limit of Word search");
  const isInsert = typeof insert_after === "string";
  if (isInsert === (typeof replace === "string")) {
    throw new Error("Give exactly one of replace or insert_after");
  }
  if (!isToolComment(comment_text)) throw new Error("comment_text must start with [R<n>.<m>]");

  return await Word.run(async (context) => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);
    const all = context.document.body.search(find, { matchCase: true });
    all.load("items");
    await context.sync();
    let hits = all.items;
    // The paragraph the hint picked, when it was read.
    let hinted = -1;
    if (hits.length > 1 && paragraph_hint) {
      const idx = findIndexById(paragraphs, paragraph_hint, idMode);
      if (idx === -1)
        return {
          ok: false,
          code: "HintNotFound",
          message: `Paragraph not found: ${paragraph_hint}`,
        };
      const inPara = paragraphs.items[idx].search(find, { matchCase: true });
      inPara.load("items");
      await context.sync();
      hits = inPara.items;
      hinted = idx;
    }
    if (hits.length === 0) return { ok: false, code: "NotFound", message: `"${find}" not found` };
    if (hits.length > 1) {
      return {
        ok: false,
        code: "Ambiguous",
        hit_count: hits.length,
        message: `"${find}" occurs ${hits.length} times; pass paragraph_hint or a longer find`,
      };
    }
    const hit = hits[0];
    const para = hit.paragraphs.getFirst();
    para.load(idMode === "uniqueLocalId" ? "text, uniqueLocalId" : "text");
    await context.sync();
    const paragraphIndex =
      hinted !== -1 ? hinted : (await paragraphIndexes(context, paragraphs, idMode, [para]))[0];
    const before = normalizeParaText(para.text);
    // after is computed from the full find/replace: narrowing cuts only
    // words both sides share, so the sentence comes out the same.
    const computed = computeAfterText(before, find, isInsert ? { insert_after } : { replace });
    if (computed.error) return { ok: false, code: "AmbiguousInParagraph", message: computed.error };

    let target = hit;
    let editFind = find;
    let editReplace = replace;
    let narrowed = false;
    if (!isInsert && narrow && typeof narrow.find === "string" && narrow.find) {
      const sub = hit.search(narrow.find, { matchCase: true });
      sub.load("items");
      await context.sync();
      if (sub.items.length === 1) {
        target = sub.items[0];
        editFind = narrow.find;
        editReplace = narrow.replace;
        narrowed = true;
      }
    }

    return await withTrackChanges(context, true, async () => {
      // The last read before the edit, in this Word.run: the next batch
      // that changes the document is the edit itself.
      para.load("text");
      const reviewed = await getReviewedTextBatch(context, [para.getRange("Whole")]);
      const now = paragraphHash(reviewed ? reviewed[0] : para.text);
      const sent = typeof expected_before_hash === "string" && expected_before_hash !== "";
      if (!sent || now !== expected_before_hash) {
        const seen = expected_paragraph ? ` (read at index ${expected_paragraph.index})` : "";
        return {
          ok: false,
          code: "StaleTarget",
          paragraph_index: paragraphIndex,
          paragraph_id: idMode === "uniqueLocalId" ? para.uniqueLocalId : null,
          message: sent
            ? `The paragraph of "${find}" changed since it was read${seen}; nothing was changed`
            : `The paragraph of "${find}" was not checked before the edit; nothing was changed`,
        };
      }
      const inserted = isInsert
        ? hit.insertText(insert_after, Word.InsertLocation.after)
        : target.insertText(editReplace, Word.InsertLocation.replace);
      // A failure of the sync that carries the edit is treated as not
      // applied (the batch failed) and propagates.
      await context.sync();
      // From here on the edit is in the document. Every later failure comes
      // back as applied:true so the daemon records the change and tells the
      // model not to repeat it (afterEditApplied).
      const applied = {
        paragraph_index: paragraphIndex,
        paragraph_id: idMode === "uniqueLocalId" ? para.uniqueLocalId : null,
        before,
        after: computed.after,
        narrowed,
        // The paragraph was compared with expected_before_hash. The daemon
        // flags an applied result without this (an older task pane).
        target_checked: true,
      };
      // The comment the edit may have got although its sync threw: looked
      // for before the second try, so that try adds no second comment.
      const commentThere = () => existingComment(context, para, comment_text);
      return await afterEditApplied(applied, commentThere, async (at) => {
        let anchor;
        if (isInsert) {
          anchor = inserted;
        } else {
          // The tracked-deleted original is still searchable (spike b). When
          // the replacement contains the find text, skip the hit inside the
          // insertion. A narrowed find is short and may occur elsewhere in
          // the paragraph, so it is searched inside the original hit first.
          let again = (narrowed ? hit : para).search(editFind, { matchCase: true });
          again.load("items");
          await context.sync();
          if (narrowed && again.items.length === 0) {
            again = para.search(editFind, { matchCase: true });
            again.load("items");
            await context.sync();
          }
          if (again.items.length > 1 && editReplace) {
            at("anchor_compare");
            const where = again.items.map((r) => r.compareLocationWith(inserted));
            await context.sync();
            anchor = again.items.find(
              (r, i) => !["Inside", "Equal", "InsideStart", "InsideEnd"].includes(where[i].value),
            );
          } else {
            anchor = again.items[0];
          }
        }
        if (!anchor) {
          return {
            ok: false,
            code: "CommentAnchorMissing",
            applied: true,
            message: "Edit applied but no range to anchor the comment",
          };
        }
        at("insert_comment");
        const c = anchor.insertComment(comment_text);
        c.load("id");
        await context.sync();
        return { ok: true, word_comment_id: c.id };
      });
    });
  });
}

// Review tab, "Add rationale comment": the comment of a change that was
// applied without one (the comment step of toolTrackedEdit failed twice).
// kind, find, text, paragraph_id, paragraph_index: the change as the daemon
// recorded it; comment_text: the text its comment would have had.
// The anchor, as in toolTrackedEdit: for a replace or a deletion the old
// words (find), which are still in the paragraph while the change is
// pending and are the place a comment survives accept and reject; when they
// are gone (the change was accepted), the new words (text) of a replace. A
// pure insertion: its own words. The anchor must be the only such place in
// the paragraph; otherwise no comment is added (CommentAnchorMissing). A
// comment with this text that is already on the paragraph is returned as it
// is (existed: true), so a second press adds no second comment. Nothing of
// the document text is changed.
export async function addChangeComment({
  kind,
  find,
  text,
  paragraph_id,
  paragraph_index,
  comment_text,
}) {
  if (!isToolComment(comment_text)) throw new Error("comment_text must start with [R<n>.<m>]");
  return await Word.run(async (context) => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);
    const k =
      idMode === "uniqueLocalId" && paragraph_id
        ? findIndexById(paragraphs, paragraph_id, idMode)
        : Number.isInteger(paragraph_index) && paragraph_index < paragraphs.items.length
          ? paragraph_index
          : -1;
    if (k === -1) {
      return {
        ok: false,
        code: "ParagraphNotFound",
        message: "The paragraph of the change was not found",
      };
    }
    const para = paragraphs.items[k];
    const there = await existingComment(context, para, comment_text);
    if (there) return { ...there, existed: true };

    const hitsOf = async (q) => {
      if (typeof q !== "string" || !q.trim() || q.length > 255) return [];
      const found = para.search(q, { matchCase: true });
      found.load("items");
      await context.sync();
      return found.items;
    };
    const news = kind === "delete" ? [] : await hitsOf(text);
    let olds = kind === "insert" ? [] : await hitsOf(find);
    // The old words inside the new ones ("degrees" in "degrees Celsius")
    // are not the old words.
    if (olds.length && news.length === 1) {
      const where = olds.map((r) => r.compareLocationWith(news[0]));
      await context.sync();
      olds = olds.filter(
        (r, i) => !["Inside", "Equal", "InsideStart", "InsideEnd"].includes(where[i].value),
      );
    }
    const anchor =
      olds.length === 1 ? olds[0] : olds.length === 0 && news.length === 1 ? news[0] : null;
    if (!anchor) {
      return {
        ok: false,
        code: "CommentAnchorMissing",
        message:
          "The words of the change are not in its paragraph exactly once; no place for the comment",
      };
    }
    const insert = async () => {
      const c = anchor.insertComment(comment_text);
      c.load("id");
      await context.sync();
      return { ok: true, word_comment_id: c.id, existed: false };
    };
    try {
      return await insert();
    } catch {
      // Once more, as in afterEditApplied: the comment may be there after all.
      const now = await existingComment(context, para, comment_text);
      return now ? { ...now, existed: true } : await insert();
    }
  });
}

// The comment with this text on the paragraph, as { ok, word_comment_id },
// or null.
async function existingComment(context, para, comment_text) {
  const comments = para.getComments();
  comments.load("items/id, items/content");
  await context.sync();
  const want = normalizeParaText(comment_text);
  const c = comments.items.find((x) => normalizeParaText(x.content) === want);
  return c ? { ok: true, word_comment_id: c.id } : null;
}

// What goes to the daemon log about a failed comment step: the step, the
// Office error code and place, and the message of the error. Word's errors
// name no text of the document; the statement lists of debugInfo, which can
// (a search string, the comment text), are left out.
function commentError(e, step) {
  return {
    step,
    code: e?.code ?? e?.debugInfo?.code ?? null,
    location: e?.debugInfo?.errorLocation ?? null,
    message: String(e?.message ?? e).slice(0, 200),
  };
}

// The steps after a tracked edit is in the document (anchor search, comment
// insert and its sync). step(at) returns { ok, ... } without the edit fields
// and names the part it is in with at(name); `applied` holds the edit fields
// (paragraph_index, paragraph_id, before, after, narrowed), which the daemon
// records the change from.
// Word's comment insert fails now and then on an edit it accepts the next
// time (the s25 acceptance run: the same edit, one run without its comment,
// the next with it). So a step that threw is tried once more. Before the
// second try found() looks for the comment in the document: a sync can
// throw after Word did the insert, and a second insert would then leave two
// comments on one change. Two throws become the partial result
// CommentInsertFailed, applied:true, like CommentAnchorMissing. Each throw is
// described in comment_errors (commentError), with comment_attempts, also
// when the second try worked: the daemon logs them.
export async function afterEditApplied(applied, found, step) {
  const errors = [];
  for (let attempt = 1; attempt <= 2; attempt++) {
    let where = "anchor_search";
    try {
      if (attempt === 2) where = "find_existing";
      const there = attempt === 2 ? await found() : null;
      where = "anchor_search";
      const r = there ?? (await step((name) => (where = name)));
      return {
        ...r,
        ...applied,
        ...(errors.length ? { comment_attempts: attempt, comment_errors: errors } : {}),
      };
    } catch (e) {
      errors.push(commentError(e, where));
    }
  }
  return {
    ok: false,
    code: "CommentInsertFailed",
    applied: true,
    message: `Edit applied but the rationale comment could not be added: ${errors.at(-1).message}`,
    comment_attempts: 2,
    comment_errors: errors,
    ...applied,
  };
}

function isoOrNull(d) {
  const t = d ? new Date(d) : null;
  return t && !Number.isNaN(t.getTime()) ? t.toISOString() : null;
}

// Paragraphs + comments for status refresh and ledger checks. replies: also
// read every reply of each thread (Word.Comment.replies, WordApi 1.4, the
// same set as getComments) for the "Delete ALL comments" backup.
export async function collectReviewSnapshot({ replies = false } = {}) {
  return await Word.run(async (context) => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);
    const oox = paragraphs.items.map((p) => p.getOoxml());
    const comments = context.document.body.getComments();
    comments.load("items/id, items/content, items/resolved, items/authorName");
    await context.sync();
    const probes = comments.items.map((c) => {
      const r = c.getRange();
      r.load("text");
      const ps = r.paragraphs;
      ps.load(idMode === "uniqueLocalId" ? "items/uniqueLocalId" : "items/text");
      if (replies) c.replies.load("items/id, items/content, items/authorName, items/creationDate");
      return { c, r, ps };
    });
    await context.sync();
    const reviewed = await getReviewedTextBatch(
      context,
      paragraphs.items.map((p) => p.getRange("Whole")),
    );
    const anchors = await paragraphIndexes(
      context,
      paragraphs,
      idMode,
      probes.map(({ ps }) => ps.items[0] ?? null),
    );
    return {
      paragraphs: paragraphs.items.map((p, i) => ({
        index: i,
        id: idMode === "uniqueLocalId" ? p.uniqueLocalId : null,
        text: p.text,
        reviewed_text: reviewed ? reviewed[i] : null,
        ...countRevisionMarks(oox[i].value),
        // Which text is still tracked, run by run (revision-classify.mjs).
        tracked: trackedRuns(oox[i].value),
      })),
      comments: probes.map(({ c, r, ps }, n) => {
        const first = ps.items[0];
        return {
          id: String(c.id),
          content: c.content,
          author: c.authorName,
          resolved: c.resolved,
          anchor_text: r.text,
          paragraph_id: idMode === "uniqueLocalId" && first ? first.uniqueLocalId : null,
          paragraph_index: anchors[n] === -1 ? null : anchors[n],
          ...(replies
            ? {
                replies: c.replies.items.map((x) => ({
                  id: String(x.id),
                  author: x.authorName,
                  content: x.content,
                  created_at: isoOrNull(x.creationDate),
                })),
              }
            : {}),
        };
      }),
    };
  });
}

// Paragraphs under the current selection, for "Propose changes".
export async function selectionParagraphs(max = 10) {
  return await Word.run(async (context) => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);
    const sel = context.document.getSelection().paragraphs;
    sel.load(idMode === "uniqueLocalId" ? "items/uniqueLocalId, items/text" : "items/text");
    await context.sync();
    const picked = sel.items.slice(0, max);
    const indexes = await paragraphIndexes(context, paragraphs, idMode, picked);
    return picked.map((p, n) => ({
      id: idMode === "uniqueLocalId" ? p.uniqueLocalId : null,
      index: indexes[n],
      text: normalizeParaText(p.text),
    }));
  });
}

export async function selectCommentById(commentId) {
  return await Word.run(async (context) => {
    const comments = await getCommentsLoaded(context, "items/id");
    const c = comments.items.find((x) => String(x.id) === String(commentId));
    if (!c) throw new Error(`Comment not found: ${commentId}`);
    c.getRange().select();
    await context.sync();
    return { selected: true };
  });
}

export async function selectParagraphByIndex(index) {
  return await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    await context.sync();
    const p = paragraphs.items[index];
    if (!p) throw new Error(`Paragraph ${index} not found`);
    p.select();
    await context.sync();
    return { selected: true };
  });
}

export async function readAllComments() {
  const snap = await collectReviewSnapshot({ replies: true });
  return snap.comments;
}

// "Delete ALL comments", second press: delete only the backed-up threads
// still equal to their backup (planBackupDelete). Reading, deciding and
// deleting share one Word.run, so a comment added or changed while the
// backup was written is kept.
export async function deleteBackedUpComments(backupRows) {
  return await Word.run(async (context) => {
    const comments = context.document.body.getComments();
    comments.load("items/id, items/content, items/resolved, items/authorName");
    await context.sync();
    const probes = comments.items.map((c) => {
      const r = c.getRange();
      r.load("text");
      c.replies.load("items/content, items/authorName");
      return { c, r };
    });
    await context.sync();
    const current = probes.map(({ c, r }) => ({
      id: String(c.id),
      author: c.authorName,
      content: c.content,
      resolved: c.resolved,
      anchor_text: r.text,
      replies: c.replies.items.map((x) => ({ author: x.authorName, content: x.content })),
    }));
    const plan = planBackupDelete(backupRows, current);
    const doomed = new Set(plan.delete_ids);
    for (const { c } of probes) if (doomed.has(String(c.id))) c.delete();
    await context.sync();
    return plan;
  });
}

// Delete only "[R<n>.<m>]" comments; returns counts before and after and
// deleted_ids, the ids of the tool comments that are gone.
export async function deleteToolComments() {
  return await Word.run(async (context) => {
    const comments = await getCommentsLoaded(context, "items/id, items/content");
    const before = { tool: 0, user: 0 };
    const doomed = [];
    for (const c of comments.items) {
      if (isToolComment(c.content)) {
        before.tool += 1;
        doomed.push(String(c.id));
        c.delete();
      } else {
        before.user += 1;
      }
    }
    await context.sync();
    const left = await getCommentsLoaded(context, "items/id, items/content");
    const after = { tool: 0, user: 0 };
    for (const c of left.items) after[isToolComment(c.content) ? "tool" : "user"] += 1;
    const still = new Set(left.items.map((c) => String(c.id)));
    return { before, after, deleted_ids: doomed.filter((id) => !still.has(id)) };
  });
}
