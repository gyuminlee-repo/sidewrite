// Which manuscript sentences a response quotes, and when a saved response no
// longer matches the manuscript. Pure: no I/O.
//
// Code picks the quotes from the recorded changes of a comment; the model
// never writes a quote, a section name or a line number.
//
// Inputs
//   comment.changes: revision.json changes. paragraph_index and state must be
//     current, i.e. classifyRevision + applyStates ran on this same snapshot.
//   snapshot: { paragraphs: [{ index, text, reviewed_text? }] } of the task pane.
//   headings: [{ index, level, text }] of the manuscript, in document order.

import { createHash } from "node:crypto";
import { normalizeParaText } from "./revision-classify.mjs";
import { splitSentences } from "./review-split.mjs";

// Only changes that are in the manuscript now are cited.
export const CITABLE = new Set(["accepted", "user_edited"]);

// Only figure changes the author adopted are reported and printed.
export const adoptedFigures = (comment) =>
  (comment?.figure_changes ?? []).filter((f) => f.state === "adopted");

// The accepted view: the paragraph as the reviewer will read it. (The change
// classifier reads p.text on purpose; quotes must not.)
export const viewText = (p) => normalizeParaText(p?.reviewed_text ?? p?.text);

const squeeze = (s) => String(s ?? "").replace(/\s+/g, "");
const ENDS_SENTENCE = /[.!?]["')\]]*$/;

// One value for "the manuscript text the line numbers were read from".
export function snapshotHash(snapshot) {
  const h = createHash("sha256");
  for (const p of snapshot?.paragraphs ?? []) h.update(viewText(p)).update("\n");
  return h.digest("hex");
}

// Text of the nearest heading at or above paragraph `index`; with `level`,
// only headings of that level or above (1 is the top). null: no heading.
export function sectionOf(index, headings = [], level = null) {
  let best = null;
  for (const h of headings ?? []) {
    if (h.index > index || (level != null && h.level > level)) continue;
    if (!best || h.index > best.index) best = h;
  }
  return best ? normalizeParaText(best.text) : null;
}

// The text without whitespace, and for each kept character its index in s.
function squeezeMap(s) {
  const at = [];
  let text = "";
  for (let i = 0; i < s.length; i++) {
    if (/\s/.test(s[i])) continue;
    text += s[i];
    at.push(i);
  }
  return { text, at };
}

const WORD_CHAR = /[\p{L}\p{N}]/u;

// Start indexes of `form` in the squeezed text m.text. A side given as
// `wordStart` / `wordEnd` must not cut a word of the source text `src`:
// context words are whole words, the changed text itself may be part of one.
function hitsOf(src, m, form, { wordStart = false, wordEnd = false } = {}) {
  const out = [];
  if (!form) return out;
  for (let p = m.text.indexOf(form); p !== -1; p = m.text.indexOf(form, p + 1)) {
    const a = m.at[p];
    const b = m.at[p + form.length - 1];
    if (wordStart && WORD_CHAR.test(src[a]) && a > 0 && WORD_CHAR.test(src[a - 1])) continue;
    if (wordEnd && WORD_CHAR.test(src[b]) && WORD_CHAR.test(src[b + 1] ?? "")) continue;
    out.push(p);
  }
  return out;
}

// Offset in `text` of the place the change touched, or null when that place
// cannot be told. A wrong quote is worse than none, so the recorded context
// words (up to three on each side of the changed range) decide, not the new
// text on its own:
//
//   0. The old text standing between its context words (for an insertion:
//      the two context sides next to each other) means the change is not in
//      effect there. null, also when the new form turns up somewhere else.
//   1. The new text between both context sides, exactly once: that place.
//   2. Otherwise every piece of context that is in the paragraph exactly once
//      names a place: left + new text, new text + right, left alone, right
//      alone (the author rewrote the new text). They must all lie in one
//      sentence; that sentence is the place. If they disagree, or no piece of
//      the context is left (the sentence was rewritten as a whole), null.
//   3. The new text alone is trusted in one case only: the change recorded no
//      context on either side AND its state is "accepted". Without context
//      the classifier calls a change accepted only when the paragraph reads
//      exactly as it did right after the change, so a single occurrence of
//      the new text is the change itself. In any other state a single
//      occurrence may be another sentence that happens to use the same
//      words, and it is not taken.
export function changeOffset(text, change) {
  const src = String(text ?? "");
  const m = squeezeMap(src);
  if (!m.at.length) return null;
  const del = change.kind === "delete";
  const needle = squeeze(del ? "" : change.text);
  const left = squeeze(change.context?.left);
  const right = squeeze(change.context?.right);
  const at = (i) => m.at[Math.min(Math.max(i, 0), m.at.length - 1)];
  const hits = (form, wordStart, wordEnd) => hitsOf(src, m, form, { wordStart, wordEnd });

  if (!left && !right) {
    const alone = hits(needle, false, false);
    return change.state === "accepted" && alone.length === 1 ? at(alone[0]) : null;
  }

  const full = left + needle + right;
  const find = squeeze(change.find);
  const old = change.kind === "insert" ? left + right : find ? left + find + right : "";
  // An old form that is part of the new one ("30" -> "30 °C" at a paragraph
  // end) says nothing.
  if (old && !full.includes(old) && hits(old, !!left, !!right).length) return null;

  const whole = hits(full, !!left, !!right);
  if (whole.length === 1) return at(whole[0] + left.length);

  const places = [];
  const one = (form, wordStart, wordEnd, place) => {
    const h = hits(form, wordStart, wordEnd);
    if (h.length === 1) places.push(at(place(h[0])));
  };
  if (needle && left) one(left + needle, true, false, (p) => p + left.length);
  if (needle && right) one(needle + right, false, true, (p) => p);
  if (left) one(left, true, false, (p) => p + left.length);
  // The right side starts after the change: the character before it is the
  // end of the changed text (it may end a sentence). After a deletion there
  // is no such text, and the right side itself is the place.
  if (right) one(right, false, true, (p) => (del ? p : p - 1));
  if (!places.length) return null;
  const sentence = sentenceAt(src, places[0]);
  return places.every((p) => sentenceAt(src, p) === sentence) ? places[0] : null;
}

// [{ start, end }] of the sentences of `text`, as ranges of the text itself.
// splitSentences is the splitter for review comments: it leaves a typed list
// number or bullet off the sentence ("1. ", "- ", "(a) ") and drops units
// without a letter. A manuscript quote must be the manuscript's own
// characters, so what the splitter left out in front of a sentence goes back
// to that sentence, and what it left out after the last one to the last.
function sentenceSpans(text) {
  const spans = [];
  let cursor = 0;
  for (const s of splitSentences(text)) {
    const at = text.indexOf(s, cursor);
    if (at === -1) return [];
    const lead = text.slice(cursor, at).search(/\S/);
    spans.push({ start: lead === -1 ? at : cursor + lead, end: at + s.length });
    cursor = at + s.length;
  }
  if (spans.length && /\S/.test(text.slice(cursor)))
    spans[spans.length - 1].end = text.trimEnd().length;
  return spans;
}

// The sentence of `text` that holds `offset`, as it stands in the text, or null.
export function sentenceAt(text, offset) {
  const span = sentenceSpans(text).find((s) => s.start <= offset && offset < s.end);
  return span ? text.slice(span.start, span.end) : null;
}

// A deletion that took whole sentences: it starts where a sentence starts
// and ends where one ends. What is left around it did not change, so there
// is nothing to quote.
function deletedWholeSentences(change) {
  if (change.kind !== "delete") return false;
  const left = String(change.context?.left ?? "").trim();
  return (
    ENDS_SENTENCE.test(String(change.find ?? "").trim()) && (!left || ENDS_SENTENCE.test(left))
  );
}

// One citation per quoted sentence (several changes in one sentence share
// it) and one per removal. A change that cannot be placed gets `problem`
// instead of a guess.
export function citationsFor(comment, snapshot, headings = [], { level = null } = {}) {
  const byIndex = new Map((snapshot?.paragraphs ?? []).map((p) => [p.index, p]));
  const out = [];
  for (const ch of comment?.changes ?? []) {
    if (!CITABLE.has(ch.state)) continue;
    const p = ch.paragraph_index == null ? null : (byIndex.get(ch.paragraph_index) ?? null);
    const base = {
      change_ids: [ch.id],
      kind: "revised",
      section: p ? sectionOf(p.index, headings, level) : null,
      paragraph_index: p ? p.index : null,
      quote: "",
      lines: null,
      lines_pdf: null,
      problem: null,
    };
    if (!p) {
      out.push({ ...base, problem: "unlocated" });
      continue;
    }
    if (deletedWholeSentences(ch)) {
      out.push({ ...base, kind: "removed" });
      continue;
    }
    const text = viewText(p);
    const off = changeOffset(text, ch);
    const quote = off == null ? null : sentenceAt(text, off);
    if (!quote) {
      out.push(
        ch.kind === "delete"
          ? { ...base, kind: "removed" }
          : { ...base, problem: "sentence not found" },
      );
      continue;
    }
    const same = out.find(
      (c) => c.kind === "revised" && c.paragraph_index === p.index && c.quote === quote,
    );
    if (same) same.change_ids.push(ch.id);
    else out.push({ ...base, quote });
  }
  return out;
}

// ---- supplementary documents (revision-parts.mjs partsForLetter) -------------
// A part is a supplementary Word file of the revision as the manuscript reads
// it: { name, label, revision, paragraphs, confirmed, why }. Everything about
// it comes from its own review file, as its own pane wrote it at its last
// status refresh: the states of its changes, the text of the file
// (paragraphs) and the citations of its accepted changes (section and quote;
// its pane computes them, because only it knows the headings of the file).

// What the response letter calls a supplementary file: the label of its
// link, or its file name without the extension.
export const partLabel = (label, name) =>
  String(label ?? "").trim() || String(name ?? "").replace(/\.[^.]+$/, "");

// The changes of comment `ref` in a part.
export const partChanges = (part, ref) =>
  (part?.revision?.comments ?? []).find((c) => c.ref === ref)?.changes ?? [];

// Those that are in the file now: only they answer a request and are cited.
export const partCitable = (part, ref) =>
  partChanges(part, ref).filter((ch) => CITABLE.has(ch.state));

// The citations a part gives the letter for comment `ref`, as citations of
// the manuscript's review file (doc: the file name; no line numbers). What
// the part stored is not taken on trust:
//   - a stored citation is used only when every change it names is accepted
//     in the part now (the states are the part's own, of the same refresh);
//   - a quote must stand in the paragraph of the part it names, as the part
//     stored its text; otherwise the citation carries a problem and is not
//     printed (the check fails on it);
//   - an accepted change no stored citation names (the pane of the part did
//     not get to write its citations) gets a citation with a problem, so it
//     is never counted as cited in silence.
export function partCitations(part, ref) {
  const accepted = new Set(partCitable(part, ref).map((ch) => ch.id));
  const stored =
    (part?.revision?.comments ?? []).find((c) => c.ref === ref)?.response?.citations ?? [];
  const out = [];
  const covered = new Set();
  for (const cit of stored) {
    const ids = cit.change_ids ?? [];
    if (cit.doc || !ids.length || !ids.every((id) => accepted.has(id))) continue;
    ids.forEach((id) => covered.add(id));
    const revised = (cit.kind ?? "revised") === "revised";
    const quote = String(cit.quote ?? "");
    const here = part.paragraphs?.[cit.paragraph_index] ?? "";
    const problem =
      cit.problem ??
      (revised && (!quote || !here.includes(quote))
        ? "the quote is not in the file as it was last read"
        : null);
    out.push({
      change_ids: [...ids],
      kind: revised ? "revised" : "removed",
      section: cit.section ?? null,
      paragraph_index: cit.paragraph_index ?? null,
      quote: problem ? "" : quote,
      lines: null,
      lines_pdf: null,
      problem,
      doc: part.name,
    });
  }
  for (const id of accepted) {
    if (covered.has(id)) continue;
    out.push({
      change_ids: [id],
      kind: "revised",
      section: null,
      paragraph_index: null,
      quote: "",
      lines: null,
      lines_pdf: null,
      problem: "no citation on record; open the file and press Refresh",
      doc: part.name,
    });
  }
  return out;
}

// Fresh citations keep the line range of an old one with the same quote in
// the same paragraph.
export function keepLines(fresh, old = []) {
  return fresh.map((c) => {
    const o = (old ?? []).find(
      (x) => x.kind === c.kind && x.quote === c.quote && x.paragraph_index === c.paragraph_index,
    );
    return o ? { ...c, lines: o.lines ?? null, lines_pdf: o.lines_pdf ?? null } : c;
  });
}

// Why the saved response of a comment is out of date. Call it with the
// citations as saved, before they are replaced by fresh ones.
//   quote_changed  a quoted sentence is no longer in the paragraph of its
//                  change. The same words in another paragraph do not count:
//                  the answer is about the sentence that was changed. The
//                  paragraph is the one the change is in now (applyStates
//                  follows it when paragraphs are put in or taken out above);
//                  a citation without a change id, or a change that records
//                  no paragraph at all, falls back to the index saved with
//                  the citation. (A change that is not found any more keeps
//                  its last index; it is change_state that tells that one.)
//   change_state   the cited changes are not the accepted changes any more,
//                  or the adopted figure changes are not the ones saved with
//                  the answer at the last refresh (figure_ids). A response
//                  without that list (the citations a pane says it showed,
//                  response_save) is not judged on figures.
//   lines_old      the manuscript text changed since the PDF the line
//                  numbers were read from (render: revision.response_render)
// parts: the supplementary files (partsForLetter). A citation that names a
// file (doc) is judged against that file as its review file has it: its
// quote against the stored text of its paragraph, its changes against the
// accepted changes of that file for this comment. A file that is no longer
// a part makes its citations change_state. null: the parts are not known to
// this caller, and citations with a doc are not judged at all (the citations
// a pane says it showed, response_save).
export function staleReasons(comment, snapshot, render, parts = []) {
  const r = comment?.response;
  if (!r) return [];
  const out = new Set();
  const all = r.citations ?? [];
  const cits = all.filter((c) => !c.doc);
  if (parts) {
    const byName = new Map(parts.map((p) => [p.name, p]));
    for (const c of all) {
      if (!c.doc || c.kind !== "revised" || !c.quote) continue;
      const here = byName.get(c.doc)?.paragraphs?.[c.paragraph_index] ?? "";
      if (!here.includes(c.quote)) out.add("quote_changed");
    }
    const docs = new Set([...parts.map((p) => p.name), ...all.flatMap((c) => c.doc ?? [])]);
    for (const doc of docs) {
      const cited = new Set(all.filter((c) => c.doc === doc).flatMap((c) => c.change_ids ?? []));
      const now = new Set(partCitable(byName.get(doc), comment.ref).map((ch) => ch.id));
      if (cited.size !== now.size || [...cited].some((id) => !now.has(id))) out.add("change_state");
    }
  }
  const texts = new Map((snapshot?.paragraphs ?? []).map((p) => [p.index, viewText(p)]));
  const changes = new Map((comment.changes ?? []).map((ch) => [ch.id, ch]));
  for (const c of cits) {
    if (c.kind !== "revised" || !c.quote) continue;
    const ids = c.change_ids?.length ? c.change_ids : [null];
    const here = ids.map((id) => {
      const ch = changes.get(id);
      return ch && ch.paragraph_index !== undefined ? ch.paragraph_index : c.paragraph_index;
    });
    if (!here.some((i) => (texts.get(i) ?? "").includes(c.quote))) out.add("quote_changed");
  }
  const cited = new Set(cits.flatMap((c) => c.change_ids ?? []));
  const now = new Set(
    (comment.changes ?? []).filter((ch) => CITABLE.has(ch.state)).map((ch) => ch.id),
  );
  if (cited.size !== now.size || [...cited].some((id) => !now.has(id))) out.add("change_state");
  if (Array.isArray(r.figure_ids)) {
    const figures = adoptedFigures(comment).map((f) => f.id);
    if (r.figure_ids.length !== figures.length || figures.some((id) => !r.figure_ids.includes(id)))
      out.add("change_state");
  }
  // (cits: the citations of the manuscript itself; a supplement has no lines.)
  if (cits.length && render && render.snapshot_hash !== snapshotHash(snapshot))
    out.add("lines_old");
  return ["quote_changed", "change_state", "lines_old"].filter((x) => out.has(x));
}

// The stale marks of every saved response, judged against the citations as
// saved. Returns a new revision; the input is not changed.
//   - quote_changed and change_state stay until the answer is written again
//     (the text was written for the old sentence or the old set of changes),
//     and are set only when there is an answer
//   - lines_old follows the PDF: set while the manuscript differs from the
//     one the line numbers were read from
// Run it after applyStates on the same snapshot. Both the Review refresh and
// the Response tab mark with this function.
//
// Where quote_changed and change_state are taken off again (all in
// response-service.mjs). The rule: only writing the answer again against the
// citations as they are now clears them; the line numbers never do.
//
//   response_draft         cleared when what the draft was written from (the
//                          quoted sentences, the changes and their states) is
//                          still what is stored when it is saved; if another
//                          request read a changed document in between (the
//                          citations changed, or the Review refresh found
//                          them not to fit another text: stale_at below),
//                          what moved is marked
//   response_save, text    judged by the citations the pane says its card
//                          showed (`shown`), against the document of the
//                          request, whatever is stored by then: cleared when
//                          they still fit; a shown sentence edited or a cited
//                          change taken back is marked instead (a card that
//                          showed no citation has nothing to compare: cleared).
//                          without `shown` or without the document: not
//                          judged, the marks stay as they are stored
//   response_save, other   (declined, declined_requirements) not touched
//   response_locate        not touched; it takes off lines_old only
//   response_translate     not touched
//   profile_set            not touched; a redraft in another language is a
//                          response_draft like any other
//   refresh (this fn)      kept; set when the saved citations no longer fit
// parts: the supplementary files of the manuscript (staleReasons).
export const STALE_ORDER = ["quote_changed", "change_state", "lines_old"];
export function markStale(rev, snapshot, parts = []) {
  const next = structuredClone(rev);
  let text = null; // snapshotHash, when a reason is found
  for (const c of next.comments ?? []) {
    if (!c.response) continue;
    const now = staleReasons(c, snapshot, next.response_render, parts);
    const kept = c.response.stale ?? [];
    const hasText = !!c.response.text?.trim();
    c.response.stale = STALE_ORDER.filter((s) =>
      s === "lines_old" ? now.includes(s) : hasText && (kept.includes(s) || now.includes(s)),
    );
    // A mark that stays says nothing about a second edit: the text in which
    // the saved citations were found not to fit is kept with each reason.
    const found = now.filter((s) => s !== "lines_old");
    if (found.length) {
      text ??= snapshotHash(snapshot);
      c.response.stale_at = {
        ...c.response.stale_at,
        ...Object.fromEntries(found.map((s) => [s, text])),
      };
    }
  }
  return next;
}
