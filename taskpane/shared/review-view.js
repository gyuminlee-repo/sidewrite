// Pure helpers for the Review tab. No Office.js and no DOM here, so
// node --test can exercise them (tests/review-view.test.mjs).

// Same rule as daemon/revision-classify.mjs normalizeParaText; the test
// runs both on one corpus so they cannot drift apart.
export function normalizeParaText(s) {
  return String(s ?? "")
    .replace(/\u0005/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

// Same function as daemon/doc-diff.mjs paragraphHash (FNV-1a 32 bit over the
// UTF-16 code units of the normalized text, 8 hex digits); the test runs
// both on one corpus. The daemon sends the hash of the paragraph it judged a
// tracked edit on, the task pane compares it with the paragraph it is about
// to edit (tools-word.js toolTrackedEdit).
export function paragraphHash(text) {
  const s = normalizeParaText(text);
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h.toString(16).padStart(8, "0");
}

// The paragraph as it reads once this one change is accepted. Computed from
// the text before the edit, because while the change is pending the Word
// paragraph.text holds both the inserted and the deleted text.
export function computeAfterText(before, find, { replace, insert_after } = {}) {
  const b = normalizeParaText(before);
  // Keep leading/trailing spaces of `find`: " with shaking" deletes the space too.
  const f = String(find ?? "")
    .replace(/\u0005/g, "")
    .replace(/\s+/g, " ");
  if (!f.trim()) return { error: "find is empty" };
  const i = b.indexOf(f);
  if (i === -1) return { error: "find is not in the paragraph text" };
  if (b.indexOf(f, i + 1) !== -1) return { error: "find occurs more than once in the paragraph" };
  if (typeof insert_after === "string") {
    return {
      after: normalizeParaText(b.slice(0, i + f.length) + insert_after + b.slice(i + f.length)),
    };
  }
  if (typeof replace === "string") {
    return { after: normalizeParaText(b.slice(0, i) + replace + b.slice(i + f.length)) };
  }
  return { error: "give replace or insert_after" };
}

// `[\s>/]` after the name excludes w:instrText, w:delText, w:delInstrText.
const W_INS = /<w:ins[\s>/]/g;
const W_DEL = /<w:del[\s>/]/g;

// getOoxml() of a paragraph is a pkg:package. Only the document part holds the
// paragraph's own revision marks: a rationale comment inserted while tracking
// was on carries a w:ins in /word/comments.xml that accepting the change in
// Word does not remove, and it made an accepted insertion read as pending.
const DOC_PART = '<pkg:part pkg:name="/word/document.xml"';

export function countRevisionMarks(xml) {
  let s = String(xml ?? "");
  const at = s.indexOf(DOC_PART);
  if (at !== -1) {
    const end = s.indexOf("</pkg:part>", at);
    s = s.slice(at, end === -1 ? undefined : end);
  }
  return { w_ins: (s.match(W_INS) || []).length, w_del: (s.match(W_DEL) || []).length };
}

// The text of a paragraph that is still a tracked change, run by run:
// [{ kind: "ins" | "del", text, start, marks }] in document order, from the
// document part of the paragraph's OOXML package (the tracked body of a
// rationale comment in /word/comments.xml does not count, as above).
// kind "del": w:delText, or text inside w:del / w:moveFrom (a deleted
// insertion is a deletion). kind "ins": other text inside w:ins / w:moveTo.
// start: where the run begins in the paragraph text as the daemon reads it
// (comment anchors dropped, whitespace collapsed, no leading space).
// marks: how many w:ins / w:del elements the entry stands for (neighbouring
// runs of one kind are one entry). A tracked element from which no text was
// read (a drawing, a field, an element not known here) is an entry with
// empty text: the daemon must know that something tracked is there even
// when it cannot say where (revision-classify.mjs placedRuns).
// The paragraph totals (countRevisionMarks) cannot say WHICH change is still
// pending when a paragraph holds several (Codex review 16); this can.
//
// Run elements that are text without being w:t, with the character Word
// gives them in Paragraph.text (measured in Word 16 on 2026-10-01 with the
// document of scripts/e2e/make_tracked_special.py; the acceptance run
// repeats the measurement): the extracted text must equal the paragraph
// text, or the run cannot be found in it.
export const RUN_CHAR = {
  noBreakHyphen: "\u001e",
  softHyphen: "\u001f",
  tab: "\t",
  ptab: "\t",
  br: "\u000b",
  cr: "\r",
  // Any symbol, whatever its font and code: Word reports "(".
  sym: "(",
};
const XML_TOKEN =
  /<w:(ins|del|moveTo|moveFrom)\b[^>]*?(\/?)>|<\/w:(ins|del|moveTo|moveFrom)>|<w:(?:t|delText)\b[^>]*?\/>|<w:(t|delText)\b[^>]*>([\s\S]*?)<\/w:\4>|<w:(noBreakHyphen|softHyphen|tab|ptab|br|cr|sym)\b([^>]*?)\/?>/g;
const XML_ENTITY = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };
export const xmlText = (s) =>
  s.replace(/&(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);/g, (m, e) =>
    e[0] === "#"
      ? String.fromCodePoint(e[1] === "x" ? parseInt(e.slice(2), 16) : Number(e.slice(1)))
      : XML_ENTITY[e],
  );

export function trackedRuns(xml) {
  const s = String(xml ?? "");
  const at = s.indexOf(DOC_PART);
  if (at === -1) return [];
  const end = s.indexOf("</pkg:part>", at);
  const doc = s.slice(at, end === -1 ? undefined : end);
  const side = { ins: "ins", moveTo: "ins", del: "del", moveFrom: "del" };
  const open = []; // the tracked elements we are inside: { kind, texted }
  const runs = [];
  let raw = "";
  let last = null; // the entry the previous text went into, with where it ended
  const place = () =>
    raw
      .replace(/\u0005/g, "")
      .replace(/\s+/g, " ")
      .trimStart().length;
  for (const m of doc.matchAll(XML_TOKEN)) {
    if (m[1]) {
      if (!m[2]) open.push({ kind: side[m[1]], texted: false });
      continue;
    }
    if (m[3]) {
      const k = open.map((w) => w.kind).lastIndexOf(side[m[3]]);
      if (k === -1) continue;
      const [w] = open.splice(k, 1);
      if (!w.texted) runs.push({ kind: w.kind, text: "", start: place(), marks: 1 });
      continue;
    }
    // A w:tab with a position is a tab stop of the paragraph, not a character.
    if (m[6] === "tab" && /\bw:pos=/.test(m[7])) continue;
    const text = m[4] ? xmlText(m[5]) : (RUN_CHAR[m[6]] ?? "");
    if (!text) continue;
    const inDel = m[4] === "delText" || open.some((w) => w.kind === "del");
    const kind = inDel ? "del" : open.some((w) => w.kind === "ins") ? "ins" : null;
    if (kind) {
      const w = open.findLast((x) => x.kind === kind);
      const mark = w && !w.texted ? 1 : 0;
      if (w) w.texted = true;
      if (last && last.kind === kind && last.end === raw.length) {
        last.run.text += text;
        last.run.marks += mark;
        last.end += text.length;
      } else {
        const run = { kind, text, start: place(), marks: mark };
        runs.push(run);
        last = { kind, run, end: raw.length + text.length };
      }
    }
    raw += text;
  }
  return runs;
}

// Same prefix as daemon/review-tools.mjs TOOL_COMMENT_PREFIX.
export const TOOL_COMMENT_RE = /^\s*\[(?:R\d+|E)\.\d+\]/;

// Reviewer number of a comment ref: R<n>.<m> -> n, E.<m> (editor) -> 0,
// anything else -> null.
export function reviewerOfRef(ref) {
  const m = /^(?:R(\d+)|E)\.\d+$/.exec(String(ref ?? ""));
  if (!m) return null;
  return m[1] === undefined ? 0 : Number(m[1]);
}

export function isToolComment(content) {
  return TOOL_COMMENT_RE.test(String(content ?? ""));
}

export function countComments(comments) {
  let tool = 0;
  for (const c of comments ?? []) if (isToolComment(c.content)) tool += 1;
  return { tool, user: (comments?.length ?? 0) - tool, total: comments?.length ?? 0 };
}

// Result line for "Remove tool comments": tool comments must reach 0 and
// the comments the user wrote must not change.
export function cleanupReport(before, after) {
  const ok = after.tool === 0 && after.user === before.user;
  return {
    ok,
    text:
      `Tool comments ${before.tool} -> ${after.tool}, your comments ${before.user} -> ${after.user}` +
      (ok ? " (unchanged)" : " (MISMATCH)"),
  };
}

// Deleting a comment deletes its replies, so the backup keeps them.
// count is the number of parent comments (what Word lists and deletes);
// reply_count is the number of replies across all threads.
export function buildCommentsBackup(comments, { manuscript = null, now = new Date() } = {}) {
  const rows = comments.map((c) => ({
    id: String(c.id),
    author: c.author ?? null,
    content: c.content ?? "",
    resolved: !!c.resolved,
    anchor_text: c.anchor_text ?? "",
    paragraph_index: c.paragraph_index ?? null,
    replies: (c.replies ?? []).map((r) => ({
      id: r.id == null ? null : String(r.id),
      author: r.author ?? null,
      content: r.content ?? "",
      created_at: r.created_at ?? null,
    })),
  }));
  return {
    version: 1,
    manuscript,
    created_at: now.toISOString(),
    count: rows.length,
    reply_count: rows.reduce((n, c) => n + c.replies.length, 0),
    comments: rows,
  };
}

// "Delete ALL comments" shows a list on the first press and deletes on the
// second. Before deleting, the pane reads the comments again and compares:
// any comment added, removed, or changed (text, author, resolved, anchor,
// replies) since the list was shown means no delete and a new confirmation.
function commentKey(c) {
  return JSON.stringify([
    c.author ?? null,
    c.content ?? "",
    !!c.resolved,
    c.anchor_text ?? "",
    (c.replies ?? []).map((r) => [r.author ?? null, r.content ?? ""]),
  ]);
}

export function compareCommentLists(shown, current) {
  const before = new Map((shown ?? []).map((c) => [String(c.id), commentKey(c)]));
  const now = new Map((current ?? []).map((c) => [String(c.id), commentKey(c)]));
  let added = 0;
  let edited = 0;
  for (const [id, key] of now) {
    if (!before.has(id)) added += 1;
    else if (before.get(id) !== key) edited += 1;
  }
  let removed = 0;
  for (const id of before.keys()) if (!now.has(id)) removed += 1;
  const same = added + removed + edited === 0;
  const parts = [];
  if (added) parts.push(`${added} added`);
  if (removed) parts.push(`${removed} removed`);
  if (edited) parts.push(`${edited} changed`);
  return { same, added, removed, edited, text: parts.join(", ") };
}

// The second "Delete ALL" press deletes only the threads in the backup, and
// only those still equal to their backed-up copy (text, author, resolved,
// anchor, replies). The Office function reads the comments and deletes in
// one Word.run with this decision in between. new_ids: comments added after
// the backup (kept); changed_ids: backed-up threads that changed (kept, the
// user presses again); missing_ids: backed-up threads already gone.
export function planBackupDelete(backupRows, current) {
  const backed = new Map((backupRows ?? []).map((c) => [String(c.id), commentKey(c)]));
  const plan = { delete_ids: [], changed_ids: [], new_ids: [], missing_ids: [] };
  const seen = new Set();
  for (const c of current ?? []) {
    const id = String(c.id);
    if (!backed.has(id)) {
      plan.new_ids.push(id);
      continue;
    }
    seen.add(id);
    (backed.get(id) === commentKey(c) ? plan.delete_ids : plan.changed_ids).push(id);
  }
  for (const id of backed.keys()) if (!seen.has(id)) plan.missing_ids.push(id);
  return plan;
}

// Verdict for "Delete ALL comments": the backup holds every comment shown;
// every backed-up thread is deleted or was already gone; none of the deleted
// ids is still in the document (stillThere, read again afterwards).
// Comments added after the backup are kept and named, not a mismatch. A
// backed-up thread that changed after the backup is kept and needs another
// press, so the result is not ok.
export function deleteAllReport({
  shown,
  backedUp,
  deleted,
  changed = 0,
  missing = 0,
  addedAfter = 0,
  stillThere = 0,
}) {
  const ok = backedUp === shown && changed === 0 && deleted + missing === shown && stillThere === 0;
  const parts = [`Backed up ${backedUp}, deleted ${deleted} of ${shown}`];
  if (missing) parts.push(`${missing} already gone`);
  if (changed) parts.push(`${changed} changed after the backup and not deleted (press again)`);
  if (stillThere) parts.push(`${stillThere} deleted comment(s) still present`);
  let text = parts.join(", ") + (ok ? "" : " (MISMATCH)");
  if (addedAfter) text += `. ${addedAfter} comment(s) added after the backup were not deleted`;
  return { ok, text };
}

// Recorded changes vs tool comments now in the document. Every change keeps
// its comment (a deleted-range comment survives accept and reject), except a
// rejected pure insertion, whose comment Word removes with the text. An
// extra [R comment means an edit that was not recorded.
export function invariantReport(revision, comments) {
  const ids = new Set(
    (comments ?? []).filter((c) => isToolComment(c.content)).map((c) => String(c.id)),
  );
  const changes = (revision?.comments ?? []).flatMap((c) => c.changes ?? []);
  // No comment is expected for a rejected insertion (Word removes it) or
  // for a change whose comment the tool comment cleanup removed.
  const expected = changes.filter(
    (ch) => !(ch.kind === "insert" && ch.state === "rejected") && !ch.comment_cleaned_at,
  );
  const recorded = new Set(changes.map((ch) => String(ch.word_comment_id)));
  // A change the task pane applied without its rationale comment
  // (comment_missing, no comment id) never had one: it is not a comment that
  // went missing, and it is named apart. It is not ok either: every change
  // needs its rationale comment.
  const bare = (ch) => ch.word_comment_id == null && !!ch.comment_missing;
  const uncommented = expected.filter(bare);
  const missing = expected.filter((ch) => !bare(ch) && !ids.has(String(ch.word_comment_id)));
  const orphans = [...ids].filter((id) => !recorded.has(id));
  // The parts of a rewritten paragraph share one comment: comments are
  // counted, not changes.
  const wanted =
    new Set(expected.filter((ch) => !bare(ch)).map((ch) => String(ch.word_comment_id))).size +
    uncommented.length;
  return {
    changes: changes.length,
    expected_comments: wanted,
    tool_comments: ids.size,
    missing_comment: missing.map((ch) => ch.id),
    uncommented: uncommented.map((ch) => ch.id),
    unrecorded_comment_ids: orphans,
    ok:
      wanted === ids.size &&
      missing.length === 0 &&
      uncommented.length === 0 &&
      orphans.length === 0,
  };
}

// The end of the status line of a refresh: "." when the report is ok, else
// what does not match. The word MISMATCH stays in every such line (the
// acceptance judge reads it).
export function invariantText(inv) {
  if (inv.ok) return ".";
  const parts = [];
  if (inv.missing_comment.length || inv.unrecorded_comment_ids.length || !inv.uncommented.length) {
    parts.push(
      `missing ${inv.missing_comment.join(",") || "-"}, unrecorded ${inv.unrecorded_comment_ids.join(",") || "-"}`,
    );
  }
  if (inv.uncommented.length) {
    parts.push(
      `no rationale comment on ${inv.uncommented.join(",")}: applied in Word, the comment could not be added`,
    );
  }
  return ` (MISMATCH: ${parts.join("; ")})`;
}

// What the author should do next, after a status refresh, in one sentence.
// The full counts go under "Details" (the pane writes them there).
export function countsSentence(c, inv, { outside = 0, unresolved = 0 } = {}) {
  const n = (k, one, many) => `${k} ${k === 1 ? one : many}`;
  let line;
  if (!inv.ok) {
    line =
      "The changes on record and the comments in Word do not match. Open Details to see which, then check those comments in Word.";
  } else if (c.pending > 0) {
    line = `${n(c.pending, "change is", "changes are")} waiting for your decision in Word.`;
  } else if (inv.changes > 0) {
    line = "Every change is decided in Word.";
  } else {
    line = "No changes yet. Ask for changes on a comment to begin.";
  }
  if (c.unlocated > 0) {
    line += ` ${n(c.unlocated, "change was", "changes were")} not found in the document.`;
  }
  if (outside > 0) {
    line += ` ${n(outside, "change was", "changes were")} made outside the tool; see "Changes made outside the tool".`;
  }
  if (unresolved > 0) {
    line += ` Word would not mark ${n(unresolved, "rationale comment", "rationale comments")} as resolved; that happens in a document in compatibility mode.`;
  }
  return line;
}

// Design 0 C1: once the author accepts or rejects a change, its rationale
// comment is marked resolved. Returns ids of open comments to resolve.
export function commentsToResolve(rows, revision, comments) {
  const open = new Set((comments ?? []).filter((c) => !c.resolved).map((c) => String(c.id)));
  const byId = new Map();
  for (const c of revision?.comments ?? []) for (const ch of c.changes ?? []) byId.set(ch.id, ch);
  // A comment the parts of a rewrite share is resolved when every part is decided.
  const decided = new Map();
  for (const r of rows ?? []) {
    const id = byId.get(r.change_id)?.word_comment_id;
    if (id == null) continue;
    const done = r.state === "accepted" || r.state === "rejected";
    decided.set(String(id), (decided.get(String(id)) ?? true) && done);
  }
  return [...decided].filter(([id, all]) => all && open.has(id)).map(([id]) => id);
}

export const STATE_LABELS = {
  pending: "Pending",
  accepted: "Accepted",
  rejected: "Rejected",
  user_edited: "Edited by you",
  unlocated: "Not found",
};

export const STATUS_LABELS = { pending: "To do", in_progress: "In progress", done: "Done" };

// Ledger status shown on an item row. A skipped item is outside the
// manuscript scope and is checked at its own stage, not counted as clean.
export const SKIPPED_SCOPE_LABELS = {
  response: "Checked at the response stage",
  supplement: "Checked at the supplement stage",
};

const unsureText = (item) =>
  `unsure about ${item.unsure.length} paragraph${item.unsure.length === 1 ? "" : "s"}`;

const places = (n) => `${n} place${n === 1 ? "" : "s"}`;

export function ledgerStatusText(item) {
  if (item?.status === "violations")
    return item.unsure?.length
      ? `Found in ${places(item.violation_count)}, ${unsureText(item)}`
      : `Found in ${places(item.violation_count)}`;
  if (item?.status === "unsure") return unsureText(item);
  // The check did not run to its end (a model reply that could not be used
  // or a failed call, a regex that does not compile): no paragraph of it is
  // taken as checked.
  if (item?.status === "error") return "check failed: nothing counts as passed";
  if (item?.status === "skipped") {
    if (item.check?.type === "count") return "Checked in the Tools tab";
    return SKIPPED_SCOPE_LABELS[item.scope] ?? `Checked at the ${item.scope} stage`;
  }
  if (item?.status === "unchecked") return "Not checked yet";
  if (item?.status === "ok") return "No problem found";
  return item?.status;
}

// How a rule is checked, in words (the type and the pattern stay in Advanced).
export function ledgerHowText(item) {
  const k = item?.check?.type;
  if (k === "regex") return "looks for the exact words";
  if (k === "llm") return "Claude reads each paragraph";
  if (k === "count") return "counted in the Tools tab";
  return "";
}

// Parse the ledger form into the ledger item shape of the daemon.
export function ledgerItemFromForm({ kind, text, scope, checkType, pattern, flags, expect }) {
  const t = String(text ?? "").trim();
  if (!t) return { error: "Write the instruction text." };
  if (checkType === "regex") {
    const p = String(pattern ?? "");
    if (!p) return { error: "A regex check needs a pattern." };
    try {
      new RegExp(p, String(flags ?? "").replace(/g/g, ""));
    } catch (e) {
      return { error: `Invalid regex: ${e.message}` };
    }
    return {
      item: {
        kind,
        text: t,
        scope: scope || "manuscript",
        check: {
          type: "regex",
          pattern: p,
          flags: String(flags ?? "").replace(/g/g, ""),
          expect: expect || "absent",
        },
      },
    };
  }
  return { item: { kind, text: t, scope: scope || "manuscript", check: { type: "llm" } } };
}

// ---- L1-L3 badges, L5 requirements, L6 ledger drafts, L7, L9 ----------------

// Short labels shown on a change row (L1 lost facts, L2 unsourced, L3 narrowed)
// and a change applied without its rationale comment.
export function changeFlags(ch) {
  const out = [];
  if (ch?.comment_missing) out.push({ kind: "comment-missing", text: "No rationale comment" });
  // lost_facts (daemon/edit-facts.mjs lostFacts): numbers with units, versions,
  // parenthesised text and names that the deleted text had and the new text
  // does not.
  if (ch?.lost_facts?.length) {
    out.push({
      kind: "lost",
      text: `The old wording had ${ch.lost_facts.join("; ")}. The new wording does not. Reject the change in Word if that is not meant.`,
    });
  }
  if (ch?.unsourced_terms?.length)
    out.push({ kind: "unsourced", text: `No source: ${ch.unsourced_terms.join(", ")}` });
  if (ch?.from_review?.length)
    out.push({ kind: "unsourced", text: `From the review: ${ch.from_review.join(", ")}` });
  // narrowed_from (daemon/edit-facts.mjs narrowEdit, review-tools.mjs recordApplied):
  // the task pane cut the words the old and the new text share at both ends.
  if (ch?.narrowed_from) {
    out.push({
      kind: "narrowed",
      text: "Only the words that differ are marked in Word, not the whole phrase Claude gave. Nothing to do.",
    });
  }
  if (ch?.rewrite_id) out.push({ kind: "narrowed", text: "Rewrite" });
  return out;
}

// One requirement per non-empty line of the text box.
export function requirementsFromText(text) {
  return String(text ?? "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
}

// L5. The state of every request of a comment, in list order. The first
// that holds:
//   changed            a change in Word that is not rejected names it
//   figure_changed     an adopted figure change names it (figures page)
//   changed_elsewhere  an accepted change in another file of the revision
//                      names it (file: the name of that file)
//   needs_author       a question with its id has no answer yet
//   figure_waiting     a figure change that names it waits for a decision
//   waiting_elsewhere  a change in another file of the revision names it and
//                      waits for the author's decision there (file)
//   outside_this_file  the model noted it as belonging to another file (why)
//   open               none of these: nobody spoke of it
// A turn has to end every request as changed, needs_author or
// outside_this_file; open is what a request is until then, so a request
// passed over in silence shows. A rejected figure change counts for nothing,
// like a rejected change, in this file or in another.
// elsewhere: the changes on record for this comment in the other files of
// the revision (changesElsewhere), each in the state the last status refresh
// of that file gave it. The two states they give are never saved.
export function requirementStates(comment, elsewhere = []) {
  const used = new Set();
  for (const ch of comment?.changes ?? []) {
    if (ch.state === "rejected") continue;
    for (const id of ch.rationale?.requirement_ids ?? []) used.add(id);
  }
  const figureIds = (state) =>
    new Set(
      (comment?.figure_changes ?? [])
        .filter((f) => f.state === state)
        .flatMap((f) => f.requirement_ids ?? []),
    );
  const adopted = figureIds("adopted");
  const proposed = figureIds("proposed");
  const asked = new Set(
    (comment?.questions ?? []).filter((q) => !q.answer).map((q) => q.requirement_id),
  );
  // A note with a part covers only that part of its request: the request
  // gets its state from the rest, and the part a row of its own after it.
  const notes = comment?.requirements_outside ?? [];
  const outside = new Map(notes.filter((n) => !n.part).map((n) => [n.id, n.why]));
  const partOf = new Map(notes.filter((n) => n.part).map((n) => [n.id, n]));
  // id -> the file of the first change elsewhere that names it, decided and waiting.
  const away = (states) => {
    const files = new Map();
    for (const x of elsewhere) {
      if (!states.includes(x.state)) continue;
      for (const id of x.requirement_ids ?? []) if (!files.has(id)) files.set(id, x.file);
    }
    return files;
  };
  const changedAway = away(DECIDED_ELSEWHERE);
  const waitingAway = away(WAITING_ELSEWHERE);
  return (comment?.requirements ?? []).flatMap((r) => {
    const state = used.has(r.id)
      ? "changed"
      : adopted.has(r.id)
        ? "figure_changed"
        : changedAway.has(r.id)
          ? "changed_elsewhere"
          : asked.has(r.id)
            ? "needs_author"
            : proposed.has(r.id)
              ? "figure_waiting"
              : waitingAway.has(r.id)
                ? "waiting_elsewhere"
                : outside.has(r.id)
                  ? "outside_this_file"
                  : "open";
    const row = {
      ...r,
      state,
      ...(state === "outside_this_file" ? { why: outside.get(r.id) } : {}),
      ...(state === "changed_elsewhere" ? { file: changedAway.get(r.id) } : {}),
      ...(state === "waiting_elsewhere" ? { file: waitingAway.get(r.id) } : {}),
    };
    const p = partOf.get(r.id);
    if (!p || state === "outside_this_file") return [row];
    return [row, { ...r, state: "outside_this_file", why: p.why, part: p.part }];
  });
}

// A change in another file of the revision, by its stored state: decided in
// favour of the change, or still waiting for the author in that file. A
// rejected one counts for nothing.
const DECIDED_ELSEWHERE = ["accepted", "user_edited"];
const WAITING_ELSEWHERE = ["pending", "unlocated"];
// The changes on record for one comment in the other files of the revision
// (linked.changes of the last load: [{ file, ref, id, state, requirement_ids }]).
export const changesElsewhere = (linked, ref) =>
  (linked?.changes ?? []).filter((x) => x.ref === ref);
// The files in which changes of a comment wait for the author, and those in
// which one was accepted: { waiting: [{ file, n }], changed: [file] }.
function elsewhereSummary(elsewhere = []) {
  const waiting = new Map();
  const changed = new Set();
  for (const x of elsewhere) {
    if (WAITING_ELSEWHERE.includes(x.state)) waiting.set(x.file, (waiting.get(x.file) ?? 0) + 1);
    else if (DECIDED_ELSEWHERE.includes(x.state)) changed.add(x.file);
  }
  return { waiting: [...waiting].map(([file, n]) => ({ file, n })), changed: [...changed] };
}

const inState = (comment, state, elsewhere = []) =>
  requirementStates(comment, elsewhere).filter((r) => r.state === state);

// The address of the figures page (daemon/figure-http.mjs) for a manuscript,
// at one comment. origin: where the daemon serves the pane and that page.
// doc: the document as this pane named it to the daemon (hello active_doc);
// the page finds the manuscript by that name. Neither is needed.
export function figuresUrl(origin, doc = null, ref = null) {
  const query = new URLSearchParams();
  if (doc) query.set("doc", doc);
  if (ref) query.set("ref", ref);
  const q = query.toString();
  return `${String(origin ?? "").replace(/\/+$/, "")}/figures/${q ? `?${q}` : ""}`;
}

// Has this comment something on the figures page: a listed request that
// belongs to another file, or a figure change in any state.
export function hasFigureWork(comment) {
  const listed = new Set((comment?.requirements ?? []).map((r) => r.id));
  return (
    (comment?.figure_changes ?? []).length > 0 ||
    (comment?.requirements_outside ?? []).some((n) => listed.has(n.id))
  );
}

// Figure changes of a comment that are neither adopted nor rejected, and the
// words for them: "1 figure or file waits for your decision".
const proposedFigures = (comment) =>
  (comment?.figure_changes ?? []).filter((f) => f.state === "proposed").length;
const figuresWaitingText = (n) =>
  `${n} figure${n === 1 ? "" : "s"} or file${n === 1 ? "" : "s"} ${n === 1 ? "waits" : "wait"} for your decision`;

// Requests of a comment nobody spoke of (open).
export function unaddressedRequirements(comment, elsewhere = []) {
  return inState(comment, "open", elsewhere).map(({ state, ...r }) => r);
}

// What a request is, in words, for the list "What the reviewer asks".
export function requestStateText(r) {
  if (r.state === "changed") return "Changed";
  if (r.state === "figure_changed") return "Changed in a figure or other file";
  if (r.state === "changed_elsewhere") return `Changed in ${r.file}`;
  if (r.state === "needs_author") return "Needs your answer";
  if (r.state === "figure_waiting") return "A figure or file waits for your decision";
  if (r.state === "waiting_elsewhere") return `Waiting for your decision in ${r.file}`;
  if (r.state === "outside_this_file" && r.part)
    return `This part is in another file: ${r.part}${r.why ? ` (${r.why})` : ""}`;
  if (r.state === "outside_this_file")
    return r.why ? `In another file: ${r.why}` : "In another file";
  return "Not yet";
}

// After a turn: "2 requests of R1.2 have no change yet: ...", the requests in words; "" when none is.
export function openRequestsText(comment, elsewhere = []) {
  const open = inState(comment, "open", elsewhere);
  if (!open.length) return "";
  return `${open.length} request${open.length === 1 ? "" : "s"} of ${comment.ref} ${open.length === 1 ? "has" : "have"} no change yet: ${open.map((r) => `\u201c${r.text}\u201d`).join("; ")}.`;
}

// ---- "Check revised paragraphs" (daemon/paragraph-check.mjs) ---------------
const CONSISTENCY_KIND = {
  contradiction: "the paragraph contradicts itself.",
  conflict: "changes of different comments work against each other.",
  review_as_source: "a sentence takes the review as its source.",
  claim_left: "a claim a change withdrew is still made.",
  open_question: "a sentence rests on what an open question to the author is about.",
};

// The comment whose turn rewrites a paragraph: the one with the most changes
// in it (the first of them when equal). Change ids are "<ref>-<n>".
function rewriteRef(changeIds) {
  const n = new Map();
  for (const id of changeIds) {
    const ref = String(id).replace(/-\d+$/, "");
    n.set(ref, (n.get(ref) ?? 0) + 1);
  }
  return [...n].sort((a, b) => b[1] - a[1])[0]?.[0] ?? null;
}

// The message line after a check. model: the label of what it ran on, or null.
export function consistencyMessage(c, model) {
  const n = c.paragraphs.length;
  const more = (c.warnings ?? []).length ? ` ${c.warnings.join(" ")}` : "";
  if (!n) return `No paragraph holds two or more changes: nothing to check.${more}`;
  const k = c.findings.length;
  return (
    `Checked ${n} revised paragraph${n === 1 ? "" : "s"}${model ? ` with ${model}` : ""}: ` +
    `${k ? `${k} problem${k === 1 ? "" : "s"}` : "no problem found"}.${more}`
  );
}

// The lines under the button: what was found, then the paragraphs that hold
// three or more changes; those carry rewrite_ref, the comment whose turn
// would write the paragraph anew. [{ paragraph_index, text, rewrite_ref? }]
export function consistencyLines(c) {
  if (!c) return [];
  return [
    ...c.findings.map((f) => ({
      paragraph_index: f.paragraph_index,
      text: `${CONSISTENCY_KIND[f.kind] ?? f.kind} ${f.detail}`,
    })),
    ...c.paragraphs
      .filter((p) => p.many)
      .map((p) => ({
        paragraph_index: p.index,
        text: `This paragraph holds ${p.change_ids.length} changes: read it as a whole and rewrite it if it no longer reads as one paragraph.`,
        rewrite_ref: rewriteRef(p.change_ids),
      })),
  ];
}

// What is on record about the paragraphs a comment changed, for its card
// and for its next turn: findings of the last paragraph check that say the
// paragraph contradicts itself or still makes a withdrawn claim, and
// rewrites Word stopped part way. Derived from the review file; nothing is
// stored for it. A finding counts for a comment when the checked paragraph
// held one of its changes; a rewrite stays until all of its changes are
// rejected.
export const SURFACED_FINDINGS = ["contradiction", "claim_left"];

export function paragraphFindings(comment, consistency) {
  const own = new Set(
    (comment?.changes ?? []).filter((ch) => ch.state !== "rejected").map((ch) => ch.id),
  );
  if (!consistency || !own.size) return [];
  const held = new Map((consistency.paragraphs ?? []).map((p) => [p.index, p.change_ids ?? []]));
  return (consistency.findings ?? []).filter(
    (f) =>
      SURFACED_FINDINGS.includes(f.kind) &&
      [...(held.get(f.paragraph_index) ?? []), ...(f.change_ids ?? [])].some((id) => own.has(id)),
  );
}

export function stoppedRewrites(comment) {
  const rejected = new Set(
    (comment?.changes ?? []).filter((ch) => ch.state === "rejected").map((ch) => ch.id),
  );
  return (comment?.rewrites_stopped ?? []).filter(
    (w) => !w.change_ids.length || w.change_ids.some((id) => !rejected.has(id)),
  );
}

const stoppedText = (w) =>
  `The rewrite of this paragraph stopped after ${w.applied} of ${w.parts} parts (${w.code}): ${w.change_ids.join(", ") || "its changes"} ${w.change_ids.length === 1 ? "is" : "are"} in the document and the paragraph may read unevenly.`;

// [{ kind: "finding" | "rewrite_stopped", paragraph_index, paragraph_id, text }]
export function commentNotices(comment, consistency) {
  return [
    ...paragraphFindings(comment, consistency).map((f) => ({
      kind: "finding",
      paragraph_index: f.paragraph_index,
      paragraph_id: f.paragraph_id ?? `p${f.paragraph_index}`,
      text: `The paragraph check found: ${CONSISTENCY_KIND[f.kind] ?? f.kind} ${f.detail}`,
    })),
    ...stoppedRewrites(comment).map((w) => ({
      kind: "rewrite_stopped",
      paragraph_index: w.paragraph_index,
      paragraph_id: w.paragraph_id ?? (w.paragraph_index == null ? null : `p${w.paragraph_index}`),
      text: stoppedText(w),
    })),
  ];
}

// What is still open when a comment is marked Done; the tab asks once more
// when the list is not empty. ledger: revision.ledger (the confirmed rules,
// with the status of their last check).
//   - requests nobody spoke of, and requests that wait for the author (L5)
//   - figure changes not decided on the figures page
//   - changes not decided in Word: pending, or not found in the document
//   - changes applied without their rationale comment (invariantReport's
//     `uncommented`), and edits Word did not answer (outcome_unknown)
//   - ledger rules whose last check found violations
export function doneBlockers(comment, ledger, elsewhere = [], consistency = null) {
  const out = [];
  const open = unaddressedRequirements(comment, elsewhere);
  if (open.length) out.push(`${open.map((r) => `"${r.text}"`).join(", ")} has no change`);
  // Changes of this comment in another file of the revision that wait there.
  for (const w of elsewhereSummary(elsewhere).waiting) {
    out.push(`${w.n} ${w.n === 1 ? "change is" : "changes are"} not decided in ${w.file}`);
  }
  const asked = inState(comment, "needs_author", elsewhere);
  if (asked.length)
    out.push(
      `${asked.map((r) => `"${r.text}"`).join(", ")} ${asked.length === 1 ? "waits" : "wait"} for the author's answer`,
    );
  const figures = proposedFigures(comment);
  if (figures) out.push(`${figuresWaitingText(figures)} on the figures page`);
  const n = (state) => (comment?.changes ?? []).filter((ch) => ch.state === state).length;
  const pending = n("pending");
  const unlocated = n("unlocated");
  if (pending + unlocated) {
    const parts = [pending && `${pending} pending`, unlocated && `${unlocated} not found`];
    const total = pending + unlocated;
    out.push(
      `${total} ${total === 1 ? "change is" : "changes are"} not decided in Word (${parts.filter(Boolean).join(", ")})`,
    );
  }
  const bare = (comment?.changes ?? []).filter(
    (ch) => ch.word_comment_id == null && !!ch.comment_missing,
  );
  if (bare.length)
    out.push(
      `${bare.length} ${bare.length === 1 ? "change has" : "changes have"} no rationale comment in Word (${bare.map((ch) => ch.id).join(", ")})`,
    );
  const unknown = (comment?.blocked ?? []).filter((b) => b.state === "outcome_unknown");
  if (unknown.length)
    out.push(
      `${unknown.length} ${unknown.length === 1 ? "edit" : "edits"} got no answer from Word and may be applied (${unknown.map((b) => b.id).join(", ")})`,
    );
  for (const n of commentNotices(comment, consistency)) {
    const at = n.paragraph_id ? `paragraph ${n.paragraph_id}` : "a paragraph";
    out.push(
      n.kind === "finding"
        ? `the paragraph check found a problem in ${at} this comment changed`
        : `the rewrite of ${at} stopped part way`,
    );
  }
  const dirty = (ledger ?? []).filter((i) => i.status === "violations");
  if (dirty.length)
    out.push(
      `the ledger check has violations (${dirty.map((i) => `${i.id}: ${i.violation_count}`).join(", ")})`,
    );
  return out;
}

const escapeRegex = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

// L6. Ledger form fields drafted from a change the author rejected: new text
// that was rejected becomes a forbid rule; a rejected deletion becomes a
// "must stay" rule on the deleted text. The author edits the draft before Add.
export function ledgerDraftFromChange(ch) {
  const newText = String(ch?.text ?? "").trim();
  const tag = ch?.id ? ` (rejected ${ch.id})` : "";
  if (ch?.kind === "delete" || !newText) {
    const kept = String(ch?.find ?? "").trim();
    return {
      kind: "instruction",
      text: `Keep "${kept}"${tag}`,
      scope: "manuscript",
      checkType: "regex",
      pattern: escapeRegex(kept),
      flags: "",
      expect: "present",
    };
  }
  return {
    kind: "forbid",
    text: `Do not write "${newText}"${tag}`,
    scope: "manuscript",
    checkType: "regex",
    pattern: escapeRegex(newText),
    flags: "i",
    expect: "absent",
  };
}

// L6. Form fields of an existing item, for editing it (ledger_update).
export function ledgerFormFromItem(item) {
  const c = item?.check ?? {};
  return {
    kind: item?.kind ?? "instruction",
    text: item?.text ?? "",
    scope: item?.scope ?? "manuscript",
    checkType: c.type === "regex" ? "regex" : "llm",
    pattern: c.type === "regex" ? c.pattern : "",
    flags: c.type === "regex" ? (c.flags ?? "") : "",
    expect: c.type === "regex" ? (c.expect ?? "absent") : "absent",
  };
}

// L6. Draft a check from what the author typed: a quoted phrase ("...",
// '...', or corner brackets) becomes a literal regex; without one the draft
// is an LLM check. Whether the phrase is forbidden or must stay is left to
// the author (the form keeps its current choice).
export function ledgerCheckFromText(text) {
  const m = /"([^"]+)"|'([^']+)'|「([^」]+)」|“([^”]+)”/.exec(String(text ?? ""));
  const phrase = m ? (m[1] ?? m[2] ?? m[3] ?? m[4]).trim() : "";
  return phrase
    ? { checkType: "regex", pattern: escapeRegex(phrase), flags: "i" }
    : { checkType: "llm", pattern: "", flags: "" };
}

// L7. Everything that waits for the author, in one list.
// A normal edit that is with Word right now: the daemon writes it to the
// file before the call (review-service.mjs onAttempt) and takes it out with
// the answer. Nothing for the author to decide, so it is not shown. One
// that got no answer is outcome_unknown and is shown.
export const inFlightEdit = (b) => b.state === "applying" && (b.blocked_by ?? []).length === 0;

// linked: the link of the document as the last load gave it (its changes
// tell where a request stands in the other files of the revision).
export function pendingDecisions(revision, linked = null) {
  const blocked = [];
  const requirements = [];
  const ruleText = (id) => (revision?.ledger ?? []).find((i) => i.id === id)?.text ?? id;
  const quoted = (t) => `“${String(t).replace(/[.\s]+$/, "")}”`;
  const named = (i) => quoted(i.text ?? i.id);
  for (const c of revision?.comments ?? []) {
    const who = `${c.ref}${c.section ? ` (${c.section})` : ""}`;
    for (const b of c.blocked ?? []) {
      if (inFlightEdit(b)) continue;
      const newText = b.args?.replace ?? b.args?.insert_after ?? "";
      blocked.push({
        ref: c.ref,
        id: b.id,
        kind: "edit_blocked",
        // No forbid rule: a normal edit Word did not answer (outcome unknown).
        text: b.blocked_by.length
          ? `${who}: an edit (${quoted(newText)}) was stopped by the rule ${b.blocked_by.map((r) => quoted(ruleText(r))).join(", ")}. Open the comment to decide.`
          : `${who}: Word did not answer an edit (${quoted(newText)}). Open the comment to check it.`,
      });
    }
    for (const r of requirementStates(c, changesElsewhere(linked, c.ref))) {
      const kind = {
        open: "request_open",
        needs_author: "request_asked",
        figure_waiting: "request_figure",
      }[r.state];
      if (!kind) continue;
      requirements.push({
        ref: c.ref,
        id: r.id,
        kind,
        text:
          r.state === "open"
            ? `${who}: nothing changes yet for ${quoted(r.text)}.`
            : r.state === "needs_author"
              ? `${who}: Claude waits for your answer about ${quoted(r.text)}.`
              : `${who}: a figure or file for ${quoted(r.text)} waits for your decision on the figures page.`,
      });
    }
  }
  // An unsure paragraph (the LLM check abstained) is the author's call.
  const ledger = (revision?.ledger ?? [])
    .filter((i) => i.status === "unchecked" || i.status === "error" || i.unsure?.length)
    .map((i) => ({
      id: i.id,
      kind: "rule",
      text: i.unsure?.length
        ? `The rule ${named(i)} is ${unsureText(i)}. Look at them in Word.`
        : i.status === "error"
          ? `The check of the rule ${named(i)} failed.`
          : `The rule ${named(i)} is not checked yet.`,
    }));
  return {
    blocked,
    ledger,
    requirements,
    total: blocked.length + ledger.length + requirements.length,
  };
}

// ---- The list of comments: one compact row each --------------------------------

// m:ss from a start time to `now` (ms, or an ISO string for the start).
export function clockText(startedAt, now = Date.now()) {
  const t0 = typeof startedAt === "string" ? Date.parse(startedAt) : startedAt;
  const s = Math.max(0, Math.floor((now - t0) / 1000));
  return `${Math.floor(s / 60)}:${String(s % 60).padStart(2, "0")}`;
}

// The first line of a comment, cut at `max` characters on a word end.
export function firstLine(text, max = 100) {
  const line =
    String(text ?? "")
      .split(/\r?\n/)
      .map((l) => l.trim())
      .find(Boolean) ?? "";
  if (line.length <= max) return line;
  const cut = line.slice(0, max);
  return `${cut.slice(0, Math.max(cut.lastIndexOf(" "), max >> 1))}...`;
}

// Where a comment stands, in words, from what the pane already knows (the
// review file, plus what it holds in memory for this session). Nothing is
// stored for it.
//   working       a turn of this comment runs now (working: { startedAt })
//   done          the comment is marked done
//   needs_answer  a question to the author has no answer, or an edit is
//                 blocked and waits for a decision
//   to_review     changes wait for the author's decision in Word, or a
//                 figure change waits for it on the figures page
//   no_change     the last turn of this session proposed nothing (noChange),
//                 or every request belongs to another file
//   decided       changes (or adopted figure changes) exist and none waits
//   not_started   none of the above
// attention: the author has something to do here.
// elsewhere: the changes of this comment in the other files of the revision
// (changesElsewhere). One that waits there is to_review, named with its
// file; an accepted one is decided when this file has nothing of its own.
export function commentState(
  comment,
  { working = null, now = Date.now(), noChange = false, elsewhere = [] } = {},
) {
  if (working) {
    return {
      kind: "working",
      label: `Working ${clockText(working.startedAt, now)}`,
      attention: false,
    };
  }
  if (comment?.status === "done") return { kind: "done", label: "Done", attention: false };
  const asked = (comment?.questions ?? []).filter((q) => !q.answer).length;
  const blocked = (comment?.blocked ?? []).filter((b) => !inFlightEdit(b)).length;
  if (asked) return { kind: "needs_answer", label: "Needs your answer", attention: true };
  if (blocked) return { kind: "needs_answer", label: "Needs your decision", attention: true };
  const changes = comment?.changes ?? [];
  const waiting = changes.filter((ch) => ch.state === "pending" || ch.state === "unlocated").length;
  if (waiting) {
    return {
      kind: "to_review",
      label: `${waiting} change${waiting === 1 ? "" : "s"} to review in Word`,
      attention: true,
    };
  }
  const figures = proposedFigures(comment);
  if (figures) return { kind: "to_review", label: figuresWaitingText(figures), attention: true };
  const away = elsewhereSummary(elsewhere);
  if (away.waiting.length) {
    return {
      kind: "to_review",
      label: away.waiting
        .map((w) => `${w.n} change${w.n === 1 ? "" : "s"} to review in ${w.file}`)
        .join(", "),
      attention: true,
    };
  }
  if (changes.length || (comment?.figure_changes ?? []).some((f) => f.state === "adopted"))
    return { kind: "decided", label: "All changes decided", attention: false };
  if (away.changed.length) {
    return { kind: "decided", label: `Changed in ${away.changed.join(", ")}`, attention: false };
  }
  if (noChange) return { kind: "no_change", label: "No change proposed", attention: true };
  const reqs = requirementStates(comment, elsewhere);
  if (reqs.length && reqs.every((r) => r.state === "outside_this_file")) {
    return { kind: "no_change", label: "Belongs to another file", attention: false };
  }
  return { kind: "not_started", label: "Not started", attention: false };
}

// "20 comments, 2 done. 3 need you." The header of the list.
export function summaryText(comments, states) {
  const n = comments.length;
  if (!n) return "";
  const need = states.filter((s) => s.attention).length;
  const done = states.filter((s) => s.kind === "done").length;
  return (
    `${n} comment${n === 1 ? "" : "s"}` +
    (done ? `, ${done} done` : "") +
    `. ${need ? `${need} ${need === 1 ? "needs" : "need"} you.` : "None needs you right now."}`
  );
}

// "Review: 5 comments from 1 reviewer", the line the paste box shrinks to.
export function reviewLine(comments) {
  const n = comments.length;
  const refs = comments.map((c) => reviewerOfRef(c.ref));
  const reviewers = new Set(refs.filter((r) => r > 0)).size;
  const editor = refs.includes(0);
  const from = [
    reviewers ? `${reviewers} reviewer${reviewers === 1 ? "" : "s"}` : "",
    editor ? "the editor" : "",
  ].filter(Boolean);
  return `Review: ${n} comment${n === 1 ? "" : "s"}${from.length ? ` from ${from.join(" and ")}` : ""}`;
}

// L9. "R1.3 running 1:05. 2 of 3 comments not done."
export function progressText({ ref, startedAt, now = Date.now(), comments = [] }) {
  const clock = clockText(startedAt, now);
  const open = comments.filter((c) => c.status !== "done").length;
  return `${ref} running ${clock}. ${open} of ${comments.length} comments not done.`;
}

// P15. The comments "Propose all" runs, in list order: every one not done.
export function queueRefs(revision) {
  return (revision?.comments ?? []).filter((c) => c.status !== "done").map((c) => c.ref);
}

// ---- P15: review text from a file --------------------------------------------

export const REVIEW_FILE_MAX_BYTES = 2 * 1024 * 1024;

// A picked review file ({ name, size, text }) as text for the paste box:
// .txt or .md only, at most 2 MB, BOM removed, line ends as \n.
export function reviewTextFromFile({ name, size, text }) {
  if (!/\.(txt|md)$/i.test(String(name ?? ""))) {
    return { ok: false, error: `${name}: choose a .txt or .md file.` };
  }
  if (size > REVIEW_FILE_MAX_BYTES) {
    return { ok: false, error: `${name} is larger than 2 MB.` };
  }
  const t = String(text ?? "")
    .replace(/^\uFEFF/, "")
    .replace(/\r\n?/g, "\n");
  if (!t.trim()) return { ok: false, error: `${name} is empty.` };
  return { ok: true, text: t };
}

// ---- P15: ledger file --------------------------------------------------------

// The text of a picked ledger file as the ledger_import payload. Only the
// outer shape is checked here ({ items: [...] }, an exported ledger or a
// hand-written list); the daemon checks every rule.
export function ledgerImportFromJson(text) {
  let data;
  try {
    data = JSON.parse(String(text ?? "").replace(/^\uFEFF/, ""));
  } catch (e) {
    return { ok: false, error: `The ledger file is not JSON (${e.message}).` };
  }
  if (!Array.isArray(data?.items)) {
    return { ok: false, error: 'The ledger file needs an "items" list.' };
  }
  if (!data.items.length) return { ok: false, error: "The ledger file has no items." };
  return { ok: true, doc: { items: data.items }, count: data.items.length };
}

// ---- P15: ledger rule candidates from the review text ---------------------------

// The rule fields of a candidate (ledger_suggest), as ledger_import takes them.
export function candidateRule({ kind, text, scope, check }) {
  return { kind, text, scope, check };
}

// "Forbidden, must not match /novel/i: <sentence>"
export function candidateLabel(c) {
  const kind = c.kind === "forbid" ? "Forbidden" : "Instruction";
  const k = c.check ?? {};
  const check =
    k.type === "regex"
      ? `${k.expect === "present" ? "must keep" : "must not match"} /${k.pattern}/${k.flags ?? ""}`
      : "LLM check";
  return `${kind}, ${check}: ${c.text}`;
}

// What the card of a comment says after an edit of its turn was taken back
// and Word took an earlier pending change of the paragraph back with it
// (review_event undo_took_more).
export const UNDO_TOOK_MORE_TEXT =
  "While taking back an edit that would have damaged a citation, Word also took back an earlier pending change in the same paragraph. Check the paragraph and ask for that change again if you still want it.";

// ---- "Before submission": the checklist (daemon/submission-check.mjs) -------

export const SUBMISSION_STATE_LABELS = {
  ready: "Ready",
  not_ready: "Not ready",
  not_applicable: "Not needed",
};

// The count in the header of the section once a check has run.
export const submissionCount = (rows) =>
  `(${(rows ?? []).filter((r) => r.state === "not_ready").length} not ready)`;

const TAB_NAMES = { review: "Review", response: "Response", tools: "Tools" };

// What the button of a row says, by where the row is fixed ("" for a row
// without one). here: the tab the checklist is on.
export function submissionGoLabel(go, here = "review") {
  if (!go) return "";
  if (Number.isInteger(go.paragraph)) return "Show in Word";
  if (!go.tab || !go.target) return "";
  return go.tab === here ? "Go there" : `Go to the ${TAB_NAMES[go.tab] ?? go.tab} tab`;
}

// ---- a supplementary file in the revision of a manuscript -------------------------
// (daemon/revision-parts.mjs linkedSummary: part_of is the manuscript of this
// file, parts are the supplementary files of this manuscript.)

// The button of a manuscript this file could join.
export const joinLabel = (c) =>
  `Join the revision of ${c.name}, ${c.comments} comment${c.comments === 1 ? "" : "s"}`;
// The line of a manuscript that has supplementary files.
export const linkedFilesText = (parts) => `Linked files: ${parts.map((p) => p.name).join(", ")}`;
// The line of a supplementary file, and the note where something of the
// manuscript is shown.
export const partOfText = (main) =>
  `This file is part of the revision of ${main.name}. The comments, requests and rules below are edited in ${main.name}.`;
export const editedInText = (main) => `edited in ${main.name}`;
