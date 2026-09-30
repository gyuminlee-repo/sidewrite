// Decide what happened to each tracked edit the review flow made, from a
// snapshot of the document as it is now. Pure: no Office.js, no I/O.
//
// Word evidence (2026-09-30, spike session f2748a6f, spike tool outputs):
//   - A pending tracked replace shows BOTH the inserted and the deleted text
//     in paragraph.text, separated by U+0005 comment-anchor characters, and
//     the paragraph OOXML has w:ins/w:del elements.
//   - After accept or reject every w:ins/w:del in the paragraph is gone and
//     the text (minus U+0005) equals the "after" or the "before" sentence.
//   - A comment on the deleted range survives accept (empty range) and
//     reject. A comment on the inserted range is removed by reject.
//
// Snapshot shape (built by the task pane):
//   { paragraphs: [{ index, id, text, w_ins, w_del }],
//     comments:   [{ id, paragraph_index, paragraph_id }] }

export const CHANGE_STATES = ["pending", "accepted", "rejected", "user_edited", "unlocated"];

export function normalizeParaText(s) {
  return String(s ?? "")
    .replace(/\u0005/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function byIndex(snapshot) {
  const m = new Map();
  for (const p of snapshot.paragraphs ?? []) m.set(p.index, p);
  return m;
}

// Where does this change live now? Tried in order:
//   1. paragraph_id (Word uniqueLocalId recorded at edit time)
//   2. the paragraph that holds the rationale comment of the change
//   3. the only paragraph whose text equals the after or the before sentence
//   4. the recorded paragraph_index, if that paragraph still matches
// Returns { index, method }; index is null when nothing fits or the text
// match is ambiguous (identical paragraphs), so a caller never guesses.
export function locateParagraph(change, snapshot) {
  const paras = snapshot.paragraphs ?? [];
  const idx = byIndex(snapshot);
  if (change.paragraph_id) {
    const p = paras.find((x) => x.id && x.id === change.paragraph_id);
    if (p) return { index: p.index, method: "paragraph_id" };
  }
  if (change.word_comment_id != null) {
    const c = (snapshot.comments ?? []).find(
      (x) => String(x.id) === String(change.word_comment_id),
    );
    if (c && c.paragraph_index != null && idx.has(c.paragraph_index)) {
      return { index: c.paragraph_index, method: "comment" };
    }
  }
  const before = normalizeParaText(change.before);
  const after = normalizeParaText(change.after);
  const same = paras.filter((p) => {
    const t = normalizeParaText(p.text);
    return t === after || t === before;
  });
  if (same.length === 1) return { index: same[0].index, method: "text" };
  if (same.length > 1) return { index: null, method: "ambiguous" };
  const rec = idx.get(change.paragraph_index);
  if (rec && (rec.w_ins ?? 0) + (rec.w_del ?? 0) > 0) {
    return { index: rec.index, method: "recorded_index" };
  }
  return { index: null, method: "not_found" };
}

export function classifyChange(change, snapshot) {
  const loc = locateParagraph(change, snapshot);
  const commentIds = new Set((snapshot.comments ?? []).map((c) => String(c.id)));
  const commentPresent =
    change.word_comment_id != null && commentIds.has(String(change.word_comment_id));
  if (loc.index === null) {
    // A rejected pure insertion loses its comment (Word removes comments on
    // a rejected inserted range) and its paragraph may no longer be unique.
    if (change.kind === "insert" && !commentPresent) {
      return { state: "rejected", paragraph_index: null, method: `${loc.method}+comment_gone` };
    }
    return { state: "unlocated", paragraph_index: null, method: loc.method };
  }
  const p = byIndex(snapshot).get(loc.index);
  const t = normalizeParaText(p.text);
  const before = normalizeParaText(change.before);
  const after = normalizeParaText(change.after);
  const out = (state) => ({ state, paragraph_index: loc.index, method: loc.method });
  const revisions = (p.w_ins ?? 0) + (p.w_del ?? 0);
  if (revisions > 0) {
    // A pending replace reads as inserted+deleted text, i.e. neither side.
    // A pending pure insertion reads as "after" and a pending pure deletion
    // as "before", so for those kinds any revision mark means pending.
    if (change.kind !== "replace" || (t !== before && t !== after)) return out("pending");
  }
  if (t === after) return out("accepted");
  if (t === before) return out("rejected");
  if (change.kind === "insert" && !commentPresent) return out("rejected");
  return out("user_edited");
}

// Classify every change in a revision. Returns rows the store can apply.
export function classifyRevision(revision, snapshot) {
  const rows = [];
  for (const c of revision.comments ?? []) {
    for (const ch of c.changes ?? []) {
      rows.push({ ref: c.ref, change_id: ch.id, ...classifyChange(ch, snapshot) });
    }
  }
  return rows;
}

export function summarizeStates(rows) {
  const counts = Object.fromEntries(CHANGE_STATES.map((s) => [s, 0]));
  for (const r of rows) counts[r.state] = (counts[r.state] ?? 0) + 1;
  return counts;
}
