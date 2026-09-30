// Paragraph-level comparison of two document snapshots. Pure.
//   L4 outsideChanges: paragraphs that changed since the last snapshot and
//      that no recorded review change explains (Codex, direct edits in Word).
//   L10 paragraphHash: the key of an incremental ledger check.
// Texts are the reviewed view (Word getReviewedText: pending changes counted
// as accepted), normalized like revision-classify.

import { normalizeParaText } from "./revision-classify.mjs";
import { lostFacts } from "./edit-facts.mjs";

// FNV-1a 32 bit over UTF-16 code units, 8 hex digits. Not cryptographic: it
// only tells whether a paragraph changed since the last check.
export function paragraphHash(text) {
  const s = normalizeParaText(text);
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h.toString(16).padStart(8, "0");
}

// Snapshot paragraphs -> normalized texts in index order.
export function snapshotTexts(snapshot) {
  return [...(snapshot?.paragraphs ?? [])]
    .sort((a, b) => a.index - b.index)
    .map((p) => normalizeParaText(p.reviewed_text ?? p.text));
}

// Longest common subsequence of equal texts. Returns [{ old, new }] in
// document order; old or new is null for a paragraph only on one side.
// Paragraphs between two matches are paired in order as "modified".
export function alignParagraphs(oldTexts, newTexts, { maxCells = 25_000_000 } = {}) {
  const n = oldTexts.length;
  const m = newTexts.length;
  if (n * m > maxCells) {
    // Too large for the table: pair by index.
    const out = [];
    for (let i = 0; i < Math.max(n, m); i++)
      out.push({ old: i < n ? i : null, new: i < m ? i : null });
    return out;
  }
  const w = m + 1;
  const L = new Int32Array((n + 1) * w);
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      L[i * w + j] =
        oldTexts[i] === newTexts[j]
          ? L[(i + 1) * w + j + 1] + 1
          : Math.max(L[(i + 1) * w + j], L[i * w + j + 1]);
    }
  }
  const out = [];
  let gapOld = [];
  let gapNew = [];
  const flush = () => {
    const k = Math.max(gapOld.length, gapNew.length);
    for (let x = 0; x < k; x++) out.push({ old: gapOld[x] ?? null, new: gapNew[x] ?? null });
    gapOld = [];
    gapNew = [];
  };
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (oldTexts[i] === newTexts[j]) {
      flush();
      out.push({ old: i, new: j, same: true });
      i += 1;
      j += 1;
    } else if (L[(i + 1) * w + j] >= L[i * w + j + 1]) {
      gapOld.push(i++);
    } else {
      gapNew.push(j++);
    }
  }
  while (i < n) gapOld.push(i++);
  while (j < m) gapNew.push(j++);
  flush();
  return out;
}

// States of a recorded change that account for a new paragraph text: the
// tool edit itself (pending) or the author accepting or rejecting it.
const EXPLAINED = new Set(["pending", "accepted", "rejected"]);

// oldTexts/newTexts: snapshotTexts of the last and the current snapshot.
// rows: classifyRevision rows for the current snapshot.
// Returns [{ kind: modified|inserted|deleted, paragraph_index, old_index,
// before, after, lost_facts }], the ones that lose facts first.
export function outsideChanges(oldTexts, newTexts, rows = []) {
  const byPara = new Map();
  for (const r of rows) {
    if (r.paragraph_index == null) continue;
    if (!byPara.has(r.paragraph_index)) byPara.set(r.paragraph_index, []);
    byPara.get(r.paragraph_index).push(r.state);
  }
  const explained = (j) => {
    const states = byPara.get(j) ?? [];
    return states.length > 0 && states.every((s) => EXPLAINED.has(s));
  };
  const out = [];
  for (const a of alignParagraphs(oldTexts, newTexts)) {
    if (a.same) continue;
    const before = a.old == null ? "" : oldTexts[a.old];
    const after = a.new == null ? "" : newTexts[a.new];
    if (before === after) continue;
    if (a.new != null && explained(a.new)) continue;
    out.push({
      kind: a.old == null ? "inserted" : a.new == null ? "deleted" : "modified",
      paragraph_index: a.new,
      old_index: a.old,
      before,
      after,
      lost_facts: lostFacts(before, after, { sentenceStart: true }),
    });
  }
  return out.sort((x, y) => (y.lost_facts.length > 0) - (x.lost_facts.length > 0));
}
