// Paragraph-level comparison of two document snapshots. Pure.
//   L4 outsideChanges: paragraphs that changed since the last snapshot and
//      that no recorded review change explains (Codex, direct edits in Word).
//   L10 paragraphHash: the key of an incremental ledger check.
// Texts are the reviewed view (Word getReviewedText: pending changes counted
// as accepted), normalized like revision-classify.

import { normalizeParaText, boundedMatches } from "./revision-classify.mjs";
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

// Their task pane ids in the same order (null where the pane sent none).
export function snapshotIds(snapshot) {
  return [...(snapshot?.paragraphs ?? [])]
    .sort((a, b) => a.index - b.index)
    .map((p) => p.id ?? null);
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

const CONTEXT_WORDS = 3;
const collapse = (s) =>
  String(s ?? "")
    .replace(/\u0005/g, "")
    .replace(/\s+/g, " ");
const tailWords = (s, k) => (k > 0 ? (s.match(/\S+\s*/g) ?? []).slice(-k).join("") : "");
const headWords = (s, k) => (k > 0 ? (s.match(/\s*\S+/g) ?? []).slice(0, k).join("") : "");

// Put one recorded change into the form its state gives in the reviewed
// view (pending and accepted: the new words, rejected: the old ones). Both
// snapshots are the reviewed view, so a paragraph holds either the new form
// (L+I+R) or the old form (L+D+R) of the change; which one is found in
// `text` is swapped for the target. Context words shrink 3..0 like
// classifyByContext. Returns null when neither form, both forms or a
// repeated form is found: the change cannot be placed and explains nothing.
function applyRecorded(text, ch) {
  const kind = ch.kind;
  if (!["replace", "insert", "delete"].includes(kind)) return null;
  // Changes recorded before context existed: an insertion sits right after
  // its find, the other kinds have no context words.
  const ctx = ch.context ?? (kind === "insert" ? { left: ch.find, right: "" } : null);
  const left = collapse(ctx?.left);
  const right = collapse(ctx?.right);
  const I = kind === "delete" ? "" : collapse(ch.text);
  const D = kind === "insert" ? "" : collapse(ch.find);
  const once = (f) => boundedMatches(text, f).length === 1;
  for (let k = CONTEXT_WORDS; k >= 0; k--) {
    const L = tailWords(left, k);
    const R = headWords(right, k);
    const fI = normalizeParaText(L + I + R);
    const fD = normalizeParaText(L + D + R);
    if (fI === fD) continue;
    const hasI = once(fI);
    const hasD = once(fD);
    if (hasI === hasD) continue;
    const [from, to] = ch.state === "rejected" ? [fI, fD] : [fD, fI];
    if (hasI !== (from === fI)) return text;
    const [at] = boundedMatches(text, from);
    return normalizeParaText(text.slice(0, at) + to + text.slice(at + from.length));
  }
  return null;
}

// The old paragraph with every recorded change of the new paragraph put in
// its current form: what the paragraph would read if only the tool changes
// and the author's accept or reject happened. null when any change is not
// in an EXPLAINED state or cannot be placed.
function expectedText(before, changes) {
  let t = before;
  for (const ch of changes) {
    if (!EXPLAINED.has(ch.state)) return null;
    t = applyRecorded(t, ch);
    if (t == null) return null;
  }
  return t;
}

// oldTexts/newTexts: snapshotTexts of the last and the current snapshot.
// changes: classifyRevision rows joined to their recorded change, i.e.
// [{ paragraph_index, state, kind, find, text, context }].
// A paragraph with recorded changes is reported when its text differs from
// the old text with those changes applied (expectedText); `before` is then
// that expected text and lost_facts are what the rest of the edit dropped.
// Returns [{ kind: modified|inserted|deleted, paragraph_index, old_index,
// before, after, lost_facts }], the ones that lose facts first.
export function outsideChanges(oldTexts, newTexts, changes = []) {
  const byPara = new Map();
  for (const c of changes) {
    if (c.paragraph_index == null) continue;
    if (!byPara.has(c.paragraph_index)) byPara.set(c.paragraph_index, []);
    byPara.get(c.paragraph_index).push(c);
  }
  const out = [];
  for (const a of alignParagraphs(oldTexts, newTexts)) {
    if (a.same) continue;
    let before = a.old == null ? "" : oldTexts[a.old];
    const after = a.new == null ? "" : newTexts[a.new];
    if (before === after) continue;
    const own = a.new == null ? [] : (byPara.get(a.new) ?? []);
    if (own.length) {
      const expected = expectedText(before, own);
      if (expected === after) continue;
      if (expected != null) before = expected;
    }
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
