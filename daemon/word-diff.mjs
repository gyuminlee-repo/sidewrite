// A paragraph rewritten as a whole, turned into the tracked edits that make
// it: the words that differ, and only those, so the author sees in Word what
// changed instead of one paragraph struck out and another put in.
//
// rewriteSegments(before, after) gives the edits in document order, each a
// { find, replace } for office_tracked_edit. They are made to be applied
// from the last to the first: an edit then never moves the text of one that
// is still to come. For each edit, at the moment it is applied:
//   - `find` is in the paragraph exactly once. The differing words are
//     widened with the unchanged words before them (then after them) until
//     that holds; the same words go into `replace`, and the tool narrows the
//     tracked range back to what differs (edit-facts.mjs narrowEdit). Two
//     neighbours that cannot be told apart are made one edit.
//   - `find` starts and ends on a word or a mark, not on white space, and is
//     at most `max` characters (the limit of Word's search). A longer one is
//     cut at word boundaries: its first piece is replaced by the whole new
//     text, the others are deleted.
// Returns { segments } or { error } when the rewrite cannot be put that way
// (a piece that occurs more than once and cannot be widened).
//
// Words are compared as written (case and all); white space counts as one
// space. Differences separated by nothing but white space or one mark are
// one edit. Assumed: paragraphs of some hundred words (the table of the
// longest common subsequence has words(before) x words(after) cells).

import { tokens } from "./edit-facts.mjs";

const norm = (s) =>
  String(s ?? "")
    .replace(/\s+/g, " ")
    .trim();
// A find shorter than this is widened when it can be ("C" for "30 C").
const MIN_FIND = 4;
const isWord = (t) => /[\p{L}\p{N}_]/u.test(t);

// Runs of tokens that differ: [{ a0, a1, b0, b1 }] over token indexes.
function hunks(a, b) {
  const n = a.length;
  const m = b.length;
  const L = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      L[i][j] = a[i] === b[j] ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
    }
  }
  const out = [];
  let i = 0;
  let j = 0;
  let open = null;
  const close = () => {
    if (open) out.push({ ...open, a1: i, b1: j });
    open = null;
  };
  while (i < n || j < m) {
    if (i < n && j < m && a[i] === b[j]) {
      close();
      i += 1;
      j += 1;
    } else {
      open ??= { a0: i, b0: j };
      if (j < m && (i === n || L[i][j + 1] >= L[i + 1][j])) j += 1;
      else i += 1;
    }
  }
  close();
  // Differences with no word between them are one.
  const merged = [];
  for (const h of out) {
    const last = merged.at(-1);
    if (last && !a.slice(last.a1, h.a0).some(isWord)) {
      last.a1 = h.a1;
      last.b1 = h.b1;
    } else merged.push({ ...h });
  }
  return merged;
}

// Occurrences of needle in hay, overlapping ones too ("up up" is twice in
// "up up up").
function count(hay, needle) {
  if (!needle) return 0;
  let n = 0;
  for (let i = hay.indexOf(needle); i !== -1; i = hay.indexOf(needle, i + 1)) n += 1;
  return n;
}

export function rewriteSegments(before, after, { max = 255 } = {}) {
  const from = norm(before);
  const to = norm(after);
  if (from === to) return { segments: [] };
  const a = tokens(from);
  const b = tokens(to);
  const list = hunks(a, b);
  const segments = [];
  let k = list.length - 1;
  while (k >= 0) {
    const h = list[k];
    const low = k > 0 ? list[k - 1].a1 : 0; // not into the hunk before
    const high = k < list.length - 1 ? list[k + 1].a0 : a.length; // nor the one after
    // The paragraph as it reads when this hunk is applied: the original up
    // to its end, then the new text of everything after it.
    const now = a.slice(0, h.a1).join("") + b.slice(h.b1).join("");
    let s = h.a0;
    let e = h.a1;
    let least = MIN_FIND;
    const ok = () => {
      const f = a.slice(s, e).join("");
      return f.length >= least && f === f.trim() && count(now, f) === 1;
    };
    // The smallest widening that is found once: w more tokens in all, as
    // many of them before the hunk as fit, then fewer. A find of a letter
    // or two is not taken while a longer one can be had.
    const room = h.a0 - low + (high - h.a1);
    const widen = () => {
      for (let w = 0; w <= room; w++) {
        for (let left = Math.min(w, h.a0 - low); left >= 0; left--) {
          if (w - left > high - h.a1) break;
          s = h.a0 - left;
          e = h.a1 + (w - left);
          if (ok()) return true;
        }
      }
      return false;
    };
    if (!widen()) {
      least = 1;
      widen();
    }
    if (!ok()) {
      if (k === 0) return { error: "a changed part of the paragraph cannot be found on its own" };
      // Not to be told apart from the hunk before it: the two are one edit.
      // The edits already made for the hunks after it stay as they are.
      const prev = list[k - 1];
      list.splice(k - 1, 2, { a0: prev.a0, b0: prev.b0, a1: h.a1, b1: h.b1 });
      k -= 1;
      continue;
    }
    const find = a.slice(s, e).join("");
    const replace =
      a.slice(s, h.a0).join("") + b.slice(h.b0, h.b1).join("") + a.slice(h.a1, e).join("");
    k -= 1;
    if (find.length <= max) {
      segments.unshift({ find, replace });
      continue;
    }
    // Over the search limit: pieces cut between words (a piece may end on a
    // space), each found once when its turn comes: the pieces after it are
    // gone by then.
    const pieces = [];
    let piece = "";
    for (const t of a.slice(s, e)) {
      if (piece && (piece + t).length > max) {
        pieces.push(piece);
        piece = "";
      }
      piece += t;
    }
    pieces.push(piece);
    const head = a.slice(0, s).join("");
    const rest = now.slice(a.slice(0, e).join("").length);
    for (let p = pieces.length - 1; p >= 0; p--) {
      const here = head + pieces.slice(0, p + 1).join("") + rest;
      if (pieces[p].length > max || !pieces[p].trim() || count(here, pieces[p]) !== 1) {
        return { error: "a changed part is too long to be found in Word in pieces" };
      }
    }
    segments.unshift(
      { find: pieces[0], replace },
      ...pieces.slice(1).map((x) => ({ find: x, replace: "" })),
    );
  }
  return { segments };
}

// The segments applied from the last to the first, as Word applies them:
// the text that comes out, or null when a find is not there exactly once.
export function applySegments(before, segments) {
  let text = norm(before);
  for (const s of [...segments].reverse()) {
    if (count(text, s.find) !== 1) return null;
    text = text.replace(s.find, () => s.replace);
  }
  return norm(text);
}
