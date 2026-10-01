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
//   quote_changed  a quoted sentence is no longer in the manuscript
//   change_state   the cited changes are not the accepted changes any more
//   lines_old      the manuscript text changed since the PDF the line
//                  numbers were read from (render: revision.response_render)
export function staleReasons(comment, snapshot, render) {
  const r = comment?.response;
  if (!r) return [];
  const out = new Set();
  const cits = r.citations ?? [];
  const texts = (snapshot?.paragraphs ?? []).map(viewText);
  for (const c of cits)
    if (c.kind === "revised" && c.quote && !texts.some((t) => t.includes(c.quote)))
      out.add("quote_changed");
  const cited = new Set(cits.flatMap((c) => c.change_ids ?? []));
  const now = new Set(
    (comment.changes ?? []).filter((ch) => CITABLE.has(ch.state)).map((ch) => ch.id),
  );
  if (cited.size !== now.size || [...cited].some((id) => !now.has(id))) out.add("change_state");
  if (cits.length && render && render.snapshot_hash !== snapshotHash(snapshot))
    out.add("lines_old");
  return ["quote_changed", "change_state", "lines_old"].filter((x) => out.has(x));
}
