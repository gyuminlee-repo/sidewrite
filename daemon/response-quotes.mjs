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

const hitsOf = (hay, needle) => {
  const out = [];
  if (!needle) return out;
  for (let p = hay.indexOf(needle); p !== -1; p = hay.indexOf(needle, p + 1)) out.push(p);
  return out;
};

// Offset in `text` of the place the change touched, or null when no form of
// it is there exactly once. Forms, most specific first: the new text between
// its recorded context words, with one side of the context, alone, then the
// context alone (the author rewrote the new text).
export function changeOffset(text, change) {
  const m = squeezeMap(String(text ?? ""));
  if (!m.at.length) return null;
  const needle = squeeze(change.kind === "delete" ? "" : change.text);
  const left = squeeze(change.context?.left);
  const right = squeeze(change.context?.right);
  const forms = [
    [left + needle + right, left.length],
    [left + needle, left.length],
    [needle + right, 0],
    [needle, 0],
    [left, left.length],
    [right, 0],
  ];
  const tried = new Set();
  for (const [form, skip] of forms) {
    if (!form || tried.has(`${form}\n${skip}`)) continue;
    tried.add(`${form}\n${skip}`);
    const hits = hitsOf(m.text, form);
    if (hits.length === 1) return m.at[Math.min(hits[0] + skip, m.at.length - 1)];
  }
  return null;
}

// The sentence of `text` that holds `offset`, or null.
export function sentenceAt(text, offset) {
  let cursor = 0;
  for (const s of splitSentences(text)) {
    const start = text.indexOf(s, cursor);
    if (start === -1) return null;
    if (start > offset) return null;
    cursor = start + s.length;
    if (offset < cursor) return s;
  }
  return null;
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
