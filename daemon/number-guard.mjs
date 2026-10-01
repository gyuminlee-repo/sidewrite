// Numbers a tracked edit would put into the manuscript.
//
// A number in a paper is a measurement, a count or a setting. Whether it is
// right cannot be told from the paper: when two places disagree, or a value
// can be computed another way, only the author's records say which holds.
// A review turn therefore may not write a number that is not already given
// to it; it asks the author (office_ask_author) instead. This module finds
// the numbers; review-tools.mjs applyTrackedEdit refuses the edit.
//
// What counts as a number: a run of digits with decimal or thousands marks
// that does not follow a letter or digit ("CO2", "K12" and "H2O" hold none).
// "3,360" and "3360" are the same number; "0,5" is "0.5". No other form is
// folded: "30" and "30.0" differ, a number written as a word is not seen.
// Not counted: citation brackets ("[3]", "[4, 5]", "[7-9]") and the number of
// a figure, table, section, equation or reference ("Fig. 2", "Tables 2 and
// 3"): they point into the paper and state nothing measured.
// Assumed: English prose. A year in an author-year citation counts as a number.

const CITATION = /\[\s*\d+(?:\s*[,;–-]\s*\d+)*\s*\]/g;
const LABEL = String.raw`S?\d+[A-Za-z]?`;
const POINTER = new RegExp(
  String.raw`\b(?:Figs?|Figures?|Tables?|Sections?|Eqs?|Equations?|Refs?|References?|Schemes?)\.?\s*${LABEL}(?:\s*(?:,|and|to|–|-)\s*${LABEL}(?![\p{L}\p{N}]))*`,
  "giu",
);
const NUMBER = /(?<![\p{L}\p{N}])\d+(?:[.,]\d+)*/gu;

function norm(raw) {
  // Groups of three after a comma: thousands. Any other comma: a decimal mark.
  if (/^\d{1,3}(?:,\d{3})+(?:\.\d+)?$/.test(raw)) return raw.replace(/,/g, "");
  return raw.replace(/,/g, ".");
}

function tokens(text) {
  const s = String(text ?? "")
    .replace(CITATION, " ")
    .replace(POINTER, " ");
  return (s.match(NUMBER) ?? []).map((raw) => ({ raw, norm: norm(raw) }));
}

// The numbers a text states, normalized, as a Set in order of appearance.
export function numbersIn(text) {
  return new Set(tokens(text).map((t) => t.norm));
}

// The numbers a change adds, as written in the new text: those of `text`
// that are neither in `find` (the text it replaces, or its anchor) nor in
// one of `sources`. A deletion adds none.
export function newNumbers({ kind, find, text, sources = [] }) {
  if (kind === "delete") return [];
  const known = new Set([find, ...sources].flatMap((s) => [...numbersIn(s)]));
  const out = [];
  const seen = new Set();
  for (const t of tokens(text)) {
    if (known.has(t.norm) || seen.has(t.norm)) continue;
    seen.add(t.norm);
    out.push(t.raw);
  }
  return out;
}
