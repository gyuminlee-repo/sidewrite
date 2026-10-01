// What an office_tracked_edit removes, adds, and how small it can be.
// Pure: no Office.js, no I/O. Checks L1-L3 against common failure modes of LLM-assisted revisions:
//   L1 lostFacts: numbers (with units), versions, text in parentheses and
//      capitalized name candidates that are in the deleted text and not in
//      the replacement.
//   L2 unsourcedTerms: facts and content words of the new text that appear
//      neither in the manuscript snapshot nor in the reviewer comment.
//   L3 narrowEdit: cut the common words at both ends of find/replace so the
//      tracked change covers only the words that change.

const WORD = String.raw`\p{L}\p{N}`;

// Units after a number. A unit must end at a non-letter, so "3 mice" is the
// number 3 and not "3 m".
export const UNIT = String.raw`(?:%|°\s?[CF]|[×x]\s?g|degrees?(?:\s(?:Celsius|C))?|kDa|Da|kb|bp|nt|rpm|fold|min|h|[pnµμmck]?(?:mol|g|L|l|M|m|s|Hz|V|W|J)|K)`;

const RE = {
  paren: /\(([^()]{1,120})\)/gu,
  version: new RegExp(
    String.raw`(?<![${WORD}_.])(?:v\d+(?:\.\d+)+|\d+(?:\.\d+){2,})(?![${WORD}]|\.\d)`,
    "gu",
  ),
  number: new RegExp(
    String.raw`(?<![${WORD}_.,]|[${WORD}]-)[-+±−]?\d+(?:[.,]\d+)*(?:\s?${UNIT})?(?![${WORD}]|[.,]\d)`,
    "gu",
  ),
  name: new RegExp(String.raw`(?<![${WORD}_-])\p{Lu}[${WORD}]*(?:-[${WORD}]+)*`, "gu"),
};

// Capitalized words that start sentences or name figure parts, not facts.
const NAME_STOP = new Set(
  (
    "A An The This That These Those It Its In On At For With By From To Of Into And But Or Nor " +
    "However Here There Their They Our As After Before During Although Because Since When While " +
    "Where Each All Both No Not Only Further Moreover Therefore Thus Then Such Some One Two Three " +
    "Figure Fig Table Supplementary Section Eq"
  ).split(" "),
);

function mask(text, start, end) {
  return text.slice(0, start) + " ".repeat(end - start) + text.slice(end);
}

function atSentenceStart(text, index, textStartsSentence) {
  const before = text.slice(0, index);
  if (!before.trim()) return textStartsSentence;
  return /[.!?]["')\]]?\s+["'(\[]?$/.test(before);
}

// Does `find` start a sentence in one of these paragraph texts? Name
// candidates at a sentence start are skipped unless they look like names.
export function startsSentence(find, texts = []) {
  const f = String(find ?? "")
    .replace(/\s+/g, " ")
    .trim();
  if (!f) return false;
  for (const t of texts) {
    const s = String(t ?? "").replace(/\s+/g, " ");
    const i = s.indexOf(f);
    if (i !== -1) return atSentenceStart(s, i, true);
  }
  return false;
}

// Name candidate at a sentence start counts only when it looks like a name
// anyway: a digit, a hyphen, or a second capital ("Boltz-2", "MaxQuant").
const strongName = (w) => /\d|-|.\p{Lu}/u.test(w);

// facts: [{ kind: "paren"|"version"|"number"|"name", text, index }] in text
// order. Parentheses are taken whole and masked, then versions, numbers and
// names, so "(v2.1)" is one fact, not three.
// opts.sentenceStart: the text starts a sentence (default false: `find` is
// usually a phrase inside a sentence).
export function extractFacts(text, { sentenceStart = false } = {}) {
  let t = String(text ?? "").replace(/\s+/g, " ");
  const out = [];
  for (const m of [...t.matchAll(RE.paren)]) {
    const inner = m[1].trim();
    if (inner) out.push({ kind: "paren", text: inner, index: m.index });
    t = mask(t, m.index, m.index + m[0].length);
  }
  for (const kind of ["version", "number"]) {
    for (const m of [...t.matchAll(RE[kind])]) {
      out.push({ kind, text: m[0].trim(), index: m.index });
      t = mask(t, m.index, m.index + m[0].length);
    }
  }
  for (const m of t.matchAll(RE.name)) {
    const w = m[0];
    if (w.length < 2 || NAME_STOP.has(w)) continue;
    if (atSentenceStart(t, m.index, sentenceStart) && !strongName(w)) continue;
    out.push({ kind: "name", text: w, index: m.index });
  }
  out.sort((a, b) => a.index - b.index);
  const seen = new Set();
  return out.filter((f) => {
    const k = canon(f.text);
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });
}

// Compare form: lower case, one space, no space between a number and its
// unit, and "degrees", "degrees Celsius", "° C" all written "°c" (a unit
// written another way is not a lost fact).
export function canon(s) {
  return String(s ?? "")
    .toLowerCase()
    .replace(/−/g, "-")
    .replace(/\bdegrees?(?:\s(?:celsius|c)\b)?|°\s?c\b/g, "°c")
    .replace(new RegExp(String.raw`(\d)\s+(?=${UNIT}(?![\p{L}]))`, "giu"), "$1")
    .replace(/\s+/g, " ")
    .trim();
}

const escapeRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

// Is this fact in `text` as a whole token? "3 mM" is not in "30 mM" or "13 mM",
// "3" is not in "3.5", "30 °C" is in "30°C".
export function factPresent(fact, text) {
  const key = canon(fact);
  if (!key) return true;
  const re = new RegExp(String.raw`(?<![${WORD}]|\d[.,])${escapeRe(key)}(?![${WORD}]|[.,]\d)`, "u");
  // Also the text as written: canon() folds a unit written out ("degrees
  // Celsius") into its sign, and the word itself would then seem gone.
  return (
    re.test(canon(text)) ||
    re.test(
      String(text ?? "")
        .toLowerCase()
        .replace(/\s+/g, " "),
    )
  );
}

// L1. Facts of the deleted text that the replacement does not keep.
export function lostFacts(deleted, inserted = "", opts = {}) {
  return extractFacts(deleted, opts)
    .filter((f) => !factPresent(f.text, inserted))
    .map((f) => f.text);
}

// ---- L2 ---------------------------------------------------------------------

// Function words and generic verbs are never reported as unsourced.
export const STOPWORDS = new Set(
  (
    "a an the this that these those it its is are was were be been being has have had do does did " +
    "and or but nor not no so as at by for from in into of on onto to with within without than then " +
    "also only both each all any some such very more most less least much many few other another same " +
    "which who whom whose what when where while whether how why there here their they them we our " +
    "can could may might must shall should will would about above after again against before below " +
    "between during over under up down out off per via among across along around because since until " +
    "although however therefore thus hence moreover further furthermore indeed instead rather " +
    "used using use shown show shows showed described describe found observed performed obtained " +
    "added respectively approximately significantly markedly slightly further previously here"
  ).split(" "),
);

const words = (s) =>
  (String(s ?? "").match(/\p{L}[\p{L}\p{N}']*/gu) ?? []).map((w) =>
    w.toLowerCase().replace(/'.*$/, ""),
  );

// "filters" and "filter", "assays" and "assay", "studies" and "study" count
// as one word. Nothing else is stemmed.
export function wordForms(w) {
  const f = new Set([w, `${w}s`, `${w}es`]);
  if (w.endsWith("ies")) f.add(`${w.slice(0, -3)}y`);
  if (w.endsWith("es")) f.add(w.slice(0, -2));
  if (w.endsWith("s")) f.add(w.slice(0, -1));
  if (w.endsWith("y")) f.add(`${w.slice(0, -1)}ies`);
  return f;
}

// L2. sources: texts the new words may come from (manuscript paragraphs,
// the reviewer comment). Returns the new facts and content words (4+
// letters, not stopwords) found in none of them, in text order.
export function unsourcedTerms(inserted, sources = [], opts = {}) {
  const all = sources.map((s) => String(s ?? "")).join("\n");
  const vocab = new Set(words(all));
  const out = [];
  const seen = new Set();
  const push = (t) => {
    const k = canon(t);
    if (!seen.has(k)) {
      seen.add(k);
      out.push(t);
    }
  };
  const facts = extractFacts(inserted, opts);
  const factWords = new Set();
  for (const f of facts) {
    for (const w of words(f.text)) factWords.add(w);
    if (!factPresent(f.text, all)) push(f.text);
  }
  for (const w of words(inserted)) {
    if (w.length < 4 || STOPWORDS.has(w) || factWords.has(w)) continue;
    if ([...wordForms(w)].some((x) => vocab.has(x))) continue;
    push(w);
  }
  return out;
}

// ---- L3 ---------------------------------------------------------------------

const tokens = (s) => String(s ?? "").match(/\s+|[\p{L}\p{N}_]+|[^\s\p{L}\p{N}_]/gu) ?? [];

function occurrences(hay, needle) {
  let n = 0;
  for (let i = hay.indexOf(needle); i !== -1; i = hay.indexOf(needle, i + 1)) n += 1;
  return n;
}

// L3. The smallest word-aligned part of `find` that changes, with its
// replacement. Returns null when the full find/replace must be used:
//   - nothing to cut, or only white space differs;
//   - the cut find occurs more than once inside the original find (the task
//     pane searches it inside the located hit, so it must be unique there).
// A pure insertion ("30 degrees" -> "30 degrees Celsius") keeps the word
// before it: find "degrees", replace "degrees Celsius". The rationale
// comment then sits on the deleted "degrees" (design: deleted range first).
export function narrowEdit(find, replace) {
  const a = tokens(find);
  const b = tokens(replace);
  let p = 0;
  while (p < a.length && p < b.length && a[p] === b[p]) p += 1;
  let q = 0;
  while (q < a.length - p && q < b.length - p && a[a.length - 1 - q] === b[b.length - 1 - q])
    q += 1;
  if (p === a.length - q) {
    // Pure insertion: take in the word before (or after) the insertion point.
    const back = (i) => (i > 0 && /^\s+$/.test(a[i - 1]) ? i - 2 : i - 1);
    if (p > 0) {
      const np = Math.max(back(p), 0);
      p = /^\s+$/.test(a[np]) ? np + 1 : np;
    } else if (q > 0) {
      q -= 1;
    } else {
      return null;
    }
  }
  const nf = a.slice(p, a.length - q).join("");
  const nr = b.slice(p, b.length - q).join("");
  if (!nf.trim() || nf === String(find)) return null;
  if (!nr.trim() && nr !== "") return null;
  if (occurrences(String(find), nf) !== 1) return null;
  return { find: nf, replace: nr };
}
