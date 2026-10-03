// The words a reviewer names, and every place the manuscript has them.
//
// A reviewer who objects to a word ("the term 'convergence' is not
// justified") objects to it wherever it stands. A judged revision changed
// such words in the body and left them in figure legends, the methods and
// section headings. So the turn of a comment is given, for each term the
// comment names, every paragraph that holds it, and after the turn the card
// of the comment shows where it still stands.
//
// One implementation for both sides: the daemon imports this file for the
// prompt (daemon/review-prompt.mjs), the task pane for the card. No imports,
// nothing of the browser or of node.
//
// Which terms a comment names:
//   - what it puts in quotation marks ("...", '...', and the curly forms),
//     2 to 60 characters and at most 6 words; "A/B" in quotes is A and B;
//   - the word after "the term", "the word", "the phrase", "the expression"
//     (and a second one after "and" / "or");
//   - comment.named_terms: what the requirement labelling (a model) gave,
//     kept only when the comment text holds it.
// Assumed: English. At most MAX_TERMS for a comment, in that order.
//
// Where a term occurs: as a whole word, any case. A single word of six
// letters or more also matches its other forms by its first letters
// ("upregulated" finds "upregulation"; "convergence" finds "convergent"),
// which can take in a different word that starts the same; the turn decides
// what to change, this only shows where to look.
// A phrase of several words also stands in a paragraph that has one of its
// content words (five letters or more) in another form of the same family:
// "network densification" is found where only "densified network" or
// "densification" stands. Families are compared by a stem made by taking
// off a few suffixes of English inflection and derivation (familyStem), as
// whole words: a stem inside a longer word ("distribution" in
// "redistribution") is no hit. A content word whose family stands in more
// than a fifth of the paragraphs (never fewer than 3) is vocabulary of the
// paper and is not searched this way (as in residue.js): a general word
// would list most of the paper. Such a place carries the form found.
//
// What a paragraph is (a guess from its text, for the label next to its id):
//   caption  starts like "Figure 2", "Fig. S1", "Table 3", "Scheme 1"
//   heading  at most 12 words and no sentence end
//   methods  under a heading that says methods, materials or experimental,
//            up to the next heading of another main part
//   body     everything else

export const MAX_TERMS = 8;
const OPEN_CLOSE = [
  ["“", "”"],
  ["‘", "’"],
  ['"', '"'],
  ["'", "'"],
];
const WORD = String.raw`[\p{L}\p{N}][\p{L}\p{N}'’/-]*`;

// After "the term": words that name no term ("the term used in line 3").
const NOT_A_TERM = new Set(
  "a an and are as can for has have here in is it itself may of or should that the this to used was were which".split(
    " ",
  ),
);

const clean = (t) =>
  String(t ?? "")
    .replace(/\s+/g, " ")
    .replace(/^[\s.,;:!?()]+|[\s.,;:!?()]+$/g, "")
    .trim();

function usable(t) {
  return t.length >= 2 && t.length <= 60 && /\p{L}/u.test(t) && t.split(" ").length <= 6;
}

// The terms the text of a comment names itself, in the order they stand.
export function quotedTerms(text) {
  const s = String(text ?? "");
  const found = [];
  for (const [open, close] of OPEN_CLOSE) {
    // A straight single quote opens only after a space or at the start and
    // closes only before a space, a mark or the end: an apostrophe does neither.
    const before = open === "'" ? String.raw`(?<![\p{L}\p{N}])` : "";
    const after = close === "'" ? String.raw`(?![\p{L}\p{N}])` : "";
    const re = new RegExp(`${before}${open}([^${open}${close}\\n]{2,80})${close}${after}`, "gu");
    for (const m of s.matchAll(re)) found.push({ at: m.index, text: m[1] });
  }
  const named = new RegExp(
    String.raw`\bthe (?:terms?|words?|phrases?|expressions?) (${WORD})(?: (?:and|or) (${WORD}))?`,
    "giu",
  );
  for (const m of s.matchAll(named)) {
    if (NOT_A_TERM.has(m[1].toLowerCase())) continue;
    found.push({ at: m.index, text: m[1] });
    if (m[2] && !NOT_A_TERM.has(m[2].toLowerCase())) found.push({ at: m.index + 1, text: m[2] });
  }
  const out = [];
  for (const f of found.sort((a, b) => a.at - b.at)) {
    for (const part of f.text.split("/")) {
      const t = clean(part);
      if (usable(t) && !out.some((x) => x.toLowerCase() === t.toLowerCase())) out.push(t);
    }
  }
  return out;
}

// Every term of a comment: its own, then the stored ones its text holds.
export function namedTerms(comment) {
  const text = String(comment?.text ?? "");
  const lower = text.toLowerCase();
  const out = quotedTerms(text);
  for (const raw of comment?.named_terms ?? []) {
    const t = clean(raw);
    if (!usable(t) || !lower.includes(t.toLowerCase())) continue;
    if (!out.some((x) => x.toLowerCase() === t.toLowerCase())) out.push(t);
  }
  return out.slice(0, MAX_TERMS);
}

const CAPTION = /^(?:supplementary\s+)?(?:fig(?:ure|\.)?|table|scheme)\s*\.?\s*S?\d/i;
const METHODS = /\b(methods?|materials|experimental)\b/i;
const OTHER_PART =
  /^(?:\d+(?:\.\d+)*\.?\s*)?(?:abstract|introduction|results?|discussion|conclusions?|references|acknowledg|supplementa|figure legends?|tables?$)/i;

// "caption" | "heading" | "methods" | "body" for each paragraph.
export function paragraphKinds(paragraphs) {
  let inMethods = false;
  return (paragraphs ?? []).map((raw) => {
    const t = String(raw ?? "").trim();
    if (!t) return "body";
    if (CAPTION.test(t)) return "caption";
    const heading = t.split(/\s+/).length <= 12 && !/[.!?:;]$/.test(t);
    if (heading) {
      if (METHODS.test(t)) inMethods = true;
      else if (OTHER_PART.test(t)) inMethods = false;
      return "heading";
    }
    return inMethods ? "methods" : "body";
  });
}

const escapeRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

function termRe(term) {
  const single = /^\p{L}{6,}$/u.test(term);
  const core = single
    ? `${escapeRe(term.slice(0, Math.max(5, term.length - 3)))}\\p{L}*`
    : escapeRe(term).replace(/ /g, "\\s+");
  return new RegExp(`(?<![\\p{L}\\p{N}])${core}(?![\\p{L}\\p{N}])`, "iu");
}

// The family stem of a word: lower case, a plural taken off, then one
// derivational or inflectional ending, which must leave four letters or
// more; a final "e" goes. "densification", "densified", "densifies" ->
// "densify"; "density" and "dense" stay apart from them; "station" stays
// whole (what is left of "-ation" is too short), so "state" is not of it.
// Assumed: English.
const FAMILY_ENDINGS = [
  [/(?:ification|ified|ifying|ify)$/, "ify"],
  [/(?:[iy][sz]ation|[iy][sz]ed|[iy][sz]ing|[iy][sz]e)$/, "ize"],
  [/(?:ation|ated|ating|ate)$/, ""],
  [/(?:ution|uted|uting|ute)$/, "ut"],
  [/(?<=[st])ion$/, ""],
  [/(?:ment|ing|ed)$/, ""],
];
const trimE = (w) => (w.length > 4 && w.endsWith("e") ? w.slice(0, -1) : w);
export function familyStem(word) {
  let w = String(word ?? "").toLowerCase();
  if (w.length > 5 && w.endsWith("ies")) w = `${w.slice(0, -3)}y`;
  else if (w.length > 4 && w.endsWith("s") && !/(?:ss|us|is)$/.test(w)) w = w.slice(0, -1);
  for (const [re, add] of FAMILY_ENDINGS) {
    const m = re.exec(w);
    if (!m) continue;
    const base = w.slice(0, m.index);
    return base.length < 4 ? trimE(w) : `${add ? base : trimE(base)}${add}`;
  }
  return trimE(w);
}

const FAMILY_SKIP = new Set(
  "about above after among based being between could other since their there these those under using where which while would within without".split(
    " ",
  ),
);
const wordsOf = (t) => String(t ?? "").match(/\p{L}+/gu) ?? [];

// The rare content words of a phrase, by stem; null for a single word.
// stems: per paragraph, Map stem -> the first form it has.
function familyStems(term, stems) {
  const parts = term.split(/\s+/);
  if (parts.length < 2) return null;
  const filled = stems.filter((m) => m.size).length;
  const most = Math.max(3, Math.ceil(filled / 5));
  const out = new Set();
  for (const w of parts) {
    if (!/^\p{L}{5,}$/u.test(w) || FAMILY_SKIP.has(w.toLowerCase())) continue;
    const s = familyStem(w);
    if (stems.filter((m) => m.has(s)).length <= most) out.add(s);
  }
  return out;
}

// [{ term, places: [{ index, id, kind, form? }] }] for the terms of a
// comment, in the paragraphs given (ids: their task pane ids, or null).
// form: the word found when only a word of the family of the phrase stands.
export function termPlaces(comment, paragraphs, ids = null) {
  const texts = (paragraphs ?? []).map((t) => String(t ?? ""));
  const kinds = paragraphKinds(texts);
  let stems = null;
  const stemsOf = () =>
    (stems ??= texts.map((t) => {
      const m = new Map();
      for (const w of wordsOf(t)) {
        const s = familyStem(w);
        if (!m.has(s)) m.set(s, w);
      }
      return m;
    }));
  return namedTerms(comment).map((term) => {
    const re = termRe(term);
    const family = /\s/.test(term) ? familyStems(term, stemsOf()) : null;
    return {
      term,
      places: texts.flatMap((t, index) => {
        const at = { index, id: ids?.[index] ?? `p${index}`, kind: kinds[index] };
        if (re.test(t)) return [at];
        const hit = family && [...family].find((s) => stems[index].has(s));
        return hit ? [{ ...at, form: stems[index].get(hit) }] : [];
      }),
    };
  });
}

const placeText = (p) => `${p.id} (${p.kind}${p.form ? `, as "${p.form}"` : ""})`;

// The lines of the prompt block. Most places of a term shown: `most`.
export function termBlock(comment, paragraphs, ids = null, { most = 30 } = {}) {
  const rows = termPlaces(comment, paragraphs, ids);
  if (!rows.length) return [];
  return [
    "Terms this comment names, and every place they occur (paragraph id, and what the paragraph is):",
    ...rows.map((r) =>
      r.places.length
        ? `- "${r.term}": ${r.places.slice(0, most).map(placeText).join(", ")}${r.places.length > most ? `, and ${r.places.length - most} more` : ""}`
        : `- "${r.term}": not in the manuscript as written`,
    ),
    "Go to every place: change it there, or say in your end report why it stays. Headings, figure and table legends and the methods count like the body.",
  ];
}

// The lines of the card after a turn: the terms that still stand somewhere.
// snapshot: revision.snapshot ({ paragraphs, ids }).
export function namedTermLines(comment, snapshot) {
  return termPlaces(comment, snapshot?.paragraphs, snapshot?.ids)
    .filter((r) => r.places.length)
    .map((r) => {
      const n = r.places.length;
      const shown = r.places.slice(0, 12).map(placeText).join(", ");
      return `The reviewer names "${r.term}": still present in ${n} place${n === 1 ? "" : "s"}: ${shown}${n > 12 ? `, and ${n - 12} more` : ""}.`;
    });
}
