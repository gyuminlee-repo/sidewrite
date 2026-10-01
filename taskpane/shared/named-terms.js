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

// [{ term, places: [{ index, id, kind }] }] for the terms of a comment, in
// the paragraphs given (ids: their task pane ids, or null).
export function termPlaces(comment, paragraphs, ids = null) {
  const texts = (paragraphs ?? []).map((t) => String(t ?? ""));
  const kinds = paragraphKinds(texts);
  return namedTerms(comment).map((term) => {
    const re = termRe(term);
    return {
      term,
      places: texts.flatMap((t, index) =>
        re.test(t) ? [{ index, id: ids?.[index] ?? `p${index}`, kind: kinds[index] }] : [],
      ),
    };
  });
}

const placeText = (p) => `${p.id} (${p.kind})`;

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
