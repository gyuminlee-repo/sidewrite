// The ledger's "count" check: word limits of a section, the number of
// keywords, figures and tables, and the order of sections. Pure: no model, no
// Word, no file.
//
// Paragraph shape (task pane snapshot): { index, text, reviewed_text? }.
// Headings: [{ index, level, text }], the paragraphs that carry a heading
// style. A manuscript without heading styles is cut by a regex instead
// (profile.body_section_head). With neither there are no sections, and a
// check that needs one is an error, never a pass.
//
// Words are whitespace-separated tokens of the text with pending tracked
// changes accepted. Word's own word count can differ by a few words
// (hyphens, numbers with units, field results).
//
// Check: { type: "count", target, section?, min?, max?, order? } with target
// one of COUNT_TARGETS. The result has the shape checkRegexItem returns:
// { id, status: "ok" | "violations" | "error", violations: [{ paragraph_index,
// start, match, count, min, max }], count?, error? }.
//
// Not wired yet. The ledger check needs (ledger-check.mjs, checkLedger):
//   item.check.type === "count" ? checkCountItem(item, paragraphs, deps) : ...
// and the incremental check must read the whole document for a count item
// (needsFull). revision-store.mjs needs a CountCheck in the check union and
// count, min, max on Violation, or the store strips them.

import { normalizeParaText } from "./revision-classify.mjs";

export const COUNT_TARGETS = ["section_words", "keywords", "figures", "tables", "section_order"];

// The sections a rule can name, and the words that name each one, in a
// guideline sentence as well as in a heading of the manuscript. This is the
// one table: the guideline extractor (guideline-extract.mjs) reads it to find
// a section in a sentence, and the heading patterns below are made from it.
export const SECTION_TERMS = {
  abstract: ["abstract", "summary"],
  introduction: ["introduction", "background"],
  methods: ["methods", "materials and methods", "material and methods", "experimental"],
  results: ["results"],
  discussion: ["discussion"],
  conclusion: ["conclusion", "conclusions"],
};

// The heading pattern of each section: one of its words at the start of the
// title, as a whole word, after an optional number ("2. Methods").
const words = (terms) =>
  [...terms]
    .sort((a, b) => b.length - a.length)
    .map((t) => t.replace(/ /g, "\\s+"))
    .join("|");
export const SECTION_PATTERNS = Object.fromEntries(
  Object.entries(SECTION_TERMS).map(([name, terms]) => [
    name,
    `^(?:[\\d.]+\\s+)?(?:${words(terms)})\\b`,
  ]),
);

const view = (p) => normalizeParaText(p.reviewed_text ?? p.text);

export const wordCount = (text) => {
  const t = String(text ?? "").trim();
  return t ? t.split(/\s+/).length : 0;
};

const KEYWORDS = /^key\s?words\s*[:.]/i;

// Sections in document order: { title, index, level, body }. body is every
// paragraph up to the next heading of the same or a higher level, without the
// subsection headings themselves.
export function sectionsOf(paragraphs, { headings = [], headRe = null } = {}) {
  let level = new Map(headings.map((h) => [h.index, h.level ?? 1]));
  if (!level.size && headRe) {
    const re = new RegExp(headRe, "i");
    level = new Map(paragraphs.filter((p) => re.test(view(p))).map((p) => [p.index, 1]));
  }
  const out = [];
  let open = [];
  for (const p of paragraphs) {
    const lv = level.get(p.index);
    if (lv === undefined) {
      for (const s of open) s.body.push(p);
      continue;
    }
    open = open.filter((s) => s.level < lv);
    const section = { title: view(p), index: p.index, level: lv, body: [] };
    out.push(section);
    open.push(section);
  }
  return out;
}

const CAPTION = {
  figures: /^Fig(?:ure|\.)?\s*(\d+)\s*[.:|]/i,
  tables: /^Table\s+(\d+)\s*[.:|]/i,
};
const LEGEND_KEY = { figures: "main", tables: "table" };

const isCount = (n) => Number.isInteger(n) && n >= 0;

export function checkCountItem(item, paragraphs, { headings = [], profile = {} } = {}) {
  const c = item.check ?? {};
  const min = c.min ?? null;
  const max = c.max ?? null;
  const error = (message) => ({ id: item.id, status: "error", error: message, violations: [] });
  const judge = (count, paragraph_index) =>
    (max !== null && count > max) || (min !== null && count < min)
      ? {
          id: item.id,
          status: "violations",
          violations: [{ paragraph_index, start: null, match: null, count, min, max }],
        }
      : { id: item.id, status: "ok", violations: [], count };

  if (!COUNT_TARGETS.includes(c.target)) return error(`unknown count target ${c.target}`);
  let sections;
  try {
    sections = sectionsOf(paragraphs, { headings, headRe: profile.body_section_head });
  } catch (e) {
    return error(e.message);
  }

  if (c.target === "section_order") {
    if (!Array.isArray(c.order) || !c.order.length) return error("section_order needs an order");
    // A name of SECTION_TERMS is found by its heading pattern, like the
    // section of a word limit; any other name by its text in the title.
    const at = c.order.map((name) => {
      const key = String(name).toLowerCase();
      const re = Object.hasOwn(SECTION_PATTERNS, key)
        ? new RegExp(SECTION_PATTERNS[key], "i")
        : null;
      return sections.findIndex((s) =>
        re ? re.test(s.title) : s.title.toLowerCase().includes(key),
      );
    });
    const missing = c.order.filter((_, i) => at[i] === -1);
    if (missing.length) return error(`no section: ${missing.join(", ")}`);
    const bad = at.findIndex((v, i) => i > 0 && v < at[i - 1]);
    if (bad === -1) return { id: item.id, status: "ok", violations: [] };
    return {
      id: item.id,
      status: "violations",
      violations: [
        {
          paragraph_index: sections[at[bad]].index,
          start: null,
          match: null,
          count: null,
          min: null,
          max: null,
        },
      ],
    };
  }

  if ((min === null && max === null) || ![min, max].every((n) => n === null || isCount(n))) {
    return error("a count check needs a whole-number min or max");
  }

  if (c.target === "section_words") {
    if (!c.section) return error("section_words needs a section");
    let re;
    try {
      re = new RegExp(c.section, "i");
    } catch (e) {
      return error(e.message);
    }
    const section = sections.find((s) => re.test(s.title));
    if (!section) return error(`no section matches ${c.section}`);
    const count = section.body
      .filter((p) => !KEYWORDS.test(view(p)))
      .reduce((n, p) => n + wordCount(view(p)), 0);
    return judge(count, section.index);
  }

  if (c.target === "keywords") {
    const line = paragraphs.find((p) => KEYWORDS.test(view(p)));
    if (!line) return error("no Keywords line");
    const count = view(line)
      .replace(KEYWORDS, "")
      .split(/[;,]/)
      .filter((k) => k.trim()).length;
    return judge(count, line.index);
  }

  // figures, tables: distinct caption numbers.
  let caption = CAPTION[c.target];
  const own = profile.legend_heads?.[LEGEND_KEY[c.target]];
  if (own) {
    try {
      caption = new RegExp(own, "i");
    } catch (e) {
      return error(e.message);
    }
  }
  const numbers = new Set();
  for (const p of paragraphs) {
    const m = caption.exec(view(p));
    if (m) numbers.add(m[1] ?? m[0]);
  }
  return judge(numbers.size, null);
}
