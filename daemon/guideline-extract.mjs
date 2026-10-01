// Draft ledger rules from a journal's author guidelines.
//
// The model never writes a rule. The code cuts the guideline into sentences
// (the splitter is passed in), the model labels each sentence number with a
// type and, for a countable limit, a few parameters from closed lists. A rule
// candidate is the guideline sentence as written, and it cites that sentence.
// A count rule is kept only when its own sentence names the section, has
// the number, counts the same thing with it and says in which direction the
// number limits it, and, when the sentence names several sections, when the
// number stands in the clause of that section alone; otherwise the sentence stays as a model-checked rule and
// a warning says what was missing. A reply that is malformed, partial or carries any text
// yields no candidate at all.
//
// Pure apart from the injected functions: oneshot({ system, prompt }) ->
// { ok, json?, error? } (daemon/oneshot.mjs), splitSentences(text) ->
// string[], fetchFn(url, { signal }) -> Response.
//
// Not wired yet. The caller passes splitSentences from review-split.mjs and
// the one-shot function of the chosen provider, and shows candidates to the
// author before any of them enters the ledger.

import { COUNT_TARGETS } from "./count-check.mjs";
import { FETCH_TIMEOUT_MS, deadline } from "./fetch-timeout.mjs";

export const GUIDELINE_SYSTEM =
  "You sort the sentences of a journal's author guidelines. Label every sentence exactly once. Do not return any text of the guidelines. Reply with JSON only.";

export const GUIDELINE_TYPES = ["count", "style", "format", "other"];

// Format rules of the journal profile that a guideline sentence can turn on
// without a list of names: setting -> the profile keys it sets and what the
// sentence must say.
export const FORMAT_SETTINGS = {
  "italics.statistics": {
    profile: { italics: { statistics: true } },
    says: "statistics symbols are set in italics",
    needs: ["statistics", "italic"],
  },
};

// The words the code looks for in such a sentence (whole words, any case). A
// negation anywhere in the sentence keeps the setting from being proposed.
export const FORMAT_WORDS = {
  italic: ["italic", "italics", "italicize", "italicized", "italicise", "italicised"],
  statistics: [
    "statistic",
    "statistics",
    "statistical",
    "p value",
    "p values",
    "p-value",
    "p-values",
  ],
  negation: ["not", "never", "no", "nor", "without", "roman", "upright", "non-italic"],
};

export const GUIDELINE_WARNING =
  "model could not label the guideline sentences; add the rules by hand";

// Section names the model may pick, and the heading pattern each one becomes
// (a heading may carry a number: "1. Introduction").
const HEAD = "^(?:[\\d.]+\\s+)?";
export const SECTION_PATTERNS = {
  abstract: `${HEAD}(?:abstract|summary)\\b`,
  introduction: `${HEAD}(?:introduction|background)\\b`,
  methods: `${HEAD}(?:materials?\\s+and\\s+methods|methods|experimental)\\b`,
  results: `${HEAD}results\\b`,
  discussion: `${HEAD}discussion\\b`,
  conclusion: `${HEAD}conclusions?\\b`,
};
const SECTIONS = Object.keys(SECTION_PATTERNS);
const BOUNDED = COUNT_TARGETS.filter((t) => t !== "section_order");

export function guidelinePrompt(sentences) {
  return [
    "Each line below is one sentence of a journal's author guidelines, numbered [s1], [s2], ...",
    'Return {"sentences":[{"s":<number>,"type":"count"|"style"|"other","rule":{...}}]} with one entry for every sentence.',
    '"count": the sentence sets a countable limit on the manuscript. Give "rule" with',
    `  "target": one of ${BOUNDED.map((t) => `"${t}"`).join(", ")} (with "min" and/or "max", whole numbers that stand in the sentence as digits), or "section_order" (with "order", a list of section names);`,
    `  "section" (only for "section_words", and required there): one of ${SECTIONS.map((t) => `"${t}"`).join(", ")}.`,
    `"format": the sentence sets one of these format rules. Give "rule" with "setting": one of ${Object.entries(
      FORMAT_SETTINGS,
    )
      .map(([k, v]) => `"${k}" (${v.says})`)
      .join(", ")}.`,
    '"style": the sentence is another rule an author must follow in the manuscript text. No "rule".',
    '"other": anything else (scope of the journal, submission steps, fees, contacts). No "rule".',
    "Do not add any other key and do not repeat or rewrite the sentences.",
    "Sentences:",
    ...sentences.map((s, i) => `[s${i + 1}] ${s}`),
  ].join("\n");
}

const isPlain = (x) => x !== null && typeof x === "object" && !Array.isArray(x);
const onlyKeys = (x, keys) => Object.keys(x).every((k) => keys.includes(k));
const isCount = (n) => Number.isInteger(n) && n >= 0;

// The label of one sentence as [type, rule | null], or null when it is not
// one of the allowed shapes.
function sentenceLabel(x) {
  if (!isPlain(x) || !onlyKeys(x, ["s", "type", "rule"])) return null;
  if (!GUIDELINE_TYPES.includes(x.type)) return null;
  if (x.type === "format") {
    const ok = isPlain(x.rule) && onlyKeys(x.rule, ["setting"]);
    return ok && Object.hasOwn(FORMAT_SETTINGS, x.rule.setting)
      ? ["format", { setting: x.rule.setting }]
      : null;
  }
  if (x.type !== "count") return x.rule === undefined ? [x.type, null] : null;
  const r = x.rule;
  if (!isPlain(r) || !COUNT_TARGETS.includes(r.target)) return null;
  if (r.target === "section_order") {
    if (!onlyKeys(r, ["target", "order"])) return null;
    if (!Array.isArray(r.order) || r.order.length < 2) return null;
    if (!r.order.every((name) => SECTIONS.includes(name))) return null;
    return ["count", { target: r.target, order: [...r.order] }];
  }
  const words = r.target === "section_words";
  if (!onlyKeys(r, words ? ["target", "section", "min", "max"] : ["target", "min", "max"])) {
    return null;
  }
  if (words && !SECTIONS.includes(r.section)) return null;
  if (r.min === undefined && r.max === undefined) return null;
  for (const n of [r.min, r.max]) if (n !== undefined && !isCount(n)) return null;
  if (r.min !== undefined && r.max !== undefined && r.min > r.max) return null;
  return [
    "count",
    {
      target: r.target,
      ...(words && { section: r.section }),
      ...(r.min !== undefined && { min: r.min }),
      ...(r.max !== undefined && { max: r.max }),
    },
  ];
}

// ---- what a count rule must find in its own sentence
//
// The number is not enough. The sentence must name the section (as a whole
// word), the number must count the thing the rule counts ("250 words", not
// "250 characters"; "8 figures", not "6 panels"), and the sentence must say
// in which direction the number limits it. Whatever is not found keeps the
// sentence as a model-checked rule, with a warning.

// The words that name each section in running text.
export const SECTION_TERMS = {
  abstract: ["abstract", "summary"],
  introduction: ["introduction", "background"],
  methods: ["methods", "materials and methods", "experimental"],
  results: ["results"],
  discussion: ["discussion"],
  conclusion: ["conclusion", "conclusions"],
};

// When is a number confirmed as the limit of a section? The unit stands next
// to the number and the direction right before or after it (below), so those
// two are tied to the number. The number is tied to the section when
//   - the sentence names no other section, or
//   - the clause that holds the number names the rule's section and no other.
// A clause is the stretch between two of these breaks. A break does not count
// inside a number ("1,000"), inside a section name ("materials and methods"),
// between the two numbers of a range ("150 and 250"), or where it only joins
// section names ("the abstract and the introduction", "Introduction, Methods,
// and Results"): those sections share one clause, the sentence does not say
// whether the limit holds for each or for all together, and nothing is
// confirmed. Anything not confirmed stays a model-checked rule.
export const CLAUSE_BREAKS = {
  marks: [";", ","],
  words: ["and", "but", "while", "whereas", "or"],
};

// When is a section order confirmed? The sentence names every section of the
// order, the order given is the order in which the sentence first names them,
// and the sentence has none of the words below. After such a word the order of
// naming may not be the order asked for ("Results should follow Methods",
// "Before the Results, describe the Methods", "Methods last"), or the sentence
// forbids or offers a choice instead of asking. Telling these apart needs the
// grammar of the sentence, so the sentence stays a model-checked rule. The
// phrases in ORDER_KEEPS hold such a word and keep the order of naming; they
// are taken out before the words are looked for.
export const ORDER_DOUBT = [
  "after",
  "before",
  "behind",
  "follow",
  "follows",
  "followed",
  "following",
  "precede",
  "precedes",
  "preceded",
  "preceding",
  "first",
  "last",
  "beginning",
  "end",
  "not",
  "never",
  "no",
  "nor",
  "or",
  "either",
  "except",
  "unless",
  "without",
];
export const ORDER_KEEPS = ["followed by", "the following", "as follows"];

// The words that name what each target counts. A number other than 1 needs
// the plural ("6 figure panels" does not count figures).
export const UNIT_WORDS = {
  section_words: { plural: ["words"], singular: ["word"] },
  keywords: { plural: ["keywords", "key words"], singular: ["keyword", "key word"] },
  figures: { plural: ["figures"], singular: ["figure"] },
  tables: { plural: ["tables"], singular: ["table"] },
};

// The words that say a number is an upper (max) or a lower (min) limit, the
// number included. `before` stands right before the number ("at most 8",
// "a maximum of 8"), `after` right after what is counted ("8 figures at
// most."). Strict forms ("fewer than", "under", "more than") are left out on
// purpose: they name a limit one off from the number.
export const DIRECTION = {
  max: {
    before: [
      "not exceed",
      "no more than",
      "not more than",
      "at most",
      "up to",
      "maximum",
      "limit",
      "limited to",
      "no longer than",
      "not longer than",
      "within",
    ],
    after: ["at most", "maximum", "or fewer", "or less"],
  },
  min: {
    before: [
      "at least",
      "minimum",
      "no fewer than",
      "not fewer than",
      "no less than",
      "not less than",
    ],
    after: ["at least", "minimum", "or more"],
  },
};

// A direction word says nothing when it is denied ("no maximum", "not limited
// to", "without a limit"): one of these words right before it, with at most
// an article in between. Such a sentence stays a model-checked rule.
export const LIMIT_NEGATION = ["no", "not", "never", "without", "nor"];

// A number given as advice is not a limit ("about 250 words", "250 words is a
// suggestion"). One of these words in the clause of the number (CLAUSE_BREAKS)
// keeps the sentence a model-checked rule.
export const ADVISORY = [
  "about",
  "approximately",
  "around",
  "roughly",
  "typically",
  "usually",
  "generally",
  "normally",
  "ideally",
  "ideal",
  "preferably",
  "suggest",
  "suggests",
  "suggested",
  "suggestion",
  "recommend",
  "recommends",
  "recommended",
  "recommendation",
  "aim",
];

const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/ /g, "\\s+");
const alt = (list) =>
  [...list]
    .sort((a, b) => b.length - a.length)
    .map(esc)
    .join("|");
const wholeWord = (list, tail = "") => new RegExp(`(?<![a-z])(?:${alt(list)})${tail}(?![a-z])`);

const NUMBER = /(?<![\d.,])(?:\d{1,3}(?:[,\s]\d{3})+|\d+)(?!\d|[.,]\d)/g;

// The whole numbers a sentence writes in digits, with their places.
function numberSpans(sentence) {
  return [...String(sentence ?? "").matchAll(NUMBER)].map((m) => ({
    value: Number(m[0].replace(/[,\s]/g, "")),
    start: m.index,
    end: m.index + m[0].length,
  }));
}

// The whole numbers a sentence writes in digits. "1,000", "1 000" (also with
// a no-break space) and "1000" are one number; its parts ("1", "000", "100")
// are not numbers of the sentence. A number with a decimal point or comma
// ("2.5", "1,5", "1.000") is left out: it cannot be read as a whole number
// with certainty.
export function numbersIn(sentence) {
  return numberSpans(sentence).map((s) => s.value);
}

const sectionNamed = (section, lower) => wholeWord(SECTION_TERMS[section], "s?").test(lower);

// Every place where the sentence names a section, in reading order.
function sectionMentions(lower) {
  const out = [];
  for (const section of SECTIONS) {
    const re = new RegExp(wholeWord(SECTION_TERMS[section], "s?").source, "g");
    for (const m of lower.matchAll(re)) {
      out.push({ section, start: m.index, end: m.index + m[0].length });
    }
  }
  return out.sort((a, b) => a.start - b.start);
}

const MARKS = CLAUSE_BREAKS.marks.map(esc).join("");
const BREAK = new RegExp(`[${MARKS}]|(?<![a-z])(?:${alt(CLAUSE_BREAKS.words)})(?![a-z])`, "g");
const JOINERS = alt([...CLAUSE_BREAKS.words, "the", "a", "an"]);
const ONLY_JOINS = new RegExp(`^(?:\\s|[${MARKS}/&]|(?<![a-z])(?:${JOINERS})(?![a-z]))*$`);

// [start, end] of the clause that holds the number at `span` (see
// CLAUSE_BREAKS for what a clause is).
function clauseOf(span, spans, mentions, lower) {
  const within = (list, a, b) => list.some((x) => a >= x.start && b <= x.end);
  const joins = (list, a, b) =>
    list.some((x, i) => {
      const next = list[i + 1];
      return (
        next && a >= x.end && b <= next.start && ONLY_JOINS.test(lower.slice(x.end, next.start))
      );
    });
  let from = 0;
  let to = lower.length;
  for (const m of lower.matchAll(BREAK)) {
    const [a, b] = [m.index, m.index + m[0].length];
    if (within(spans, a, b) || within(mentions, a, b)) continue;
    if (joins(spans, a, b) || joins(mentions, a, b)) continue;
    if (b <= span.start) from = b;
    else if (a >= span.end) {
      to = a;
      break;
    }
  }
  return [from, to];
}

// The sections named in the clause that holds the number at `span`.
function sectionsInClause(span, spans, mentions, lower) {
  const [from, to] = clauseOf(span, spans, mentions, lower);
  const inside = mentions.filter((x) => x.start >= from && x.end <= to);
  return [...new Set(inside.map((x) => x.section))];
}

// Does the number at `span` count the target's unit? "together" when figures
// and tables are counted as one.
function countsUnit(target, span, lower) {
  const unit = UNIT_WORDS[target];
  const after = lower.slice(span.end);
  const units = alt(span.value === 1 ? [...unit.plural, ...unit.singular] : unit.plural);
  const range = "(?:\\s*(?:to|and|or|-|\\u2013)\\s*(?:[a-z]+\\s+){0,3}?\\d[\\d,\\s]*)?";
  const next = new RegExp(`^${range}(?:[\\s-]+[a-z]+)?[\\s-]+(?:${units})(?![a-z])`).exec(after);
  if (next) {
    const joined = /^\s*(?:and|or|&|\/)\s*(?:figures?|tables?)(?![a-z])/.test(
      after.slice(next[0].length),
    );
    return joined && (target === "figures" || target === "tables") ? "together" : true;
  }
  // "a 250-word limit"
  if (new RegExp(`^-(?:${alt(unit.singular)})(?![a-z])`).test(after)) return true;
  // "The number of figures must not exceed 8."
  const named = wholeWord([...unit.plural, ...unit.singular]).test(lower.slice(0, span.start));
  return named && /^\s*(?:[.;,)\]]|$)/.test(after);
}

// [lower, upper] of every range the sentence writes, as the two number spans:
// "4 to 6", "4-6", "between 150 and 250".
function rangesIn(spans, lower) {
  const out = [];
  for (let i = 0; i + 1 < spans.length; i++) {
    const gap = lower.slice(spans[i].end, spans[i + 1].start).trim();
    const between = gap === "and" && /(?<![a-z])between\s*$/.test(lower.slice(0, spans[i].start));
    if ((between || ["to", "-", "\u2013"].includes(gap)) && spans[i].value < spans[i + 1].value) {
      out.push([spans[i], spans[i + 1]]);
    }
  }
  return out;
}

const FILLER = /(?:[\s:(]|(?<![a-z])(?:of|a|an|the|total|is|be)(?![a-z]))*$/;

const DENIED = new RegExp(`(?<![a-z])(?:${alt(LIMIT_NEGATION)})(?:\\s+(?:a|an|the|any))?\\s+$`);

// Does the sentence say that the number at `span` is a limit on `side`
// ("max" or "min")? A direction word that is denied says nothing.
function saysLimit(side, span, spans, lower) {
  const other = side === "max" ? "min" : "max";
  // The number itself must stand in the range: the same value elsewhere in
  // the sentence is another number.
  if (rangesIn(spans, lower).some(([lo, hi]) => span === (side === "max" ? hi : lo))) {
    return true;
  }
  const lead = lower.slice(0, span.start).replace(FILLER, "");
  const before = new RegExp(`(?<![a-z])(?:${alt(DIRECTION[side].before)})$`).exec(lead);
  if (before) return !DENIED.test(lead.slice(0, before.index));
  const trail = new RegExp(
    `^(?:[\\s-]+[a-z]+){0,2}\\s+(?:${alt(DIRECTION[side].after)})\\s*(?:[.;,)\\]]|$)`,
  );
  if (trail.test(lower.slice(span.end))) return true;
  // One number in the sentence and words of one direction only, none of them
  // denied: "The maximum length of the introduction is 600 words."
  const said = (s) => {
    const re = new RegExp(wholeWord([...DIRECTION[s].before, ...DIRECTION[s].after]).source, "g");
    return [...lower.matchAll(re)].map((m) => DENIED.test(lower.slice(0, m.index)));
  };
  const [mine, theirs] = [said(side), said(other)];
  return spans.length === 1 && mine.length > 0 && !mine.includes(true) && theirs.length === 0;
}

// What a count rule claims that its sentence does not say, or null.
function unsupported(rule, sentence) {
  const lower = sentence.toLowerCase();
  if (rule.target === "section_order") {
    const absent = rule.order.filter((name) => !sectionNamed(name, lower));
    if (absent.length) return `section ${absent.join(", ")} is not in the sentence`;
    const plain = lower.replace(new RegExp(`(?<![a-z])(?:${alt(ORDER_KEEPS)})(?![a-z])`, "g"), " ");
    const doubt = wholeWord(ORDER_DOUBT).exec(plain);
    if (doubt) {
      return `the sentence has "${doubt[0]}", so the order in which it names the sections may not be the order it asks for`;
    }
    const first = (name) => sectionMentions(lower).find((x) => x.section === name).start;
    const inOrder = rule.order.every(
      (name, i) => i === 0 || first(rule.order[i - 1]) < first(name),
    );
    return inOrder
      ? null
      : `the sentence does not name the sections in the order ${rule.order.join(", ")}`;
  }
  const spans = numberSpans(sentence);
  for (const n of [rule.min, rule.max]) {
    if (n !== undefined && !spans.some((s) => s.value === n)) {
      return `number ${n} is not in the sentence`;
    }
  }
  if (rule.section && !sectionNamed(rule.section, lower)) {
    return `the section "${rule.section}" is not named in the sentence`;
  }
  const what = UNIT_WORDS[rule.target].plural[0];
  const mentions = sectionMentions(lower);
  const several = rule.section ? [...new Set(mentions.map((x) => x.section))] : [];
  for (const [side, n] of [
    ["min", rule.min],
    ["max", rule.max],
  ]) {
    if (n === undefined) continue;
    let here = spans.filter((s) => s.value === n);
    if (several.length > 1) {
      here = here.filter((s) => {
        const named = sectionsInClause(s, spans, mentions, lower);
        return named.length === 1 && named[0] === rule.section;
      });
      if (!here.length) {
        return `the sentence names several sections (${several.join(", ")}) and does not tie ${n} to the ${rule.section} alone`;
      }
    }
    const counted = here.map((s) => countsUnit(rule.target, s, lower));
    if (counted.includes("together") && !counted.includes(true)) {
      return "the sentence counts figures and tables together";
    }
    const bound = here.filter((_, i) => counted[i] === true);
    if (!bound.length) return `${n} does not count ${what} in the sentence`;
    const advice = bound.map((s) => {
      const [from, to] = clauseOf(s, spans, mentions, lower);
      return wholeWord(ADVISORY).exec(lower.slice(from, to))?.[0];
    });
    if (advice.every(Boolean)) {
      return `the sentence gives ${n} as advice ("${advice[0]}"), not as a limit`;
    }
    const firm = bound.filter((_, i) => !advice[i]);
    if (!firm.some((s) => saysLimit(side, s, spans, lower))) {
      const limit = side === "max" ? "an upper limit" : "a lower limit";
      return `the sentence does not say that ${n} is ${limit}`;
    }
  }
  return null;
}

const none = () => ({ candidates: [], warnings: [GUIDELINE_WARNING] });

// sentences: the guideline sentences that were numbered in the prompt.
// json: the model's reply. `offset` only shifts the numbers named in warnings.
export function candidatesFromLabels(sentences, json, { offset = 0 } = {}) {
  const list = isPlain(json) ? json.sentences : null;
  if (!Array.isArray(list) || list.length !== sentences.length) return none();
  const labels = new Array(sentences.length);
  for (const x of list) {
    const s = x?.s;
    if (!Number.isInteger(s) || s < 1 || s > labels.length || labels[s - 1]) return none();
    const label = sentenceLabel(x);
    if (!label) return none();
    labels[s - 1] = label;
  }
  const candidates = [];
  const warnings = [];
  labels.forEach(([type, rule], i) => {
    if (type === "other") return;
    const sentence = sentences[i];
    let check = { type: "llm" };
    if (type === "count") {
      const gap = unsupported(rule, sentence);
      if (gap) {
        warnings.push(
          `s${offset + i + 1}: ${gap}; kept as a rule the model checks, not as a count`,
        );
      } else {
        check = {
          type: "count",
          ...rule,
          ...(rule.section && { section: SECTION_PATTERNS[rule.section] }),
        };
      }
    }
    let profile = null;
    if (type === "format") {
      const setting = FORMAT_SETTINGS[rule.setting];
      const lower = sentence.toLowerCase();
      const said =
        setting.needs.every((words) => wholeWord(FORMAT_WORDS[words]).test(lower)) &&
        !wholeWord(FORMAT_WORDS.negation).test(lower);
      if (said) profile = structuredClone(setting.profile);
      else {
        warnings.push(
          `s${offset + i + 1}: the sentence does not say that ${setting.says}; kept as a rule the model checks, without a profile setting`,
        );
      }
    }
    candidates.push({
      kind: "instruction",
      text: sentence,
      scope: "manuscript",
      check,
      source: { kind: "guideline", sentence },
      ...(profile && { profile }),
    });
  });
  return { candidates, warnings };
}

export async function extractGuidelineRules(
  text,
  { oneshot, splitSentences, batchSize = 100 } = {},
) {
  if (typeof oneshot !== "function" || typeof splitSentences !== "function") {
    throw new TypeError("extractGuidelineRules needs oneshot and splitSentences");
  }
  const sentences = splitSentences(text);
  const candidates = [];
  const warnings = [];
  for (let at = 0; at < sentences.length; at += batchSize) {
    const batch = sentences.slice(at, at + batchSize);
    const r = await oneshot({ system: GUIDELINE_SYSTEM, prompt: guidelinePrompt(batch) });
    if (!r?.ok) {
      return {
        candidates: [],
        warnings: [`guideline labelling failed: ${r?.error}; add the rules by hand`],
      };
    }
    const part = candidatesFromLabels(batch, r.json, { offset: at });
    if (part.warnings[0] === GUIDELINE_WARNING) return none();
    candidates.push(...part.candidates);
    warnings.push(...part.warnings);
  }
  return { candidates, warnings };
}

const ENTITIES = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: " " };

export function htmlToText(html) {
  return String(html ?? "")
    .replace(/<!--[\s\S]*?-->/g, "")
    .replace(/<(script|style)\b[\s\S]*?<\/\1\s*>/gi, "")
    .replace(/<\/(?:p|div|li|tr|h[1-6]|ul|ol|table|section|article)\s*>|<br\s*\/?>/gi, "\n")
    .replace(/<[^>]*>/g, "")
    .replace(/&(#x?[0-9a-f]+|[a-z]+);/gi, (m, name) => {
      if (name[0] !== "#") return ENTITIES[name.toLowerCase()] ?? m;
      const code = /^#x/i.test(name) ? parseInt(name.slice(2), 16) : Number(name.slice(1));
      return code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : m;
    })
    .split("\n")
    .map((line) => line.replace(/\s+/g, " ").trim())
    .filter(Boolean)
    .join("\n");
}

// Only called when the author typed a URL. One request, http(s) only.
export async function fetchGuidelineText(
  url,
  { fetchFn, maxBytes = 2_000_000, timeoutMs = FETCH_TIMEOUT_MS } = {},
) {
  let parsed;
  try {
    parsed = new URL(url);
  } catch {
    return { ok: false, error: "not a URL" };
  }
  if (!["http:", "https:"].includes(parsed.protocol)) {
    return { ok: false, error: "only http and https addresses are read" };
  }
  try {
    const res = await fetchFn(url, { signal: deadline(timeoutMs) });
    if (!res.ok) return { ok: false, error: `the page answered ${res.status}` };
    const body = await res.text();
    if (Buffer.byteLength(body) > maxBytes) {
      return { ok: false, error: `the page is larger than ${maxBytes} bytes; paste the text` };
    }
    const html = /html/i.test(res.headers?.get?.("content-type") ?? "");
    return { ok: true, text: html ? htmlToText(body) : body };
  } catch (e) {
    return { ok: false, error: String(e?.message ?? e) };
  }
}
