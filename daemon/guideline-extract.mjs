// Draft ledger rules from a journal's author guidelines.
//
// The model never writes a rule. The code cuts the guideline into sentences
// (the splitter is passed in), the model labels each sentence number with a
// type and, for a countable limit, a few parameters from closed lists. A rule
// candidate is the guideline sentence as written, and it cites that sentence.
// A count rule is kept only when its own sentence names the section, has
// the number, counts the same thing with it and says in which direction the
// number limits it; otherwise the sentence stays as a model-checked rule and
// a warning says what was missing. A reply that is malformed, partial or carries any text
// yields no candidate at all.
//
// Pure apart from the injected functions: oneshot({ system, prompt }) ->
// { ok, json?, error? } (daemon/oneshot.mjs), splitSentences(text) ->
// string[], fetchFn(url) -> Response.
//
// Not wired yet. The caller passes splitSentences from review-split.mjs and
// the one-shot function of the chosen provider, and shows candidates to the
// author before any of them enters the ledger.

import { COUNT_TARGETS } from "./count-check.mjs";

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

// [lower, upper] of every range the sentence writes: "4 to 6", "4-6",
// "between 150 and 250".
function rangesIn(spans, lower) {
  const out = [];
  for (let i = 0; i + 1 < spans.length; i++) {
    const gap = lower.slice(spans[i].end, spans[i + 1].start).trim();
    const between = gap === "and" && /(?<![a-z])between\s*$/.test(lower.slice(0, spans[i].start));
    if ((between || ["to", "-", "\u2013"].includes(gap)) && spans[i].value < spans[i + 1].value) {
      out.push([spans[i].value, spans[i + 1].value]);
    }
  }
  return out;
}

const FILLER = /(?:[\s:(]|(?<![a-z])(?:of|a|an|the|total|is|be)(?![a-z]))*$/;

// Does the sentence say that the number at `span` is a limit on `side`
// ("max" or "min")?
function saysLimit(side, span, spans, lower) {
  const other = side === "max" ? "min" : "max";
  if (rangesIn(spans, lower).some(([lo, hi]) => span.value === (side === "max" ? hi : lo))) {
    return true;
  }
  const lead = lower.slice(0, span.start).replace(FILLER, "");
  if (new RegExp(`(?<![a-z])(?:${alt(DIRECTION[side].before)})$`).test(lead)) return true;
  const trail = new RegExp(
    `^(?:[\\s-]+[a-z]+){0,2}\\s+(?:${alt(DIRECTION[side].after)})\\s*(?:[.;,)\\]]|$)`,
  );
  if (trail.test(lower.slice(span.end))) return true;
  // One number in the sentence and words of one direction only: "The maximum
  // length of the introduction is 600 words."
  const has = (s) => wholeWord([...DIRECTION[s].before, ...DIRECTION[s].after]).test(lower);
  return spans.length === 1 && has(side) && !has(other);
}

// What a count rule claims that its sentence does not say, or null.
function unsupported(rule, sentence) {
  const lower = sentence.toLowerCase();
  if (rule.target === "section_order") {
    const absent = rule.order.filter((name) => !sectionNamed(name, lower));
    return absent.length ? `section ${absent.join(", ")} is not in the sentence` : null;
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
  for (const [side, n] of [
    ["min", rule.min],
    ["max", rule.max],
  ]) {
    if (n === undefined) continue;
    const here = spans.filter((s) => s.value === n);
    const counted = here.map((s) => countsUnit(rule.target, s, lower));
    if (counted.includes("together") && !counted.includes(true)) {
      return "the sentence counts figures and tables together";
    }
    const bound = here.filter((_, i) => counted[i] === true);
    if (!bound.length) return `${n} does not count ${what} in the sentence`;
    if (!bound.some((s) => saysLimit(side, s, spans, lower))) {
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
export async function fetchGuidelineText(url, { fetchFn, maxBytes = 2_000_000 } = {}) {
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
    const res = await fetchFn(url);
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
