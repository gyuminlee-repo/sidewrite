// Draft ledger rules from a journal's author guidelines.
//
// The model never writes a rule. The code cuts the guideline into sentences
// (the splitter is passed in), the model labels each sentence number with a
// type and, for a countable limit, a few parameters from closed lists. A rule
// candidate is the guideline sentence as written, and it cites that sentence.
// A count rule whose number or section is not in its own sentence is not
// kept as a count rule: the sentence stays as a model-checked rule and a
// warning names it. A reply that is malformed, partial or carries any text
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

export const GUIDELINE_TYPES = ["count", "style", "other"];

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

// What a count rule claims that its sentence does not say, or null.
function unsupported(rule, sentence) {
  if (rule.target === "section_order") {
    const lower = sentence.toLowerCase();
    const absent = rule.order.filter((name) => !lower.includes(name));
    return absent.length ? `section ${absent.join(", ")}` : null;
  }
  const written = numbersIn(sentence);
  for (const n of [rule.min, rule.max]) {
    if (n !== undefined && !written.includes(n)) return `number ${n}`;
  }
  return null;
}

// The whole numbers a sentence writes in digits. "1,000", "1 000" (also with
// a no-break space) and "1000" are one number; its parts ("1", "000", "100")
// are not numbers of the sentence. A number with a decimal point or comma
// ("2.5", "1,5", "1.000") is left out: it cannot be read as a whole number
// with certainty.
export function numbersIn(sentence) {
  const re = /(?<![\d.,])(?:\d{1,3}(?:[,\s]\d{3})+|\d+)(?!\d|[.,]\d)/g;
  return [...String(sentence ?? "").matchAll(re)].map((m) => Number(m[0].replace(/[,\s]/g, "")));
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
          `s${offset + i + 1}: ${gap} is not in the sentence; kept as a rule the model checks, not as a count`,
        );
      } else {
        check = {
          type: "count",
          ...rule,
          ...(rule.section && { section: SECTION_PATTERNS[rule.section] }),
        };
      }
    }
    candidates.push({
      kind: "instruction",
      text: sentence,
      scope: "manuscript",
      check,
      source: { kind: "guideline", sentence },
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
