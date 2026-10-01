// Ledger rule candidates from the review text (P15). The author confirms
// them in the task pane; nothing here is stored.
//
// A candidate is a sentence of the review as written, with a drafted check:
//   keep cue + quoted phrase    -> instruction, regex, the phrase must stay
//   forbid cue + quoted phrase  -> forbid, regex, the phrase must be absent
//   a cue without a quote       -> instruction, LLM check
// With a one-shot model the sentences no cue matched are labelled by index
// (forbid, instruction or none). The model returns numbers and types only,
// so a candidate's text is always text of the review (the pattern of
// review-split.mjs splitPrompt and requirementPrompt).

import { splitSentences } from "./review-split.mjs";

const KEEP_RE = /\b(keep|retain|do not (?:delete|remove)|don't (?:delete|remove))\b/i;
const FORBID_RE = /\b(do not|don't|avoid|must not|should not|never|refrain from)\b/i;
// The quote forms of review-view.js ledgerCheckFromText. A single quote next
// to a letter is an apostrophe ("don't", "authors'"), not a quote.
const QUOTED = /"([^"]+)"|(?<!\p{L})'([^']+)'(?!\p{L})|「([^」]+)」|“([^”]+)”/u;
const LIST_LINE = /^\s*(?:\(?\d+[.)]|\(?[a-z][.)]|[-*•])\s+/i;
const ENDS_SENTENCE = /[.!?:]["')\]]*\s*$/;

const escapeRegex = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

// Sentences of the whole review. splitSentences joins every line break (a
// comment's text wraps), which would glue a heading such as "Reviewer 1" to
// the sentence under it, so each line is split on its own. A line is joined
// to the one before only when that one does not end a sentence and this one
// does not start with a capital or a list mark (a wrapped line).
export function reviewSentences(text) {
  const lines = [];
  for (const line of String(text ?? "")
    .replace(/\r\n?/g, "\n")
    .split("\n")) {
    const prev = lines.at(-1);
    const wrapped =
      !!prev?.trim() &&
      !!line.trim() &&
      !ENDS_SENTENCE.test(prev) &&
      !/^\s*\p{Lu}/u.test(line) &&
      !LIST_LINE.test(line);
    if (wrapped) lines[lines.length - 1] = `${prev} ${line.trim()}`;
    else lines.push(line);
  }
  return lines.flatMap((l) => splitSentences(l));
}

function quotedPhrase(sentence) {
  const m = QUOTED.exec(sentence);
  return m ? (m[1] ?? m[2] ?? m[3] ?? m[4]).trim() : "";
}

// type: "keep" | "forbid" | "instruction".
function candidate(sentence, index, type, why) {
  const phrase = type === "instruction" ? "" : quotedPhrase(sentence);
  const base = { text: sentence, scope: "manuscript", why, sentence_index: index };
  if (!phrase) return { ...base, kind: "instruction", check: { type: "llm" } };
  return type === "keep"
    ? {
        ...base,
        kind: "instruction",
        check: { type: "regex", pattern: escapeRegex(phrase), flags: "", expect: "present" },
      }
    : {
        ...base,
        kind: "forbid",
        check: { type: "regex", pattern: escapeRegex(phrase), flags: "i", expect: "absent" },
      };
}

function regexCandidates(sentences) {
  const out = [];
  sentences.forEach((s, i) => {
    if (KEEP_RE.test(s)) out.push(candidate(s, i, "keep", "keep cue"));
    else if (FORBID_RE.test(s)) out.push(candidate(s, i, "forbid", "forbid cue"));
  });
  return out;
}

export function suggestLedgerRegex(text) {
  return regexCandidates(reviewSentences(text));
}

export const SUGGEST_SYSTEM =
  "You find standing rules in a peer review: things the authors must never write, or must keep doing, in the whole manuscript. Label numbered sentences. Do not return any text of the review. Reply with JSON only.";

export const SUGGEST_TYPES = ["forbid", "instruction", "none"];

// indexes: the sentences to label, shown under their index in `sentences`.
export function suggestPrompt(sentences, indexes) {
  return [
    'Return {"labels":[{"i":<sentence number>,"type":"<type>"}]} with every sentence number exactly once.',
    'Types: "forbid" (a word, phrase or claim that must not appear anywhere in the manuscript), "instruction" (a rule to follow in the whole manuscript, such as a term, a unit or a style), "none" (a request about one place, a question, praise, a heading or anything else).',
    "Sentences:",
    ...indexes.map((i) => `${i}: ${sentences[i]}`),
  ].join("\n");
}

// Returns { candidates, method, warnings, asked, input }: method "regex" or
// "regex+llm"; asked and input (the prompt) are for the decision receipt.
// A failed call or a label that cannot be used leaves the regex candidates
// and adds a warning.
export async function suggestLedger(text, { oneshot } = {}) {
  const sentences = reviewSentences(text);
  const found = regexCandidates(sentences);
  const taken = new Set(found.map((c) => c.sentence_index));
  const open = sentences.map((_, i) => i).filter((i) => !taken.has(i));
  const regexOnly = { candidates: found, method: "regex", warnings: [], asked: false, input: null };
  if (!oneshot || !open.length) return regexOnly;
  const input = suggestPrompt(sentences, open);
  const r = await oneshot({ system: SUGGEST_SYSTEM, prompt: input });
  if (!r.ok) {
    return {
      ...regexOnly,
      warnings: [`model suggestions failed: ${r.error}; cue-based candidates only`],
      asked: true,
      input,
    };
  }
  const labels = r.json?.labels;
  if (!Array.isArray(labels)) {
    return {
      ...regexOnly,
      warnings: ["model reply had no labels; cue-based candidates only"],
      asked: true,
      input,
    };
  }
  const warnings = [];
  const added = [];
  const seen = new Set();
  for (const x of labels) {
    const i = x?.i;
    if (!Number.isInteger(i) || !open.includes(i) || seen.has(i)) {
      warnings.push(
        `model label for sentence ${JSON.stringify(i)} ignored: not a sentence it was asked about`,
      );
      continue;
    }
    if (!SUGGEST_TYPES.includes(x.type)) {
      warnings.push(`model label for sentence ${i} ignored: unknown type`);
      continue;
    }
    seen.add(i);
    if (x.type !== "none") added.push(candidate(sentences[i], i, x.type, "model label"));
  }
  return {
    candidates: [...found, ...added].sort((a, b) => a.sentence_index - b.sentence_index),
    method: "regex+llm",
    warnings,
    asked: true,
    input,
  };
}
