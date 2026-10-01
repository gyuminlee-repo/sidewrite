// Split a pasted review into reviewer comments with refs R{reviewer}.{n}.
// Regex first (no model call); a single LLM one-shot only when the regex
// finds at most one comment. The user edits the result in the task pane.
// The one-shot recovers document structure: it labels each numbered line of
// the review and returns no text, and the code builds the comments from the
// source lines, so no line is lost or rewritten.
// L5: every comment carries requirements [{ id, text }], the separate things
// the reviewer asks for. Both splits leave them empty; extractRequirements
// fills them afterwards the same way: the code splits the comment sentences
// and marks where a sentence could be cut (clauseCuts), and one one-shot
// labels each sentence and each mark, so a requirement is always a sentence
// of the comment, or a part of one, as written.

import { numberRequirements } from "./revision-store.mjs";

// Defaults are stricter than the plain design patterns so body lines do not
// start a new comment: "Reviewer 2 raised the same point" (reviewer word at
// line start, followed by prose), "Q10 values" (temperature coefficient),
// "3.5 mM" (decimal at line start).
export const DEFAULT_PATTERNS = {
  reviewerHead: String.raw`^\s*Reviewer\s*#?\s*(\d+)\b\s*(?:[:.)\-].*)?$`,
  commentHead: String.raw`^\s*(?:(?:Comment|Point)\s*#?\s*(\d+)\b|Q\s*#?\s*(\d+)\s*[:.)])`,
  listItem: String.raw`^\s*(\d+)[.)]\s`,
};

function firstNumber(m) {
  for (let i = 1; i < m.length; i++) if (m[i] !== undefined) return Number(m[i]);
  return null;
}

const HEAD_TRIM = /^[\s:.)\-]+/;

// A journal's own label on the reviewer head line, not the reviewer's words.
const STOCK_LABEL = /^(?:comments?|remarks?)\s+(?:to|for)\s+(?:the\s+)?(?:authors?|editors?)\W*$/i;

// Text on a reviewer head line after the reviewer number ("Reviewer 1: Please
// report n." gives "Please report n."; a stock label gives ""). The regex
// needs the d flag.
function reviewerRest(m, line) {
  const end = m.indices?.[1]?.[1];
  const rest = end == null ? "" : line.slice(end).replace(HEAD_TRIM, "").trim();
  return STOCK_LABEL.test(rest) ? "" : rest;
}

export function splitReviewRegex(text, patterns = DEFAULT_PATTERNS) {
  return splitBlocks(text, patterns).comments;
}

// Regex split plus the refs of reviewer blocks kept whole and of kept
// preambles. Text after a reviewer head on its line is the block's first
// body line. A block opened by a reviewer head with body text but no comment
// head is one comment R{reviewer}.1 (no numbered comment of that block can
// share the ref), so a reviewer who did not number the points is not
// dropped. When the review has a reviewer head, the text before the first
// one (the editor letter) is one comment E.1 with reviewer 0, and a
// reviewer preamble before the block's first comment head is R{reviewer}.0.
// Both are kept even when they only greet; the user deletes them. Without
// any reviewer head the text before the first comment head is dropped.
function splitBlocks(text, patterns = DEFAULT_PATTERNS) {
  const p = { ...DEFAULT_PATTERNS, ...patterns };
  const reReviewer = new RegExp(p.reviewerHead, "id");
  const reComment = new RegExp(p.commentHead, "i");
  const reList = new RegExp(p.listItem);
  const lines = String(text ?? "")
    .replace(/\r\n?/g, "\n")
    .split("\n");

  const blocks = [];
  let cur = { reviewer: 1, lines: [], headed: false };
  for (const line of lines) {
    const m = reReviewer.exec(line);
    if (m) {
      blocks.push(cur);
      const rest = reviewerRest(m, line);
      cur = { reviewer: Number(m[1]), lines: rest ? [rest] : [], headed: true };
    } else {
      cur.lines.push(line);
    }
  }
  blocks.push(cur);

  const out = [];
  for (const b of blocks) {
    if (!b.headed && blocks.length > 1) {
      out.push({ reviewer: 0, editor: true, number: 1, lines: b.lines, kept: true });
      continue;
    }
    // Explicit "Comment n" heads win; then numbered-list items are body text.
    const head = b.lines.some((l) => reComment.test(l)) ? reComment : reList;
    if (b.headed && b.lines.join("").trim() && !b.lines.some((l) => head.test(l))) {
      out.push({ reviewer: b.reviewer, number: 1, lines: b.lines, whole: true });
      continue;
    }
    let item = null;
    let seq = 0;
    const preamble = [];
    for (const line of b.lines) {
      const m = head.exec(line);
      if (m && !item && b.headed) {
        out.push({ reviewer: b.reviewer, number: 0, lines: preamble, kept: true });
      }
      if (m) {
        if (item) out.push(item);
        seq += 1;
        item = {
          reviewer: b.reviewer,
          number: firstNumber(m) ?? seq,
          lines: [line.slice(m[0].length)],
        };
      } else if (item) {
        item.lines.push(line);
      } else {
        preamble.push(line);
      }
    }
    if (item) out.push(item);
  }
  const rows = out
    .map((c) => ({
      whole: !!c.whole,
      kept: !!c.kept,
      comment: {
        ref: c.editor ? `E.${c.number}` : `R${c.reviewer}.${c.number}`,
        reviewer: c.reviewer,
        text: c.lines.join("\n").replace(HEAD_TRIM, "").trim(),
        requirements: [],
      },
    }))
    .filter((x) => x.comment.text);
  return {
    comments: rows.map((x) => x.comment),
    unsplit: rows.filter((x) => x.whole).map((x) => x.comment.ref),
    preserved: rows.filter((x) => x.kept).map((x) => x.comment),
  };
}

// late: refs of editor text that comes after a reviewer (label split only).
const preservedWarnings = (comments, late = []) =>
  comments.map((c) =>
    late.includes(c.ref)
      ? `Editor text after a reviewer is kept as ${c.ref}. Delete it if it asks for nothing.`
      : c.reviewer === 0
        ? `Text before the first reviewer (usually the editor letter) is kept as ${c.ref}. Delete it if it asks for nothing.`
        : `Reviewer ${c.reviewer} text before the first numbered comment is kept as ${c.ref}. Delete it if it asks for nothing.`,
  );

const unsplitWarnings = (refs) =>
  refs.map(
    (ref) =>
      `Reviewer ${ref.slice(1, ref.indexOf("."))} text was not split into numbered comments; it is kept whole as ${ref}. Split it by hand if it holds several requests.`,
  );

function duplicateWarnings(comments) {
  const seen = new Set();
  const w = [];
  for (const c of comments) {
    if (seen.has(c.ref)) w.push(`duplicate ref ${c.ref}: renumber it before saving`);
    seen.add(c.ref);
  }
  return w;
}

const squash = (s) =>
  String(s ?? "")
    .replace(/\s+/g, " ")
    .trim();

// Sentence units for the coverage check: split at line breaks and sentence
// ends, drop a leading list marker (a list head loses it in its comment) and keep
// units with a letter. An abbreviation ("e.g.", "Fig. 2") only splits a
// sentence into shorter pieces, each still inside a comment that holds it.
const LIST_MARK = /^(?:\(?\d+[.)]|\(?[a-z][.)]|[-*\u2022])\s+/i;
function sentenceUnits(text) {
  return String(text ?? "")
    .split(/\n|(?<=[.!?])\s+/)
    .map((u) => squash(u).replace(LIST_MARK, ""))
    .filter((u) => /\p{L}/u.test(u));
}

const quoteUnits = (units, max = 80) =>
  units
    .slice(0, 3)
    .map((u) => `"${u.length > max ? `${u.slice(0, max - 3)}...` : u}"`)
    .join(", ") + (units.length > 3 ? ` and ${units.length - 3} more` : "");

// Sentences of the review, without reviewer heads and comment heads (the
// text after either on its line stays), that are inside none of the comments.
export function uncoveredText(text, comments, patterns) {
  const p = { ...DEFAULT_PATTERNS, ...patterns };
  const reReviewer = new RegExp(p.reviewerHead, "id");
  const reComment = new RegExp(p.commentHead, "i");
  const body = String(text ?? "")
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .map((l) => {
      const r = reReviewer.exec(l);
      if (r) return reviewerRest(r, l);
      const m = reComment.exec(l);
      return m ? l.slice(m[0].length).replace(HEAD_TRIM, "") : l;
    })
    .join("\n");
  const texts = comments.map((c) => squash(c.text));
  return sentenceUnits(body).filter((u) => !texts.some((t) => t.includes(u)));
}

export const SPLIT_SYSTEM =
  "You recover the structure of a peer review. Give every numbered line exactly one type. Do not return any text of the review. Reply with JSON only.";

export const LINE_TYPES = ["editor", "reviewer_head", "comment_start", "continue", "drop"];

// lines: the non-empty lines of the review, numbered from 1 in the prompt.
export function splitPrompt(lines) {
  return [
    'Return {"lines":[{"n":<line number>,"type":"<type>"}]} with every line number exactly once.',
    'Types: "editor" (the editor letter), "reviewer_head" (a new reviewer starts; add "reviewer":<int>), "comment_start" (a new comment starts), "continue" (same comment as the line before), "drop" (only a page header, page footer or rule line).',
    "Lines:",
    ...lines.map((l, i) => `${i + 1}: ${l}`),
  ].join("\n");
}

// The model reply as one label per line, or null unless every line number
// 1..count is there exactly once with a known type, and every reviewer_head
// has a positive integer reviewer.
function validLabels(json, count) {
  const items = json?.lines;
  if (!Array.isArray(items) || items.length !== count) return null;
  const out = new Array(count);
  for (const it of items) {
    const n = it?.n;
    if (!Number.isInteger(n) || n < 1 || n > count || out[n - 1]) return null;
    if (!LINE_TYPES.includes(it.type)) return null;
    if (it.type === "reviewer_head" && !(Number.isInteger(it.reviewer) && it.reviewer > 0)) {
      return null;
    }
    out[n - 1] = it;
  }
  return out;
}

// Comments from the source lines and one label per non-empty line. A blank
// line joins the block before it. Only recognized heads lose text: the
// reviewer label of a reviewer_head line (regex reviewerHead, as in the regex
// split) and the comment marker of a comment_start line (commentHead or
// listItem, whose number becomes the ref). Every other line not dropped is in
// exactly one comment. Refs follow the regex split: E.{k} for each editor run,
// R{n}.0 for reviewer text before its first comment, R{n}.1 for a reviewer
// with no comment start, otherwise the marker number or the next free one.
function commentsFromLabels(src, labels, patterns) {
  const p = { ...DEFAULT_PATTERNS, ...patterns };
  const reReviewer = new RegExp(p.reviewerHead, "id");
  const reComment = new RegExp(p.commentHead, "i");
  const reList = new RegExp(p.listItem);
  const counts = Object.fromEntries(LINE_TYPES.map((t) => [t, 0]));
  const groups = [];
  const dropped = [];
  let cur = null;
  let reviewer = 1;
  let editors = 0;
  let n = 0;
  const open = (g) => groups.push((cur = g));
  for (const line of src) {
    if (!line.trim()) {
      cur?.lines.push(line);
      continue;
    }
    const label = labels[n++];
    counts[label.type] += 1;
    const type = label.type === "continue" && !cur ? "comment_start" : label.type;
    if (type === "drop") {
      dropped.push(line.trim());
    } else if (type === "continue" || (type === "editor" && cur?.editor)) {
      cur.lines.push(line);
    } else if (type === "editor") {
      const late = groups.some((g) => !g.editor);
      open({ reviewer: 0, editor: true, late, number: ++editors, lines: [line] });
    } else if (type === "reviewer_head") {
      reviewer = label.reviewer;
      const m = reReviewer.exec(line);
      open({ reviewer, number: 0, pre: true, lines: [m ? reviewerRest(m, line) : line] });
    } else {
      const m = reComment.exec(line) ?? reList.exec(line);
      const used = groups.filter((g) => !g.editor && g.reviewer === reviewer).map((g) => g.number);
      const number = (m && firstNumber(m)) ?? Math.max(0, ...used) + 1;
      open({
        reviewer,
        number,
        lines: [m ? line.slice(m[0].length).replace(HEAD_TRIM, "") : line],
      });
    }
  }
  const rows = groups
    .map((g) => {
      const whole =
        !!g.pre && !groups.some((c) => !c.pre && !c.editor && c.reviewer === g.reviewer);
      const number = whole ? 1 : g.number;
      return {
        whole,
        kept: !!g.editor || (!!g.pre && !whole),
        late: !!g.late,
        comment: {
          ref: g.editor ? `E.${number}` : `R${g.reviewer}.${number}`,
          reviewer: g.reviewer,
          text: g.lines.join("\n").trim(),
          requirements: [],
        },
      };
    })
    .filter((x) => x.comment.text)
    .sort((a, b) => a.comment.reviewer - b.comment.reviewer);
  const comments = rows.map((x) => x.comment);
  const warnings = [
    ...preservedWarnings(
      rows.filter((x) => x.kept).map((x) => x.comment),
      rows.filter((x) => x.late).map((x) => x.comment.ref),
    ),
    ...unsplitWarnings(rows.filter((x) => x.whole).map((x) => x.comment.ref)),
  ];
  if (dropped.length) {
    warnings.push(
      `${dropped.length} line(s) were dropped as page headers, footers or rules: ${quoteUnits(dropped, 40)}. Add them back if they ask for something.`,
    );
  }
  if (dropped.length * 2 > labels.length) {
    warnings.push(
      `The model dropped more than half of the lines (${dropped.length} of ${labels.length}); check the split against the review.`,
    );
  }
  return { comments, warnings: [...warnings, ...duplicateWarnings(comments)], labels: counts };
}

// oneshot: (req) => Promise<{ ok, json, error }> (daemon/oneshot.mjs).
// The regex result stands when its comment heads found more than one
// comment; blocks kept whole and kept preambles do not count, so they cannot
// hide a review the regex barely split. Kept-whole blocks and kept preambles
// always come back as warnings.
// A regex result ends with a check of the review text that is in no comment
// (without a reviewer head the regex drops the text before the first comment
// head); it is named, not added back. A result built from line labels needs
// no check: only dropped lines are left out, and a warning names them.
export async function splitReview(text, { patterns, oneshot } = {}) {
  const r = await splitReviewInner(text, { patterns, oneshot });
  if (r.labels) return r;
  const lost = uncoveredText(text, r.comments, patterns);
  if (!lost.length) return r;
  return {
    ...r,
    warnings: [
      ...r.warnings,
      `${lost.length} sentence(s) of the review are in no comment: ${quoteUnits(lost)}. Add them if they ask for something.`,
    ],
  };
}

async function splitReviewInner(text, { patterns, oneshot } = {}) {
  const { comments: regexOut, unsplit, preserved } = splitBlocks(text, patterns);
  const whole = [...preservedWarnings(preserved), ...unsplitWarnings(unsplit)];
  if (regexOut.length - unsplit.length - preserved.length > 1) {
    return {
      method: "regex",
      comments: regexOut,
      warnings: [...whole, ...duplicateWarnings(regexOut)],
    };
  }
  if (!oneshot) {
    return {
      method: "regex",
      comments: regexOut,
      warnings: ["regex found at most one comment; no LLM fallback configured", ...whole],
    };
  }
  const src = String(text ?? "")
    .replace(/\r\n?/g, "\n")
    .split("\n");
  const numbered = src.filter((l) => l.trim());
  if (!numbered.length) return { method: "regex", comments: regexOut, warnings: whole };
  const r = await oneshot({ system: SPLIT_SYSTEM, prompt: splitPrompt(numbered) });
  if (!r.ok) {
    return {
      method: "llm",
      comments: regexOut,
      warnings: [`LLM split failed: ${r.error}`, ...whole],
    };
  }
  // A reply that does not label every line once (empty, another JSON shape,
  // a line missing, doubled or unknown): the regex split stands, with a warning.
  const labels = validLabels(r.json, numbered.length);
  if (!labels) {
    return {
      method: "regex",
      comments: regexOut,
      warnings: [
        "model split returned no usable line labels; kept the regex split",
        ...whole,
        ...duplicateWarnings(regexOut),
      ],
    };
  }
  return { method: "llm", ...commentsFromLabels(src, labels, patterns) };
}

// ---- L5 requirement extraction ----------------------------------------------

// A period after one of these does not end a sentence ("e.g.", "Fig. 2",
// "et al. Nature"). A lowercase letter after the space does not either.
const ABBREV =
  /(?:^|[\s(])(?:e\.g|i\.e|et al|vs|cf|Figs?|Eqs?|Refs?|Tabs?|Nos?|Suppl|approx|resp|ca|Dr|Prof|Sec)\.$/i;
// A bare list number at the start of a unit or after a colon ("points: 1.").
const LIST_NUMBER = /(?:^|[:;]\s*)\(?\d{1,2}[.)]$/;
const LIST_LINE = /^\s*(?:\(?\d+[.)]|\(?[a-z][.)]|[-*\u2022])\s+/i;
const BOUNDARY = /[.!?]+["')\]]*\s+/g;

// Sentences of a comment, whitespace squashed, a leading list marker left
// off, only units with a letter. A line break is a space (pasted text wraps
// mid-sentence) except before a list line and at a blank line. Every
// sentence is a substring of the comment text with whitespace squashed.
export function splitSentences(text) {
  const chunks = [];
  let cur = [];
  for (const line of String(text ?? "")
    .replace(/\r\n?/g, "\n")
    .split("\n")) {
    if (!line.trim() || LIST_LINE.test(line)) {
      if (cur.length) chunks.push(cur.join(" "));
      cur = line.trim() ? [line] : [];
    } else {
      cur.push(line);
    }
  }
  if (cur.length) chunks.push(cur.join(" "));
  const out = [];
  for (const chunk of chunks) {
    const t = squash(chunk);
    let start = 0;
    for (const m of t.matchAll(BOUNDARY)) {
      const end = m.index + m[0].length;
      const head = t.slice(start, end).trimEnd();
      const next = t[end] ?? "";
      if (/\p{Ll}/u.test(next) || ABBREV.test(head) || LIST_NUMBER.test(head)) continue;
      out.push(head);
      start = end;
    }
    if (start < t.length) out.push(t.slice(start));
  }
  return out.map((u) => u.replace(LIST_MARK, "").trim()).filter((u) => /\p{L}/u.test(u));
}

// S2-30: one sentence may ask for two things ("please state how growth was
// followed, and describe the carbon source as novel"). The code marks the
// places where it could be cut; the model only says split or keep at each
// mark and gives each part a type, so a requirement is still text of the
// comment as written. A mark is a coordinator outside brackets: "; ",
// ", and ", " and also ", or " and " before an imperative verb or "please".
const IMPERATIVE =
  "add|address|avoid|change|check|cite|clarify|comment|compare|consider|correct|define|delete|describe|discuss|elaborate|explain|expand|give|include|indicate|justify|list|mention|move|name|note|perform|present|provide|quantify|remove|rename|rephrase|replace|report|revise|rewrite|show|specify|state|summarise|summarize|test|update|use";
const CUT = new RegExp(
  `;\\s+|,\\s+and\\s+(?:also\\s+)?|\\s+and\\s+also\\s+|\\s+and\\s+(?=(?:please\\s+)?(?:${IMPERATIVE})\\b)`,
  "gi",
);

// Cut points of one sentence: [{ end, start }], the part before ends at
// `end` and the part after starts at `start`; the coordinator between them
// belongs to no part. Only at bracket depth 0, and only where both sides
// hold a letter.
export function clauseCuts(sentence) {
  const t = String(sentence ?? "");
  const depth = [];
  let d = 0;
  // By UTF-16 index, as matchAll reports it.
  for (let i = 0; i < t.length; i++) {
    const ch = t[i];
    if (ch === "(" || ch === "[") d += 1;
    depth.push(d);
    if ((ch === ")" || ch === "]") && d > 0) d -= 1;
  }
  const cuts = [];
  let from = 0;
  for (const m of t.matchAll(CUT)) {
    if (depth[m.index]) continue;
    const end = m.index;
    const start = m.index + m[0].length;
    if (!/\p{L}/u.test(t.slice(from, end)) || !/\p{L}/u.test(t.slice(start))) continue;
    cuts.push({ end, start });
    from = start;
  }
  return cuts;
}

// Parts of a sentence for one "split"/"keep" per cut point: a kept cut joins
// its neighbours with the coordinator as written. Each part is a substring
// of the sentence (trailing ",;" and spaces left off).
export function requirementParts(sentence, decisions) {
  const t = String(sentence ?? "");
  const cuts = clauseCuts(t);
  const parts = [];
  let start = 0;
  cuts.forEach((c, i) => {
    if (decisions[i] !== "split") return;
    parts.push(t.slice(start, c.end));
    start = c.start;
  });
  parts.push(t.slice(start));
  return parts.map((x) => x.replace(/[\s,;]+$/u, "").trim());
}

export const REQUIREMENT_SYSTEM =
  "You sort the sentences of peer-review comments. Label every sentence exactly once: one type, or cuts and types where cut marks are shown. Do not return any text of the review. Reply with JSON only.";

export const REQUIREMENT_TYPES = ["request", "other"];

export const REQUIREMENT_WARNING = "model could not label requirement sentences; add them by hand";

// A sentence with its cut points marked <c1>, <c2>, ... where each part
// before the coordinator ends.
function markCuts(sentence) {
  let out = "";
  let at = 0;
  clauseCuts(sentence).forEach((c, i) => {
    out += `${sentence.slice(at, c.end)}<c${i + 1}>`;
    at = c.end;
  });
  return out + sentence.slice(at);
}

// items: [{ ref, sentences }], sentences numbered from 1 per ref.
export function requirementPrompt(items) {
  return [
    'Return {"sentences":[{"ref":"R1.2","s":1,"type":"request"}, ...]} with every ref and sentence number exactly once.',
    'Types: "request" (asks the authors to do, change, add, explain or answer something; questions too), "other" (praise, summary, background, greeting).',
    'A sentence with marks <c1>, <c2>, ... may ask for more than one thing. For it give "cuts" instead of "type": one "split" or "keep" per mark in order, and "types": one type per part after the splits, e.g. {"ref":"R1.3","s":1,"cuts":["split"],"types":["request","request"]}. Split only where the parts ask for different things; keep when they are one request.',
    "Sentences:",
    ...items.flatMap((it) => it.sentences.map((t, i) => `[${it.ref} s${i + 1}] ${markCuts(t)}`)),
  ].join("\n");
}

// One label of a sentence: a type for a sentence without cut points; for
// one with k cut points, k decisions and one type per part. null otherwise.
function sentenceLabel(x, sentence) {
  const k = clauseCuts(sentence).length;
  if (!k) return REQUIREMENT_TYPES.includes(x.type) && !("cuts" in x) ? [[], [x.type]] : null;
  const { cuts, types } = x;
  if (!Array.isArray(cuts) || cuts.length !== k) return null;
  if (!cuts.every((c) => c === "split" || c === "keep")) return null;
  const parts = 1 + cuts.filter((c) => c === "split").length;
  if (!Array.isArray(types) || types.length !== parts) return null;
  if (!types.every((t) => REQUIREMENT_TYPES.includes(t))) return null;
  return [cuts, types];
}

// The request texts per ref as { ref: [texts] }, or null unless every asked
// sentence is there exactly once with a valid label and nothing else is.
function validSentenceLabels(json, items) {
  const list = json?.sentences;
  const total = items.reduce((n, it) => n + it.sentences.length, 0);
  if (!Array.isArray(list) || list.length !== total) return null;
  const labels = Object.fromEntries(items.map((it) => [it.ref, new Array(it.sentences.length)]));
  const sentences = Object.fromEntries(items.map((it) => [it.ref, it.sentences]));
  for (const x of list) {
    const slots = Object.hasOwn(labels, x?.ref) ? labels[x.ref] : null;
    const s = x?.s;
    if (!slots || !Number.isInteger(s) || s < 1 || s > slots.length || slots[s - 1]) return null;
    const label = sentenceLabel(x, sentences[x.ref][s - 1]);
    if (!label) return null;
    slots[s - 1] = label;
  }
  return Object.fromEntries(
    items.map((it) => [
      it.ref,
      it.sentences.flatMap((sentence, i) => {
        const [cuts, types] = labels[it.ref][i];
        return requirementParts(sentence, cuts).filter((_, n) => types[n] === "request");
      }),
    ]),
  );
}

// comments: the split result. A comment that already has requirements, has
// no sentence or shares its ref is left as it is. One one-shot labels the
// sentences of the rest; a request sentence, or a request part of a sentence
// the model split at a cut point, becomes a requirement as written. A failed
// call or a reply that does not label every sentence once leaves the
// requirements empty, with a warning; the comments are otherwise unchanged.
// Returns { comments, warnings, asked, input, requests }: input is the prompt
// (for a fingerprint), requests the count of requirements per ref.
export async function extractRequirements(comments, { oneshot } = {}) {
  const none = { comments, warnings: [], asked: false, input: null, requests: null };
  if (!oneshot) return none;
  // A doubled ref (named in a split warning) is left for the user to renumber.
  const once = (ref) => comments.filter((c) => c.ref === ref).length === 1;
  const items = comments
    .filter((c) => !(c.requirements ?? []).length && once(c.ref))
    .map((c) => ({ ref: c.ref, sentences: splitSentences(c.text) }))
    .filter((it) => it.sentences.length);
  if (!items.length) return none;
  const input = requirementPrompt(items);
  const r = await oneshot({ system: REQUIREMENT_SYSTEM, prompt: input });
  const failed = (warning) => ({ ...none, warnings: [warning], asked: true, input });
  if (!r.ok) return failed(`requirement extraction failed: ${r.error}; add them by hand`);
  const texts = validSentenceLabels(r.json, items);
  if (!texts) return failed(REQUIREMENT_WARNING);
  return {
    comments: comments.map((c) =>
      texts[c.ref] ? { ...c, requirements: numberRequirements([], texts[c.ref]) } : c,
    ),
    warnings: [],
    asked: true,
    input,
    requests: Object.fromEntries(items.map((it) => [it.ref, texts[it.ref].length])),
  };
}
