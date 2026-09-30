// Split a pasted review into reviewer comments with refs R{reviewer}.{n}.
// Regex first (no model call); a single LLM one-shot only when the regex
// finds at most one comment. The user edits the result in the task pane.
// The one-shot recovers document structure: it labels each numbered line of
// the review and returns no text, and the code builds the comments from the
// source lines, so no line is lost or rewritten.
// L5: every comment carries requirements [{ id, text }], the separate things
// the reviewer asks for. Both splits leave them empty for the user to fill.

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

const preservedWarnings = (comments) =>
  comments.map((c) =>
    c.reviewer === 0
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
      open({ reviewer: 0, editor: true, number: ++editors, lines: [line] });
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
    ...preservedWarnings(rows.filter((x) => x.kept).map((x) => x.comment)),
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
