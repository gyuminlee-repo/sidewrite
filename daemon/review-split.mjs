// Split a pasted review into reviewer comments with refs R{reviewer}.{n}.
// Regex first (no model call); a single LLM one-shot only when the regex
// finds at most one comment. The user edits the result in the task pane.
// L5: every comment carries requirements [{ id, text }], the separate things
// the reviewer asks for. The regex split leaves them empty for the user to
// fill; the LLM split asks the same one-shot for them.

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

export function splitReviewRegex(text, patterns = DEFAULT_PATTERNS) {
  const p = { ...DEFAULT_PATTERNS, ...patterns };
  const reReviewer = new RegExp(p.reviewerHead, "i");
  const reComment = new RegExp(p.commentHead, "i");
  const reList = new RegExp(p.listItem);
  const lines = String(text ?? "")
    .replace(/\r\n?/g, "\n")
    .split("\n");

  const blocks = [];
  let cur = { reviewer: 1, lines: [] };
  for (const line of lines) {
    const m = reReviewer.exec(line);
    if (m) {
      blocks.push(cur);
      cur = { reviewer: Number(m[1]), lines: [] };
    } else {
      cur.lines.push(line);
    }
  }
  blocks.push(cur);

  const out = [];
  for (const b of blocks) {
    // Explicit "Comment n" heads win; then numbered-list items are body text.
    const head = b.lines.some((l) => reComment.test(l)) ? reComment : reList;
    let item = null;
    let seq = 0;
    for (const line of b.lines) {
      const m = head.exec(line);
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
      }
    }
    if (item) out.push(item);
  }
  return out
    .map((c) => ({
      ref: `R${c.reviewer}.${c.number}`,
      reviewer: c.reviewer,
      text: c.lines.join("\n").replace(HEAD_TRIM, "").trim(),
      requirements: [],
    }))
    .filter((c) => c.text);
}

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

export const SPLIT_SYSTEM =
  "You split peer-review text into individual reviewer comments. Copy the text of each comment verbatim from the input. Do not summarize, merge, or invent comments. For each comment also list the separate things the reviewer asks the authors to do. Reply with JSON only.";

export const MAX_REQUIREMENTS = 10;

// LLM requirement list -> [{ id: q1.., text }]: strings only, trimmed,
// at most MAX_REQUIREMENTS of at most 300 characters, no duplicates.
export function requirementsFrom(list) {
  const out = [];
  for (const r of Array.isArray(list) ? list : []) {
    const text = String(typeof r === "string" ? r : "")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 300);
    if (!text || out.some((x) => x.text === text)) continue;
    out.push({ id: `q${out.length + 1}`, text });
    if (out.length === MAX_REQUIREMENTS) break;
  }
  return out;
}

export function splitPrompt(text) {
  return [
    'Return {"comments":[{"reviewer":<int>,"number":<int>,"text":"<verbatim comment text>","requirements":["<one thing the reviewer asks for>"]}]}.',
    "requirements: one short item per separate request in the comment (a comment asking for two changes has two items).",
    "Number comments per reviewer from 1 when the review does not number them. Use reviewer 1 when reviewers are not named.",
    "Review text:",
    "<<<",
    String(text ?? ""),
    ">>>",
  ].join("\n");
}

// oneshot: (req) => Promise<{ ok, json, error }> (daemon/oneshot.mjs).
export async function splitReview(text, { patterns, oneshot } = {}) {
  const regexOut = splitReviewRegex(text, patterns);
  if (regexOut.length > 1) {
    return { method: "regex", comments: regexOut, warnings: duplicateWarnings(regexOut) };
  }
  if (!oneshot) {
    return {
      method: "regex",
      comments: regexOut,
      warnings: ["regex found at most one comment; no LLM fallback configured"],
    };
  }
  const r = await oneshot({ system: SPLIT_SYSTEM, prompt: splitPrompt(text) });
  if (!r.ok) {
    return { method: "llm", comments: regexOut, warnings: [`LLM split failed: ${r.error}`] };
  }
  const items = Array.isArray(r.json?.comments) ? r.json.comments : [];
  const source = squash(text);
  const kept = [];
  const warnings = [];
  items.forEach((c, i) => {
    const body = String(c?.text ?? "").trim();
    if (!body) return;
    if (!source.includes(squash(body))) {
      warnings.push(`dropped LLM item ${i + 1}: its text is not in the review verbatim`);
      return;
    }
    const reviewer = Number.isInteger(c?.reviewer) && c.reviewer > 0 ? c.reviewer : 1;
    const number = Number.isInteger(c?.number) && c.number > 0 ? c.number : i + 1;
    kept.push({
      ref: `R${reviewer}.${number}`,
      reviewer,
      text: body,
      requirements: requirementsFrom(c?.requirements),
    });
  });
  return { method: "llm", comments: kept, warnings: [...warnings, ...duplicateWarnings(kept)] };
}
