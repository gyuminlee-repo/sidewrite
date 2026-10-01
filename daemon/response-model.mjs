// The response letter as a list of paragraphs, built from revision.json and
// the journal profile. The docx (tools/response_docx.py) and the plain text
// are both written from this one model. Pure: no I/O.
//
//   { paragraphs: [{ kind, text, color, ref, reviewer, removed? }] }
//   kind: reviewer_head | section_head | comment_head | comment | response |
//         gloss | declined | location | quote
//   color: RRGGBB or null (the reviewer's own words carry no colour)
//
// Only the `response` and `gloss` text is prose a person or the model wrote.
// Headings, location lines, quotes and line ranges are computed.

import { locationLine, renderFrame } from "./response-lines.mjs";

const linesOf = (s) =>
  String(s ?? "")
    .split(/\r?\n/)
    .filter((t) => t.trim());

const squash = (s) =>
  String(s ?? "")
    .replace(/\s+/g, " ")
    .trim();

// The prose the letter prints as the answer to one comment, a paragraph each,
// in the order printed: the lines of the answer text, then the author's
// reasons (they are part of the answer): the reason for declining the
// comment, then one for each declined request, in the order of the requests.
// A reason the answer text already gives word for word is not repeated; a
// reason of a request that is no longer listed is not written.
//
// The letter (buildLetterModel) and the checks of the letter's rules
// (response-check.mjs, the LLM rules in response-service.mjs) all read this
// one list, so no answer paragraph is printed that was not checked.
//
// The other prose the letter prints for a comment is the translation
// paragraph (glossLine below). It is in another language, so only the rules
// that do not depend on the language are read in it: the forbidden
// expressions (regex) and the other-reviewer rule. A phrase an answer must
// have, the position rule, the word cap, the opening and the LLM rules are
// about the answer in the letter's language and are not asked of it.
export function answerProse(c) {
  const r = c.response;
  const out = linesOf(r?.text);
  const said = [squash(r?.text)];
  const byRequest = new Map((r?.declined_requirements ?? []).map((d) => [d.id, d.reason]));
  const reasons = [c.declined?.reason, ...(c.requirements ?? []).map((q) => byRequest.get(q.id))];
  for (const reason of reasons.map(squash)) {
    if (!reason || said.some((s) => s.includes(reason))) continue;
    said.push(reason);
    out.push(reason);
  }
  return out;
}

// The translation paragraph of a comment as the letter prints it (one line,
// the profile's prefix in front), or "" when the answer has none. L:
// profile.response_letter. Whether it is printed is the caller's question
// (gloss_enabled).
export const glossLine = (c, L) =>
  c.response?.gloss
    ? `${L.gloss_prefix ?? ""}${c.response.gloss}`.replace(/\s*\r?\n\s*/g, " ")
    : "";

// R1.12 -> "12", E.2 -> "2".
const commentNumber = (ref) => String(ref).slice(String(ref).lastIndexOf(".") + 1);

// profile: with defaults filled (journal-profile.mjs withDefaults or
// effectiveProfile). gloss: also write the translation paragraphs.
export function buildLetterModel(rev, profile, { gloss = false } = {}) {
  const L = profile.response_letter;
  const body = profile.colors.response_body ?? null;
  const revised = profile.colors.response_revised ?? null;
  const paragraphs = [];
  // Editor (reviewer 0) first, then the reviewers; the stored order is kept
  // inside each of them (sort is stable).
  const comments = [...(rev.comments ?? [])].sort((a, b) => a.reviewer - b.reviewer);
  let reviewer = null;
  let section = null;
  for (const c of comments) {
    const add = (kind, text, color, extra = {}) =>
      paragraphs.push({ kind, text, color, ref: c.ref, reviewer: c.reviewer, ...extra });
    if (c.reviewer !== reviewer) {
      reviewer = c.reviewer;
      section = null;
      paragraphs.push({
        kind: "reviewer_head",
        text: c.reviewer === 0 ? L.editor_label : renderFrame(L.reviewer_label, { n: c.reviewer }),
        color: null,
        ref: null,
        reviewer: c.reviewer,
      });
    }
    // The name of a group of comments in the review ("Major comments"). It is
    // not the manuscript section of a citation.
    const group = c.section ? String(c.section) : null;
    if (group && group !== section && L.section_head)
      add("section_head", renderFrame(L.section_head, { section: group }), null);
    section = group;

    add("comment_head", renderFrame(L.comment_label, { n: commentNumber(c.ref) }), null);
    for (const t of linesOf(c.text)) add("comment", t, null);

    const r = c.response;
    const citations = (r?.citations ?? []).filter((cit) => !cit.problem);
    for (const t of answerProse(c)) add("response", t, body);
    const translation = gloss ? glossLine(c, L) : "";
    if (translation) add("gloss", translation, profile.colors.response_kr ?? null);
    if (c.declined && !citations.length) add("declined", L.declined_line, body);
    for (const cit of citations) {
      const removed = cit.kind === "removed";
      add("location", locationLine(cit, L), revised, { removed });
      if (!removed) add("quote", cit.quote, revised);
    }
  }
  return { paragraphs };
}

// The answers to one reviewer (0: the editor) as plain text for a portal
// text box: the reviewer heading is left out, one paragraph per line.
export function letterPlainText(model, reviewer) {
  return model.paragraphs
    .filter((p) => p.reviewer === reviewer && p.kind !== "reviewer_head")
    .map((p) => p.text)
    .join("\n");
}
