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

import { fmtLines, renderFrame } from "./response-lines.mjs";

const linesOf = (s) =>
  String(s ?? "")
    .split(/\r?\n/)
    .filter((t) => t.trim());

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
    const answer = linesOf(r?.text);
    for (const t of answer.length ? answer : linesOf(c.declined?.reason)) add("response", t, body);
    if (gloss && r?.gloss)
      add(
        "gloss",
        `${L.gloss_prefix ?? ""}${r.gloss}`.replace(/\s*\r?\n\s*/g, " "),
        profile.colors.response_kr ?? null,
      );
    if (c.declined && !citations.length) add("declined", L.declined_line, body);
    for (const cit of citations) {
      const removed = cit.kind === "removed";
      add(
        "location",
        renderFrame(removed ? L.removed_line : L.location_line, {
          section: cit.section,
          lines: fmtLines(cit.lines, L),
        }),
        revised,
        { removed },
      );
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
