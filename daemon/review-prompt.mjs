// The one user turn sent to the drafting session for "propose changes for
// comment R1.3". Pure: the daemon logs which ledger items went in.

import { inManuscriptScope } from "./ledger-check.mjs";

export const REVIEW_RULES = [
  "Change the manuscript only with office_tracked_edit, one call per change. Every other document-writing tool is disabled in this turn.",
  "Make only the changes this comment asks for. Do not delete or add anything the reviewer did not request.",
  'Do not write the revision history into the manuscript (no "as suggested by the reviewer", no "in this revised version"). That belongs in the response letter only.',
  "Base every change and every rationale on facts in the manuscript or the review. If a rationale is an inference, say so.",
  "Write the rationale in {lang}: what = what you changed, why = why, addresses = which part of the comment it answers.",
  "If office_tracked_edit returns blocked_by, do not retry the same wording. Rephrase without the forbidden expression, or stop and explain.",
  "Follow every author instruction below. They are decisions the author already made.",
];

// The comment as every review prompt shows it: who wrote it, the section of
// the review it stands under, and its text between <<< and >>>. Shared with
// the proposal prompt (proposal.mjs).
export function commentBlock(comment) {
  return [
    `${comment.reviewer === 0 ? "Editor" : `Reviewer ${comment.reviewer}`}, comment ${comment.ref} (verbatim):`,
    // The split takes a section label ("Abstract", "Round 2 / Minor comments")
    // off its line; without it "this section" in the comment means nothing.
    ...(comment.section ? [`Section of the review this comment is under: ${comment.section}`] : []),
    "<<<",
    comment.text,
    ">>>",
  ];
}

// The author's rules for the manuscript: instructions and the forbid list
// (left out when includeForbid is false). lines, and the ids that went in.
export function ledgerBlock(ledger, { includeForbid = true } = {}) {
  const inScope = (ledger ?? []).filter(inManuscriptScope);
  const instructions = inScope.filter((i) => i.kind === "instruction");
  const forbidAll = inScope.filter((i) => i.kind === "forbid");
  const forbids = includeForbid ? forbidAll : [];
  const lines = ["Author instructions:"];
  if (instructions.length) for (const i of instructions) lines.push(`- [${i.id}] ${i.text}`);
  else lines.push("- (none)");
  lines.push("", "Forbidden (an edit containing these is refused before it is applied):");
  if (forbids.length) for (const i of forbids) lines.push(`- [${i.id}] ${i.text}`);
  else lines.push("- (none listed)");
  return {
    lines,
    instructionIds: instructions.map((i) => i.id),
    forbidIds: forbids.map((i) => i.id),
    omittedForbidIds: includeForbid ? [] : forbidAll.map((i) => i.id),
  };
}

export function buildReviewPrompt({
  comment,
  paragraphs = [],
  ledger = [],
  lang = "Korean",
  includeForbid = true,
}) {
  const rules = ledgerBlock(ledger, { includeForbid });
  const lines = [`[Review turn ${comment.ref}]`, ...commentBlock(comment), ""];
  // L5: the separate requests of this comment; every change names the ones it answers.
  const reqs = comment.requirements ?? [];
  if (reqs.length) {
    lines.push(
      "What this comment asks for (put the ids a change answers in rationale.requirement_ids):",
    );
    for (const r of reqs) lines.push(`- [${r.id}] ${r.text}`);
    lines.push(
      "Put in rationale.requirement_ids the id of every request this change handles. If it handles only part of a request, still put its id. Leave it empty only when the change answers none of them.",
    );
    lines.push("");
  }
  if (paragraphs.length) {
    lines.push("Paragraphs the author selected (id, index, text):");
    for (const p of paragraphs)
      lines.push(`- id=${p.id ?? "?"} index=${p.index ?? "?"}: ${p.text}`);
  } else {
    lines.push(
      "No paragraphs were selected. Locate the text with office_find or office_read_paragraphs first.",
    );
  }
  lines.push("", ...rules.lines);
  lines.push("", "Rules:");
  REVIEW_RULES.forEach((r, n) => lines.push(`${n + 1}. ${r.replace("{lang}", lang)}`));
  lines.push("", `Call office_tracked_edit with comment_ref "${comment.ref}".`);
  return {
    text: lines.join("\n"),
    instructionIds: rules.instructionIds,
    forbidIds: rules.forbidIds,
    omittedForbidIds: rules.omittedForbidIds,
    section: comment.section || null,
  };
}

export function formatPromptLog(ref, built, { paragraphs = 0, preview = false } = {}) {
  const ids = (a) => (a.length ? a.join(",") : "-");
  return (
    `[review] prompt ${ref} instructions=${ids(built.instructionIds)} forbid=${ids(built.forbidIds)}` +
    ` omitted_forbid=${ids(built.omittedForbidIds)} paragraphs=${paragraphs}` +
    `${built.section ? ` section=${JSON.stringify(built.section)}` : ""}${preview ? " (preview)" : ""}`
  );
}
