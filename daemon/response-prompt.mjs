// The one request that drafts the answer to a reviewer comment for the
// response letter. Pure: no I/O.
//
// The model writes prose only. It is told what was changed (the changes that
// are in the manuscript now, with their rationale) and what it must not
// claim. Location lines, line numbers, quotes and the author's reasons for
// declining are added by the tool (response-model.mjs), so the prompt gives
// the model none of them to copy. The same holds for a figure change the
// author adopted on the figures page: the model is told its label and what
// the author said changed; the figure line with the file name is the tool's.

import { inScope } from "./ledger-check.mjs";
import { CITABLE, adoptedFigures } from "./response-quotes.mjs";

export const RESPONSE_SYSTEM =
  'You draft one answer in an author\'s response letter to a journal reviewer. Reply with JSON only: {"response": "<text>"}.';

export const RESPONSE_RULES = [
  "Write in {lang}.",
  "Answer this comment only. Say what was changed and why, using only the changes and rationales listed below.",
  "Never claim a change that is not listed as in the manuscript.",
  "If no change is listed as in the manuscript, do not say that anything was revised.",
  "Do not give reasons for a declined request or a declined comment. The tool prints the author's reasons after your text.",
  "Do not write location lines, line numbers, section names or quotes of the manuscript. The tool adds them after your text.",
  "Use numbers only as they appear below.",
];

// The two rules about what may be claimed, for a comment with figure changes.
const FIGURE_RULES = new Map([
  [
    RESPONSE_RULES[2],
    "Never claim a change that is not listed as in the manuscript or as an adopted figure change.",
  ],
  [
    RESPONSE_RULES[3],
    "If no change is listed as in the manuscript and no figure change as adopted, do not say that anything was revised.",
  ],
]);
const FIGURE_RULE =
  "Say what changed in a figure or other file only as the list of adopted figure changes says it. Do not write file names. The tool adds the figure line after your text.";
const FIGURE_NOT_DONE = {
  proposed: "not decided by the author yet",
  rejected: "rejected by the author",
};

// Why a change that is not in the manuscript must not be reported.
const NOT_DONE = {
  pending: "not decided in Word yet",
  rejected: "rejected by the author",
  unlocated: "not found in the manuscript",
};

const q = (s) => `<<<${String(s ?? "")}>>>`;
function edit(ch) {
  if (ch.kind === "delete") return `deleted ${q(ch.find)}`;
  if (ch.kind === "insert") return `inserted ${q(ch.text)} after ${q(ch.find)}`;
  return `replaced ${q(ch.find)} with ${q(ch.text)}`;
}

// comment: a revision.json comment whose change states are current.
// profile: with defaults and the response language filled (effectiveProfile).
// ledger: revision.ledger. previous: an earlier answer to keep the facts of.
export function buildResponsePrompt({ comment, profile, ledger = [], previous = null }) {
  const L = profile.response_letter;
  const mine = ledger.filter((i) => inScope(i, "response"));
  const instructions = mine.filter((i) => i.kind === "instruction");
  const forbids = mine.filter((i) => i.kind === "forbid");
  const lines = [
    `[Response draft ${comment.ref}]`,
    `${comment.reviewer === 0 ? "Editor" : `Reviewer ${comment.reviewer}`}, comment ${comment.ref} (verbatim):`,
    ...(comment.section ? [`Section of the review this comment is under: ${comment.section}`] : []),
    "<<<",
    comment.text,
    ">>>",
    "",
  ];

  const declined = new Set((comment.response?.declined_requirements ?? []).map((d) => d.id));
  const reqs = comment.requirements ?? [];
  if (reqs.length) {
    lines.push("What this comment asks for:");
    for (const r of reqs)
      lines.push(`- [${r.id}] ${r.text}${declined.has(r.id) ? " (declined by the author)" : ""}`);
    lines.push("");
  }
  if (comment.declined) lines.push("The author declined this comment as a whole.", "");

  const changes = comment.changes ?? [];
  const done = changes.filter((ch) => CITABLE.has(ch.state));
  lines.push(
    "Changes in the manuscript for this comment (accepted, or revised by the author; the only changes you may report):",
  );
  for (const ch of done) {
    const ids = ch.rationale?.requirement_ids ?? [];
    const answers = ids.length ? ` (answers ${ids.join(", ")})` : "";
    // A change the author rewrote by hand no longer reads as the tool wrote it.
    const what =
      ch.state === "user_edited" ? `revised by the author by hand at ${q(ch.find)}` : edit(ch);
    lines.push(`- [${ch.id}] ${what}${answers}`);
    if (ch.rationale) lines.push(`  What: ${ch.rationale.what} Why: ${ch.rationale.why}`);
  }
  if (!done.length) lines.push("- (none)");
  const notDone = changes.filter((ch) => !CITABLE.has(ch.state));
  if (notDone.length) {
    lines.push("", "Not in the manuscript (do not report any of these):");
    for (const ch of notDone)
      lines.push(
        `- [${ch.id}] ${NOT_DONE[ch.state] ?? "not in the manuscript"}; do not claim: ${edit(ch)}`,
      );
  }

  // Figure changes: only for a comment that has one, so every other prompt
  // reads as before.
  const figures = comment.figure_changes ?? [];
  const adopted = adoptedFigures(comment);
  const figureLine = (f) => {
    const ids = f.requirement_ids ?? [];
    // The label says what it is ("Figure 2", "Table S3").
    const what = f.kind === "new" ? "new" : "revised";
    return `[${f.id}] ${f.figure} (${what})${ids.length ? ` (answers ${ids.join(", ")})` : ""}`;
  };
  if (adopted.length) {
    lines.push(
      "",
      "Figure changes adopted for this comment (figures or other files replaced or added by the author outside the Word file; you may report them):",
    );
    for (const f of adopted) {
      lines.push(`- ${figureLine(f)}`);
      lines.push(`  What changed, in the words of the author: ${q(f.summary || "(not said)")}`);
    }
  }
  const open = figures.filter((f) => f.state !== "adopted");
  if (open.length) {
    lines.push("", "Figure changes that are not adopted (do not report any of these):");
    for (const f of open)
      lines.push(`- ${figureLine(f)}: ${FIGURE_NOT_DONE[f.state] ?? "not adopted"}`);
  }

  lines.push("", "Author instructions for the letter:");
  for (const i of instructions) lines.push(`- [${i.id}] ${i.text}`);
  for (const rule of L.rules ?? []) lines.push(`- ${rule}`);
  if (!instructions.length && !(L.rules ?? []).length) lines.push("- (none)");
  lines.push("", "Forbidden in the letter:");
  for (const i of forbids) lines.push(`- [${i.id}] ${i.text}`);
  if (!forbids.length) lines.push("- (none)");

  if (previous)
    lines.push(
      "",
      "Earlier draft (keep its facts unless a change above contradicts them):",
      "<<<",
      String(previous),
      ">>>",
    );

  lines.push("", "Rules:");
  const rules = [
    ...(figures.length
      ? [...RESPONSE_RULES.map((r) => FIGURE_RULES.get(r) ?? r), FIGURE_RULE]
      : RESPONSE_RULES),
    ...(L.max_words ? [`At most ${L.max_words} words.`] : []),
  ];
  rules.forEach((r, n) => lines.push(`${n + 1}. ${r.replace("{lang}", L.lang)}`));
  lines.push("", 'Reply with JSON only: {"response": "<text>"}');
  return {
    system: RESPONSE_SYSTEM,
    prompt: lines.join("\n"),
    changeIds: done.map((ch) => ch.id),
    figureIds: adopted.map((f) => f.id),
    instructionIds: instructions.map((i) => i.id),
    forbidIds: forbids.map((i) => i.id),
  };
}

// r: the result of a one-shot request. { text } or { error }.
export function parseDraft(r) {
  if (!r?.ok) return { error: r?.error ?? "The model gave no answer." };
  const text = typeof r.json?.response === "string" ? r.json.response.trim() : "";
  return text ? { text } : { error: "The model returned no response text." };
}

// Translation of an answer for the gloss paragraph: plain text, no JSON.
export const buildGlossPrompt = (text, lang) => ({
  system: `Translate into ${lang}. Reply with the translation only.`,
  prompt: text,
  expectJson: false,
});
