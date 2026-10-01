// What replay.mjs computes without Word: its arguments, the ledger file made
// from the author's instruction lines, and the reports written at the end
// (changes.json, responses.json, summary.txt) from revision.json, the
// headings of the document and the daemon log. Pure functions.

import { sectionOf } from "../../daemon/response-quotes.mjs";
import { requirementStates } from "../../taskpane/shared/review-view.js";
import { stillPresent } from "../../taskpane/shared/residue.js";

const REF = /^(?:R\d+|E)\.\d+$/;

// "--only R1.2,R2.1" -> ["R1.2", "R2.1"]; nothing given -> [].
export function parseOnly(spec) {
  const refs = String(spec ?? "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
  for (const r of refs)
    if (!REF.test(r)) throw new Error(`--only takes refs like R1.2 or E.1, not ${r}`);
  return refs;
}

// The draft rows the pane keeps (pane-driver.cjs edit-draft keep_ref): the
// refs of --only, or every reviewer and editor comment.
export const keepRefPattern = (only) =>
  only.length ? `^(?:${only.map((r) => r.replace(".", "\\.")).join("|")})$` : REF.source;

// The lines of a text file that say something: no empty line, no "#" line.
export const ruleLines = (text) =>
  String(text ?? "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l && !l.startsWith("#"));

// The ledger file the pane step ledger-add reads: one rule per line, checked
// by the model (no pattern is guessed here).
export function ledgerFromLines({ instructions = "", forbid = "" } = {}) {
  const item = (kind) => (text) => ({
    kind,
    text,
    scope: "manuscript",
    check: { type: "llm", pattern: "", flags: "", expect: "absent" },
  });
  return {
    items: [
      ...ruleLines(instructions).map(item("instruction")),
      ...ruleLines(forbid).map(item("forbid")),
    ],
  };
}

// The model of the agent session and the end of the review turn, read from
// what the daemon logged while the turn ran.
//   [agent] init session <id> (model: <model>)
//   [review] turn done <ref>            (or: ... failed: <error>)
export function turnFromLog(log, ref) {
  const text = String(log ?? "");
  const models = [...text.matchAll(/^\[agent\] init session \S+ \(model: ([^)]+)\)/gm)];
  const done = new RegExp(
    `^\\[review\\] turn done ${ref.replace(".", "\\.")}( failed: (.*))?$`,
    "m",
  ).exec(text);
  return {
    model: models.length ? models[models.length - 1][1] : null,
    ended: !!done,
    ok: !!done && !done[1],
    error: done?.[2] ?? null,
  };
}

// The edits a guard of the daemon refused while a turn ran, by code, read
// from the lines review-tools.mjs logs (they hold counts and ids, no text):
//   [review] number needs author <ref> count=<n>   NumberNeedsAuthor
//   [review] blocked <ref> by=<ledger ids>         ForbidBlocked
//   [review] refused <ref>: <reason>               Refused
const REFUSALS = [
  ["NumberNeedsAuthor", "number needs author"],
  ["ForbidBlocked", "blocked"],
  ["Refused", "refused"],
];
export function refusalsFromLog(log, ref) {
  const at = ref.replace(".", "\\.");
  const out = {};
  for (const [code, words] of REFUSALS) {
    const n = (String(log ?? "").match(new RegExp(`^\\[review\\] ${words} ${at}[ :]`, "gm")) ?? [])
      .length;
    if (n) out[code] = n;
  }
  return out;
}

// How often the model was replaced after a refusal while a turn ran
// (daemon/model-list.mjs refusalFallback):
//   [agent] model refusal fallback original=<m> fallback=<m> scope=<s> category=<c>
export const fallbacksFromLog = (log) =>
  [
    ...String(log ?? "").matchAll(
      /^\[agent\] model refusal fallback original=(\S+) fallback=(\S+)/gm,
    ),
  ].map((m) => ({ original: m[1], fallback: m[2] }));

// Did the edits stay on another model than the one asked for? A fallback
// after a refusal that is gone by the next turn is not that; two turns in a
// row on another model are. turns: the turn results in the order they ran.
export const editsStayedOff = (turns) =>
  turns.some((t, i) => i > 0 && t.other_model && turns[i - 1].other_model);

// One comment of the run: { step, log, seconds } of its propose step.
// outcome: "ok", "failed" (the turn ended with an error: limit, refusal,
// crash), "timeout" (the pane step gave up), "no_turn" (no turn was seen).
// line: the pane's model line right after the turn (readModelLine).
export function turnResult(ref, { step, log, seconds, model = null, line = null }) {
  const t = turnFromLog(log, ref);
  let outcome = "ok";
  let reason = null;
  if (step?.ok === false) {
    outcome = /still running after/.test(step.error ?? "") ? "timeout" : "failed";
    reason = step.error ?? null;
  }
  if (t.ended && !t.ok) {
    outcome = "failed";
    reason = t.error;
  } else if (!t.ended && outcome === "ok") {
    outcome = "no_turn";
    reason = "the daemon logged no end of a review turn for this comment";
  }
  const edits = line === null ? null : (readModelLine(line).edits ?? null);
  return {
    outcome,
    reason,
    seconds,
    model: t.model ?? model,
    model_line: edits?.text ?? null,
    other_model: edits?.other ?? false,
    fallbacks: fallbacksFromLog(log),
    refused: refusalsFromLog(log, ref),
  };
}

// The line of the pane that says what each kind of work last ran on
// (taskpane/shared/model-line.js modelLineText), read apart:
//   "Edits: Claude · Opus   Writing: asked for X, ran on Y   Judgments: Z (not used yet)"
// -> { edits: { text, used, other, failed }, writing: ..., judgments: ... }
// used: a request of that kind was sent; other: it ran on another model than
// the one asked for; failed: its last request failed.
export function readModelLine(line) {
  const out = {};
  for (const part of String(line ?? "").split(/\s{3,}/)) {
    const m = /^(Edits|Writing|Judgments): (.*)$/.exec(part.trim());
    if (!m) continue;
    out[m[1].toLowerCase()] = {
      text: m[2],
      used: !/\(not used yet\)$/.test(m[2]),
      other: /^asked for .+, ran on /.test(m[2]),
      failed: /\(last request failed\)$/.test(m[2]),
    };
  }
  return out;
}

// The kinds of work that ran on another model than the one asked for.
export const otherModels = (line) =>
  Object.entries(readModelLine(line))
    .filter(([, v]) => v.other)
    .map(([k, v]) => `${k}: ${v.text}`);

// changes.json: per comment its requests with the state of each (changed,
// needs_author, outside_this_file, open: review-view.js requirementStates),
// the changes that were recorded, the questions put to the author, the
// requests declared to lie outside this file, the terms the changes took out
// that other paragraphs still hold (residue.js), and what the guards refused.
//   revision: revision.json as the daemon saved it
//   headings: [{ index, level, text }] of the document after the run
//   turns: { <ref>: turnResult(...) }
// The heading of a change is the one above its paragraph in the document as
// it is after the run; a change whose paragraph index is unknown has none.
export function changesReport(revision, headings = [], turns = {}) {
  return (revision?.comments ?? []).map((c) => {
    const changes = (c.changes ?? []).map((x) => ({
      change_id: x.id,
      paragraph_index: x.paragraph_index,
      heading: Number.isInteger(x.paragraph_index) ? sectionOf(x.paragraph_index, headings) : null,
      kind: x.kind,
      before: x.before,
      after: x.after,
      what: x.rationale?.what ?? null,
      why: x.rationale?.why ?? null,
      addresses: x.rationale?.addresses ?? null,
      requirement_ids: x.rationale?.requirement_ids ?? [],
      state: x.state,
    }));
    const requirements = requirementStates(c);
    return {
      ref: c.ref,
      reviewer: c.reviewer,
      section: c.section ?? null,
      comment: c.text,
      requirements,
      // The requests no change names (whatever else was said about them).
      unmet_requirements: requirements.filter((q) => q.state !== "changed").map((q) => q.id),
      questions: (c.questions ?? []).map((q) => ({
        id: q.id,
        requirement_id: q.requirement_id ?? null,
        question: q.question,
        why: q.why ?? "",
        answer: q.answer ?? "",
      })),
      outside: (c.requirements_outside ?? []).map((n) => ({
        requirement_id: n.id,
        why: n.why ?? "",
      })),
      still_present: stillPresent(c, revision?.snapshot?.paragraphs),
      blocked_edits: (c.blocked ?? []).length,
      refused: turns[c.ref]?.refused ?? {},
      turn: turns[c.ref] ?? null,
      changes,
    };
  });
}

// responses.json: per comment the answer and what it cites.
export function responsesReport(revision) {
  return (revision?.comments ?? [])
    .filter((c) => c.response)
    .map((c) => ({
      ref: c.ref,
      text: c.response.text,
      source: c.response.source,
      citations: (c.response.citations ?? []).map((x) => ({
        change_ids: x.change_ids,
        kind: x.kind,
        section: x.section,
        paragraph_index: x.paragraph_index,
        quote: x.quote,
        problem: x.problem,
      })),
    }));
}

// summary.txt: counts and times only, no text of the manuscript.
//   models: { asked: { edits, writing, judgments }, line } (line: the pane's
//   own line at the end of the run); source: how the manuscript was opened;
//   paragraphCheck: revision.consistency after "Check revised paragraphs".
export function summaryText(
  report,
  { seconds, saved, models, source = null, responses = null, paragraphCheck = null },
) {
  const turns = report.filter((c) => c.turn);
  const bad = turns.filter((c) => c.turn.outcome !== "ok");
  const asked = models?.asked ?? {};
  const other = otherModels(models?.line).filter((x) => !x.startsWith("edits: "));
  const drifted = turns.filter((c) => c.turn.other_model);
  const stayedOff = editsStayedOff(turns.map((c) => c.turn));
  const sum = (f) => report.reduce((n, c) => n + f(c), 0);
  const requests = report.flatMap((c) => c.requirements);
  const state = (name) => requests.filter((q) => q.state === name).length;
  const codes = {};
  for (const c of report)
    for (const [code, n] of Object.entries(c.refused ?? {})) codes[code] = (codes[code] ?? 0) + n;
  const fallbacks = turns.filter((c) => c.turn.fallbacks?.length);
  const withRequests = report.filter((c) => c.requirements.length).length;
  const lines = [
    `comments: ${report.length}`,
    `turns run: ${turns.length}`,
    `changes: ${sum((c) => c.changes.length)}`,
    `comments without a change: ${report.filter((c) => !c.changes.length).length}`,
    `comments with extracted requests: ${withRequests}; without: ${report.length - withRequests}`,
    `requests by state: changed ${state("changed")}, needs_author ${state("needs_author")}, outside_this_file ${state("outside_this_file")}, open ${state("open")}`,
    `requirements no change names: ${
      requests.length
        ? sum((c) => c.unmet_requirements.length)
        : "not measured (no requests were extracted)"
    }`,
    `questions for the author: ${sum((c) => c.questions.length)}`,
    `requests declared outside this file: ${sum((c) => c.outside.length)}`,
    `terms still present elsewhere: ${sum((c) => c.still_present.length)}`,
    `edits a guard refused: ${
      Object.entries(codes)
        .map(([code, n]) => `${code} ${n}`)
        .join(", ") || "none"
    }`,
    `turns that did not end well: ${bad.length}${bad.length ? ` (${bad.map((c) => `${c.ref} ${c.turn.outcome}`).join(", ")})` : ""}`,
    `models asked for: edits ${asked.edits ?? "-"}, writing ${asked.writing ?? "(the pane's own)"}, judgments ${asked.judgments ?? "(the pane's own)"}`,
    `models that ran, as the pane says: ${models?.line || "-"}`,
    `models that ran, from the daemon log: ${[...new Set(turns.map((c) => c.turn.model).filter(Boolean))].join(", ") || "-"}`,
    `model replaced after a refusal: ${sum((c) => c.turn?.fallbacks?.length ?? 0)} times in ${fallbacks.length} turns${fallbacks.length ? ` (${fallbacks.map((c) => c.ref).join(", ")})` : ""}`,
    `turns on another model than asked: ${drifted.length}${drifted.length ? ` (${drifted.map((c) => c.ref).join(", ")})` : ""}`,
    `another model than asked: ${
      other.length || stayedOff
        ? `FAILED (${[...other, ...(stayedOff ? ["edits stayed on another model for more than one turn"] : [])].join("; ")})`
        : "no"
    }`,
    `paragraph check: ${
      paragraphCheck
        ? `${paragraphCheck.findings?.length ?? 0} findings in ${paragraphCheck.paragraphs?.length ?? 0} paragraphs`
        : "not run"
    }`,
    `turn time: ${turns.reduce((n, c) => n + (c.turn.seconds ?? 0), 0)} s`,
    `total time: ${seconds} s`,
    `copy saved with its tracked changes: ${saved ? "yes" : "NO"}`,
  ];
  if (source) lines.push(`manuscript: ${source}`);
  if (responses !== null) lines.push(`responses drafted: ${responses}`);
  return `${lines.join("\n")}\n`;
}
