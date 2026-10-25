// What replay.mjs computes without Word: its arguments, the ledger file made
// from the author's instruction lines, and the reports written at the end
// (changes.json, responses.json, summary.txt) from revision.json, the
// headings of the document and the daemon log. Pure functions.

import { sectionOf } from "../../daemon/response-quotes.mjs";
import { requirementStates } from "../../taskpane/shared/review-view.js";
import { stillPresent } from "../../taskpane/shared/residue.js";
import { termPlaces } from "../../taskpane/shared/named-terms.js";

const escapeRe = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

const REF = /^(?:R\d+|E)\.\d+$/;

export function replayExitCode({ failed, stopped, saved, otherModel, cleanup = {} }) {
  if (failed || !saved || cleanup.failed || cleanup.daemon_error) return 1;
  return stopped ? 3 : otherModel ? 1 : 0;
}

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
  only.length ? `^(?:${only.map((r) => escapeRe(r)).join("|")})$` : REF.source;

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
  const done = new RegExp(`^\\[review\\] turn done ${escapeRe(ref)}( failed: (.*))?$`, "m").exec(
    text,
  );
  return {
    model: models.length ? models[models.length - 1][1] : null,
    ended: !!done,
    ok: !!done && !done[1],
    error: done?.[2] ?? null,
  };
}

// The paragraphs a turn wrote anew (review-tools.mjs applyTrackedRewrite):
//   [review] rewrite <ref> paragraph=<i> parts=<n> applied=<k>[ stopped=<code>]
// parts: the runs of words that differ; applied: those that went in.
export const rewritesFromLog = (log, ref) =>
  [
    ...String(log ?? "").matchAll(
      new RegExp(
        `^\\[review\\] rewrite ${escapeRe(ref)} paragraph=(\\d+) parts=(\\d+) applied=(\\d+)(?: stopped=(\\S+))?$`,
        "gm",
      ),
    ),
  ].map((m) => ({
    paragraph: Number(m[1]),
    parts: Number(m[2]),
    applied: Number(m[3]),
    stopped: m[4] ?? null,
  }));

// What a follow-through turn did (review-service.mjs followEnded):
//   [review] follow <ref> candidates=<n> changes=<n> questions=<n>
// candidates: the places it was given; changes and questions: what it added.
// null when the log of the turn has no such line.
export function followFromLog(log, ref) {
  const m = new RegExp(
    `^\\[review\\] follow ${escapeRe(ref)} candidates=(\\d+) changes=(\\d+) questions=(\\d+)$`,
    "m",
  ).exec(String(log ?? ""));
  return m ? { candidates: Number(m[1]), changes: Number(m[2]), questions: Number(m[3]) } : null;
}

// The edits a guard of the daemon refused while a turn ran, by code, read
// from the lines review-tools.mjs logs (they hold counts and ids, no text).
// The code is the one the tool result of that refusal carries; the block by
// a forbid rule has none there and is named ForbidBlocked here. One line is
// one refused edit or rewrite (the count= of a number line is not added up;
// the number and the broken-rule lines are the same for both tools).
//   [review] number needs author <ref> count=<n>          NumberNeedsAuthor
//   [review] blocked <ref> by=<ledger ids>                ForbidBlocked
//   [review] rewrite blocked <ref> by=<ledger ids>        RewriteBlocked
//   [review] refused <ref>: forbid rule(s) ...            ForbidRuleBroken
//   [review] refused <ref>: task pane without ...         TaskPaneOutdated
//   [review] refused <ref>: anything else                 Refused
//   [review] field edit refused <ref> why=<reason> ...    FieldNotEditable
//   [review] fields unread <ref> paragraph=<i> ...        FieldsUnread
//   [review] not a follow place <ref> paragraph=<i>       NotAFollowPlace
//   [review] whole author insertion <ref> paragraph=<i>   WholeAuthorInsertion
//   [review] rewrite <ref> ... applied=0 stopped=<code>   <code> (ParagraphNotPlain, ...)
//   [review] rewrite <ref> ... applied=<k> stopped=...    RewriteStopped
// A rewrite that cannot be written as tracked changes (RewriteNotExpressible)
// is refused without a log line and is not counted.
const REFUSALS = [
  ["NumberNeedsAuthor", "number needs author", " count="],
  ["ForbidBlocked", "blocked", " by="],
  ["RewriteBlocked", "rewrite blocked", " by="],
  ["ForbidRuleBroken", "refused", ": forbid rule"],
  ["TaskPaneOutdated", "refused", ": task pane without "],
  ["Refused", "refused", ": (?!forbid rule|task pane without )"],
  ["FieldNotEditable", "field edit refused", " why="],
  ["FieldsUnread", "fields unread", " paragraph="],
  ["NotAFollowPlace", "not a follow place", " paragraph="],
  ["WholeAuthorInsertion", "whole author insertion", " paragraph="],
];
export function refusalsFromLog(log, ref) {
  const at = escapeRe(ref);
  const out = {};
  for (const [code, words, next] of REFUSALS) {
    const n = (
      String(log ?? "").match(new RegExp(`^\\[review\\] ${words} ${at}${next}`, "gm")) ?? []
    ).length;
    if (n) out[code] = n;
  }
  for (const w of rewritesFromLog(log, ref)) {
    if (!w.stopped) continue;
    const code = w.applied || w.stopped === "refused" ? "RewriteStopped" : w.stopped;
    out[code] = (out[code] ?? 0) + 1;
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

// The chosen model declined twice in a row and the document stays on the
// other one (daemon/index.mjs):
//   [agent] model stays on <fallback>: <original> declined <n> turns in a row (category=<c>)
export const staysFromLog = (log) =>
  [
    ...String(log ?? "").matchAll(
      /^\[agent\] model stays on (\S+): (\S+) declined (\d+) turns in a row \(category=([^)]*)\)/gm,
    ),
  ].map((m) => ({
    fallback: m[1],
    original: m[2],
    count: Number(m[3]),
    category: m[4] === "-" ? null : m[4],
  }));

// How often a turn started on the chosen model again after a fallback:
//   [agent] model restored to <model> after the fallback to <fallback>
export const restoredFromLog = (log) =>
  (String(log ?? "").match(/^\[agent\] model restored to \S+ after the fallback to /gm) ?? [])
    .length;

// The batches of the request extraction (daemon/review-split.mjs), read from
// the whole log of the run:
//   [review] requirement extract batch <i>/<m> comments=<k> sentences=<s> labelled=<a>/<k>[ unlabelled=<refs>]
//   [review] requirement extract batch <i>/<m> comments=<k> sentences=<s> failed: <error>
export const extractBatchesFromLog = (log) =>
  [
    ...String(log ?? "").matchAll(
      /^\[review\] requirement extract batch (\d+)\/(\d+) comments=(\d+) sentences=(\d+)(?: failed: (.*)| labelled=(\d+)\/\d+(?: unlabelled=(\S+))?)$/gm,
    ),
  ].map((m) => ({
    batch: Number(m[1]),
    of: Number(m[2]),
    comments: Number(m[3]),
    sentences: Number(m[4]),
    ok: m[5] === undefined,
    error: m[5] ?? null,
    labelled: m[6] === undefined ? null : Number(m[6]),
    unlabelled: m[7] ? m[7].split(",") : [],
  }));

// --answers: what the author would answer when asked, per comment:
// { "<ref>": "<answer text>", ... } (the file as JSON.parse read it). Every
// open question of a comment gets the answer of its ref. Returns the same
// with the texts trimmed; throws on anything else. What it says names a
// valid ref at most: an answer is text of the author and is never printed,
// and a key that is no ref may be one (key and value the wrong way round).
export function parseAnswers(json) {
  const rows =
    json && typeof json === "object" && !Array.isArray(json) ? Object.entries(json) : null;
  if (!rows?.length)
    throw new Error('--answers takes a JSON object { "<ref>": "<answer text>", ... }');
  const out = {};
  for (const [n, [ref, text]] of rows.entries()) {
    if (!REF.test(ref))
      throw new Error(
        `--answers: entry ${n + 1} of ${rows.length} has no comment ref like R1.3 as its key`,
      );
    if (typeof text !== "string" || !text.trim())
      throw new Error(`--answers: the answer of ${ref} is not a text`);
    out[ref] = text.trim();
  }
  return out;
}

// Which comments get their answer. revision: revision.json after the turns
// of the comments; answers: parseAnswers(...); refs: the comments of the
// run, in its order.
// { due: [{ ref, open: [question ids] }], skipped: [{ ref, reason }] }
export function answerPlan(revision, answers, refs = []) {
  const due = [];
  const skipped = [];
  for (const ref of refs) {
    if (!(ref in answers)) continue;
    const c = (revision?.comments ?? []).find((x) => x.ref === ref);
    const open = (c?.questions ?? []).filter((q) => !q.answer).map((q) => q.id);
    if (open.length) due.push({ ref, open });
    else skipped.push({ ref, reason: "no open question" });
  }
  for (const ref of Object.keys(answers))
    if (!refs.includes(ref)) skipped.push({ ref, reason: "not a comment of this run" });
  return { due, skipped };
}

// How many of the questions that were open (answerPlan: open) the review
// file now holds an answer to: what the pane step saved, read from the file
// and not from what the step says it did.
export function answeredOf(revision, ref, open = []) {
  const now = (revision?.comments ?? []).find((c) => c.ref === ref)?.questions ?? [];
  return open.filter((id) => now.find((q) => q.id === id)?.answer).length;
}

// replies.json: the answer of the model in one turn, for reading a run
// afterwards. kind: the turn of a comment ("ordinary"), its turn after the
// answers of --answers ("answer"), "follow-through" or "rewrite". step: what
// the pane step of the turn gave; its `reply` is the last answer in the chat
// when the turn put one there (pane-driver.cjs readReply), else there is none.
export const REPLY_KINDS = ["ordinary", "answer", "follow-through", "rewrite"];
export function replyRow(ref, kind, step) {
  if (!REPLY_KINDS.includes(kind)) throw new Error(`unknown kind of turn: ${kind}`);
  return { ref, kind, text: typeof step?.reply === "string" ? step.reply : "" };
}

// Whether a turn is run once more. attempts: the turnResult of each try so
// far. Only after the first try, and only when the turn ended on an API
// error (daemon/limit-detect.mjs: an authentication or server answer in
// place of the model's; a replay lost seven turns in a row to one). A turn
// the model declined, one that ran out of time and one no end was seen of
// are not tried again: the same request would end the same way, or a turn
// may still be running.
const API_ERROR = /^The turn ended on an API error/;
const USAGE_LIMIT = /^Claude usage limit reached\.?(?: resets at (\S+))?/;
export const retryTurn = (attempts) =>
  attempts.length === 1 &&
  attempts[0].outcome === "failed" &&
  API_ERROR.test(attempts[0].reason ?? "");

// Whether a turn ends the run: it failed on the usage limit of the account
// (daemon/limit-detect.mjs: "Claude usage limit reached." and, when the
// daemon knows it, " resets at <time>"). Every turn after it would end the
// same way; a replay went on for an hour like that. turn: its turnResult.
// { ref, kind, resets } (resets: the time as the daemon gave it, or null),
// or null for any other end: an API error is run again (retryTurn), a
// refusal and a turn out of time say nothing about the next comment.
export function limitStop(ref, kind, turn) {
  const m = turn?.outcome === "failed" ? USAGE_LIMIT.exec(turn.reason ?? "") : null;
  return m ? { ref, kind, resets: m[1] ?? null } : null;
}

// The final outcome and plan, with the time and recorded work of every try.
export function mergeAttempts(attempts) {
  if (attempts.length < 2) return attempts[0];
  const merged = {
    ...attempts.at(-1),
    seconds: attempts.reduce((n, t) => n + (t.seconds ?? 0), 0),
    attempts: attempts.map((t) => ({
      outcome: t.outcome,
      reason: t.reason ?? null,
      seconds: t.seconds ?? 0,
    })),
  };
  for (const field of ["fallbacks", "stays", "rewrites"]) {
    if (attempts.some((t) => t[field])) merged[field] = attempts.flatMap((t) => t[field] ?? []);
  }
  if (attempts.some((t) => t.refused)) {
    merged.refused = {};
    for (const t of attempts) {
      for (const [code, count] of Object.entries(t.refused ?? {})) {
        merged.refused[code] = (merged.refused[code] ?? 0) + count;
      }
    }
  }
  const follows = attempts.map((t) => t.follow).filter(Boolean);
  if (follows.length) {
    merged.follow = {
      ...follows.at(-1),
      changes: follows.reduce((n, f) => n + f.changes, 0),
      questions: follows.reduce((n, f) => n + f.questions, 0),
    };
  }
  return merged;
}

// The fixed class of an error text, or null when it is none of these. Only
// the class goes into summary.txt: the text itself may quote the manuscript
// or the answer of the model (daemon/oneshot.mjs puts both in its errors).
export function errorClass(text) {
  const s = String(text ?? "");
  if (USAGE_LIMIT.test(s)) return "usage stop";
  if (API_ERROR.test(s)) return "api error";
  if (/^The model declined/.test(s)) return "refusal";
  if (/^timed out\b|still running after/.test(s)) return "timeout";
  return null;
}

// Why a turn did not end well, as one of a fixed set: timeout, no turn end,
// pane error, driver error, usage stop, api error, refusal, daemon error.
// reason_kind (turnResult) says where the reason came from.
export function reasonClass(turn) {
  if (turn?.outcome === "timeout") return "timeout";
  if (turn?.outcome === "no_turn") return "no turn end";
  if (turn?.reason_kind === "pane") return "pane error";
  if (turn?.reason_kind === "driver") return "driver error";
  return errorClass(turn?.reason) ?? "daemon error";
}

// One comment of the run: { step, log, seconds } of its propose step.
// outcome: "ok", "failed" (the turn ended with an error: limit, refusal,
// crash, or a pane error when the daemon logged no end of the turn),
// "timeout" (the pane step gave up), "no_turn" (no turn was seen).
// reason_kind: where the reason came from ("driver": the pane step,
// "pane": the message line of the Review tab, "daemon": the daemon log).
// pane_message: the error the message line showed. The line is shared by
// every request of the tab, so after a turn the daemon logged as ended
// well it does not fail the turn; it is kept here apart (changes.json).
// line: the pane's model line right after the turn (readModelLine).
export function turnResult(ref, { step, log, seconds, model = null, line = null }) {
  const t = turnFromLog(log, ref);
  let outcome = "ok";
  let reason = null;
  let reason_kind = null;
  const pane_message = step?.message_error === true && step.message ? step.message : null;
  if (step?.ok === false) {
    outcome = /still running after/.test(step.error ?? "") ? "timeout" : "failed";
    reason = step.error ?? null;
    reason_kind = "driver";
  } else if (pane_message && !(t.ended && t.ok)) {
    outcome = "failed";
    reason = pane_message;
    reason_kind = "pane";
  }
  if (t.ended && !t.ok) {
    outcome = "failed";
    reason = t.error;
    reason_kind = "daemon";
  } else if (!t.ended && outcome === "ok") {
    outcome = "no_turn";
    reason = "the daemon logged no end of a review turn for this comment";
    reason_kind = "daemon";
  }
  const edits = line === null ? null : (readModelLine(line).edits ?? null);
  const follow = followFromLog(log, ref);
  return {
    outcome,
    reason,
    reason_kind,
    ...(pane_message && outcome === "ok" ? { pane_message } : {}),
    seconds,
    model: t.model ?? model,
    model_line: edits?.text ?? null,
    other_model: edits?.other ?? false,
    declined: edits?.declined ?? null,
    fallbacks: fallbacksFromLog(log),
    stays: staysFromLog(log),
    rewrites: rewritesFromLog(log, ref),
    refused: refusalsFromLog(log, ref),
    // Only a follow-through turn has it.
    ...(follow ? { follow } : {}),
  };
}

// The line of the pane that says what each kind of work last ran on
// (taskpane/shared/model-line.js modelLineText), read apart:
//   "Edits: Claude · Opus   Writing: asked for X, ran on Y   Judgments: Z (not used yet)"
// -> { edits: { text, used, other, failed }, writing: ..., judgments: ... }
// used: a request of that kind was sent; other: it ran on another model than
// the one asked for and the line gives no reason; failed: its last request
// failed; declined: { on, original, category } when the line says the chosen
// model declines this manuscript and the document stays on another one
//   "Edits: Claude · Opus 5 (Opus 5.5 declines this manuscript: bio)"
const declinedOf = (text) => {
  const m = /^(.*) \((.+) declines this manuscript(?:: (.+))?\)$/.exec(text);
  return m ? { on: m[1], original: m[2], category: m[3] ?? null } : null;
};
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
      declined: declinedOf(m[2]),
    };
  }
  return out;
}

// The kinds of work that ran on another model than the one asked for,
// without a reason in the line.
export const otherModels = (line) =>
  Object.entries(readModelLine(line))
    .filter(([, v]) => v.other)
    .map(([k, v]) => `${k}: ${v.text}`);

// What fails the run: the line at the end names another model than the one
// asked for and gives no reason. A document that stays on another model
// because the chosen one declines it says so in the line and is no failure.
// Nor are the edits when the last turn of the run was itself the one the
// model was replaced in after a refusal (lastTurn: its turnResult): no turn
// came after it that could start on the chosen model again.
export const modelFailures = (line, lastTurn = null) =>
  otherModels(line).filter((x) => !(x.startsWith("edits: ") && lastTurn?.fallbacks?.length));

// changes.json: per comment its requests with the state of each (changed,
// needs_author, outside_this_file, open: review-view.js requirementStates),
// the changes that were recorded, the questions put to the author, the
// requests declared to lie outside this file, the terms the changes took out
// that other paragraphs still hold (residue.js), the terms the reviewer
// named with every paragraph that holds them now (named-terms.js), where the
// turn itself said the same term is left, the paragraphs written anew, what
// the paragraph check found on its changes, what the guards refused, and
// what the follow-through turn of the comment did (follow: the places it was
// given as paragraph indexes of the snapshot, the changes it added, the
// number of its questions; null when none ran to its end).
//   revision: revision.json as the daemon saved it
//   headings: [{ index, level, text }] of the document after the run
//   turns: { <ref>: turnResult(...) }
//   rewriteTurns: [{ ref, paragraph_index, ...turnResult(...) }] of
//     --rewrite-stacked
//   followTurns: [{ ref, ...turnResult(...) }] of --follow-through
//   answers: { answered: { <ref>: n }, turns: { <ref>: turnResult(...) } } of
//     --answers: the questions answered from the file and the turn after
//     them. A question says whether it has an answer (`answered`); the
//     answer itself is text of the author and stays in revision.json.
export function changesReport(
  revision,
  headings = [],
  turns = {},
  rewriteTurns = [],
  followTurns = [],
  answers = {},
) {
  const snapshot = revision?.snapshot ?? null;
  // A place of a follow-through turn, as a paragraph index: by the task pane
  // id, else by the index an id like "p12" names; left out when neither.
  const placeIndex = (id) => {
    const i = (snapshot?.ids ?? []).indexOf(id);
    if (i !== -1) return i;
    const m = /^p(\d+)$/.exec(id);
    return m ? Number(m[1]) : null;
  };
  const findings = revision?.consistency?.findings ?? [];
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
      // A part of a paragraph written anew: the id its parts share.
      rewrite_id: x.rewrite_id ?? null,
      // New words and quantities the review has and the manuscript did not.
      from_review: x.from_review ?? [],
    }));
    const requirements = requirementStates(c);
    const rewrites = [];
    for (const x of changes) {
      if (!x.rewrite_id) continue;
      let row = rewrites.find((w) => w.rewrite_id === x.rewrite_id);
      if (!row)
        rewrites.push(
          (row = { rewrite_id: x.rewrite_id, paragraph_index: x.paragraph_index, change_ids: [] }),
        );
      row.change_ids.push(x.change_id);
    }
    const own = new Set(changes.map((x) => x.change_id));
    return {
      ref: c.ref,
      reviewer: c.reviewer,
      section: c.section ?? null,
      comment: c.text,
      requirements,
      // The requests no change names (whatever else was said about them).
      // A row of the part of a request that is in another file is not the request.
      unmet_requirements: requirements
        .filter((q) => !q.part && q.state !== "changed")
        .map((q) => q.id),
      questions: (c.questions ?? []).map((q) => ({
        id: q.id,
        requirement_id: q.requirement_id ?? null,
        question: q.question,
        why: q.why ?? "",
        answered: !!q.answer,
      })),
      answered: answers.answered?.[c.ref] ?? 0,
      answer_turn: answers.turns?.[c.ref] ?? null,
      outside: (c.requirements_outside ?? []).map((n) => ({
        requirement_id: n.id,
        why: n.why ?? "",
        ...(n.part ? { part: n.part } : {}),
      })),
      still_present: stillPresent(c, snapshot?.paragraphs),
      // [{ term, places: [{ index, id, kind }] }]: no place, the term is gone.
      named_terms: termPlaces(c, snapshot?.paragraphs, snapshot?.ids ?? null),
      same_term_at: c.same_term_at ?? [],
      rewrites,
      paragraph_findings: findings
        .filter((f) => (f.change_ids ?? []).some((id) => own.has(id)))
        .map((f) => ({
          kind: f.kind,
          paragraph_index: f.paragraph_index,
          paragraph_id: f.paragraph_id ?? null,
          detail: f.detail,
          change_ids: f.change_ids,
        })),
      blocked_edits: (c.blocked ?? []).length,
      refused: turns[c.ref]?.refused ?? {},
      turn: turns[c.ref] ?? null,
      rewrite_turns: rewriteTurns.filter((t) => t.ref === c.ref),
      follow: c.follow_through
        ? {
            candidates: c.follow_through.candidate_ids.map(placeIndex).filter((i) => i !== null),
            changes: c.follow_through.change_ids,
            questions: c.follow_through.questions,
          }
        : null,
      follow_turns: followTurns.filter((t) => t.ref === c.ref),
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

// summary.txt: counts and times only, no text of the manuscript. Why a turn
// or a batch failed is written as a fixed class (reasonClass, errorClass);
// the text of the reason stays in changes.json and the daemon log.
//   models: { asked: { edits, writing, judgments }, line } (line: the pane's
//   own line at the end of the run); source: how the manuscript was opened;
//   paragraphCheck: revision.consistency after "Check revised paragraphs".
//   stopped: limitStop(...) of the turn the run stopped at, or null.
export function summaryText(
  report,
  {
    seconds,
    saved,
    models,
    source = null,
    responses = null,
    paragraphCheck = null,
    extraction = [],
    stays = [],
    restored = 0,
    lastTurn = null,
    stopped = null,
  },
) {
  const turns = report.filter((c) => c.turn);
  const rewriteTurns = report.flatMap((c) => c.rewrite_turns ?? []);
  const followTurns = report.flatMap((c) => c.follow_turns ?? []);
  const answerTurns = report
    .filter((c) => c.answer_turn)
    .map((c) => ({ ref: c.ref, ...c.answer_turn }));
  const allTurns = [
    ...turns.map((c) => ({ ...c.turn, label: c.ref })),
    ...answerTurns.map((t) => ({ ...t, label: `${t.ref} after answers` })),
    ...followTurns.map((t) => ({ ...t, label: `${t.ref} follow-through` })),
    ...rewriteTurns.map((t) => ({ ...t, label: `${t.ref} rewrite p${t.paragraph_index}` })),
  ];
  const bad = [
    ...turns.filter((c) => c.turn.outcome !== "ok").map((c) => `${c.ref} ${c.turn.outcome}`),
    ...answerTurns
      .filter((t) => t.outcome !== "ok")
      .map((t) => `${t.ref} after answers ${t.outcome}`),
    ...followTurns
      .filter((t) => t.outcome !== "ok")
      .map((t) => `${t.ref} follow-through ${t.outcome}`),
    ...rewriteTurns
      .filter((t) => t.outcome !== "ok")
      .map((t) => `${t.ref} rewrite p${t.paragraph_index} ${t.outcome}`),
  ];
  const again = [
    ...turns.filter((c) => c.turn.attempts).map((c) => c.ref),
    ...answerTurns.filter((t) => t.attempts).map((t) => `${t.ref} after answers`),
    ...followTurns.filter((t) => t.attempts).map((t) => `${t.ref} follow-through`),
  ];
  const paneErrors = allTurns.filter((t) => t.outcome === "ok" && t.pane_message);
  const asked = models?.asked ?? {};
  const failures = modelFailures(models?.line, lastTurn);
  const drifted = turns.filter((c) => c.turn.other_model);
  const sum = (f) => report.reduce((n, c) => n + f(c), 0);
  const requests = report.flatMap((c) => c.requirements);
  const state = (name) => requests.filter((q) => q.state === name).length;
  const codes = {};
  const refused = [
    ...report.map((c) => c.refused ?? {}),
    ...answerTurns.map((t) => t.refused ?? {}),
    ...followTurns.map((t) => t.refused ?? {}),
    ...rewriteTurns.map((t) => t.refused ?? {}),
  ];
  for (const r of refused)
    for (const [code, n] of Object.entries(r)) codes[code] = (codes[code] ?? 0) + n;
  const fallbacks = allTurns.filter((t) => t.fallbacks?.length);
  const withRequests = report.filter((c) => c.requirements.length).length;
  const named = report.flatMap((c) => c.named_terms ?? []);
  const rewrites = report.flatMap((c) => c.rewrites ?? []);
  const logged = allTurns.flatMap((t) => t.rewrites ?? []);
  const followWork = followTurns.reduce(
    (n, t) => ({
      changes: n.changes + (t.follow?.changes ?? 0),
      questions: n.questions + (t.follow?.questions ?? 0),
    }),
    { changes: 0, questions: 0 },
  );
  // The latest completed checks are durable state, not all work of this
  // run. Failed attempts can record edits/questions without saving a check.
  // Show their logged work separately; never add the same durable IDs twice.
  const followed = report.flatMap((c) =>
    c.changes.filter((x) => (c.follow?.changes ?? []).includes(x.change_id)),
  );
  const kinds = {};
  for (const f of paragraphCheck?.findings ?? []) kinds[f.kind] = (kinds[f.kind] ?? 0) + 1;
  const list = (o) =>
    Object.entries(o)
      .map(([k, n]) => `${k} ${n}`)
      .join(", ");
  const declined = readModelLine(models?.line).edits?.declined ?? null;
  const stay = declined
    ? `edits stayed on ${declined.on} because ${declined.original} declines this manuscript (${declined.category ?? "-"})`
    : stays.length
      ? `edits stayed on ${stays.at(-1).fallback} because ${stays.at(-1).original} declines this manuscript (${stays.at(-1).category ?? "-"})`
      : null;
  const failedBatches = extraction.filter((b) => !b.ok).length;
  const lines = [
    `comments: ${report.length}`,
    `turns run: ${turns.length}`,
    `changes: ${sum((c) => c.changes.length)}`,
    `comments without a change: ${report.filter((c) => !c.changes.length).length}`,
    extraction.length
      ? `request extraction: ${extraction.length} batches, ok ${extraction.length - failedBatches}, failed ${failedBatches}`
      : "request extraction: no batch in the log",
    ...extraction.map(
      (b) =>
        `  batch ${b.batch}/${b.of}: ${b.comments} comments, ${b.sentences} sentences, ${
          b.ok
            ? `labelled ${b.labelled}/${b.comments}${b.unlabelled.length ? `, unlabelled ${b.unlabelled.join(",")}` : ""}`
            : `FAILED: ${errorClass(b.error) ?? "error"}`
        }`,
    ),
    `comments with extracted requests: ${withRequests}; without: ${report.length - withRequests}`,
    `requests by state: changed ${state("changed")}, needs_author ${state("needs_author")}, outside_this_file ${state("outside_this_file")}, open ${state("open")}`,
    `requirements no change names: ${
      requests.length
        ? sum((c) => c.unmet_requirements.length)
        : "not measured (no requests were extracted)"
    }`,
    `questions for the author: ${sum((c) => c.questions.length)}`,
    `questions answered from the answers file: ${sum((c) => c.answered ?? 0)} in ${report.filter((c) => c.answered).length} comments`,
    `turns run after answers: ${answerTurns.length}`,
    `requests declared outside this file: ${sum((c) => c.outside.length)}`,
    `terms still present elsewhere: ${sum((c) => c.still_present.length)}`,
    `terms the reviewers named: ${named.length} in ${report.filter((c) => c.named_terms?.length).length} comments; still in the text: ${named.filter((t) => t.places.length).length}`,
    `changes with words from the review: ${sum((c) => c.changes.filter((x) => x.from_review?.length).length)}`,
    `paragraphs written anew: ${rewrites.length} rewrites of ${new Set(rewrites.map((w) => w.paragraph_index)).size} paragraphs, ${rewrites.reduce((n, w) => n + w.change_ids.length, 0)} parts recorded; stopped: ${logged.filter((w) => w.stopped).length}`,
    `rewrite turns run: ${rewriteTurns.length}`,
    `follow-through turns run: ${followTurns.length}`,
    `follow-through logged work (all attempts): changes ${followWork.changes}, questions ${followWork.questions}`,
    `latest completed follow-through checks: changes ${followed.length} in ${new Set(followed.map((x) => x.paragraph_index)).size} paragraphs, questions ${sum((c) => c.follow?.questions ?? 0)}`,
    `edits a guard refused: ${list(codes) || "none"}`,
    `turns that did not end well: ${bad.length}${bad.length ? ` (${bad.join(", ")})` : ""}`,
    ...allTurns
      .filter((t) => t.outcome !== "ok")
      .map((t) => `  ${t.label} ${t.outcome}: ${reasonClass(t)}`),
    `turns that ended well with a pane error shown: ${paneErrors.length}${paneErrors.length ? ` (${paneErrors.map((t) => t.label).join(", ")})` : ""}`,
    `turns run again after an API error: ${again.length}${again.length ? ` (${again.join(", ")})` : ""}`,
    `models asked for: edits ${asked.edits ?? "-"}, writing ${asked.writing ?? "(the pane's own)"}, judgments ${asked.judgments ?? "(the pane's own)"}`,
    `models that ran, as the pane says: ${models?.line || "-"}`,
    `models that ran, from the daemon log: ${[...new Set(turns.map((c) => c.turn.model).filter(Boolean))].join(", ") || "-"}`,
    `model replaced after a refusal: ${fallbacks.reduce((n, t) => n + t.fallbacks.length, 0)} times in ${fallbacks.length} turns${fallbacks.length ? ` (${fallbacks.map((t) => t.label).join(", ")})` : ""}; restored ${restored} times`,
    ...(stay ? [stay] : []),
    `turns on another model than asked: ${drifted.length}${drifted.length ? ` (${drifted.map((c) => c.ref).join(", ")})` : ""}`,
    `another model than asked: ${failures.length ? `FAILED (${failures.join("; ")})` : "no"}`,
    `paragraph check: ${
      paragraphCheck
        ? `${paragraphCheck.findings?.length ?? 0} findings in ${paragraphCheck.paragraphs?.length ?? 0} paragraphs${Object.keys(kinds).length ? ` (${list(kinds)})` : ""}`
        : "not run"
    }`,
    `turn time: ${[...turns.map((c) => c.turn), ...answerTurns, ...followTurns, ...rewriteTurns].reduce((n, t) => n + (t.seconds ?? 0), 0)} s`,
    `total time: ${seconds} s`,
    `copy saved with its tracked changes: ${saved ? "yes" : "NO"}`,
  ];
  if (stopped)
    lines.push(
      `stopped: usage limit reached at ${stopped.ref} (${stopped.kind})${stopped.resets ? `, resets at ${stopped.resets}` : ""}`,
    );
  if (source) lines.push(`manuscript: ${source}`);
  if (responses !== null) lines.push(`responses drafted: ${responses}`);
  return `${lines.join("\n")}\n`;
}

// A Word window that shows no document of the user: a new blank document
// ("Document1 - Word", "문서1 - Word") or no window at all.
export const blankWordTitle = (title) =>
  !String(title ?? "").trim() || /^(?:Document|문서)\s*\d+ - Word$/.test(String(title).trim());

// New WINWORD processes after one step of the run: in after, not in before
// (both { pid, started, title } lists taken right around the step), and not
// the process the step itself started and quits (own).
export const newWords = (before, after, own = null) =>
  after.filter((p) => p.pid !== own && !before.some((b) => b.pid === p.pid));

// PID deltas, start times and window titles cannot prove ownership. Strays
// are report-only, including a user-created blank Word opened during a step.
export function sweepSelection({ processes = [], strays = [], preexisting = [], keep = [] }) {
  const close = [];
  const skip = [];
  for (const p of processes) {
    const stray = strays.some((s) => s.pid === p.pid && s.started && s.started === p.started);
    const why = preexisting.includes(p.pid)
      ? "ran before the run"
      : keep.includes(p.pid)
        ? "kept by the run"
        : !stray
          ? "left open: not started by this run"
          : "left open: ownership unproven";
    skip.push({ ...p, reason: why });
  }
  return { close, skip };
}
