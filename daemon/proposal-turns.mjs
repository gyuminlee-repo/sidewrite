// A proposal turn: a review turn for a provider that is given no tools.
//
// The provider is asked once for a list of proposed changes (proposal.mjs),
// in a run that can do nothing but answer. This program then checks every
// proposal and applies them one by one through applyTrackedEdit, the same
// path as office_tracked_edit in a Claude turn: the attempt written before
// Word is called, the paragraph hash, the forbid list, the record of the
// change. Nothing here tells applyTrackedEdit to lift a rule: it is called
// with (deps, args, turn) and no fourth argument.
//
// Stops. The first proposal that does not end as an applied change with its
// rationale comment ends the turn; the ones after it are not sent to Word.
// That includes StaleTarget (in a Claude turn the model reads the paragraph
// again; here nothing can), a change without its comment, a status refresh
// while the turn runs, and Stop. A proposal the checks of proposal.mjs
// refuse was never sent to Word and stops nothing.
//
// The task pane sees the events of a Claude turn: tool_use_announce per
// edit, then turn_complete; endTurn is the last thing a turn does.
// sendText(text, key), sendEvent(obj, key), endTurn(key, outcome) as in
// cli-turns.mjs.
//
// One receipt per turn (decision-log.mjs, kind proposal_batch), also for a
// turn that failed or was stopped: ids, codes and outcomes, no text of the
// manuscript, the comment or the reply. Log lines hold the same.

import { randomUUID } from "node:crypto";
import { applyTrackedEdit, currentTexts } from "./review-tools.mjs";
import {
  PROPOSAL_JSON_SCHEMA,
  buildProposalPrompt,
  checkProposal,
  checkProposals,
  readProposalReply,
  toTrackedEditArgs,
} from "./proposal.mjs";
import { decisionsPathFor, findComment, loadRevision } from "./revision-store.mjs";
import { appendDecisions, decisionRecord } from "./decision-log.mjs";
import { presetById } from "./providers/registry.mjs";

export const PROPOSAL_TOOL = "mcp__office__office_tracked_edit";
const STOPPED = "The turn was stopped.";

// What the author is told when the turn stops at a proposal, by code.
const HINTS = {
  Blocked: (r) =>
    `It matches forbidden rule(s) ${(r.blocked_by ?? []).join(", ")}. It is in the blocked list of the Review tab, where an exception can be approved.`,
  ForbidRuleBroken: (r) =>
    `Forbidden rule(s) ${(r.broken_rules ?? []).join(", ")} cannot be checked (invalid regex). Fix or remove them in the Review tab, then propose again.`,
  TaskPaneOutdated: () =>
    "The task pane is older than the daemon or not connected, so nothing was sent to Word. Close and reopen the task pane, then propose again.",
  StaleTarget: () =>
    "The paragraph is not as the last status refresh has it, and Word was left unchanged. Press Refresh status, then propose again.",
  OutcomeUnknown: () =>
    "The task pane did not answer: the change may or may not be in the document. Press Refresh status and check the paragraph.",
  AppliedNotRecorded: () =>
    "The change is in Word but could not be written to the review file. Press Refresh status, which records it.",
  CommentMissing: () =>
    "The change is in the document without its rationale comment. Press Add rationale comment on that change in the Review tab.",
  SnapshotChanged: () => "The status was refreshed while the turn ran. Propose again.",
};

// outcome of a proposal -> its words in the summary line, in this order.
const OUTCOME_WORDS = [
  ["applied", "applied"],
  ["applied_no_comment", "applied without its comment"],
  ["applied_not_recorded", "applied but not recorded"],
  ["unknown", "unanswered"],
  ["failed", "failed"],
  ["rejected", "rejected"],
  ["skipped", "skipped"],
  ["not_applied", "not applied"],
];
const plural = (n) => `${n} proposal${n === 1 ? "" : "s"}`;

// handleFor(choice): the generation handle (oneshot-router.mjs): { model,
//   run({ system, prompt, schema, signal, timeoutMs }) -> { ok, text, json }
//   | { ok: false, error, aborted?, tool_event? } }, never throwing.
// trackedEditDeps(key, at), call(name, args, key): review-service.mjs and the
//   bridge to the task pane.
// budgetOf(choice): the most characters of one request, or null for a
//   provider that runs no proposal turns.
export function createProposalTurns({
  handleFor,
  trackedEditDeps,
  call,
  applyEdit = applyTrackedEdit,
  load = (d) => loadRevision(d.path, d.doc),
  receipt = (d, records) => appendDecisions(decisionsPathFor(d.doc), records),
  budgetOf = (choice) => presetById(choice?.id)?.proposalBudget ?? null,
  includeForbid = () => true,
  lang = "Korean",
  timeoutMs = 300_000,
  sendText,
  sendEvent,
  endTurn,
  log = console.log,
}) {
  const live = new Map(); // pane key -> turn

  // Everything before the end of the turn. Returns nothing: what happened is
  // in t (failure, rows, stopped).
  async function work(t) {
    const { key, ref, d } = t;
    const label = presetById(t.choice?.id)?.label ?? String(t.choice?.id ?? "This provider");
    const limit = budgetOf(t.choice);
    if (!limit) return void (t.failure = `${label} does not run proposal turns.`);
    const handle = handleFor(t.choice);
    t.model = handle.model;
    const deps = trackedEditDeps(key, d);

    const rev = await load(d);
    const comment = findComment(rev, ref);
    if (!comment) return void (t.failure = `No reviewer comment ${ref}.`);
    t.snapshotAt = rev.snapshot?.taken_at ?? null;
    const ctx = await deps.getReviewContext(ref, d);
    const built = buildProposalPrompt({
      comment,
      snapshot: rev.snapshot ? { paragraphs: ctx.paragraphs, ids: ctx.paragraph_ids } : null,
      recentChanges: ctx.recent_changes,
      selected: t.selected,
      ledger: rev.ledger,
      lang,
      includeForbid: includeForbid(),
      limit,
      label,
    });
    if (!built.ok) return void (t.failure = built.message);
    t.input = { system: built.system, prompt: built.prompt };
    log(
      `[proposal] prompt ref=${ref} provider=${handle.model} scope=${built.scope} paragraphs=${built.shown.length} chars=${built.chars}`,
    );
    if (t.stopped) return;

    // One request; a reply in the wrong form is asked for once more.
    const ask = (prompt) =>
      handle.run({
        system: built.system,
        prompt,
        schema: PROPOSAL_JSON_SCHEMA,
        signal: t.abort.signal,
        timeoutMs,
      });
    let read = null;
    for (let n = 0; n < 2; n += 1) {
      const prompt = n
        ? `${built.prompt}\n\nYour reply was not usable: ${read.problem}. Reply again with only the JSON object.`
        : built.prompt;
      t.retried = n > 0;
      const r = await ask(prompt);
      if (t.stopped || r?.aborted) return void (t.stopped = true);
      if (r?.tool_event) {
        // The run did more than answer: nothing of it is used.
        t.toolEvent = r.tool_event;
        return void (t.failure = `${r.error} No change was made.`);
      }
      read = readProposalReply(r);
      if (!read) {
        return void (t.failure = `${label} gave no proposals (${r?.error ?? "no reply"}). No change was made.`);
      }
      if (read.ok) break;
      log(`[proposal] unusable reply ref=${ref} try=${n + 1}`);
    }
    if (!read.ok) {
      return void (t.failure = `${label} answered twice in a form that cannot be used (${read.problem}). No change was made.`);
    }

    // The checks of proposal.mjs, on the paragraphs the model was shown.
    const requirements = comment.requirements ?? [];
    const known = new Set(requirements.map((q) => q.id));
    const shownIds = new Set(built.shown.map((s) => s.id));
    const checkCtx = (c) => ({
      texts: currentTexts(c.paragraphs, c.recent_changes),
      ids: c.paragraph_ids,
      shown: built.shown,
      requirements,
    });
    t.rows = checkProposals(read.proposals, checkCtx(ctx)).map(({ id, proposal, code }) => ({
      id,
      proposal,
      // For the receipt: only ids this program knows, never a string the
      // model made up.
      paragraph_id: shownIds.has(proposal.paragraph_id) ? proposal.paragraph_id : null,
      requirement_ids: proposal.requirement_ids.filter((q) => known.has(q)),
      outcome: code ? "rejected" : "pending",
      code,
      change_id: null,
    }));
    sendText(`${plural(t.rows.length)} received.\n`, key);
    if (read.note) sendText(`${read.note}\n`, key);
    for (const row of t.rows.filter((x) => x.code)) {
      sendText(`${row.id} rejected (${row.code}): paragraph ${row.proposal.paragraph_id}.\n`, key);
      log(`[proposal] rejected ref=${ref} id=${row.id} code=${row.code}`);
    }

    for (const row of t.rows) {
      if (row.outcome !== "pending") continue;
      if (t.stopped) return void (t.stoppedAt = row.id);
      // The proposals were made for one snapshot. A refresh since then means
      // paragraphs the model did not see.
      const now = await load(d);
      if ((now.snapshot?.taken_at ?? null) !== t.snapshotAt) {
        return stopAt(t, row, "failed", "SnapshotChanged", {});
      }
      // As the paragraphs read now, with the edits of this turn in them.
      if (checkProposal(row.proposal, checkCtx(await deps.getReviewContext(ref, d)))) {
        row.outcome = "skipped";
        row.code = "SupersededInTurn";
        sendText(`${row.id} skipped (${row.code}): paragraph ${row.proposal.paragraph_id}.\n`, key);
        log(`[proposal] skipped ref=${ref} id=${row.id} code=${row.code}`);
        continue;
      }
      const args = toTrackedEditArgs(ref, row.proposal);
      sendEvent({ event: "tool_use_announce", tool: PROPOSAL_TOOL, input: args }, key);
      // A Stop during this call is read after it: what Word did is recorded.
      const r = await Promise.resolve()
        .then(() => applyEdit({ ...deps, call: (name, a) => call(name, a, key) }, args, { ref }))
        .then(undefined, (e) => ({
          ok: false,
          code: "ToolError",
          message: String(e?.message ?? e),
        }));
      row.change_id = r?.change_id ?? null;
      if (r?.ok && !r.comment_missing) {
        row.outcome = "applied";
        continue;
      }
      if (r?.ok) return stopAt(t, row, "applied_no_comment", "CommentMissing", r);
      const code = r?.code ?? (r?.blocked_by ? "Blocked" : "NotApplied");
      const outcome =
        code === "AppliedNotRecorded"
          ? "applied_not_recorded"
          : code === "OutcomeUnknown"
            ? "unknown"
            : "failed";
      return stopAt(t, row, outcome, code, r ?? {});
    }
    if (t.stopped) return;
    const applied = t.rows.filter((x) => x.outcome === "applied").length;
    if (t.rows.length && !applied) t.failure = "No proposal could be placed.";
  }

  // The turn ends at this proposal.
  function stopAt(t, row, outcome, code, r) {
    row.outcome = outcome;
    row.code = code;
    t.stoppedAt = row.id;
    const hint = HINTS[code]?.(r) ?? r.message ?? r.error ?? "The change was not applied.";
    t.failure = `The turn stopped at ${row.id} (${code}). ${hint}`;
    log(`[proposal] stopped ref=${t.ref} id=${row.id} code=${code}`);
  }

  async function run(t) {
    const { key, ref } = t;
    await Promise.resolve()
      .then(() => work(t))
      .then(undefined, (e) => {
        t.failure ??= String(e?.message ?? e);
      });
    // What was not started.
    const left = t.rows.filter((x) => x.outcome === "pending");
    for (const row of left) row.outcome = "not_applied";
    if (t.stopped) t.stoppedAt ??= left[0]?.id ?? null;
    if (left.length) {
      sendText(
        `${left.map((x) => x.id).join(", ")} not applied: ${
          t.stopped
            ? "the turn was stopped"
            : `the turn stopped at ${t.stoppedAt} (${t.rows.find((x) => x.id === t.stoppedAt)?.code})`
        }.\n`,
        key,
      );
    }
    const count = (outcome) => t.rows.filter((x) => x.outcome === outcome).length;
    if (t.rows.length) {
      const parts = OUTCOME_WORDS.filter(([o]) => count(o)).map(([o, w]) => `${count(o)} ${w}`);
      sendText(`${plural(t.rows.length)}: ${parts.join(", ")}.\n`, key);
    }
    const outcome = t.stopped ? "stopped" : t.failure ? "error" : "ok";
    log(
      `[proposal] done ref=${ref} proposals=${t.rows.length} applied=${count("applied")} rejected=${count("rejected")} skipped=${count("skipped")} stopped_at=${t.stoppedAt ?? "-"} outcome=${outcome}`,
    );
    await Promise.resolve()
      .then(() =>
        receipt(t.d, [
          decisionRecord({
            kind: "proposal_batch",
            input: t.input,
            model: t.model,
            output: {
              turn_id: t.id,
              ref,
              snapshot_taken_at: t.snapshotAt,
              retried: t.retried,
              stopped_at: t.stoppedAt,
              ...(t.toolEvent ? { tool_event: t.toolEvent } : {}),
              proposals: t.rows.map((x) => ({
                id: x.id,
                paragraph_id: x.paragraph_id,
                op: x.proposal.op,
                requirement_ids: x.requirement_ids,
                outcome: x.outcome,
                code: x.code,
                ...(x.change_id ? { change_id: x.change_id } : {}),
              })),
            },
          }),
        ]),
      )
      .then(undefined, (e) => log(`[proposal] receipt not written ref=${ref}: ${e?.message ?? e}`));

    live.delete(key);
    if (t.stopped) {
      sendEvent({ event: "turn_complete", interrupted: true }, key);
      endTurn(key, { ok: false, error: STOPPED });
    } else if (t.failure) {
      sendEvent({ event: "error", error: t.failure }, key);
      sendEvent({ event: "turn_complete", subtype: "error" }, key);
      endTurn(key, { ok: false, error: t.failure });
    } else {
      sendEvent({ event: "turn_complete", subtype: "success" }, key);
      endTurn(key, { ok: true });
    }
  }

  return {
    running: (key) => live.has(key),
    // Starts the turn and resolves at once: the turn is registered before
    // anything is awaited, so a Stop right after this call finds it. Rejects
    // only when nothing was started. turn: { ref }. source: { at: the
    // document handle of the pane (review-service.mjs at()), selected:
    // [{ index }] or [], choice: { id, model } of the provider }.
    async start(key, host, turn, { at, selected = [], choice = null } = {}) {
      if (!turn?.ref) throw new Error("A proposal turn needs the comment it is for.");
      if (live.has(key)) throw new Error("A proposal turn is already running in this pane.");
      const t = {
        id: randomUUID(),
        key,
        ref: turn.ref,
        d: at,
        selected,
        choice,
        abort: new AbortController(),
        stopped: false,
        failure: null,
        rows: [],
        input: { system: null, prompt: null },
        model: null,
        snapshotAt: null,
        retried: false,
        stoppedAt: null,
        toolEvent: null,
        done: null,
      };
      live.set(key, t);
      t.done = run(t);
    },
    // true when a turn of this pane was running, in whatever stage: before
    // the request, while the provider answers (its run is aborted), between
    // two proposals, or while Word works (that answer is recorded first).
    stop(key) {
      const t = live.get(key);
      if (!t) return false;
      t.stopped = true;
      t.abort.abort();
      return true;
    },
    // Daemon exit: stop every turn and wait until each has ended.
    async shutdown() {
      const all = [...live.values()];
      for (const t of all) {
        t.stopped = true;
        t.abort.abort();
      }
      await Promise.all(all.map((t) => t.done));
    },
  };
}
