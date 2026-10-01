// Review-tab message handlers and the recording side of office_tracked_edit.
// index.mjs wires the dependencies; everything here is testable with fakes.
//
// Messages (task pane -> daemon, reply type is `<type>_result` with the same
// request_id, the existing sendRequest pattern in taskpane.js):
//   review_load, review_split, review_save_comments, review_set_status,
//   review_set_requirements, review_preview, review_propose, review_refresh,
//   review_external_dismiss, review_comments_cleaned, review_override,
//   review_comment_text, review_comment_set,
//   ledger_add, ledger_update, ledger_remove, ledger_check, ledger_export,
//   ledger_import, ledger_suggest, comments_backup
// Push (daemon -> task pane, no request_id):
//   { type: "review_event", event: "change_recorded" | "edit_blocked" | "turn_done"
//     | "late_result_recorded" | "stale_target" | "record_failed" | "job_done", ... }
//   turn_done: { ref, ok, error? }. ok false: the turn ended on a usage limit,
//   a refusal, an API error, an empty turn, a stop or a crash (error says which).
// Jobs (JOB_TYPES): the reply is { ok: true, accepted: true, job_id } at once
// and the result follows as { event: "job_done", job_id, result }, result
// being what the reply would have held ({ ok: true, ... } or { ok: false,
// error }). Their model calls can outlast any request timeout of the pane.

import { randomUUID } from "node:crypto";
import { access } from "node:fs/promises";
import {
  localDocPath,
  revisionPathFor,
  decisionsPathFor,
  commentsBackupPathFor,
  loadRevision,
  withRevision,
  replaceComments,
  restoreComment,
  setCommentStatus,
  addChange,
  applyStates,
  addBlocked,
  addAttempt,
  setBlockedAttempt,
  dropBlocked,
  takeBlocked,
  markBlockedUnknown,
  claimBlocked,
  releaseBlocked,
  markStaleApplying,
  addLedgerItem,
  updateLedgerItem,
  removeLedgerItem,
  ledgerExportDoc,
  importLedgerItems,
  sidecarPath,
  findComment,
  writeJsonAtomic,
  setRequirements,
  setSnapshot,
  addExternal,
  dismissExternal,
  markCommentsCleaned,
  setChangeComment,
} from "./revision-store.mjs";
import {
  classifyRevision,
  classifyChange,
  summarizeStates,
  normalizeParaText,
} from "./revision-classify.mjs";
import { extractRequirements, splitReview } from "./review-split.mjs";
import { buildReviewPrompt, formatPromptLog } from "./review-prompt.mjs";
import { checkLedgerIncremental, brokenForbidRules } from "./ledger-check.mjs";
import { suggestLedger } from "./ledger-suggest.mjs";
import { presetById } from "./providers/registry.mjs";
import { snapshotTexts, snapshotIds, outsideChanges } from "./doc-diff.mjs";
import {
  applyTrackedEdit,
  narrowedAttempt,
  formatRationaleComment,
  TOOL_COMMENT_PREFIX,
  TARGET_HASH_CAP,
  PROPOSAL_CAP,
} from "./review-tools.mjs";
import { appendDecisions, decisionRecord, redactQuotes } from "./decision-log.mjs";

async function exists(p) {
  try {
    await access(p);
    return true;
  } catch {
    return false;
  }
}

// An override that got no answer (Blocked.state outcome_unknown) is taken as
// applied when the classifier finds its expected form in the snapshot:
// accepted, or pending for a replace or an insertion whose new text is in
// the paragraph. A pending deletion reads the same as one never applied, so
// it stays unknown. Found ones become a recorded change (with the rationale
// comment the snapshot holds, if any) and leave the blocked list.
//
// Two forms are looked for (Codex review of the recovery path): the call as
// it was asked, and the narrower form the task pane applies when it can
// (attempt.change.narrow; review-tools.mjs narrowedAttempt). A narrowed edit
// that is still tracked leaves the old words of the full call in the
// paragraph, which reads "rejected" for the full form. The full form is
// tried first: it is what an accepted edit reads as either way. The narrowed
// form counts, when pending, with its rationale comment in the document or,
// for a replace, with its new and old words next to each other (Word's
// tracked replace); its new words alone are too few to say anything. An
// attempt is dropped as not applied only when every form reads rejected.
export function resolveUnknownOverrides(rev, snapshot) {
  let next = rev;
  const resolved = [];
  const dropped = [];
  const used = new Set(
    rev.comments.flatMap((c) => c.changes.map((ch) => ch.word_comment_id).filter(Boolean)),
  );
  const paras = snapshot.paragraphs ?? [];
  for (const c of rev.comments) {
    for (const b of c.blocked) {
      if (b.state !== "outcome_unknown" || !b.attempt?.change?.context) continue;
      const own = b.blocked_by.length === 0;
      const comment = (snapshot.comments ?? []).find(
        (x) =>
          !used.has(String(x.id)) &&
          normalizeParaText(x.content) === normalizeParaText(b.attempt.comment_text),
      );
      const { narrow: _narrow, ...full } = b.attempt.change;
      const judged = [full, narrowedAttempt(b.attempt.change)].filter(Boolean).map((form) => {
        const change = { ...form, word_comment_id: comment ? String(comment.id) : null };
        const row = classifyChange(change, snapshot);
        const p = paras.find((x) => x.index === row.paragraph_index);
        // Plain containment: Word joins pending inserted and deleted text with
        // no space, so a word-bounded match would miss it.
        const holds = (s) => !!p && normalizeParaText(p.text).includes(normalizeParaText(s));
        const newTextThere = form.narrowed_from
          ? change.kind === "replace" && holds(change.text + change.find)
          : change.kind !== "delete" && holds(change.text);
        // Its rationale comment in the document settles a pending deletion
        // too: the task pane adds the comment only after the edit.
        const applied =
          row.state === "accepted" || (row.state === "pending" && (!!comment || newTextThere));
        return { change, row, p, applied };
      });
      const hit = judged.find((j) => j.applied);
      if (!hit) {
        // A normal edit whose own range still reads the old form (in every
        // form Word may have applied), with no rationale comment: not
        // applied, nothing is left to wait for. An approved exception stays,
        // so the author can decide again.
        if (own && !comment && judged.every((j) => j.row.state === "rejected")) {
          next = takeBlocked(next, c.ref, b.id).revision;
          dropped.push({ ref: c.ref, blocked_id: b.id, tool_call_id: b.attempt.tool_call_id });
        }
        continue;
      }
      const { change, row, p } = hit;
      const added = addChange(next, c.ref, {
        ...change,
        paragraph_index: row.paragraph_index,
        paragraph_id: p?.id ?? null,
        comment_missing: comment ? null : "CommentNotFound",
      });
      next = takeBlocked(added.revision, c.ref, b.id).revision;
      if (comment) used.add(String(comment.id));
      resolved.push({
        ref: c.ref,
        blocked_id: b.id,
        change_id: added.change.id,
        tool_call_id: b.attempt.tool_call_id,
        state: row.state,
        own,
      });
    }
  }
  return { revision: next, resolved, dropped };
}

// Word comment ids of the rationale comments on record.
const recordedCommentIds = (rev) =>
  rev.comments.flatMap((c) => c.changes.map((ch) => ch.word_comment_id).filter(Boolean));

// What a ledger check result depends on: the rule, not its last result.
// review_save_comments: comments work still points at are not removed.
export const keptWarning = (refs) =>
  `Not removed because work on it is in progress: ${refs.join(", ")}`;

// Requests that may run model calls one after another (a split and its
// requirement extraction; a ledger check of up to 20 batches per rule). The
// regex-only path takes the same route: one reply shape per request type.
const JOB_TYPES = new Set(["review_split", "ledger_check", "ledger_suggest"]);

const ruleFingerprint = (item) => JSON.stringify([item.kind, item.text, item.scope, item.check]);

export function createReviewService({
  getContext,
  callTool,
  pushTurn,
  oneshot = null,
  // Model names for decision receipts: the one-shot model, and the model of
  // the drafting session of a pane (getModel(key)). null when unknown.
  oneshotModel = null,
  // P14: the one-shot of ONE JOB of a pane, when judgments go to the
  // provider chosen in that pane (providers/oneshot-router.mjs pin):
  // async (key) => { run, model }. A job (a split with its requirement
  // extraction, a ledger check with all its batches, a suggestion) asks for
  // it once, when it starts; every request of the job goes through `run`
  // and every receipt carries `model`, so a provider changed in the pane
  // while the job runs applies to the next job. The default is the two
  // arguments above, for every pane.
  oneshotJob = async () => ({ run: oneshot, model: oneshotModel }),
  // P14: asked before a review turn is built or queued; it resolves when
  // the pane's provider for review turns is confirmed and throws otherwise
  // (providers/selection.mjs ready). The request then fails with that text.
  turnGate = async () => {},
  // A provider that is given no tools and proposes (proposal-turns.mjs):
  // async (key, ref, { at, selected }) => null when the pane's provider for
  // review turns calls the tools itself (Claude), else { choice, built }:
  // the provider of this request and what would be sent to it
  // (proposal.mjs buildProposalPrompt, or { ok: false, message }). It throws
  // for a provider that runs no review turns.
  proposalPreview = async () => null,
  clock = Date.now,
  getModel = () => null,
  sendEvent = () => {},
  log = console.log,
  env = process.env,
  lateResultTtlMs = 10 * 60_000,
}) {
  // Tracked edits the task pane did not answer, by bridge call id:
  // { key, ref, record, blocked_id, timer }. blocked_id: the approved
  // exception (review_override) the edit came from, released when recorded. A late tool_result (bridge.mjs
  // onLateToolResult) is recorded through `record` once; the entry goes away
  // on first use or after lateResultTtlMs, like the bridge's own copy.
  const pendingLate = new Map();
  function forgetLate(id) {
    const p = pendingLate.get(id);
    if (!p) return null;
    clearTimeout(p.timer);
    pendingLate.delete(id);
    p.release();
    return p;
  }

  // Work still open on a comment, by token: { path, ref, comment }. Held by
  // a review turn (review_propose .. turnEnded), a tracked edit waiting for
  // Word (holdRef), and an unanswered edit whose answer may come late
  // (pendingLate). comment: { reviewer, text, requirements } when known, so
  // a comment removed anyway can be restored with its text (onEdit).
  const inflight = new Map();
  let lastToken = 0;
  function hold(d, ref, comment = null) {
    let path;
    try {
      path = d.path;
    } catch {
      return () => {};
    }
    const token = ++lastToken;
    inflight.set(token, { path, ref, comment: comment ?? remembered(path, ref) });
    return () => inflight.delete(token);
  }
  function remembered(path, ref) {
    for (const e of inflight.values()) {
      if (e.path === path && e.ref === ref && e.comment) return e.comment;
    }
    return null;
  }
  // Refs a comment-list save must keep: held above, or with an exception
  // being applied or unanswered in the file (Blocked.state).
  function heldRefs(path, rev) {
    const refs = new Set();
    for (const e of inflight.values()) if (e.path === path) refs.add(e.ref);
    for (const c of rev.comments) {
      if (c.blocked.some((b) => b.state === "applying" || b.state === "outcome_unknown")) {
        refs.add(c.ref);
      }
    }
    return refs;
  }
  // The comment a change lands under; restored when it was removed anyway,
  // so an edit Word applied is never thrown away.
  function withComment(d, rev, ref) {
    if (findComment(rev, ref)) return rev;
    log(`[review] comment ${ref} was missing; restored to keep the change`);
    return restoreComment(rev, ref, remembered(d.path, ref));
  }
  const turnHolds = new Map(); // pane key -> release of its review turn hold
  // A review_propose that has not become a turn yet, by pane key:
  // { stopped }. A Stop in that stretch (cancelStarting) is kept: the request
  // then starts nothing and fails as stopped.
  const starting = new Map();
  function cancelStarting(key) {
    const s = starting.get(key);
    if (!s) return false;
    s.stopped = true;
    return true;
  }
  // Ledger check generation by revision file and item id: each check of an
  // item takes the next one and only the check started last saves its
  // result. Two checks of one rule on two snapshots share a fingerprint, so
  // the fingerprint alone lets the older one win when it ends last.
  const checkGen = new Map();
  let lastCheckGen = 0;

  function docFor(key) {
    const doc = localDocPath(getContext(key)?.activeDoc, env);
    if (!doc) {
      throw new Error(
        "Save the document to a local or WSL-reachable folder first. The review file lives next to it.",
      );
    }
    return { doc, path: revisionPathFor(doc) };
  }
  // The document of a pane, read from its live context ONCE, when a request
  // or a tool call starts. Everything that runs after an await (the answer
  // of Word, a late answer, a job that called a model) uses this handle and
  // never the context again: a dropped socket removes the pane's context,
  // and a document saved under a new name changes it under the same key
  // (Codex review 15: the path was read again after the tool call, threw,
  // and an applied edit was lost). A pane with no usable document gives a
  // handle that throws the "save the document first" error on first use.
  function at(key) {
    let found = null;
    let error = null;
    try {
      found = docFor(key);
    } catch (e) {
      error = e;
    }
    const need = () => {
      if (error) throw error;
      return found;
    };
    return {
      key,
      // What the task pane said it can do, as of now (bridge.mjs caps).
      caps: getContext(key)?.caps ?? [],
      get doc() {
        return need().doc;
      },
      get path() {
        return need().path;
      },
    };
  }
  const load = (d) => loadRevision(d.path, d.doc);
  const update = (d, fn) => withRevision(d.path, d.doc, fn);

  // Decision receipts (decision-log.mjs). A receipt never stops the feature
  // it records: a failure only logs.
  async function receipt(d, records) {
    if (!records.length) return;
    try {
      await appendDecisions(decisionsPathFor(d.doc), records);
    } catch (e) {
      log(`[review] decision log failed: ${e?.message ?? e}`);
    }
  }

  // Overrides this daemon is running, per review file: "<ref>\0<blocked id>".
  // An applying item not in here was left by a daemon that stopped.
  const applying = new Map();
  const applyingKey = (ref, id) => `${ref}\0${id}`;
  function activeFor(d) {
    const { path } = d;
    if (!applying.has(path)) applying.set(path, new Set());
    return applying.get(path);
  }
  const clearStale = (d, rev) => markStaleApplying(rev, activeFor(d));

  // Changes Word applied whose save failed (onRecordFailed), by file:
  // attempt id -> { ref, attempt_id, change }, the change exactly as onEdit
  // was given it (narrowed or not, with its Word comment id). The attempt
  // on file was written before the call and knows neither, so while this
  // daemon runs the change is saved from here, in the next load or refresh
  // of the file, and its attempt leaves the blocked list in the same save.
  // A daemon that stopped has only the attempt (resolveUnknownOverrides).
  const unsaved = new Map();
  // In an update of the file. done: the entries to forget once it is saved.
  function recordUnsaved(d, rev) {
    let next = rev;
    const done = [];
    for (const h of unsaved.get(d.path)?.values() ?? []) {
      // Its attempt is gone: settled some other way, nothing to add.
      if (!findComment(next, h.ref)?.blocked.some((b) => b.id === h.attempt_id)) {
        done.push({ ...h, saved: null });
        continue;
      }
      const r = addChange(next, h.ref, h.change);
      next = takeBlocked(r.revision, h.ref, h.attempt_id).revision;
      done.push({ ...h, saved: r.change });
    }
    return { revision: next === rev ? null : next, done };
  }
  // After that update was saved.
  async function unsavedDone(d, key, revision, done) {
    for (const h of done) {
      unsaved.get(d.path)?.delete(h.attempt_id);
      if (!h.saved) {
        log(`[review] applied result of ${h.attempt_id} not used: its attempt is not on file`);
        continue;
      }
      log(
        `[review] change recorded ${h.saved.id} kind=${h.saved.kind} comment=${h.saved.word_comment_id} from the applied result kept in memory (attempt ${h.attempt_id})`,
      );
      await linkReceipt(d, key, h.ref, findComment(revision, h.ref), h.saved);
    }
  }
  // The requirement link of a recorded change, for the decision log.
  const linkReceipt = (d, key, ref, comment, saved) =>
    receipt(d, [
      decisionRecord({
        kind: "requirement_link",
        input: {
          comment: comment?.text ?? "",
          ...(comment?.section && { section: comment.section }),
          requirements: comment?.requirements ?? [],
          find: saved.find,
          text: saved.text,
        },
        model: getModel(key),
        output: {
          change_id: saved.id,
          comment_ref: ref,
          requirement_ids: saved.rationale?.requirement_ids ?? [],
        },
      }),
    ]);
  const backupChains = new Map();

  // begin(): the handle of one tool call, taken before Word is called;
  // applyTrackedEdit gives it back to every function below as `at`/`d`.
  function trackedEditDeps(key, fixed = null) {
    return {
      log,
      begin: () => fixed ?? at(key),
      getLedger: async (d = at(key)) => (await load(d)).ledger,
      // L2 sources and L5 ids for office_tracked_edit. recent_changes: the
      // changes recorded since the snapshot (earlier edits of this turn), so
      // the forbid check reads the paragraphs as they are now.
      getReviewContext: async (ref, d = at(key)) => {
        const rev = await load(d);
        const c = findComment(rev, ref);
        const since = rev.snapshot?.taken_at;
        return {
          requirements: c?.requirements ?? [],
          // With its section label: the split took it off the comment text.
          comment_text: [c?.section, c?.text ?? ""].filter((x) => x !== undefined).join("\n"),
          paragraphs: rev.snapshot?.paragraphs ?? [],
          paragraph_ids: rev.snapshot?.ids ?? null,
          comment_ids: recordedCommentIds(rev),
          recent_changes: since
            ? rev.comments
                .flatMap((x) => x.changes)
                .filter((ch) => ch.created_at > since)
                .sort((a, b) => (a.created_at < b.created_at ? -1 : 1))
            : [],
        };
      },
      // A normal edit is written to the file before Word is called, as a
      // blocked item no rule blocks (state applying, with the form it would
      // have in the document). The stretch between "Word applied it" and
      // "the change is saved" then always has it on file: a daemon that
      // stops there leaves an item the next one takes as unanswered
      // (clearStale) and a refresh settles. It joins `active` in the same
      // step, like an override, so a load in between does not take it for a
      // stale one. One more save per edit, in the chain of the file
      // (withRevision): it is done before Word is called, and the save of
      // the change is queued after it.
      onAttempt: async ({ ref, args, attempt, at: d = at(key) }) => {
        const active = activeFor(d);
        let saved = null;
        try {
          await update(d, (rev) => {
            const r = addAttempt(withComment(d, rev, ref), ref, { args, attempt });
            saved = r.blocked;
            active.add(applyingKey(ref, saved.id));
            return r.revision;
          });
        } catch (e) {
          if (saved) active.delete(applyingKey(ref, saved.id)); // added, then the save failed
          throw e;
        }
        return { id: saved.id, end: () => active.delete(applyingKey(ref, saved.id)) };
      },
      // Word answered that the edit is not in the document.
      onNotApplied: async ({ ref, attempt_id, at: d = at(key) }) => {
        if (attempt_id) await update(d, (rev) => dropBlocked(rev, ref, attempt_id));
      },
      // attempt_id: the attempt (or the approved exception) this change
      // settles. It leaves the blocked list in the save that adds the
      // change: no reader of the file sees both, or neither.
      onEdit: async ({ ref, change, attempt_id = null, at: d = at(key) }) => {
        let saved = null;
        let comment = null;
        await update(d, (rev) => {
          const base = withComment(d, rev, ref);
          const r = addChange(base, ref, change);
          saved = r.change;
          comment = findComment(base, ref);
          return (attempt_id && dropBlocked(r.revision, ref, attempt_id)) || r.revision;
        });
        await linkReceipt(d, key, ref, comment, saved);
        sendEvent(key, { type: "review_event", event: "change_recorded", ref, change: saved });
        return saved;
      },
      // blocked_id: the approved exception the edit came from (its item is
      // marked by review_override). attempt_id: the item a normal edit wrote
      // before Word was called (onAttempt); it turns outcome_unknown here
      // with the call id, so a refresh or a restarted daemon can still find
      // out whether Word applied it (Codex review 14).
      onOutcomeUnknown: async ({
        id,
        ref,
        record,
        attempt,
        attempt_id = null,
        blocked_id = null,
        at: d = at(key),
      }) => {
        const own = !blocked_id && !!attempt_id;
        if (own) blocked_id = attempt_id;
        // Before the save below: a late answer is recorded even if it fails.
        if (id) {
          const timer = setTimeout(() => forgetLate(id), lateResultTtlMs);
          timer.unref?.();
          pendingLate.set(id, {
            key,
            at: d,
            ref,
            record,
            blocked_id,
            own,
            timer,
            release: hold(d, ref),
          });
        }
        if (!own) return;
        const revision = await update(d, (rev) =>
          markBlockedUnknown(rev, ref, attempt_id, attempt),
        );
        const saved = findComment(revision, ref).blocked.find((b) => b.id === attempt_id);
        log(`[review] outcome unknown ${attempt_id} kept for refresh id=${id}`);
        sendEvent(key, { type: "review_event", event: "edit_blocked", ref, blocked: saved });
      },
      holdRef: (ref, d = at(key)) => hold(d, ref),
      // Does the task pane compare the paragraph before a tracked edit? One
      // that does not say so in its hello is sent no edit.
      paneChecksTarget: (d = at(key)) => d.caps.includes(TARGET_HASH_CAP),
      // Word applied the edit and its record could not be saved. change: as
      // it was applied; kept for the next load or refresh (recordUnsaved).
      onRecordFailed: async ({ ref, attempt_id, change, at: d = at(key) }) => {
        if (attempt_id && change) {
          if (!unsaved.has(d.path)) unsaved.set(d.path, new Map());
          unsaved.get(d.path).set(attempt_id, { ref, attempt_id, change });
        }
        sendEvent(key, { type: "review_event", event: "record_failed", ref, attempt_id });
      },
      // The task pane refused an edit because its paragraph is not as the
      // snapshot here has it. The snapshot comes from the task pane only
      // (review_refresh), so it is asked to refresh.
      onStaleTarget: async ({ ref }) => {
        sendEvent(key, { type: "review_event", event: "stale_target", ref });
      },
      onBlocked: async ({ ref, args, blocked_by, at: d = at(key) }) => {
        let saved = null;
        await update(d, (rev) => {
          const r = addBlocked(withComment(d, rev, ref), ref, { args, blocked_by });
          saved = r.blocked;
          return r.revision;
        });
        sendEvent(key, { type: "review_event", event: "edit_blocked", ref, blocked: saved });
        return saved;
      },
    };
  }

  // entry: { id, ok, result, error } from bridge.mjs onLateToolResult.
  // Applied (ok, or applied without its comment): recorded like an answer in
  // time. Anything else only logs.
  async function lateToolResult(key, entry) {
    // Held until the result is recorded, so a save in between keeps the comment.
    const p = pendingLate.get(entry.id);
    const held = p ? hold(p.at, p.ref) : () => {};
    try {
      return await recordLate(entry);
    } finally {
      held();
    }
  }
  async function recordLate(entry) {
    const p = forgetLate(entry.id);
    if (!p) {
      log(`[review] late result dropped ${entry.id}: no pending tracked edit`);
      return null;
    }
    // p.record saves the change and takes its item (p.blocked_id) out of
    // the blocked list in one save (onEdit).
    const out = entry.ok ? await p.record(entry.result) : null;
    if (!out?.ok) {
      // Applied, and the save of the change failed: the item stays
      // outcome_unknown in the file for the next refresh.
      if (out?.applied) {
        log(`[review] late result applied, not recorded ${entry.id}: ${p.blocked_id} kept`);
        return null;
      }
      const why = entry.ok ? (out?.code ?? out?.message ?? "not applied") : entry.error;
      log(`[review] late result not applied ${entry.id}: ${why}`);
      // Word answered that the normal edit is not in the document.
      if (p.own) await update(p.at, (rev) => dropBlocked(rev, p.ref, p.blocked_id));
      return null;
    }
    log(`[review] late result recorded ${entry.id}`);
    if (p.blocked_id) {
      log(
        `[review] ${p.own ? "edit" : "override"} ${p.blocked_id} applied late as ${out.change_id}`,
      );
    }
    sendEvent(p.key, {
      type: "review_event",
      event: "late_result_recorded",
      ref: p.ref,
      change_id: out.change_id,
      blocked_id: p.blocked_id,
    });
    return out;
  }

  // outcome: { ok: true } or { ok: false, error } (index.mjs endTurnState).
  // It goes with turn_done itself: the daemon sends the error event of a
  // usage limit after turn_done, so the pane cannot pair the two by order.
  function turnEnded(key, turn, outcome = { ok: true }) {
    turnHolds.get(key)?.();
    turnHolds.delete(key);
    const failed = outcome?.ok === false;
    const error = failed ? String(outcome.error ?? "The turn ended with an error.") : null;
    // outcome.code: what a proposal turn gives the log in place of its
    // sentence, which may quote the manuscript (proposal-turns.mjs fail).
    log(`[review] turn done ${turn.ref}${failed ? ` failed: ${outcome.code ?? error}` : ""}`);
    sendEvent(key, {
      type: "review_event",
      event: "turn_done",
      ref: turn.ref,
      ok: !failed,
      ...(failed ? { error } : {}),
      // What the author should see with a turn that worked (a request of
      // the comment left without a change, proposal-turns.mjs).
      ...(!failed && outcome?.notice ? { notice: String(outcome.notice) } : {}),
    });
  }

  const includeForbid = () => env.CC_OFFICE_REVIEW_PROMPT_FORBID !== "0";

  async function buildFor(d, ref, paragraphs) {
    const rev = await load(d);
    const comment = findComment(rev, ref);
    if (!comment) throw new Error(`No reviewer comment ${ref}. Save the comment list first.`);
    const built = buildReviewPrompt({
      comment,
      paragraphs: Array.isArray(paragraphs) ? paragraphs : [],
      ledger: rev.ledger,
      includeForbid: includeForbid(),
    });
    return { rev, comment, built };
  }

  async function propose(msg, key, host, mine) {
    const asked = clock();
    await turnGate(key);
    // The pane waits wait_ms for this answer and then shows the request as
    // failed. A turn started after that would run behind the author's
    // back, so a request that was held longer starts nothing.
    const waitMs = Number(msg.wait_ms);
    if (waitMs > 0 && clock() - asked >= waitMs) {
      throw new Error(
        `The task pane stopped waiting for this request after ${Math.round(waitMs / 1000)} s, so no turn was started. Press Propose changes again.`,
      );
    }
    const d = at(key);
    const selected = Array.isArray(msg.paragraphs) ? msg.paragraphs : [];
    const proposal = await proposalPreview(key, msg.ref, { at: d, selected });
    const { built, comment, rev } = await buildFor(d, msg.ref, msg.paragraphs);
    // A turn in which no edit can be applied is not started: every
    // office_tracked_edit of it would be refused (review-tools.mjs
    // applyTrackedEdit: TaskPaneOutdated, ForbidRuleBroken), the turn
    // would be spent, and a "Propose all" run would spend one for every
    // comment after it. As the answer of this request the refusal is
    // shown by a task pane of any age and ends its run.
    if (!d.caps.includes(TARGET_HASH_CAP)) {
      throw new Error(
        "This task pane is older than the daemon and does not check the paragraph before an edit, so no change would be applied. Reload the task pane (close and reopen it), then propose again.",
      );
    }
    const broken = brokenForbidRules(rev.ledger);
    if (broken.length) {
      throw new Error(
        `Forbidden rule(s) ${broken.map((b) => b.id).join(", ")} cannot be checked (invalid regex: ${broken.map((b) => `${b.id} ${b.error}`).join("; ")}), so no change would be applied. Fix or remove the rule in the Instruction ledger, then propose again.`,
      );
    }
    if (proposal) {
      // A task pane loaded before proposal turns existed has not told its
      // author what goes to this provider (its first-use notice).
      if (!d.caps.includes(PROPOSAL_CAP)) {
        throw new Error(
          `This task pane is older than the daemon and does not know the proposal turns of ${presetById(proposal.choice?.id)?.label ?? proposal.choice?.id}, so nothing was sent. Reload the task pane (close and reopen it), then propose again.`,
        );
      }
      // A request that cannot be built (no snapshot, too long): the
      // refusal is the answer of this request and no turn is spent.
      if (!proposal.built.ok) throw new Error(proposal.built.message);
    }
    // Nothing is awaited between this check and the queued turn.
    if (mine.stopped) throw new Error("The turn was stopped.");
    starting.delete(key);
    if (proposal) pushTurn(key, host, null, { ref: msg.ref, selected, choice: proposal.choice });
    else pushTurn(key, host, built.text, { ref: msg.ref });
    turnHolds.get(key)?.();
    turnHolds.set(
      key,
      hold(d, msg.ref, {
        reviewer: comment.reviewer,
        text: comment.text,
        section: comment.section,
        requirements: comment.requirements,
      }),
    );
    log(
      formatPromptLog(msg.ref, proposal ? proposal.built : built, {
        paragraphs: msg.paragraphs?.length ?? 0,
      }),
    );
    const revision = await update(d, (rev) =>
      findComment(rev, msg.ref).status === "pending"
        ? setCommentStatus(rev, msg.ref, "in_progress")
        : null,
    );
    return { started: true, ref: msg.ref, started_at: new Date().toISOString(), revision };
  }

  const impl = {
    async review_load(msg, key) {
      const d = at(key);
      const { path } = d;
      let done = [];
      const revision = await update(d, (rev) => {
        const r = recordUnsaved(d, rev);
        done = r.done;
        return clearStale(d, r.revision ?? rev) ?? r.revision;
      });
      await unsavedDone(d, key, revision, done);
      return { path, revision };
    },
    async review_split(msg, key) {
      const d = at(key);
      const text = String(msg.text ?? "");
      // A receipt only when the model was asked (a regex split is no LLM decision).
      // Provider and model of this job, for the split and the extraction after it.
      const job = await oneshotJob(key);
      const shot = job.run;
      const shotModel = job.model;
      let asked = false;
      const ask =
        shot &&
        ((req) => {
          asked = true;
          return shot(req);
        });
      const r = await splitReview(text, { patterns: msg.patterns, oneshot: ask });
      const records = [];
      if (asked) {
        records.push(
          decisionRecord({
            kind: "comment_split",
            input: text,
            model: shotModel,
            output: {
              method: r.method,
              refs: r.comments.map((c) => c.ref),
              labels: r.labels ?? null,
              warnings: r.warnings.map(redactQuotes),
            },
          }),
        );
      }
      // L5: once the comments are known, whichever split made them. A save of
      // the edited list does not come here, so it never runs again on it.
      const x = await extractRequirements(r.comments, { oneshot: shot });
      if (x.asked) {
        records.push(
          decisionRecord({
            kind: "requirement_extract",
            input: x.input,
            model: shotModel,
            output: { requests: x.requests },
          }),
        );
      }
      if (records.length) await receipt(d, records);
      return { ...r, comments: x.comments, warnings: [...r.warnings, ...x.warnings] };
    },
    async review_save_comments(msg, key) {
      const d = at(key);
      const incoming = Array.isArray(msg.comments) ? msg.comments : [];
      const { path } = d;
      const listed = new Set(incoming.map((c) => c?.ref));
      let kept = [];
      // Read inside the chain: what is held when the save runs, not when asked.
      const revision = await update(d, (rev) => {
        const held = heldRefs(path, rev);
        kept = rev.comments.map((c) => c.ref).filter((r) => !listed.has(r) && held.has(r));
        return replaceComments(rev, incoming, held);
      });
      if (kept.length) log(`[review] comments kept, work in progress: ${kept.join(",")}`);
      return { revision, warnings: kept.length ? [keptWarning(kept)] : [] };
    },
    async review_set_status(msg, key) {
      const d = at(key);
      return { revision: await update(d, (rev) => setCommentStatus(rev, msg.ref, msg.status)) };
    },
    async review_set_requirements(msg, key) {
      const d = at(key);
      const texts = Array.isArray(msg.requirements) ? msg.requirements : [];
      return { revision: await update(d, (rev) => setRequirements(rev, msg.ref, texts)) };
    },
    async review_preview(msg, key) {
      const d = at(key);
      const selected = Array.isArray(msg.paragraphs) ? msg.paragraphs : [];
      const proposal = await proposalPreview(key, msg.ref, { at: d, selected });
      if (proposal) {
        // What goes to the provider: its system text, then the prompt.
        const b = proposal.built;
        if (!b.ok) throw new Error(b.message);
        log(formatPromptLog(msg.ref, b, { paragraphs: selected.length, preview: true }));
        return { prompt: `${b.system}\n\n${b.prompt}` };
      }
      const { built } = await buildFor(d, msg.ref, msg.paragraphs);
      log(
        formatPromptLog(msg.ref, built, { paragraphs: msg.paragraphs?.length ?? 0, preview: true }),
      );
      return { prompt: built.text };
    },
    async review_propose(msg, key, host) {
      // From here until the turn is queued a Stop is remembered (cancelStarting).
      const mine = { stopped: false };
      starting.set(key, mine);
      try {
        return await propose(msg, key, host, mine);
      } finally {
        if (starting.get(key) === mine) starting.delete(key);
      }
    },
    async review_refresh(msg, key) {
      const d = at(key);
      const snapshot = msg.snapshot ?? { paragraphs: [], comments: [] };
      const texts = snapshotTexts(snapshot);
      let rows = [];
      let external = [];
      let resolved = [];
      let dropped = [];
      let done = [];
      const revision = await update(d, (file) => {
        // What this daemon knows Word applied goes in first, as it was applied.
        const u = recordUnsaved(d, file);
        done = u.done;
        const old = u.revision ?? file;
        const r0 = resolveUnknownOverrides(clearStale(d, old) ?? old, snapshot);
        const rev = r0.revision;
        resolved = r0.resolved;
        dropped = r0.dropped;
        rows = classifyRevision(rev, snapshot);
        let next = applyStates(rev, rows);
        // L4: changes since the last refresh that no review change explains.
        // The first refresh only stores the snapshot.
        if (rev.snapshot) {
          const changes = rows.map((row) => ({
            ...rev.comments
              .find((c) => c.ref === row.ref)
              ?.changes.find((ch) => ch.id === row.change_id),
            ...row,
          }));
          const r = addExternal(next, outsideChanges(rev.snapshot.paragraphs, texts, changes));
          next = r.revision;
          external = r.added;
        }
        return setSnapshot(next, texts, new Date(), snapshotIds(snapshot));
      });
      await unsavedDone(d, key, revision, done);
      for (const x of resolved) {
        forgetLate(x.tool_call_id);
        log(
          `[review] ${x.own ? "edit" : "override"} ${x.blocked_id} found applied in refresh as ${x.change_id} state=${x.state}`,
        );
      }
      // Not forgotten as a late result: should Word answer after all, the
      // answer is still recorded.
      for (const x of dropped) {
        log(
          `[review] edit ${x.blocked_id} not applied (old text in place, no rationale comment); dropped`,
        );
      }
      const counts = summarizeStates(rows);
      const toolComments = (snapshot.comments ?? []).filter((c) =>
        TOOL_COMMENT_PREFIX.test(String(c.content ?? "")),
      ).length;
      const userComments = (snapshot.comments ?? []).length - toolComments;
      log(
        `[review] refresh changes=${rows.length} tool_comments=${toolComments} user_comments=${userComments} ` +
          `states=${Object.entries(counts)
            .map(([k, v]) => `${k}:${v}`)
            .join(",")}` +
          ` external=${external.length}`,
      );
      for (const x of external) {
        log(
          `[review] external ${x.id} kind=${x.kind} paragraph=${x.paragraph_index ?? x.old_index} lost=${x.lost_facts.length}` +
            (x.lost_facts.length ? ` facts=${x.lost_facts.join("; ")}` : ""),
        );
      }
      return { rows, counts, external, revision };
    },
    async review_external_dismiss(msg, key) {
      const d = at(key);
      return { revision: await update(d, (rev) => dismissExternal(rev, msg.id)) };
    },
    // msg.ids: Word comment ids the task pane's comment cleanup deleted.
    async review_comments_cleaned(msg, key) {
      const d = at(key);
      const ids = Array.isArray(msg.ids) ? msg.ids : [];
      const revision = await update(d, (rev) => markCommentsCleaned(rev, ids));
      log(`[review] comments cleaned ids=${ids.length}`);
      return { revision };
    },
    // A change Word applied without its rationale comment (comment_missing)
    // gets it later, from the Review tab ("Add rationale comment"): this
    // gives the text the comment would have had and where the change is,
    // the task pane anchors the comment (tools-word.js addChangeComment),
    // and review_comment_set records its id. No text of the manuscript is
    // changed on this way.
    async review_comment_text(msg, key) {
      const d = at(key);
      const rev = await load(d);
      const ch = findComment(rev, msg.ref)?.changes.find((x) => x.id === msg.change_id);
      if (!ch) throw new Error(`No change ${msg.change_id} under ${msg.ref}`);
      if (ch.word_comment_id != null) {
        throw new Error(`${ch.id} already has its rationale comment.`);
      }
      return {
        change_id: ch.id,
        // As applyTrackedEdit composed it for the edit (the stored rationale
        // holds the note of an approved exception; lost facts on the last line).
        comment_text: formatRationaleComment(msg.ref, ch.rationale, { lost: ch.lost_facts }),
        // The task pane reuses none of these for this change.
        known_comment_ids: recordedCommentIds(rev),
        change: {
          kind: ch.kind,
          find: ch.find,
          text: ch.text,
          paragraph_id: ch.paragraph_id,
          paragraph_index: ch.paragraph_index,
        },
      };
    },
    async review_comment_set(msg, key) {
      const d = at(key);
      const revision = await update(d, (rev) =>
        setChangeComment(rev, msg.ref, msg.change_id, msg.word_comment_id),
      );
      log(`[review] comment added later ${msg.change_id} comment=${msg.word_comment_id}`);
      return { revision };
    },
    async review_override(msg, key) {
      const d = at(key);
      // Claimed inside the revision chain: an overlapping override of the
      // same item finds it applying and is refused before Word is called.
      // An applying item left by a stopped daemon is taken as outcome_unknown
      // first, so it gets the Refresh hint, not "already being applied".
      // The id joins `active` in the same step, before any other request of
      // the chain can read the file.
      let blocked = null;
      const active = activeFor(d);
      const held = applyingKey(msg.ref, msg.blocked_id);
      try {
        await update(d, (rev) => {
          const r = claimBlocked(clearStale(d, rev) ?? rev, msg.ref, msg.blocked_id, {
            force: msg.force === true,
          });
          blocked = r.blocked;
          active.add(held);
          return r.revision;
        });
      } catch (e) {
        if (blocked) active.delete(held); // claimed, then the save failed
        throw e;
      }
      // Back to the state and the attempt it had before this try.
      const restore = () =>
        update(d, (x) => releaseBlocked(x, msg.ref, blocked.id, blocked.state, blocked.attempt));
      try {
        const base = trackedEditDeps(key, d);
        let attempt = null;
        const deps = {
          ...base,
          call: (name, args) => callTool(name, args, key),
          // The claimed item is the attempt: the form the edit would have in
          // the document goes on it before Word is called, so a daemon that
          // stops mid-call leaves an item a refresh can settle.
          onAttempt: async (x) => {
            await update(d, (rev) => setBlockedAttempt(rev, msg.ref, blocked.id, x.attempt));
            return { id: blocked.id };
          },
          onNotApplied: async () => {}, // restore() below puts the item back
          onOutcomeUnknown: (x) => {
            attempt = x.attempt;
            return base.onOutcomeUnknown({ ...x, blocked_id: blocked.id });
          },
        };
        let r;
        try {
          r = await applyTrackedEdit(
            deps,
            blocked.args,
            { ref: msg.ref },
            { override: { blocked_by: blocked.blocked_by } },
          );
        } catch (e) {
          await restore();
          throw e;
        }
        if (r.code === "OutcomeUnknown" && attempt) {
          log(`[review] override outcome unknown ${blocked.id} id=${attempt.tool_call_id}`);
          const revision = await update(d, (x) =>
            markBlockedUnknown(x, msg.ref, blocked.id, attempt),
          );
          return { applied: false, outcome_unknown: true, result: r, revision };
        }
        // Word applied the edit and its record could not be saved, or the
        // task pane sent no result at all: the edit is, or may be, in the
        // document. Back to "blocked" would offer Allow exception again and
        // a second press would apply it twice. The item turns
        // outcome_unknown with the attempt written before the call, and a
        // refresh records the change.
        if (!r.ok && (r.applied === true || r.code === "EmptyResult")) {
          log(`[review] override ${blocked.id} not recorded (${r.code}); kept for refresh`);
          const revision = await update(d, (x) => {
            const b = findComment(x, msg.ref)?.blocked.find((y) => y.id === blocked.id);
            return b?.attempt ? markBlockedUnknown(x, msg.ref, blocked.id, b.attempt) : null;
          });
          return { applied: false, outcome_unknown: true, result: r, revision };
        }
        if (!r.ok) {
          const revision = await restore();
          return { applied: false, result: r, revision };
        }
        log(`[review] override ${msg.blocked_id} approved for ${blocked.blocked_by.join(",")}`);
        // The save of the change took the item out of the blocked list (onEdit).
        const revision = await update(d, (x) => dropBlocked(x, msg.ref, msg.blocked_id));
        return { applied: true, result: r, revision };
      } finally {
        active.delete(held);
      }
    },
    async ledger_add(msg, key) {
      const d = at(key);
      let item = null;
      const revision = await update(d, (rev) => {
        const r = addLedgerItem(rev, msg.item);
        item = r.item;
        return r.revision;
      });
      log(`[review] ledger add ${item.id} kind=${item.kind} check=${item.check.type}`);
      return { item, revision };
    },
    async ledger_update(msg, key) {
      const d = at(key);
      return {
        revision: await update(d, (rev) => updateLedgerItem(rev, msg.id, msg.patch ?? {})),
      };
    },
    async ledger_remove(msg, key) {
      const d = at(key);
      return { revision: await update(d, (rev) => removeLedgerItem(rev, msg.id)) };
    },
    // P15: rules only (no ids, no check results) to <stem>.ledger.json next
    // to the manuscript; an older export of the same name is replaced.
    async ledger_export(msg, key) {
      const d = at(key);
      const out = ledgerExportDoc(await load(d));
      const path = sidecarPath(d.doc, ".ledger.json");
      await writeJsonAtomic(path, out);
      log(`[review] ledger export path=${path} count=${out.items.length}`);
      return { path, count: out.items.length };
    },
    // msg.doc: the parsed file ({ items: [{ kind, text, scope, check }] }).
    // msg.mode: "append" (default; known rules are skipped) or "replace".
    async ledger_import(msg, key) {
      const d = at(key);
      const mode = msg.mode ?? "append";
      let added = [];
      let skipped = [];
      const revision = await update(d, (rev) => {
        const r = importLedgerItems(rev, msg.doc?.items, { mode });
        added = r.added;
        skipped = r.skipped;
        return r.revision;
      });
      log(`[review] ledger import added=${added.length} skipped=${skipped.length} mode=${mode}`);
      return {
        added: added.length,
        skipped: skipped.length,
        added_ids: added.map((i) => i.id),
        revision,
      };
    },
    // P15: rule candidates from the review text (ledger-suggest.mjs). Nothing
    // is stored: the author picks candidates in the pane, which sends them
    // back through ledger_import.
    async ledger_suggest(msg, key) {
      const text = String(msg.text ?? "");
      if (!text.trim()) throw new Error("Paste or load the review text first.");
      const d = at(key); // before the model call: the receipt goes next to this document
      const job = await oneshotJob(key);
      const shotModel = job.model;
      const r = await suggestLedger(text, { oneshot: job.run });
      log(
        `[review] ledger suggest method=${r.method} candidates=${r.candidates.length} warnings=${r.warnings.length}`,
      );
      // A receipt only when the model was asked; sentence numbers, no text.
      if (r.asked) {
        await receipt(d, [
          decisionRecord({
            kind: "ledger_suggest",
            input: r.input,
            model: shotModel,
            output: {
              method: r.method,
              candidates: r.candidates.length,
              labels: r.candidates.map((c) => ({
                i: c.sentence_index,
                kind: c.kind,
                check: c.check.type,
                why: c.why,
              })),
              warnings: r.warnings.length,
            },
          }),
        ]);
      }
      return { candidates: r.candidates, method: r.method, warnings: r.warnings };
    },
    async ledger_check(msg, key) {
      const d = at(key);
      // Taken on arrival, before the file read: a check that reads the file
      // later is still the older one. An older check never lowers the mark.
      const gen = ++lastCheckGen;
      const rev = await load(d);
      const wanted = Array.isArray(msg.ids) && msg.ids.length ? new Set(msg.ids) : null;
      const items = rev.ledger.filter((i) => !wanted || wanted.has(i.id));
      const { path } = d;
      const genKey = (id) => `${path}\0${id}`;
      for (const i of items) {
        if ((checkGen.get(genKey(i.id)) ?? 0) < gen) checkGen.set(genKey(i.id), gen);
      }
      const started = new Map(items.map((i) => [i.id, ruleFingerprint(i)]));
      const paragraphs = msg.snapshot?.paragraphs ?? [];
      // L10: changed paragraphs only unless msg.full (the "Full check" button).
      const decisions = [];
      // Every batch of this check runs on the provider and model of its start.
      const job = await oneshotJob(key);
      const shotModel = job.model;
      const outcomes = await checkLedgerIncremental(items, paragraphs, {
        oneshot: job.run,
        full: msg.full === true,
        onDecision: (x) =>
          decisions.push(
            decisionRecord({
              kind: "ledger_check",
              input: { instruction: x.instruction, paragraph: x.paragraph },
              model: shotModel,
              output: {
                item_id: x.item_id,
                paragraph_index: x.paragraph_index,
                verdict: x.verdict,
                ...(x.error ? { error: String(x.error).slice(0, 200) } : {}),
              },
            }),
          ),
      });
      await receipt(d, decisions);
      const checkedAt = new Date().toISOString();
      for (const r of outcomes) {
        const unsure = r.unsure.length ? ` unsure=${r.unsure.length}` : "";
        log(
          `[review] ledger check ${r.id} status=${r.status} violations=${r.violations.length}${unsure}`,
        );
        log(`[review] ledger scope ${r.id} mode=${r.mode} paragraphs=${r.checked}/${r.total}`);
      }
      const revision = await update(d, (x) => {
        let next = x;
        for (const r of outcomes) {
          // The rule may have been edited or removed while the check ran: its
          // result then belongs to the old rule and is dropped. An edited text
          // or check already left the item unchecked (updateLedgerItem), so
          // the next check reads every paragraph; a newer check's result is
          // kept.
          const now = x.ledger.find((i) => i.id === r.id);
          if (!now || ruleFingerprint(now) !== started.get(r.id)) {
            log(
              `[review] ledger check ${r.id} dropped: rule ${now ? "changed" : "removed"} during the check`,
            );
            continue;
          }
          // A check of the same rule started later owns the item's result.
          if (checkGen.get(genKey(r.id)) !== gen) {
            log(`[review] ledger check ${r.id} dropped: a newer check started`);
            continue;
          }
          next = updateLedgerItem(next, r.id, {
            status: r.status,
            violation_count: r.violations.length,
            checked_at: checkedAt,
            violations: r.violations,
            unsure: r.unsure,
            checked_hashes: r.checked_hashes,
          });
        }
        return next;
      });
      return { outcomes, revision };
    },
    async comments_backup(msg, key) {
      const d = at(key);
      const { doc } = d;
      const backup = msg.backup;
      if (!backup || !Array.isArray(backup.comments))
        throw new Error("backup.comments is required");
      if (backup.count !== backup.comments.length) {
        throw new Error(`backup count ${backup.count} != ${backup.comments.length} comments`);
      }
      // Replies are deleted with their thread; count stays the parent count.
      const replies = backup.comments.reduce(
        (n, c) => n + (Array.isArray(c.replies) ? c.replies.length : 0),
        0,
      );
      if (backup.reply_count != null && backup.reply_count !== replies) {
        throw new Error(`backup reply_count ${backup.reply_count} != ${replies} replies`);
      }
      // One backup at a time per file name: picking a free name and writing
      // it is not atomic, so an overlapping backup could take the same name.
      const base = commentsBackupPathFor(doc);
      const prev = backupChains.get(base) ?? Promise.resolve();
      const run = prev
        .catch(() => {})
        .then(async () => {
          let path = base;
          if (await exists(path)) {
            const stamp = new Date().toISOString().replace(/[:.]/g, "-");
            const stamped = (n) =>
              base.replace(
                /\.comments_backup\.json$/,
                `.comments_backup.${stamp}${n ? `-${n}` : ""}.json`,
              );
            let n = 0;
            path = stamped(n);
            while (await exists(path)) path = stamped(++n);
          }
          await writeJsonAtomic(path, backup);
          return path;
        });
      backupChains.set(base, run);
      const path = await run;
      log(
        `[review] comments_backup path=${path} count=${backup.comments.length} replies=${replies}`,
      );
      return { path, count: backup.comments.length, reply_count: replies };
    },
  };

  const handlers = {};
  for (const [type, fn] of Object.entries(impl)) {
    if (JOB_TYPES.has(type)) {
      // Returns once accepted; the job runs on and ends in job_done.
      handlers[type] = async (msg, reply, key, host) => {
        const job_id = randomUUID();
        reply({
          type: `${type}_result`,
          request_id: msg.request_id,
          ok: true,
          accepted: true,
          job_id,
        });
        const done = (result) =>
          sendEvent(key, { type: "review_event", event: "job_done", job_id, result });
        fn(msg, key, host)
          .then(
            (out) => done({ ok: true, ...out }),
            (e) => done({ ok: false, error: e?.message ?? String(e) }),
          )
          // Nobody awaits the job: a failed send must not end the daemon.
          .catch((e) => log(`[review] job_done not delivered: ${e?.message ?? e}`));
      };
      continue;
    }
    handlers[type] = async (msg, reply, key, host) => {
      try {
        const out = await fn(msg, key, host);
        reply({ type: `${type}_result`, request_id: msg.request_id, ok: true, ...out });
      } catch (e) {
        reply({
          type: `${type}_result`,
          request_id: msg.request_id,
          ok: false,
          error: e?.message ?? String(e),
        });
      }
    };
  }
  return {
    handlers,
    trackedEditDeps,
    turnEnded,
    lateToolResult,
    // For the proposal runner (index.mjs): the document handle of a pane as
    // it is now, the forbid switch of the prompt, and a Stop before a turn.
    at,
    includeForbid,
    cancelStarting,
  };
}
