// Review-tab message handlers and the recording side of office_tracked_edit.
// index.mjs wires the dependencies; everything here is testable with fakes.
//
// Messages (task pane -> daemon, reply type is `<type>_result` with the same
// request_id, the existing sendRequest pattern in taskpane.js):
//   review_load, review_split, review_save_comments, review_set_status,
//   review_set_requirements, review_question_answer, review_check_paragraphs, review_preview, review_propose, review_refresh,
//   review_external_dismiss, review_comments_cleaned, review_override,
//   review_comment_text, review_comment_set, review_comment_retry,
//   ledger_add, ledger_update, ledger_remove, ledger_check, ledger_export,
//   ledger_import, ledger_suggest, comments_backup,
//   review_link_candidates, review_link_join, review_link_leave
// Push (daemon -> task pane, no request_id):
//   { type: "review_event", event: "change_recorded" | "edit_blocked" | "turn_done"
//     | "late_result_recorded" | "stale_target" | "undo_took_more" | "record_failed" | "job_done"
//     | "linked_changed", ... }
//   linked_changed: { doc }. The link between a manuscript and a supplement, or
//   what the manuscript owns, is not as it was; it goes to the OTHER open panes
//   of the revision, which load their file again.
//   turn_done: { ref, ok, error? }. ok false: the turn ended on a usage limit,
//   a refusal, an API error, an empty turn, a stop or a crash (error says which).
// Jobs (JOB_TYPES): the reply is { ok: true, accepted: true, job_id } at once
// and the result follows as { event: "job_done", job_id, result }, result
// being what the reply would have held ({ ok: true, ... } or { ok: false,
// error }). Their model calls can outlast any request timeout of the pane.
// handler-wrap.mjs wraps both kinds.

import { access, stat } from "node:fs/promises";
import { basename, dirname, resolve } from "node:path";
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
  addQuestion,
  answerQuestion,
  authorAnswers,
  setRequirementOutside,
  addRewriteStopped,
  setSameTerm,
  setFollowThrough,
  byCreated,
  ruleFingerprint,
  mirrorFromMain,
  setPartOf,
  leaveMain,
  addPart,
  removePart,
  setDocState,
} from "./revision-store.mjs";
import {
  storedPartPath,
  resolvePart,
  linkCandidates,
  claimedBy,
  mainDocFrom,
  loadMain,
  loadParts,
  partChangeRefs,
  linkedSummary,
  refuseInPart,
  otherFiles,
  linkedChanges,
  linkedAnswers,
  outsideNotes,
  citedInMain,
  partsForLetter,
} from "./revision-parts.mjs";
import {
  classifyRevision,
  classifyChange,
  summarizeStates,
  normalizeParaText,
} from "./revision-classify.mjs";
import { extractRequirements, splitReview } from "./review-split.mjs";
import { buildReviewPrompt, formatPromptLog, sameTermReport } from "./review-prompt.mjs";
import { checkLedgerIncremental, brokenForbidRules } from "./ledger-check.mjs";
import { suggestLedger } from "./ledger-suggest.mjs";
import { checkParagraphs, followNote, followPlan, rewriteNote } from "./paragraph-check.mjs";
import { snapshotTexts, snapshotIds, snapshotRead, outsideChanges } from "./doc-diff.mjs";
import {
  applyTrackedEdit,
  currentTexts,
  narrowedAttempt,
  formatRationaleComment,
  TOOL_COMMENT_PREFIX,
  TARGET_HASH_CAP,
} from "./review-tools.mjs";
import { appendDecisions, decisionRecord, redactQuotes } from "./decision-log.mjs";
import { wrapHandlers } from "./handler-wrap.mjs";
import { effectiveProfile, rationaleLang, splitPatterns } from "./journal-profile.mjs";
import { snapshotHash } from "./response-quotes.mjs";

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
        // A part of a rewrite that adds no comment is under the comment of
        // its paragraph (form.word_comment_id).
        const change = {
          ...form,
          word_comment_id: comment ? String(comment.id) : (form.word_comment_id ?? null),
        };
        const row = classifyChange(
          change,
          snapshot,
          // With the changes resolved earlier in this loop.
          next.comments.flatMap((x) => x.changes),
        );
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
        comment_missing: comment || change.word_comment_id != null ? null : "CommentNotFound",
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
// One request of the paragraph check may hold several paragraphs and goes
// to the model of the edits, which thinks before it answers.
const PARAGRAPH_CHECK_TIMEOUT_MS = 300_000;
const JOB_TYPES = new Set([
  "review_split",
  "ledger_check",
  "ledger_suggest",
  "review_check_paragraphs",
]);

// Requests that change what the manuscript owns (the comment list, the
// requests, the ledger rules): refused in the file of a supplement.
const MAIN_ONLY = [
  "review_split",
  "review_save_comments",
  "review_set_requirements",
  "ledger_add",
  "ledger_update",
  "ledger_remove",
  "ledger_import",
  "ledger_suggest",
];
// Of those, the ones that write nothing themselves.
const MAIN_ONLY_READS = new Set(["review_split", "ledger_suggest"]);

// What a document was at a status refresh, for its review file (doc_state,
// revision-store.mjs DocState). snapshot: what the task pane collected.
// saved: what Word said about unsaved changes then (Word.Document.saved;
// anything but a boolean is "not known"). file: { mtime, bytes } of the file
// on disk, or null.
export function docState(snapshot, { saved = null, file = null, now = new Date() } = {}) {
  const comments = snapshot?.comments ?? [];
  return {
    at: now.toISOString(),
    text_hash: snapshotHash(snapshot),
    tracked_paragraphs: (snapshot?.paragraphs ?? []).filter(
      (p) => (p.w_ins ?? 0) + (p.w_del ?? 0) > 0,
    ).length,
    comments: comments.length,
    tool_comments: comments.filter((c) => TOOL_COMMENT_PREFIX.test(String(c.content ?? ""))).length,
    track_changes_on:
      typeof snapshot?.track_changes_on === "boolean" ? snapshot.track_changes_on : null,
    saved: typeof saved === "boolean" ? saved : null,
    file,
  };
}
// The changes of a review file as the other files of its revision see them.
const changeStates = (rev) =>
  JSON.stringify(rev.comments.map((c) => [c.ref, c.changes.map((ch) => [ch.id, ch.state])]));

// The comments whose stale marks are not what they were, and not empty.
function staleChanges(before, after) {
  const was = new Map(before.comments.map((c) => [c.ref, (c.response?.stale ?? []).join(",")]));
  return after.comments
    .map((c) => ({ ref: c.ref, stale: c.response?.stale ?? [] }))
    .filter((x) => x.stale.length && x.stale.join(",") !== was.get(x.ref));
}

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
  // Default language of the settings file, between the profile and the
  // environment (journal-profile.mjs rationaleLang): () => string | null.
  langDefault = null,
  // (revision, snapshot) => revision, run in a refresh once the states of
  // the changes are set (index.mjs: the stale marks of saved responses).
  afterRefresh = null,
  clock = Date.now,
  getModel = () => null,
  // The model the session of the pane runs on, for the receipt of a change
  // a review turn made (after a refusal fallback it is not the one asked for).
  getRunModel = getModel,
  // The manuscript and the session of a pane (review-prompt.mjs
  // manuscriptBlock): manuscriptSeen(key) gives the paragraph hashes the
  // session was last given, or null; manuscriptGiven(key, hashes) is called
  // when a turn that carries the manuscript is queued.
  manuscriptSeen = () => null,
  manuscriptGiven = () => {},
  sendEvent = () => {},
  // The panes that are connected now: () => [{ key, host, activeDoc }]
  // (bridge.mjs listPanes), for the other panes of a revision.
  listPanes = () => [],
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

  // ---- supplementary documents (revision-parts.mjs) ---------------------------
  // The manuscript of the document of a pane, when the document is a
  // supplement: { partOf, main }. null for any other document. Plain reads.
  // A supplement whose file lost the link (an earlier daemon saved it) gets
  // it back from the manuscript in its folder that still lists it.
  async function linkOf(d) {
    const rev = await load(d);
    const partOf = rev.part_of ?? (rev.parts?.length ? null : await claimedBy(resolve(d.doc)));
    return partOf ? { partOf, main: await loadMain(d.doc, partOf) } : null;
  }
  // The file of a supplement with its link and the copy of what the
  // manuscript owns as they are now; null when it holds both already. A
  // manuscript whose review file is gone or lists no comment is not copied
  // from: a file that was moved must not empty the list of the supplement.
  function inLine(rev, link) {
    let next = rev.part_of ? rev : setPartOf(rev, link.partOf);
    const main = link.main.revision;
    if (main?.comments.length) next = mirrorFromMain(next, main);
    return JSON.stringify(next) === JSON.stringify(rev) ? null : next;
  }
  const partEntry = (main, doc) =>
    (main.revision?.parts ?? []).find((p) => resolvePart(dirname(main.doc), p.doc) === doc);
  // What the pane is told about the link, with one sentence when the
  // manuscript cannot be read from.
  async function linkedReply(d, revision, link = null) {
    const main = link?.main;
    const note = !main
      ? null
      : !main.revision
        ? `The review file of ${main.name} was not found, so this list may be out of date. Open ${main.name} once, or leave this revision.`
        : main.revision.comments.length === 0
          ? `${main.name} lists no reviewer comments now, so this list was left as it was.`
          : null;
    // Where the requests stand in the other files of the revision, as their
    // review files have it. A file that cannot be read tells nothing.
    let changes = [];
    try {
      changes = linkedChanges(await otherFiles(d.doc, revision));
    } catch (e) {
      log(`[review] linked files not read: ${e?.message ?? e}`);
    }
    // For a supplement: which of its changes the letter cites, as the
    // manuscript's review file has them (the Response tab of the supplement
    // shows them), and its own state for the letter: was its last refresh
    // taken on a saved document.
    const cited = main?.revision ? citedInMain(main.revision, basename(d.doc)) : null;
    return {
      ...linkedSummary(d.doc, revision),
      note,
      changes,
      ...(link
        ? { cited_in_main: cited ?? [], saved_at_refresh: revision.doc_state?.saved ?? null }
        : {}),
    };
  }
  // The other open panes of a revision load their file again.
  function tellLinked(d, docs) {
    const want = new Set(docs.map((x) => resolve(x)));
    for (const p of listPanes()) {
      const doc = p.key === d.key ? null : localDocPath(p.activeDoc, env);
      if (doc && want.has(resolve(doc))) {
        sendEvent(p.key, { type: "review_event", event: "linked_changed", doc: basename(d.doc) });
      }
    }
  }
  const partDocs = (doc, rev) => (rev?.parts ?? []).map((p) => resolvePart(dirname(doc), p.doc));
  // A change of this document was recorded, decided or found in a refresh:
  // the other open panes of its revision load again and show where the
  // request stands now. Never throws: what Word applied stays recorded.
  async function tellOthers(d, rev = null) {
    try {
      const own = rev ?? (await load(d));
      if (!own.part_of && !own.parts?.length) return;
      tellLinked(
        d,
        (await otherFiles(d.doc, own)).map((f) => f.doc),
      );
    } catch (e) {
      log(`[review] linked panes not told: ${e?.message ?? e}`);
    }
  }
  // What the author answered in the other files of the revision.
  async function answersElsewhere(d, rev) {
    if (!rev.part_of && !rev.parts?.length) return [];
    try {
      return linkedAnswers(await otherFiles(d.doc, rev));
    } catch (e) {
      log(`[review] linked answers not read: ${e?.message ?? e}`);
      return [];
    }
  }

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
    if (done.some((h) => h.saved)) await tellOthers(d, revision);
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
        model: getRunModel(key),
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
          // What that refresh could read of the fields (field-text.js
          // fieldsUnread). A snapshot from before this was recorded says
          // nothing about it: unknown.
          fields_read: {
            unreviewed: rev.snapshot?.unreviewed ?? null,
            fields: rev.snapshot?.fields ?? null,
            unknown: !!rev.snapshot && !Array.isArray(rev.snapshot.unreviewed),
          },
          comment_ids: recordedCommentIds(rev),
          // What the author answered to questions of this review, in this
          // file and in the other files of its revision: a source of numbers
          // for an edit, like the comment.
          author_texts: [
            ...authorAnswers(rev),
            ...(await answersElsewhere(d, rev)).map((q) => q.answer),
          ],
          recent_changes: since
            ? rev.comments
                .flatMap((x) => x.changes)
                .filter((ch) => ch.created_at > since)
                .sort(byCreated)
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
        const revision = await update(d, (rev) => {
          const base = withComment(d, rev, ref);
          const r = addChange(base, ref, change);
          saved = r.change;
          comment = findComment(base, ref);
          return (attempt_id && dropBlocked(r.revision, ref, attempt_id)) || r.revision;
        });
        await linkReceipt(d, key, ref, comment, saved);
        sendEvent(key, { type: "review_event", event: "change_recorded", ref, change: saved });
        await tellOthers(d, revision);
        return saved;
      },
      // office_ask_author: the question goes under the comment and the pane
      // loads the file again.
      onQuestion: async ({ ref, question, at: d = at(key) }) => {
        let saved = null;
        await update(d, (rev) => {
          const r = addQuestion(withComment(d, rev, ref), ref, question);
          saved = r.question;
          return r.revision;
        });
        sendEvent(key, { type: "review_event", event: "comment_noted", ref });
        return saved;
      },
      // office_requirement_outside: the note goes under the comment.
      onRequirementOutside: async ({ ref, note, at: d = at(key) }) => {
        await update(d, (rev) => setRequirementOutside(withComment(d, rev, ref), ref, note));
        sendEvent(key, { type: "review_event", event: "comment_noted", ref });
      },
      // office_tracked_rewrite stopped after some of its parts went in.
      onRewriteStopped: async ({ ref, record, at: d = at(key) }) => {
        await update(d, (rev) => addRewriteStopped(withComment(d, rev, ref), ref, record));
        sendEvent(key, { type: "review_event", event: "comment_noted", ref });
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
      // A capability the task pane named in its hello (review-tools.mjs REWRITE_CAP).
      paneCan: (cap, d = at(key)) => d.caps.includes(cap),
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
      // The task pane took an edit back that would have damaged a field,
      // and Word took a pending change next to it back as well. That change
      // is still on record as pending: the pane refreshes (as for a stale
      // target), which reads its state from the document, and shows the
      // author one sentence on the card of the comment that was proposed.
      onUndoTookMore: async ({ ref, paragraph_index = null }) => {
        log(`[review] undo took more ${ref} paragraph=${paragraph_index ?? "?"}: refresh asked`);
        sendEvent(key, { type: "review_event", event: "undo_took_more", ref, paragraph_index });
      },
      onBlocked: async ({ ref, args, blocked_by, follow, at: d = at(key) }) => {
        let saved = null;
        await update(d, (rev) => {
          const r = addBlocked(withComment(d, rev, ref), ref, { args, blocked_by, follow });
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
    log(`[review] turn done ${turn.ref}${failed ? ` failed: ${error}` : ""}`);
    const done = () =>
      sendEvent(key, {
        type: "review_event",
        event: "turn_done",
        ref: turn.ref,
        ok: !failed,
        ...(failed ? { error } : {}),
      });
    // outcome.text: the answer of the turn (index.mjs gives it for a turn
    // that ended with a result). The paragraphs its last line names as
    // still holding the term are saved first, so the pane that loads on
    // turn_done has them; a save that fails only loses that list.
    // A follow-through turn ("Check other places") is saved before that,
    // for the same reason.
    const saves = [];
    if (turn.follow) {
      saves.push(() =>
        followEnded(turn, failed).catch((e) =>
          log(`[review] follow-through not saved ${turn.ref}: ${e?.message ?? e}`),
        ),
      );
    }
    // The answer of a follow-through turn is about the places it was given:
    // without the line on the term it leaves the list of the turn before.
    const noTermLine = turn.follow && !/Same term still at:/i.test(outcome?.text ?? "");
    if (typeof outcome?.text === "string" && !noTermLine) {
      saves.push(() =>
        Promise.resolve()
          .then(() => {
            const d = at(key);
            return update(d, (rev) =>
              setSameTerm(rev, turn.ref, sameTermReport(outcome.text, rev.snapshot)),
            );
          })
          .catch((e) => log(`[review] same-term list not saved ${turn.ref}: ${e?.message ?? e}`)),
      );
    }
    if (!saves.length) return done();
    saves.reduce((p, save) => p.then(save), Promise.resolve()).finally(done);
  }

  // The end of a follow-through turn: what it added under its comment (the
  // changes and questions that were not there when it began) goes to the
  // log as counts, and the comment is marked as checked. A turn that ended
  // with an error marks nothing, so the check can be run again. It is saved
  // to the document the turn began in (turn.follow.doc), whatever the pane
  // shows by now.
  async function followEnded(turn, failed) {
    const { candidate_ids, answer_ids = [], had, doc } = turn.follow;
    let added = { changes: 0, questions: 0 };
    await update(doc, (rev) => {
      const c = findComment(rev, turn.ref);
      if (!c) return null;
      const change_ids = c.changes.map((ch) => ch.id).filter((id) => !had.changes.includes(id));
      const questions = (c.questions ?? []).filter((x) => !had.questions.includes(x.id)).length;
      added = { changes: change_ids.length, questions };
      if (failed) return null;
      return setFollowThrough(rev, turn.ref, {
        at: new Date().toISOString(),
        candidate_ids,
        change_ids,
        questions,
        seen_change_ids: c.changes
          .filter((ch) => ch.state !== "rejected" && ch.state !== "unlocated")
          .map((ch) => ch.id),
        ...(answer_ids.length ? { seen_answer_ids: answer_ids } : {}),
      });
    });
    log(
      `[review] follow ${turn.ref} candidates=${candidate_ids.length} changes=${added.changes} questions=${added.questions}`,
    );
  }

  // The model of a receipt: the one the last request of the job ran on,
  // when the router knows it (a Claude job; another provider's name stays
  // as the job gives it, with the provider in it).
  const ranOf = (job) =>
    job.provider == null || job.provider === "claude"
      ? (job.used?.()?.ran ?? job.model)
      : job.model;

  const includeForbid = () => env.CC_OFFICE_REVIEW_PROMPT_FORBID !== "0";

  // The journal profile of the paper (journal-profile.mjs). A profile that
  // cannot be read does not stop a review: the language falls back to the
  // settings and environment default, the split to its built-in patterns.
  async function profileOf(d, rev) {
    const settingsDefault = langDefault?.() ?? null;
    try {
      const { profile } = await effectiveProfile(rev, {
        docDir: dirname(d.doc),
        env,
        settingsDefault,
      });
      return profile;
    } catch (e) {
      log(`[review] profile not loaded: ${e?.message ?? e}`);
      return { rationale_lang: rationaleLang({}, { env, settingsDefault }) };
    }
  }

  async function buildFor(d, ref, paragraphs, key) {
    // In a supplement the copy of what the manuscript owns is brought up to
    // date before every turn (and before its preview), in a save of its own;
    // the manuscript is read before that, plainly.
    const link = await linkOf(d);
    const rev = link ? await update(d, (file) => inLine(file, link)) : await load(d);
    const comment = findComment(rev, ref);
    if (!comment) throw new Error(`No reviewer comment ${ref}. Save the comment list first.`);
    const profile = await profileOf(d, rev);
    const built = buildReviewPrompt({
      comment,
      // What the other comments of the review ask and what is on record for them.
      others: rev.comments,
      // The manuscript of the last status refresh, with the changes
      // recorded since put in.
      manuscript: rev.snapshot
        ? {
            paragraphs: currentTexts(
              rev.snapshot.paragraphs,
              rev.comments
                .flatMap((c) => c.changes)
                .filter((ch) => ch.created_at > rev.snapshot.taken_at)
                .sort(byCreated),
            ),
            ids: rev.snapshot.ids,
            seen: manuscriptSeen(key),
          }
        : null,
      paragraphs: Array.isArray(paragraphs) ? paragraphs : [],
      ledger: rev.ledger,
      lang: profile.rationale_lang,
      // policy_packs of the journal profile: which rule packs the turn carries.
      profile,
      includeForbid: includeForbid(),
      // A turn in a supplement is told its file and what the turn of the
      // manuscript left for another file. It is not given the manuscript.
      part: link
        ? {
            name: basename(d.doc),
            main_name: link.main.name,
            outside: outsideNotes(link.main.revision, ref),
          }
        : null,
      linkedAnswers: await answersElsewhere(d, rev),
      consistency: rev.consistency,
    });
    if (link) {
      log(
        `[review] part turn ${ref} file=${basename(d.doc)} main=${link.main.name} outside=${built.outsideIds.join(",") || "-"}`,
      );
    }
    return { rev, comment, built };
  }

  const impl = {
    async review_load(msg, key) {
      const d = at(key);
      const { path } = d;
      // In a supplement the copy of what the manuscript owns is brought up
      // to date in the same save; the manuscript is read before, plainly.
      const link = await linkOf(d);
      let done = [];
      const revision = await update(d, (rev) => {
        const r = recordUnsaved(d, rev);
        done = r.done;
        const own = clearStale(d, r.revision ?? rev) ?? r.revision;
        return (link && inLine(own ?? rev, link)) ?? own;
      });
      await unsavedDone(d, key, revision, done);
      // The other end of the link, when the manuscript lost it: a save of
      // its own, after the one above.
      const own = resolve(d.doc);
      if (link?.main.revision && !link.main.revision.part_of && !partEntry(link.main, own)) {
        const { main } = link;
        const { label, joined_at } = revision.part_of;
        await withRevision(main.path, main.doc, (m) =>
          addPart(m, { doc: storedPartPath(dirname(main.doc), own), label, joined_at }),
        );
        log(`[review] link restored in ${main.name} for ${basename(own)}`);
        tellLinked(d, [main.doc, own]);
      }
      return { path, revision, linked: await linkedReply(d, revision, link) };
    },
    // Manuscripts in the folder of this document whose revision it could join.
    async review_link_candidates(msg, key) {
      const d = at(key);
      let own = null;
      try {
        own = await load(d);
      } catch {
        own = null; // no readable file of its own: it holds nothing
      }
      return { candidates: await linkCandidates(resolve(d.doc), own) };
    },
    // This document becomes a supplement in the revision of msg.main (the
    // path of that manuscript or of its review file). msg.label: what the
    // letter calls this file (default: its file name without the ending).
    // Its own file is saved first, then the file of the manuscript; a second
    // join of the same manuscript changes nothing.
    async review_link_join(msg, key) {
      const d = at(key);
      const own = resolve(d.doc);
      const mainDoc = await mainDocFrom(msg.main, own, env);
      const name = basename(mainDoc);
      if (revisionPathFor(mainDoc) === d.path) {
        throw new Error("This is the file you have open. Name the manuscript it belongs to.");
      }
      const partOf = {
        manuscript: storedPartPath(dirname(own), mainDoc),
        label: String(msg.label ?? "").trim() || basename(own).replace(/\.[^.]+$/, ""),
        joined_at: new Date().toISOString(),
      };
      if (String(msg.label ?? "").trim().length > 80)
        throw new Error("The name is too long (at most 80 characters).");
      const main = await loadMain(own, partOf);
      if (!main.revision?.comments.length) {
        throw new Error(
          `${name} has no saved reviewer comments. Open ${name}, paste the review and save the comment list first.`,
        );
      }
      if (main.revision.part_of) {
        throw new Error(
          `${name} is itself part of the revision of ${linkedSummary(mainDoc, main.revision).part_of.name}. Join that file.`,
        );
      }
      const listed = new Set(main.revision.comments.map((c) => c.ref));
      const revision = await update(d, (rev) => {
        if (rev.part_of && resolvePart(dirname(own), rev.part_of.manuscript) !== mainDoc) {
          throw new Error(
            `This file already belongs to the revision of ${linkedSummary(own, rev).part_of.name}. Leave that revision first.`,
          );
        }
        if (rev.parts?.length) {
          throw new Error(
            `Other files are linked to this file (${linkedSummary(own, rev)
              .parts.map((p) => p.name)
              .join(", ")}), so it cannot join another revision.`,
          );
        }
        const extra = rev.comments
          .filter((c) => c.changes.length > 0 && !listed.has(c.ref))
          .map((c) => c.ref);
        if (extra.length) {
          throw new Error(
            `This file has changes for ${extra.join(", ")}, which ${name} does not list. Add ${extra.length === 1 ? "that comment" : "those comments"} to the list in ${name} first.`,
          );
        }
        return inLine(rev, { partOf, main });
      });
      const { label, joined_at } = revision.part_of;
      const saved = await withRevision(main.path, main.doc, (m) =>
        addPart(m, { doc: storedPartPath(dirname(mainDoc), own), label, joined_at }),
      );
      log(`[review] link join ${basename(own)} -> ${name} comments=${revision.comments.length}`);
      tellLinked(d, [mainDoc, own, ...partDocs(mainDoc, saved)]);
      return { revision, linked: await linkedReply(d, revision, await linkOf(d)) };
    },
    // What the response letter calls this supplementary file ("Supplementary
    // Information"), set from its own pane after the join. Written at both
    // ends like the join: this file first, then the manuscript. An empty
    // label is the file name without its extension.
    async review_link_label(msg, key) {
      const d = at(key);
      const own = resolve(d.doc);
      const label = String(msg.label ?? "").trim();
      if (label.length > 80) throw new Error("The name is too long (at most 80 characters).");
      const before = await load(d);
      if (!before.part_of) throw new Error("This file is not part of another revision.");
      const main = await loadMain(own, before.part_of);
      const revision = await update(d, (rev) =>
        rev.part_of ? setPartOf(rev, { ...rev.part_of, label }) : null,
      );
      const entry = partEntry(main, own);
      if (entry)
        await withRevision(main.path, main.doc, (m) =>
          addPart(m, { doc: entry.doc, label, joined_at: entry.joined_at }),
        );
      log(`[review] link label ${basename(own)} set=${label ? 1 : 0}`);
      tellLinked(d, [main.doc, ...partDocs(main.doc, main.revision)]);
      return { revision, linked: await linkedReply(d, revision, await linkOf(d)) };
    },
    // The link is taken away at both ends, this file first. The copy of what
    // the manuscript owns goes with it; what is this file's own stays.
    async review_link_leave(msg, key) {
      const d = at(key);
      const own = resolve(d.doc);
      const before = await load(d);
      if (!before.part_of) throw new Error("This file is not part of another revision.");
      const main = await loadMain(own, before.part_of);
      const revision = await update(d, (rev) => (rev.part_of ? leaveMain(rev) : null));
      const entry = partEntry(main, own);
      if (entry) await withRevision(main.path, main.doc, (m) => removePart(m, entry.doc));
      log(`[review] link leave ${basename(own)} from ${main.name}`);
      tellLinked(d, [main.doc, own, ...partDocs(main.doc, main.revision)]);
      return { revision, linked: await linkedReply(d, revision) };
    },
    async review_split(msg, key) {
      const d = at(key);
      const text = String(msg.text ?? "");
      // A receipt only when the model was asked (a regex split is no LLM decision).
      // Provider and model of this job, for the split and the extraction after it.
      const job = await oneshotJob(key);
      const shot = job.run;
      let asked = false;
      const ask =
        shot &&
        ((req) => {
          asked = true;
          return shot(req);
        });
      // Head patterns: the request's own, else those of the paper profile.
      // A document without a local path has no profile.
      let patterns = msg.patterns;
      if (!patterns) {
        try {
          patterns = splitPatterns(await profileOf(d, await load(d)));
        } catch {
          patterns = undefined;
        }
      }
      const r = await splitReview(text, { patterns, oneshot: ask });
      const records = [];
      if (asked) {
        records.push(
          decisionRecord({
            kind: "comment_split",
            input: text,
            model: ranOf(job),
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
      const x = await extractRequirements(r.comments, { oneshot: shot, log });
      if (x.asked) {
        records.push(
          decisionRecord({
            kind: "requirement_extract",
            input: x.input,
            model: ranOf(job),
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
      const revision = await update(d, async (rev) => {
        // A comment that has changes in a supplement stays: those edits are
        // in that document. Its file is read, never written from here.
        // So does one a running turn or an unanswered edit of a supplement
        // holds (inflight is keyed by review file).
        const parts = rev.parts?.length ? await loadParts(d.doc, rev) : [];
        const inParts = partChangeRefs(parts);
        const names = new Map(parts.map((p) => [revisionPathFor(p.doc), p.name]));
        const heldInParts = new Map();
        for (const e of inflight.values()) {
          const name = names.get(resolve(e.path));
          if (name && !heldInParts.has(e.ref)) heldInParts.set(e.ref, name);
        }
        for (const c of rev.comments) {
          if (listed.has(c.ref)) continue;
          if (inParts.has(c.ref)) {
            throw new Error(
              `Comment ${c.ref} has recorded change(s) in ${inParts.get(c.ref)}; keep it`,
            );
          }
          if (heldInParts.has(c.ref)) {
            throw new Error(
              `Comment ${c.ref} is being worked on in ${heldInParts.get(c.ref)}; keep it`,
            );
          }
        }
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
    // The author's answer to a question of the model (the card of the comment).
    async review_question_answer(msg, key) {
      const d = at(key);
      return {
        revision: await update(d, (rev) => answerQuestion(rev, msg.ref, msg.id, msg.answer)),
      };
    },
    // "Check revised paragraphs" (paragraph-check.mjs): the paragraphs that
    // hold two or more changes, read by the model that writes the edits
    // (the router's role "turn" with the model of the pane). The pane
    // refreshes the status first. The result replaces the last one in the
    // review file; the log and the receipt hold counts and ids, no text.
    async review_check_paragraphs(msg, key) {
      const d = at(key);
      const rev = await load(d);
      let job = null;
      const r = await checkParagraphs(rev, {
        oneshot: async (req) => {
          job ??= await oneshotJob(key, { role: "turn", claudeModel: getModel(key) ?? undefined });
          if (!job.run) return { ok: false, error: "no model is available" };
          return job.run({ ...req, timeoutMs: PARAGRAPH_CHECK_TIMEOUT_MS });
        },
      });
      // The model the last request ran on, when the router knows it.
      const model = job?.used?.()?.ran ?? job?.model ?? null;
      log(
        `[review] paragraph check paragraphs=${r.paragraphs.length} findings=${r.findings.length}` +
          `${r.warnings.length ? ` warnings=${r.warnings.length}` : ""} model=${model ?? "-"}`,
      );
      if (r.asked) {
        await receipt(d, [
          decisionRecord({
            kind: "paragraph_check",
            input: r.input,
            model,
            output: {
              paragraphs: r.paragraphs.map((p) => p.index),
              findings: r.findings.map((f) => ({
                p: f.paragraph_index,
                kind: f.kind,
                change_ids: f.change_ids,
              })),
              warnings: r.warnings.length,
            },
          }),
        ]);
      }
      const consistency = {
        checked_at: new Date().toISOString(),
        model,
        paragraphs: r.paragraphs.map((p) => ({
          index: p.index,
          id: p.id,
          change_ids: p.changes.map((ch) => ch.id),
          many: p.many,
        })),
        findings: r.findings,
        warnings: r.warnings,
      };
      const revision = await update(d, (cur) => ({ ...cur, consistency }));
      return { revision, consistency };
    },
    async review_preview(msg, key) {
      const d = at(key);
      const { built } = await buildFor(d, msg.ref, msg.paragraphs, key);
      log(
        formatPromptLog(msg.ref, built, { paragraphs: msg.paragraphs?.length ?? 0, preview: true }),
      );
      return { prompt: built.text };
    },
    async review_propose(msg, key, host) {
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
      const { built, comment, rev } = await buildFor(d, msg.ref, msg.paragraphs, key);
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
      if (msg.follow_through === true && Number.isInteger(msg.rewrite_paragraph)) {
        throw new Error(
          "A turn either writes a paragraph anew or checks the other places of a comment, not both.",
        );
      }
      // "Rewrite this paragraph": the turn of this comment is told to write
      // one paragraph anew (paragraph-check.mjs rewriteNote). A paragraph
      // that does not hold two changes starts no turn.
      const rewrite = Number.isInteger(msg.rewrite_paragraph)
        ? rewriteNote(rev, msg.rewrite_paragraph, msg.ref)
        : null;
      if (rewrite) {
        log(
          `[review] rewrite turn ${msg.ref} paragraph=${msg.rewrite_paragraph} changes=${Number(/holds (\d+) changes/.exec(rewrite)?.[1] ?? 0)}`,
        );
      }
      // "Check other places": the turn of this comment is given the places
      // its changes did not touch (paragraph-check.mjs followNote), and the
      // places its answered questions name. A comment without such a place
      // starts no turn. The turn
      // carries the places it may change and, as sources of numbers, what
      // the changes of this comment put in and took out (review-tools.mjs
      // applyTrackedEdit), and what its end needs (turnEnded): the document
      // it began in and what was on record then.
      const plan = msg.follow_through === true ? followPlan(rev, msg.ref) : null;
      const follow = plan ? followNote(rev, msg.ref, undefined, plan) : null;
      const note = rewrite ?? follow;
      pushTurn(key, host, note ? `${built.text}\n\n${note}` : built.text, {
        ref: msg.ref,
        ...(plan && {
          follow: {
            candidate_ids: plan.candidates.map((p) => p.id),
            // The answered questions whose place is among them.
            answer_ids: plan.candidates.flatMap((p) => (p.answer ? [p.answer] : [])),
            sources: plan.changes.flatMap((ch) => [ch.after, ch.before]).filter(Boolean),
            doc: d,
            had: {
              changes: comment.changes.map((ch) => ch.id),
              questions: (comment.questions ?? []).map((x) => x.id),
            },
          },
        }),
      });
      // The session of the pane has the manuscript as of this turn.
      if (built.manuscript.mode !== "none") manuscriptGiven(key, built.manuscript.hashes);
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
      log(formatPromptLog(msg.ref, built, { paragraphs: msg.paragraphs?.length ?? 0 }));
      const revision = await update(d, (rev) =>
        findComment(rev, msg.ref).status === "pending"
          ? setCommentStatus(rev, msg.ref, "in_progress")
          : null,
      );
      return { started: true, ref: msg.ref, started_at: new Date().toISOString(), revision };
    },
    async review_refresh(msg, key) {
      const d = at(key);
      const snapshot = msg.snapshot ?? { paragraphs: [], comments: [] };
      const texts = snapshotTexts(snapshot);
      let rows = [];
      let external = [];
      let resolved = [];
      let dropped = [];
      let stale = [];
      let done = [];
      let moved = false;
      // The file on disk as it is now, for doc_state: read before the chain
      // of the review file, and nothing when it cannot be read.
      let onDisk = null;
      try {
        const s = await stat(d.doc);
        onDisk = { mtime: s.mtime.toISOString(), bytes: s.size };
      } catch {
        onDisk = null;
      }
      // The supplementary files of a manuscript, as their review files have
      // them (plain reads, before the chain of this file): the stale marks
      // of the saved answers judge the citations of those files against them.
      let parts = [];
      try {
        const head = await load(d);
        if (head.parts?.length) parts = await partsForLetter(d.doc, head);
      } catch (e) {
        log(`[review] linked files not read: ${e?.message ?? e}`);
      }
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
        next = afterRefresh?.(next, snapshot, parts) ?? next;
        stale = staleChanges(old, next);
        moved = changeStates(file) !== changeStates(next);
        const now = new Date();
        return setDocState(
          setSnapshot(next, texts, now, snapshotIds(snapshot), snapshotRead(snapshot)),
          docState(snapshot, { saved: msg.saved, file: onDisk, now }),
        );
      });
      await unsavedDone(d, key, revision, done);
      // A change was decided in Word, or found: the other open panes of the
      // revision show where its request stands now.
      if (moved) await tellOthers(d, revision);
      for (const x of stale) log(`[response] stale ${x.ref} ${x.stale.join(",")}`);
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
        comment_text: formatRationaleComment(msg.ref, ch.rationale, {
          lost: ch.lost_facts,
          fromReview: ch.from_review,
        }),
        // The task pane reuses none of these for this change.
        known_comment_ids: recordedCommentIds(rev),
        change: {
          kind: ch.kind,
          find: ch.find,
          text: ch.text,
          paragraph_id: ch.paragraph_id,
          paragraph_index: ch.paragraph_index,
          // For a deletion whose old words Word dropped: where the comment
          // of the edit would have gone (tools-word.js droppedDeletionAnchor).
          before: ch.before,
          narrowed_from: ch.narrowed_from,
          comment_missing: ch.comment_missing,
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
    // One try of the task pane at a rationale comment waiting for Word
    // (review-pane.js addWaitingComments, or its button): which change, what
    // started it (timer, focus, visibility, refresh, export, button), and
    // what came of it (added, AccessDenied, skipped_busy, or the code and
    // message of another failure). Only logged; nothing is stored.
    async review_comment_retry(msg) {
      const field = (v, n) =>
        String(v ?? "-")
          .replace(/\s+/g, " ")
          .trim()
          .slice(0, n) || "-";
      log(
        `[review] comment retry ${field(msg.change_id, 40)} trigger=${field(msg.trigger, 20)} outcome=${field(msg.outcome, 200)}`,
      );
      return {};
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
          // An edit of a follow-through turn is judged as in that turn: its
          // places and its sources of numbers are on the item.
          r = await applyTrackedEdit(
            deps,
            blocked.args,
            { ref: msg.ref, ...(blocked.follow ? { follow: blocked.follow } : {}) },
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
            model: ranOf(job),
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
      const outcomes = await checkLedgerIncremental(items, paragraphs, {
        oneshot: job.run,
        full: msg.full === true,
        onDecision: (x) =>
          decisions.push(
            decisionRecord({
              kind: "ledger_check",
              input: { instruction: x.instruction, paragraph: x.paragraph },
              model: ranOf(job),
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

  // In the file of a supplement what the manuscript owns is not changed. In
  // a manuscript with supplements, their open panes are told to load again
  // after a change of what they copy.
  // (An unsaved document has no file and is no supplement.)
  for (const type of MAIN_ONLY) {
    const fn = impl[type];
    impl[type] = async (msg, key, host) => {
      const d = at(key);
      let rev = null;
      try {
        rev = await load(d);
      } catch {
        /* no readable file: the request itself says so, or needs none */
      }
      refuseInPart(rev);
      const out = await fn(msg, key, host);
      if (rev?.parts?.length && !MAIN_ONLY_READS.has(type)) {
        tellLinked(d, partDocs(d.doc, rev));
      }
      return out;
    };
  }

  const handlers = wrapHandlers(impl, { jobTypes: JOB_TYPES, sendEvent, log });
  return { handlers, trackedEditDeps, turnEnded, lateToolResult };
}
