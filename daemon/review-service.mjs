// Review-tab message handlers and the recording side of office_tracked_edit.
// index.mjs wires the dependencies; everything here is testable with fakes.
//
// Messages (task pane -> daemon, reply type is `<type>_result` with the same
// request_id, the existing sendRequest pattern in taskpane.js):
//   review_load, review_split, review_save_comments, review_set_status,
//   review_set_requirements, review_preview, review_propose, review_refresh,
//   review_external_dismiss, review_comments_cleaned, review_override,
//   ledger_add, ledger_update, ledger_remove, ledger_check, comments_backup
// Push (daemon -> task pane, no request_id):
//   { type: "review_event", event: "change_recorded" | "edit_blocked" | "turn_done"
//     | "late_result_recorded" | "job_done", ... }
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
  takeBlocked,
  markBlockedUnknown,
  claimBlocked,
  releaseBlocked,
  markStaleApplying,
  addLedgerItem,
  updateLedgerItem,
  removeLedgerItem,
  findComment,
  writeJsonAtomic,
  setRequirements,
  setSnapshot,
  addExternal,
  dismissExternal,
  markCommentsCleaned,
} from "./revision-store.mjs";
import {
  classifyRevision,
  classifyChange,
  summarizeStates,
  normalizeParaText,
} from "./revision-classify.mjs";
import { extractRequirements, splitReview } from "./review-split.mjs";
import { buildReviewPrompt, formatPromptLog } from "./review-prompt.mjs";
import { checkLedgerIncremental } from "./ledger-check.mjs";
import { snapshotTexts, snapshotIds, outsideChanges } from "./doc-diff.mjs";
import { applyTrackedEdit, TOOL_COMMENT_PREFIX } from "./review-tools.mjs";
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
export function resolveUnknownOverrides(rev, snapshot) {
  let next = rev;
  const resolved = [];
  const used = new Set(
    rev.comments.flatMap((c) => c.changes.map((ch) => ch.word_comment_id).filter(Boolean)),
  );
  const paras = snapshot.paragraphs ?? [];
  for (const c of rev.comments) {
    for (const b of c.blocked) {
      if (b.state !== "outcome_unknown" || !b.attempt?.change?.context) continue;
      const comment = (snapshot.comments ?? []).find(
        (x) =>
          !used.has(String(x.id)) &&
          normalizeParaText(x.content) === normalizeParaText(b.attempt.comment_text),
      );
      const change = {
        ...b.attempt.change,
        word_comment_id: comment ? String(comment.id) : null,
      };
      const row = classifyChange(change, snapshot);
      const p = paras.find((x) => x.index === row.paragraph_index);
      // Plain containment: Word joins pending inserted and deleted text with
      // no space, so a word-bounded match would miss it.
      const newTextThere =
        !!p && normalizeParaText(p.text).includes(normalizeParaText(change.text));
      const applied =
        row.state === "accepted" ||
        (row.state === "pending" && change.kind !== "delete" && newTextThere);
      if (!applied) continue;
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
      });
    }
  }
  return { revision: next, resolved };
}

// What a ledger check result depends on: the rule, not its last result.
// review_save_comments: comments work still points at are not removed.
export const keptWarning = (refs) =>
  `Not removed because work on it is in progress: ${refs.join(", ")}`;

// Requests that may run model calls one after another (a split and its
// requirement extraction; a ledger check of up to 20 batches per rule). The
// regex-only path takes the same route: one reply shape per request type.
const JOB_TYPES = new Set(["review_split", "ledger_check"]);

const ruleFingerprint = (item) => JSON.stringify([item.kind, item.text, item.scope, item.check]);

export function createReviewService({
  getContext,
  callTool,
  pushTurn,
  oneshot = null,
  // Model names for decision receipts: the one-shot model, and the model of
  // the drafting session of a pane (getModel(key)). null when unknown.
  oneshotModel = null,
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
  function hold(key, ref, comment = null) {
    let path;
    try {
      path = docFor(key).path;
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
  function withComment(key, rev, ref) {
    if (findComment(rev, ref)) return rev;
    log(`[review] comment ${ref} was missing; restored to keep the change`);
    return restoreComment(rev, ref, remembered(docFor(key).path, ref));
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
  const load = (key) => {
    const { doc, path } = docFor(key);
    return loadRevision(path, doc);
  };
  const update = (key, fn) => {
    const { doc, path } = docFor(key);
    return withRevision(path, doc, fn);
  };

  // Decision receipts (decision-log.mjs). A receipt never stops the feature
  // it records: a failure only logs.
  async function receipt(key, records) {
    if (!records.length) return;
    try {
      await appendDecisions(decisionsPathFor(docFor(key).doc), records);
    } catch (e) {
      log(`[review] decision log failed: ${e?.message ?? e}`);
    }
  }

  // Overrides this daemon is running, per review file: "<ref>\0<blocked id>".
  // An applying item not in here was left by a daemon that stopped.
  const applying = new Map();
  const applyingKey = (ref, id) => `${ref}\0${id}`;
  function activeFor(key) {
    const { path } = docFor(key);
    if (!applying.has(path)) applying.set(path, new Set());
    return applying.get(path);
  }
  const clearStale = (key, rev) => markStaleApplying(rev, activeFor(key));
  const backupChains = new Map();

  function trackedEditDeps(key) {
    return {
      log,
      getLedger: async () => (await load(key)).ledger,
      // L2 sources and L5 ids for office_tracked_edit. recent_changes: the
      // changes recorded since the snapshot (earlier edits of this turn), so
      // the forbid check reads the paragraphs as they are now.
      getReviewContext: async (ref) => {
        const rev = await load(key);
        const c = findComment(rev, ref);
        const since = rev.snapshot?.taken_at;
        return {
          requirements: c?.requirements ?? [],
          comment_text: c?.text ?? "",
          paragraphs: rev.snapshot?.paragraphs ?? [],
          paragraph_ids: rev.snapshot?.ids ?? null,
          recent_changes: since
            ? rev.comments
                .flatMap((x) => x.changes)
                .filter((ch) => ch.created_at > since)
                .sort((a, b) => (a.created_at < b.created_at ? -1 : 1))
            : [],
        };
      },
      onEdit: async ({ ref, change }) => {
        let saved = null;
        let comment = null;
        await update(key, (rev) => {
          const base = withComment(key, rev, ref);
          const r = addChange(base, ref, change);
          saved = r.change;
          comment = findComment(base, ref);
          return r.revision;
        });
        await receipt(key, [
          decisionRecord({
            kind: "requirement_link",
            input: {
              comment: comment?.text ?? "",
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
        sendEvent(key, { type: "review_event", event: "change_recorded", ref, change: saved });
        return saved;
      },
      onOutcomeUnknown: ({ id, ref, record, blocked_id = null }) => {
        if (!id) return;
        const timer = setTimeout(() => forgetLate(id), lateResultTtlMs);
        timer.unref?.();
        pendingLate.set(id, { key, ref, record, blocked_id, timer, release: hold(key, ref) });
      },
      holdRef: (ref) => hold(key, ref),
      onBlocked: async ({ ref, args, blocked_by }) => {
        let saved = null;
        await update(key, (rev) => {
          const r = addBlocked(withComment(key, rev, ref), ref, { args, blocked_by });
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
    const held = p ? hold(p.key, p.ref) : () => {};
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
    const out = entry.ok ? await p.record(entry.result) : null;
    if (!out?.ok) {
      const why = entry.ok ? (out?.code ?? out?.message ?? "not applied") : entry.error;
      log(`[review] late result not applied ${entry.id}: ${why}`);
      return null;
    }
    log(`[review] late result recorded ${entry.id}`);
    if (p.blocked_id) {
      await update(p.key, (rev) =>
        findComment(rev, p.ref)?.blocked.some((b) => b.id === p.blocked_id)
          ? takeBlocked(rev, p.ref, p.blocked_id).revision
          : null,
      );
      log(`[review] override ${p.blocked_id} applied late as ${out.change_id}`);
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

  function turnEnded(key, turn) {
    turnHolds.get(key)?.();
    turnHolds.delete(key);
    log(`[review] turn done ${turn.ref}`);
    sendEvent(key, { type: "review_event", event: "turn_done", ref: turn.ref });
  }

  const includeForbid = () => env.CC_OFFICE_REVIEW_PROMPT_FORBID !== "0";

  async function buildFor(key, ref, paragraphs) {
    const rev = await load(key);
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

  const impl = {
    async review_load(msg, key) {
      const { path } = docFor(key);
      return { path, revision: await update(key, (rev) => clearStale(key, rev)) };
    },
    async review_split(msg, key) {
      const text = String(msg.text ?? "");
      // A receipt only when the model was asked (a regex split is no LLM decision).
      let asked = false;
      const ask =
        oneshot &&
        ((req) => {
          asked = true;
          return oneshot(req);
        });
      const r = await splitReview(text, { patterns: msg.patterns, oneshot: ask });
      const records = [];
      if (asked) {
        records.push(
          decisionRecord({
            kind: "comment_split",
            input: text,
            model: oneshotModel,
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
      const x = await extractRequirements(r.comments, { oneshot });
      if (x.asked) {
        records.push(
          decisionRecord({
            kind: "requirement_extract",
            input: x.input,
            model: oneshotModel,
            output: { requests: x.requests },
          }),
        );
      }
      if (records.length) await receipt(key, records);
      return { ...r, comments: x.comments, warnings: [...r.warnings, ...x.warnings] };
    },
    async review_save_comments(msg, key) {
      const incoming = Array.isArray(msg.comments) ? msg.comments : [];
      const { path } = docFor(key);
      const listed = new Set(incoming.map((c) => c?.ref));
      let kept = [];
      // Read inside the chain: what is held when the save runs, not when asked.
      const revision = await update(key, (rev) => {
        const held = heldRefs(path, rev);
        kept = rev.comments.map((c) => c.ref).filter((r) => !listed.has(r) && held.has(r));
        return replaceComments(rev, incoming, held);
      });
      if (kept.length) log(`[review] comments kept, work in progress: ${kept.join(",")}`);
      return { revision, warnings: kept.length ? [keptWarning(kept)] : [] };
    },
    async review_set_status(msg, key) {
      return { revision: await update(key, (rev) => setCommentStatus(rev, msg.ref, msg.status)) };
    },
    async review_set_requirements(msg, key) {
      const texts = Array.isArray(msg.requirements) ? msg.requirements : [];
      return { revision: await update(key, (rev) => setRequirements(rev, msg.ref, texts)) };
    },
    async review_preview(msg, key) {
      const { built } = await buildFor(key, msg.ref, msg.paragraphs);
      log(
        formatPromptLog(msg.ref, built, { paragraphs: msg.paragraphs?.length ?? 0, preview: true }),
      );
      return { prompt: built.text };
    },
    async review_propose(msg, key, host) {
      const { built, comment } = await buildFor(key, msg.ref, msg.paragraphs);
      pushTurn(key, host, built.text, { ref: msg.ref });
      turnHolds.get(key)?.();
      turnHolds.set(
        key,
        hold(key, msg.ref, {
          reviewer: comment.reviewer,
          text: comment.text,
          requirements: comment.requirements,
        }),
      );
      log(formatPromptLog(msg.ref, built, { paragraphs: msg.paragraphs?.length ?? 0 }));
      const revision = await update(key, (rev) =>
        findComment(rev, msg.ref).status === "pending"
          ? setCommentStatus(rev, msg.ref, "in_progress")
          : null,
      );
      return { started: true, ref: msg.ref, started_at: new Date().toISOString(), revision };
    },
    async review_refresh(msg, key) {
      const snapshot = msg.snapshot ?? { paragraphs: [], comments: [] };
      const texts = snapshotTexts(snapshot);
      let rows = [];
      let external = [];
      let resolved = [];
      const revision = await update(key, (old) => {
        const r0 = resolveUnknownOverrides(clearStale(key, old) ?? old, snapshot);
        const rev = r0.revision;
        resolved = r0.resolved;
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
      for (const x of resolved) {
        forgetLate(x.tool_call_id);
        log(
          `[review] override ${x.blocked_id} found applied in refresh as ${x.change_id} state=${x.state}`,
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
      return { revision: await update(key, (rev) => dismissExternal(rev, msg.id)) };
    },
    // msg.ids: Word comment ids the task pane's comment cleanup deleted.
    async review_comments_cleaned(msg, key) {
      const ids = Array.isArray(msg.ids) ? msg.ids : [];
      const revision = await update(key, (rev) => markCommentsCleaned(rev, ids));
      log(`[review] comments cleaned ids=${ids.length}`);
      return { revision };
    },
    async review_override(msg, key) {
      // Claimed inside the revision chain: an overlapping override of the
      // same item finds it applying and is refused before Word is called.
      // An applying item left by a stopped daemon is taken as outcome_unknown
      // first, so it gets the Refresh hint, not "already being applied".
      // The id joins `active` in the same step, before any other request of
      // the chain can read the file.
      let blocked = null;
      const active = activeFor(key);
      const held = applyingKey(msg.ref, msg.blocked_id);
      try {
        await update(key, (rev) => {
          const r = claimBlocked(clearStale(key, rev) ?? rev, msg.ref, msg.blocked_id, {
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
      const restore = () =>
        update(key, (x) => releaseBlocked(x, msg.ref, blocked.id, blocked.state));
      try {
        const base = trackedEditDeps(key);
        let attempt = null;
        const deps = {
          ...base,
          call: (name, args) => callTool(name, args, key),
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
          const revision = await update(key, (x) =>
            markBlockedUnknown(x, msg.ref, blocked.id, attempt),
          );
          return { applied: false, outcome_unknown: true, result: r, revision };
        }
        if (!r.ok) {
          const revision = await restore();
          return { applied: false, result: r, revision };
        }
        log(`[review] override ${msg.blocked_id} approved for ${blocked.blocked_by.join(",")}`);
        const revision = await update(key, (x) => takeBlocked(x, msg.ref, msg.blocked_id).revision);
        return { applied: true, result: r, revision };
      } finally {
        active.delete(held);
      }
    },
    async ledger_add(msg, key) {
      let item = null;
      const revision = await update(key, (rev) => {
        const r = addLedgerItem(rev, msg.item);
        item = r.item;
        return r.revision;
      });
      log(`[review] ledger add ${item.id} kind=${item.kind} check=${item.check.type}`);
      return { item, revision };
    },
    async ledger_update(msg, key) {
      return {
        revision: await update(key, (rev) => updateLedgerItem(rev, msg.id, msg.patch ?? {})),
      };
    },
    async ledger_remove(msg, key) {
      return { revision: await update(key, (rev) => removeLedgerItem(rev, msg.id)) };
    },
    async ledger_check(msg, key) {
      // Taken on arrival, before the file read: a check that reads the file
      // later is still the older one. An older check never lowers the mark.
      const gen = ++lastCheckGen;
      const rev = await load(key);
      const wanted = Array.isArray(msg.ids) && msg.ids.length ? new Set(msg.ids) : null;
      const items = rev.ledger.filter((i) => !wanted || wanted.has(i.id));
      const { path } = docFor(key);
      const genKey = (id) => `${path}\0${id}`;
      for (const i of items) {
        if ((checkGen.get(genKey(i.id)) ?? 0) < gen) checkGen.set(genKey(i.id), gen);
      }
      const started = new Map(items.map((i) => [i.id, ruleFingerprint(i)]));
      const paragraphs = msg.snapshot?.paragraphs ?? [];
      // L10: changed paragraphs only unless msg.full (the "Full check" button).
      const decisions = [];
      const outcomes = await checkLedgerIncremental(items, paragraphs, {
        oneshot,
        full: msg.full === true,
        onDecision: (d) =>
          decisions.push(
            decisionRecord({
              kind: "ledger_check",
              input: { instruction: d.instruction, paragraph: d.paragraph },
              model: oneshotModel,
              output: {
                item_id: d.item_id,
                paragraph_index: d.paragraph_index,
                verdict: d.verdict,
                ...(d.error ? { error: String(d.error).slice(0, 200) } : {}),
              },
            }),
          ),
      });
      await receipt(key, decisions);
      const at = new Date().toISOString();
      for (const r of outcomes) {
        const unsure = r.unsure.length ? ` unsure=${r.unsure.length}` : "";
        log(
          `[review] ledger check ${r.id} status=${r.status} violations=${r.violations.length}${unsure}`,
        );
        log(`[review] ledger scope ${r.id} mode=${r.mode} paragraphs=${r.checked}/${r.total}`);
      }
      const revision = await update(key, (x) => {
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
            checked_at: at,
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
      const { doc } = docFor(key);
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
  return { handlers, trackedEditDeps, turnEnded, lateToolResult };
}
