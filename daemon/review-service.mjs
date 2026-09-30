// Review-tab message handlers and the recording side of office_tracked_edit.
// index.mjs wires the dependencies; everything here is testable with fakes.
//
// Messages (task pane -> daemon, reply type is `<type>_result` with the same
// request_id, the existing sendRequest pattern in taskpane.js):
//   review_load, review_split, review_save_comments, review_set_status,
//   review_set_requirements, review_preview, review_propose, review_refresh,
//   review_external_dismiss, review_override, ledger_add, ledger_update,
//   ledger_remove, ledger_check, comments_backup
// Push (daemon -> task pane, no request_id):
//   { type: "review_event", event: "change_recorded" | "edit_blocked" | "turn_done", ... }

import { access } from "node:fs/promises";
import {
  localDocPath,
  revisionPathFor,
  commentsBackupPathFor,
  loadRevision,
  withRevision,
  replaceComments,
  setCommentStatus,
  addChange,
  applyStates,
  addBlocked,
  takeBlocked,
  addLedgerItem,
  updateLedgerItem,
  removeLedgerItem,
  findComment,
  writeJsonAtomic,
  setRequirements,
  setSnapshot,
  addExternal,
  dismissExternal,
} from "./revision-store.mjs";
import { classifyRevision, summarizeStates } from "./revision-classify.mjs";
import { splitReview } from "./review-split.mjs";
import { buildReviewPrompt, formatPromptLog } from "./review-prompt.mjs";
import { checkLedgerIncremental } from "./ledger-check.mjs";
import { snapshotTexts, outsideChanges } from "./doc-diff.mjs";
import { applyTrackedEdit, TOOL_COMMENT_PREFIX } from "./review-tools.mjs";

async function exists(p) {
  try {
    await access(p);
    return true;
  } catch {
    return false;
  }
}

export function createReviewService({
  getContext,
  callTool,
  pushTurn,
  oneshot = null,
  sendEvent = () => {},
  log = console.log,
  env = process.env,
}) {
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

  function trackedEditDeps(key) {
    return {
      log,
      getLedger: async () => (await load(key)).ledger,
      // L2 sources and L5 ids for office_tracked_edit.
      getReviewContext: async (ref) => {
        const rev = await load(key);
        const c = findComment(rev, ref);
        return {
          requirements: c?.requirements ?? [],
          comment_text: c?.text ?? "",
          paragraphs: rev.snapshot?.paragraphs ?? [],
        };
      },
      onEdit: async ({ ref, change }) => {
        let saved = null;
        await update(key, (rev) => {
          const r = addChange(rev, ref, change);
          saved = r.change;
          return r.revision;
        });
        sendEvent(key, { type: "review_event", event: "change_recorded", ref, change: saved });
        return saved;
      },
      onBlocked: async ({ ref, args, blocked_by }) => {
        let saved = null;
        await update(key, (rev) => {
          const r = addBlocked(rev, ref, { args, blocked_by });
          saved = r.blocked;
          return r.revision;
        });
        sendEvent(key, { type: "review_event", event: "edit_blocked", ref, blocked: saved });
        return saved;
      },
    };
  }

  function turnEnded(key, turn) {
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
      return { path, revision: await load(key) };
    },
    async review_split(msg) {
      return splitReview(String(msg.text ?? ""), { patterns: msg.patterns, oneshot });
    },
    async review_save_comments(msg, key) {
      const incoming = Array.isArray(msg.comments) ? msg.comments : [];
      return { revision: await update(key, (rev) => replaceComments(rev, incoming)) };
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
      const { built } = await buildFor(key, msg.ref, msg.paragraphs);
      pushTurn(key, host, built.text, { ref: msg.ref });
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
      const revision = await update(key, (rev) => {
        rows = classifyRevision(rev, snapshot);
        let next = applyStates(rev, rows);
        // L4: changes since the last refresh that no review change explains.
        // The first refresh only stores the snapshot.
        if (rev.snapshot) {
          const r = addExternal(next, outsideChanges(rev.snapshot.paragraphs, texts, rows));
          next = r.revision;
          external = r.added;
        }
        return setSnapshot(next, texts);
      });
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
    async review_override(msg, key) {
      const rev = await load(key);
      const blocked = findComment(rev, msg.ref)?.blocked.find((b) => b.id === msg.blocked_id);
      if (!blocked) throw new Error(`No blocked edit ${msg.blocked_id} under ${msg.ref}`);
      const deps = { ...trackedEditDeps(key), call: (name, args) => callTool(name, args, key) };
      const r = await applyTrackedEdit(
        deps,
        blocked.args,
        { ref: msg.ref },
        { override: { blocked_by: blocked.blocked_by } },
      );
      if (!r.ok) return { applied: false, result: r };
      log(`[review] override ${msg.blocked_id} approved for ${blocked.blocked_by.join(",")}`);
      const revision = await update(key, (x) => takeBlocked(x, msg.ref, msg.blocked_id).revision);
      return { applied: true, result: r, revision };
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
      const rev = await load(key);
      const wanted = Array.isArray(msg.ids) && msg.ids.length ? new Set(msg.ids) : null;
      const items = rev.ledger.filter((i) => !wanted || wanted.has(i.id));
      const paragraphs = msg.snapshot?.paragraphs ?? [];
      // L10: changed paragraphs only unless msg.full (the "Full check" button).
      const outcomes = await checkLedgerIncremental(items, paragraphs, {
        oneshot,
        full: msg.full === true,
      });
      const at = new Date().toISOString();
      for (const r of outcomes) {
        log(`[review] ledger check ${r.id} status=${r.status} violations=${r.violations.length}`);
        log(`[review] ledger scope ${r.id} mode=${r.mode} paragraphs=${r.checked}/${r.total}`);
      }
      const revision = await update(key, (x) => {
        let next = x;
        for (const r of outcomes) {
          next = updateLedgerItem(next, r.id, {
            status: r.status,
            violation_count: r.violations.length,
            checked_at: at,
            violations: r.violations,
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
      let path = commentsBackupPathFor(doc);
      if (await exists(path)) {
        const stamp = new Date().toISOString().replace(/[:.]/g, "-");
        path = path.replace(/\.comments_backup\.json$/, `.comments_backup.${stamp}.json`);
      }
      await writeJsonAtomic(path, backup);
      log(`[review] comments_backup path=${path} count=${backup.comments.length}`);
      return { path, count: backup.comments.length };
    },
  };

  const handlers = {};
  for (const [type, fn] of Object.entries(impl)) {
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
  return { handlers, trackedEditDeps, turnEnded };
}
