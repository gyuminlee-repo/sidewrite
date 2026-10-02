// Pure helpers of the acceptance runner: the Word COM plan built from
// revision.json, the expected pane texts of each step (wording of the
// stage-2 acceptance procedure), and the path rewrite that runs a
// feature_list verify_command against this run's files.

import { formatRationaleComment } from "../../daemon/review-tools.mjs";
import { normalizeParaText, placedRuns } from "../../daemon/revision-classify.mjs";

// The author's handling of the proposed changes in the acceptance procedure:
// R1.1 accepted, R1.2 rejected, R1.3 first change accepted then one inserted
// word edited, everything else accepted. R1.4 (the deletion of the flask
// sentence) is rejected, with a tracked insertion of the author pending in
// the same paragraph (siblingMatch).
export const DEFAULT_DECISIONS = {
  "R1.1": "accept",
  "R1.2": "reject",
  "R1.3": ["accept_edit", "accept"],
  "R1.4": "reject",
};

// One inserted word the author edits by hand, and what it becomes.
export const WORD_EDITS = {
  reading: "measuring",
  measuring: "recording",
  monitoring: "tracking",
  monitored: "tracked",
  followed: "tracked",
};

function decisionFor(decisions, ref, i) {
  const d = decisions[ref] ?? "accept";
  if (!Array.isArray(d)) return d;
  return d[Math.min(i, d.length - 1)];
}

// The text of the rationale comment of a change as Word was given it
// (review-tools.mjs applyTrackedEdit): the lines of lost facts and of words
// from the review belong to it. A change is found in Word by this text.
export const commentTextOf = (ref, ch) =>
  formatRationaleComment(ref, ch.rationale, {
    lost: ch.lost_facts ?? [],
    fromReview: ch.from_review ?? [],
  });

// The edit the acceptance run sends to the guarded write for S2-23: it puts
// the forbidden word of ledger rule L2 (tests/fixtures/review_ledger.json)
// into the paragraph of R1.3 of the fixture manuscript. The item it is put
// on file with names another rule (L1), so "Allow exception" lifts that one
// only and L2 has to refuse the edit.
export const FORBID_GUARD_EDIT = {
  comment_ref: "R1.3",
  find: "synthetic carbon source",
  replace: "novel synthetic carbon source",
  rationale: {
    what: "Called the carbon source novel.",
    why: "Written by the acceptance run to show the forbid guard; no model wrote it.",
    addresses: "describe the synthetic carbon source as a novel substrate",
  },
};
export const FORBID_GUARD_SEED_RULE = "L1";

// S2-23, without a model: an edit that holds a forbidden word was sent to
// the guarded write of the daemon (run-acceptance.mjs forbidGuard) and must
// have been refused before Word.
//   step      what the pane showed after "Allow exception" (view)
//   log       the daemon log written since the step began
//   revision  revision.json after the step
//   before    revision.json before it, with the seeded item (an edit of a
//             turn the guard blocked earlier is not the record of this step)
//   word, wordBefore   word-com.ps1 state after and before the step
//   seedId    the id of the seeded item (it must still be there, not applied)
// pass: the pane shows the refusal by the rule; the daemon logged the block;
// the daemon recorded the blocked edit under the rule; no change holds the
// word without an approved exception for the rule; Word has no new tracked
// change and no inserted text with the word.
export function forbidBlockEvidence({
  step,
  log,
  revision,
  before,
  word,
  wordBefore,
  seedId,
  ref = "R1.3",
  rule = "L2",
  word_re = /\bnovel(?:ty)?\b/i,
}) {
  const at = ref.replace(".", "\\.");
  const comment = (revision?.comments ?? []).find((c) => c.ref === ref);
  const holds = (b) => word_re.test(String(b.args?.replace ?? b.args?.insert_after ?? ""));
  const had = new Set(
    ((before?.comments ?? []).find((c) => c.ref === ref)?.blocked ?? []).map((b) => b.id),
  );
  const recorded = (comment?.blocked ?? []).filter(
    (b) => !had.has(b.id) && b.state === "blocked" && b.blocked_by.includes(rule) && holds(b),
  );
  const changes = (revision?.comments ?? [])
    .flatMap((c) => c.changes ?? [])
    .filter((ch) => word_re.test(String(ch.text ?? "")) && !(ch.override_of ?? []).includes(rule))
    .map((ch) => ch.id);
  const inserted = (word?.revisions ?? []).filter(
    (r) => r.type === 1 && word_re.test(r.text ?? ""),
  );
  const checks = {
    refused:
      step?.ok !== false &&
      step?.message_error === true &&
      new RegExp(`forbidden rule\\(s\\) ${rule}\\b`).test(step?.message ?? ""),
    logged: new RegExp(`^\\[review\\] blocked ${at} by=${rule}$`, "m").test(String(log ?? "")),
    recorded: recorded.length === 1,
    seed_kept: (comment?.blocked ?? []).some((b) => b.id === seedId && b.state === "blocked"),
    no_change: changes.length === 0,
    word_unchanged:
      word?.ok === true &&
      wordBefore?.ok === true &&
      (word.revisions ?? []).length === (wordBefore.revisions ?? []).length &&
      inserted.length === 0,
  };
  return {
    pass: Object.values(checks).every(Boolean),
    checks,
    message: step?.message ?? step?.error ?? null,
    recorded: recorded.map((b) => ({ id: b.id, blocked_by: b.blocked_by })),
    changes_with_the_word: changes,
    word_revisions: [(wordBefore?.revisions ?? []).length, (word?.revisions ?? []).length],
  };
}

// ops for word-com.ps1 apply. A change without a rationale comment cannot be
// found in Word and is listed in `skipped`.
export function buildComPlan(revision, decisions = DEFAULT_DECISIONS) {
  const ops = [];
  const skipped = [];
  for (const c of revision?.comments ?? []) {
    c.changes.forEach((ch, i) => {
      if (!ch.word_comment_id || ch.comment_missing || ch.comment_cleaned_at) {
        skipped.push(ch.id);
        return;
      }
      ops.push({
        change_id: ch.id,
        ref: c.ref,
        op: decisionFor(decisions, c.ref, i),
        kind: ch.kind,
        find: ch.find,
        text: ch.text,
        comment_text: commentTextOf(c.ref, ch),
        replacements: WORD_EDITS,
      });
    });
  }
  return { ops, skipped };
}

// A replace whose find and text share a first or last word did not cover only
// the changed words (L3). Inserts and deletes are minimal by construction.
export function minimalFind(ch) {
  if (ch.kind !== "replace") return true;
  const w = (s) => String(s).trim().split(/\s+/).filter(Boolean);
  const a = w(ch.find);
  const b = w(ch.text);
  if (!a.length || !b.length) return true;
  return a[0] !== b[0] && a[a.length - 1] !== b[b.length - 1];
}

// --screenshots. The file of the n-th pane step: "07-propose-R1.1.png",
// "03-show-tab-tools.png". What tells two calls of a step apart (the comment,
// the tab) is in the name; nothing else of the arguments is.
export function shotName(n, step, args = {}) {
  const part = [args.ref, args.tab].find((x) => typeof x === "string" && /^[\w.-]+$/.test(x));
  return `${String(n).padStart(2, "0")}-${step}${part ? `-${part}` : ""}.png`;
}

// Pane steps that wait for a model (the others answer from the daemon or
// from Word alone). split: only when no pattern splits the review, and for
// the requests it then extracts.
export const MODEL_STEPS = new Set([
  "split",
  "propose",
  "propose-all",
  "propose-stop",
  "check-paragraphs",
  "rewrite-paragraph",
  "suggest-ledger",
  "load-letter",
  "response-draft",
  "response-draft-all",
  "tools-translate",
  "tools-ask",
  "tools-extract",
]);

// One row of shots/timings.json. seconds: from the start of the step's
// process to its answer (about 2 s of it are the driver connecting to the
// pane); turn_seconds: what the step measured itself from the click to the
// end of the turn; waiting: what the pane said 4 s and 20 s into the wait;
// progress: the progress lines a propose step saw.
export function stepTiming({ n, step, args = {}, shot, seconds, result }) {
  const bare = step.replace(/^(?:mode14|special)-/, "");
  return {
    n,
    step,
    ...(args.ref ? { ref: args.ref } : {}),
    ...(args.tab ? { tab: args.tab } : {}),
    model_call: MODEL_STEPS.has(bare),
    ok: result?.ok !== false,
    seconds,
    turn_seconds: result?.seconds ?? null,
    shot: result?.shot === null ? null : shot,
    shot_error: result?.shot_error ?? null,
    waiting: (result?.waiting ?? []).map((w) => ({
      ...w,
      ...(w.shot ? { shot: String(w.shot).split(/[\\/]/).pop() } : {}),
    })),
    progress: result?.progress_samples ?? [],
    message: result?.message ?? result?.error ?? null,
  };
}

// Replace each `from` prefix with its `to`, longest first so a folder and a
// file inside it do not clash.
export function rewriteCommand(cmd, map) {
  const pairs = Object.entries(map).sort((x, y) => y[0].length - x[0].length);
  let out = "";
  for (let i = 0; i < cmd.length; ) {
    const hit = pairs.find(([from]) => from && cmd.startsWith(from, i));
    if (hit) {
      out += hit[1];
      i += hit[0].length;
    } else {
      out += cmd[i++];
    }
  }
  return out;
}

// "S25-30..S25-41" or "S2-18": the feature ids of one stage, in order, with
// the number padded to two digits as the feature lists write it.
export function featureIds(spec) {
  // A list of ranges ("S4-30..S4-33,S4-37"): all of one stage.
  const parts = String(spec).split(",");
  if (parts.length > 1) {
    const ids = parts.flatMap((p) => featureIds(p));
    if (new Set(ids.map((id) => id.split("-")[0])).size > 1)
      throw new Error(`--features must stay in one stage, not ${spec}`);
    return ids;
  }
  const ends = String(spec)
    .split("..")
    .map((s) => /^(S\d+)-(\d+)$/.exec(s.trim()));
  if (ends.length > 2 || ends.some((m) => !m)) {
    throw new Error(`--features takes S<stage>-<n> or S<stage>-<n>..S<stage>-<m>, not ${spec}`);
  }
  const [a, b = a] = ends;
  if (a[1] !== b[1]) throw new Error(`--features must stay in one stage, not ${spec}`);
  if (Number(b[2]) < Number(a[2])) throw new Error(`--features ${spec} ends before it starts`);
  const out = [];
  for (let i = Number(a[2]); i <= Number(b[2]); i++) {
    out.push(`${a[1]}-${String(i).padStart(2, "0")}`);
  }
  return out;
}

// The marker prefix of a feature spec: "S25-30..S25-41" -> "s25".
export const stageOf = (spec) => featureIds(spec)[0].split("-")[0].toLowerCase();

const noSlash = (p) => String(p).replace(/\/+$/, "");

// Where the markers and evidence of a stage are: the feature lists name them
// /tmp/<stage>-*, a run keeps them in its own ok folder.
export const markerMap = (stage, okDir) => ({ [`/tmp/${stage}-`]: `${noSlash(okDir)}/${stage}-` });

// A verify command of a feature list as this run must run it:
//   /tmp/<prefix>-*   markers and evidence   -> <runDir>/ok/<prefix>-*
//   planDir           the folder the list names -> runDir
//   cd <planRepo>     the checkout the list names -> cd <repo>
//   planLog, and its -knob.log               -> log, knobLog
// After this no path of the command is a marker in /tmp, so a file left
// there by an earlier run or by hand cannot make it pass.
export function rewriteVerify(
  cmd,
  { runDir, prefix, planDir, planRepo, repo, planLog, log, knobLog } = {},
) {
  if (!runDir) throw new Error("rewriteVerify needs runDir");
  if (!/^s\d+$/.test(String(prefix))) {
    throw new Error(`rewriteVerify needs a stage prefix like s25, not ${prefix}`);
  }
  const map = markerMap(prefix, `${noSlash(runDir)}/ok`);
  if (planDir) map[noSlash(planDir)] = noSlash(runDir);
  if (planRepo && repo) map[`cd ${noSlash(planRepo)} `] = `cd ${repo} `;
  if (planLog && log) map[planLog] = log;
  if (planLog && knobLog) map[planLog.replace(/\.log$/, "") + "-knob.log"] = knobLog;
  return rewriteCommand(cmd, map);
}

// File names of an ok folder that are markers or evidence of the stage. The
// runner removes them before a run, so a reused run folder starts empty.
export function staleMarkers(names, prefix) {
  return names.filter((n) => n.startsWith(`${prefix}-`) && /\.(ok|evidence\.json)$/.test(n));
}

// What to do with the Word process the runner started, from the answer of
// word-com.ps1 quit. action: "none" | "kill" | "leave"; fail: the run fails.
//   kill   only when quit ran (so every open document was in the run folder
//          and was closed unsaved) and the process is still there.
//   leave  quit refused because a document outside the run folder is open
//          in that Word, or the open documents are unknown (the COM call
//          failed): unsaved work of the user may be in it, Word stays up.
export function cleanupDecision({ pid, preexisting = [], quit, runDir, alive = true }) {
  if (!pid || preexisting.includes(pid)) return { action: "none", fail: false, reason: "" };
  const dir =
    String(runDir ?? "")
      .replace(/\\+$/, "")
      .toLowerCase() + "\\";
  const docs = Array.isArray(quit?.documents) ? quit.documents : null;
  const outside = (docs ?? []).filter((d) => !String(d).toLowerCase().startsWith(dir));
  if (quit?.refused || outside.length) {
    return {
      action: "leave",
      fail: true,
      reason: `Word ${pid} left running: documents outside the run folder are open (${(quit?.outside ?? outside).join(", ")})`,
    };
  }
  if (quit?.ok === true && docs) {
    if (quit.exited === true) return { action: "none", fail: false, reason: "" };
    return { action: "kill", fail: false, reason: `Word ${pid} still alive after quit` };
  }
  if (!alive) return { action: "none", fail: false, reason: "" };
  return {
    action: "leave",
    fail: true,
    reason: `Word ${pid} left running: its open documents are unknown (${quit?.error ?? "no answer from quit"})`,
  };
}

// Features whose verify_command can pass only when something happened that
// the run cannot force. trigger: the daemon log line of that event.
export const CONDITIONS = {
  // A turn that ended on a Claude usage limit.
  "S2-27": { trigger: /^\[agent\] turn complete \(usage_limit\)/m },
};
const TURN_LOGGED = /^\[agent\] turn complete \(/m;

// Exit code of the run. features: [{ id, exit }] of the verify commands;
// conditional: ids given with --conditional; logs: text of the daemon logs
// this run wrote. A conditional feature that failed is
// SKIPPED-CONDITION-NOT-MET only when its trigger is provably absent: the
// logs show finished turns and none with the trigger. Otherwise it counts.
export function exitDecision({ failed, rows, features, featuresError, conditional = [], logs }) {
  const reasons = [];
  if (failed) reasons.push("the run stopped on an error");
  for (const r of rows ?? []) if (r.result === "FAIL") reasons.push(`step failed: ${r.step}`);
  if (featuresError) reasons.push(`verify commands did not run: ${featuresError}`);
  const text = typeof logs === "string" ? logs : "";
  const out = (features ?? []).map((f) => {
    if (f.exit === 0) return { ...f, status: "PASS" };
    const cond = conditional.includes(f.id) ? CONDITIONS[f.id] : null;
    if (cond && TURN_LOGGED.test(text) && !cond.trigger.test(text)) {
      return { ...f, status: "SKIPPED-CONDITION-NOT-MET" };
    }
    reasons.push(`verify command failed: ${f.id} (exit ${f.exit})`);
    return { ...f, status: "FAIL" };
  });
  return { code: reasons.length ? 1 : 0, reasons, features: out };
}

const has = (s, sub) => String(s ?? "").includes(sub);

// R1.1 all Accepted, R1.2 all Rejected, R1.3 first change Edited by you.
export function badgesMatch(view) {
  const badges = (ref) =>
    view.comments.find((c) => c.ref === ref)?.changes.map((c) => c.badge) ?? [];
  const a = badges("R1.1");
  const b = badges("R1.2");
  const c = badges("R1.3");
  return {
    pass:
      a.length > 0 &&
      a.every((x) => x === "Accepted") &&
      b.length > 0 &&
      b.every((x) => x === "Rejected") &&
      c[0] === "Edited by you",
    got: `R1.1 ${a.join("/")}; R1.2 ${b.join("/")}; R1.3 ${c.join("/")}`,
  };
}

// word-com.ps1 Norm: how the script prints a comment text.
export const norm = (s) =>
  String(s ?? "")
    .replace(/[\r\n\u000b\u0007]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

// C1: after the author accepts or rejects a change and presses Refresh, its
// rationale comment is resolved. revision: revision.json after that refresh;
// word: word-com.ps1 state (comments with `done` = Comment.Done); message:
// the pane's refresh message. Edited and pending changes keep an open comment.
export function resolvedMatch(revision, word, message = "") {
  const byText = new Map((word?.comments ?? []).map((c) => [norm(c.text), c]));
  const seen = [];
  for (const c of revision?.comments ?? []) {
    for (const ch of c.changes ?? []) {
      if (ch.state !== "accepted" && ch.state !== "rejected") continue;
      if (!ch.word_comment_id || ch.comment_missing || ch.comment_cleaned_at) continue;
      const text = commentTextOf(c.ref, ch);
      const com = byText.get(norm(text));
      seen.push({ id: ch.id, done: com ? com.done === true : null });
    }
  }
  const refused = has(message, "not marked resolved");
  return {
    pass: seen.length > 0 && seen.every((x) => x.done === true) && !refused,
    expect:
      "every accepted or rejected change has its rationale comment Done in Word; no 'not marked resolved' note",
    got:
      (seen
        .map((x) => `${x.id} ${x.done === null ? "no comment" : x.done ? "done" : "open"}`)
        .join("; ") || "no accepted or rejected change with a comment") +
      (refused ? "; pane: not marked resolved" : ""),
  };
}

// Evidence of the `providers` pane step, one entry per provider id at the
// top level: enabled and note are those of the Judgments menu, turn_enabled
// and turn_note those of the Review turns menu. pass: the list loaded and
// Claude can be chosen for both.
export function providersEvidence(step) {
  const out = {};
  const turn = new Map((step?.menus?.turn ?? []).map((o) => [o.id, o]));
  const state = new Map((step?.rows ?? []).map((r) => [r.id, r.state]));
  for (const o of step?.menus?.oneshot ?? []) {
    out[o.id] = {
      enabled: !!o.enabled,
      note: o.note ?? "",
      turn_enabled: !!turn.get(o.id)?.enabled,
      turn_note: turn.get(o.id)?.note ?? "",
      state: state.get(o.id) ?? null,
    };
  }
  const pass = step?.ok !== false && !!out.claude?.enabled && !!out.claude?.turn_enabled;
  return { pass, ...out };
}

// Review turns in a daemon log: `[review] prompt <ref> ...` starts one (a
// preview does not), `[review] turn done <ref>` ends it. overlaps: starts
// that came while another turn was open.
export function turnOverlaps(log) {
  const started = [];
  const open = new Set();
  let overlaps = 0;
  for (const line of String(log ?? "").split("\n")) {
    const s = /^\[review\] prompt (\S+) /.exec(line);
    if (s && !line.endsWith("(preview)")) {
      if (open.size) overlaps++;
      open.add(s[1]);
      started.push(s[1]);
      continue;
    }
    const e = /^\[review\] turn done (\S+)/.exec(line);
    if (e) open.delete(e[1]);
  }
  return { started, overlaps };
}

// Evidence of a Propose all run. message: the pane's summary line; log: the
// daemon log lines written during the run.
export function proposeAllEvidence({ message, log }) {
  const m = /(\d+) of (\d+) proposed\./.exec(String(message ?? ""));
  const { started, overlaps } = turnOverlaps(log);
  const proposed = m ? Number(m[1]) : 0;
  const open_before = m ? Number(m[2]) : null;
  return {
    pass: !!m && open_before > 0 && proposed === open_before && overlaps === 0,
    overlaps,
    proposed,
    open_before,
    started,
    message: String(message ?? ""),
  };
}

// ---- Stage 2.5 (run-acceptance.mjs mainS25) ----
// Each builder below returns the evidence the runner writes next to a marker
// (`ok/s25-<name>.evidence.json`; the `.ok` file only when pass is true).

// What a run uses when --features is not given.
export function defaultFeatures(stage) {
  // Stage 3: S3-41 is checked after the cleanup (--after-cleanup).
  // Stage 4: only what is connected to the Tools tab (guideline rules, count
  // and format report, selection, marked copy, anonymized copy).
  const known = {
    s2: "S2-18..S2-34",
    s25: "S25-30..S25-41",
    s3: "S3-30..S3-40",
    s4: "S4-30..S4-33,S4-37,S4-38",
  };
  if (!known[stage])
    throw new Error(`no default feature range for stage ${stage}: give --features`);
  return known[stage];
}

// The name of the manuscript copy in the run folder. The stage 2.5 feature
// list reads paper.decisions.jsonl and paper.ledger.json next to it; the
// stage 3 list reads review_manuscript.revision.json and the letter files.
export const docStem = (stage) =>
  stage === "s2" || stage === "s3" ? "review_manuscript" : "paper";

const ALIASES = new Set(["sonnet", "opus", "haiku", "default"]);
const escapeRe = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
const refusal = (select) =>
  select?.refused ?? select?.message ?? select?.error ?? "the provider was not selected";

// S25-30. step: the `models` pane step; log: the daemon log of the run.
export function modelsEvidence({ step, log }) {
  const options = (step?.options ?? []).map((o) => o.value);
  const outside_aliases = options.filter((v) => v && !ALIASES.has(v));
  const source_sdk = /^\[models\] source=sdk /m.test(String(log ?? ""));
  return {
    pass: step?.ok !== false && outside_aliases.length > 0 && source_sdk,
    options,
    outside_aliases,
    source_sdk,
    // The step asked a second time because the list was late the first time.
    retried: step?.retried === true,
    ...(step?.retried ? { first_title: step.first_title ?? null } : {}),
  };
}

// S25-32. select: `select-provider` for judgments with Ollama; split: the
// `split` step; decisions: the text of <stem>.decisions.jsonl.
export function ollamaSplitEvidence({ select, split, decisions }) {
  const models = [];
  for (const line of String(decisions ?? "").split("\n")) {
    const m = /"model":"([^"]*)"/.exec(line);
    if (m && !models.includes(m[1])) models.push(m[1]);
  }
  const selected = select?.selected?.id === "ollama" && !select?.refused;
  return {
    pass:
      selected &&
      /^Split by /.test(String(split?.message ?? "")) &&
      models.some((m) => m.startsWith("ollama:")),
    selected: select?.selected ?? null,
    split_message: split?.message ?? null,
    models,
    ...(selected ? {} : { reason: refusal(select) }),
  };
}

// The stored judgment provider across a daemon restart (Codex review of
// stage 2.5). split: the `split` step clicked as soon as the pane was back;
// log: the daemon log from the restart on; decisions: the receipts written
// after it. pass: the stored choice was restored, the split's judgments ran
// on that provider, and neither the log nor a receipt shows another one
// (Claude is the default it must not fall back to).
export function restartSplitEvidence({ provider, split, log, decisions }) {
  const text = String(log ?? "");
  const judgments = {};
  for (const m of text.matchAll(/^\[providers\] judgment provider=(\S+) /gm)) {
    judgments[m[1]] = (judgments[m[1]] ?? 0) + 1;
  }
  const receipt_models = [];
  for (const line of String(decisions ?? "").split("\n")) {
    const m = /"model":"([^"]*)"/.exec(line);
    if (m && !receipt_models.includes(m[1])) receipt_models.push(m[1]);
  }
  const restored = new RegExp(
    `^\\[providers\\] restore .*oneshot=ready provider=${escapeRe(provider)}$`,
    "m",
  ).test(text);
  return {
    pass:
      /^Split by /.test(String(split?.message ?? "")) &&
      restored &&
      judgments[provider] > 0 &&
      Object.keys(judgments).every((id) => id === provider) &&
      receipt_models.length > 0 &&
      receipt_models.every((m) => m.startsWith(`${provider}:`)),
    restored,
    judgments,
    receipt_models,
    split_message: split?.message ?? null,
  };
}

// S25-38. step: the `focus-refresh` pane step after a hand edit in Word.
export function focusEvidence(step) {
  const external = step?.external ?? [];
  return {
    pass:
      step?.changed === true && /^Auto-refreshed /.test(step?.after ?? "") && external.length > 0,
    auto_line: step?.after ?? null,
    external,
  };
}

// S25-40. step: the `suggest-ledger` pane step; pick: the text the ticked
// candidate holds; before: ledger rules before the step.
export function suggestEvidence({ step, pick, before }) {
  const ticked = (step?.candidates ?? []).filter((c) => c.checked);
  const ledger = step?.ledger ?? [];
  return {
    pass:
      ticked.some((c) => c.label.includes(pick)) &&
      /^Added [1-9]\d* ledger rule\(s\)/.test(step?.message ?? "") &&
      ledger.length > before &&
      ledger.some((l) => l.includes(pick)),
    candidates: step?.candidates ?? [],
    message: step?.message ?? null,
    ledger_before: before,
    ledger_after: ledger.length,
  };
}

// S25-41. cleanup: record.cleanup of the runner; listening: daemon ports
// still open after it.
export function cleanupEvidence({ cleanup, listening }) {
  return {
    pass:
      cleanup?.own_word_gone === true &&
      cleanup?.preexisting_alive === true &&
      !cleanup?.left_running?.length &&
      !cleanup?.daemon_error &&
      !cleanup?.kept &&
      listening === 0,
    own_word_gone: cleanup?.own_word_gone ?? null,
    preexisting_alive: cleanup?.preexisting_alive ?? null,
    left_running: cleanup?.left_running ?? [],
    daemon_error: cleanup?.daemon_error ?? null,
    listening,
    ports_free: listening === 0,
  };
}

// ---- Stage 3 (run-acceptance.mjs mainS3): the response letter ----
// Each builder returns the evidence written next to a marker
// (`ok/s3-<name>.evidence.json`; the `.ok` file only when pass is true), from
// what the pane showed, revision.json, the daemon log, Word over COM and
// `pdftotext`. The keys are the ones the stage 3 feature list reads.

// The author's decisions of the stage 3 run. As stage 2, but the deletion
// asked for in R1.4 is accepted (the letter then cites a removal) and nothing
// is left pending: a pending tracked change is a warning of the letter check.
export const S3_DECISIONS = {
  "R1.1": "accept",
  "R1.2": "reject",
  "R1.3": ["accept_edit", "accept"],
  "R1.4": "accept",
};

// S3-30. view: the Response tab; revision: revision.json. The comment of a
// card is the reviewer's text, character for character.
export function openEvidence(view, revision) {
  const cards = view?.cards ?? [];
  const comments = revision?.comments ?? [];
  const verbatim =
    cards.length === comments.length &&
    cards.every((c, i) => c.ref === comments[i].ref && c.comment === comments[i].text);
  const profile = view?.profile ?? null;
  return {
    pass: cards.length > 0 && verbatim && profile === "generic",
    cards: cards.length,
    verbatim,
    profile,
  };
}

// S3-31. rows: the proposal lines of the pane after the letter was read;
// before: profile_overrides of revision.json at that moment (nothing saved
// yet); after: the same once "Apply to this paper" was pressed.
export function proposalEvidence({ rows, before, after }) {
  const counted = (colour) =>
    (rows ?? []).some((r) => new RegExp(`-> ${colour} \\(\\d+ paragraphs?\\)`).test(r));
  return {
    pass:
      Object.keys(before ?? {}).length === 0 &&
      counted("C00000") &&
      counted("0070C0") &&
      after?.colors?.response_body === "C00000" &&
      after?.colors?.response_revised === "0070C0",
    before: before ?? null,
    after: after ?? null,
    rows: rows ?? [],
  };
}

// Text as two readers of one PDF can agree on it: compatibility forms
// (ligatures) folded, and no spaces or hyphens, which a line break adds.
const squeeze = (s) =>
  String(s ?? "")
    .normalize("NFKC")
    .replace(/[\s\-­‐‑]/g, "");

// S3-33. A second reader of the rendered PDF. layout: `pdftotext -layout`
// text; a numbered line is "<up to 3 spaces><n><spaces><text>", and a line
// without a number (a page number, a running head, a blank line, a form
// feed) is not manuscript text. citations: [{ id, quote, start, end?,
// match_count }] with the line range the daemon read (PyMuPDF).
//
// The lines from `start` on are joined, as long as their numbers go on
// without a gap. A citation agrees when the whole quote is in that text and
// begins inside the start line: at its end, however short the piece (the
// last word of the line), or anywhere in it. A quote that begins on the next
// line does not agree, and neither does one whose text is not all there.
// With `end`, the last character of the quote must be on that line.
export function secondReader(layout, citations) {
  const lines = new Map();
  for (const raw of String(layout ?? "").split(/\n|\f/)) {
    const m = /^\s{0,3}(\d+)\s+(.*)$/.exec(raw);
    if (m && !lines.has(Number(m[1]))) lines.set(Number(m[1]), m[2].trim());
  }
  const rows = (citations ?? []).map((c) => {
    const line_text = lines.get(c.start) ?? "";
    const quote = squeeze(c.quote);
    const head = squeeze(line_text).length;
    // The joined text and where each line ends in it.
    let text = "";
    const ends = [];
    for (let n = c.start; lines.has(n) && text.length < head + quote.length; n++) {
      text += squeeze(lines.get(n));
      ends.push([n, text.length]);
    }
    let agrees = false;
    for (
      let at = text.indexOf(quote);
      quote && at !== -1 && at < head;
      at = text.indexOf(quote, at + 1)
    ) {
      const last = at + quote.length - 1;
      const endLine = ends.find(([, upTo]) => last < upTo)?.[0];
      if (c.end == null || endLine === c.end) agrees = true;
    }
    return { id: c.id, start: c.start, line_text, agrees };
  });
  const unique = (citations ?? []).every((c) => c.match_count === 1);
  const second_reader_agrees = rows.length > 0 && rows.every((r) => r.agrees);
  return { pass: unique && second_reader_agrees, unique, second_reader_agrees, rows };
}

// S3-35. step: the export step of the pane; read: tools/response_read.py on
// the written letter (the colours it finds and the roles it gives them);
// exists: the docx is there.
export function exportEvidence(step, read, path, exists) {
  const body = read?.overrides?.colors?.response_body ?? null;
  const revised = read?.overrides?.colors?.response_revised ?? null;
  const colours = Object.keys(read?.evidence?.colors ?? {}).sort();
  return {
    pass:
      exists === true &&
      /^Exported/.test(step?.output ?? "") &&
      !!body &&
      !!revised &&
      colours.length === 2 &&
      colours.includes(body) &&
      colours.includes(revised),
    body,
    revised,
    path,
    colours,
  };
}

// S3-36. Where the one-line paragraph goes: in front of the paragraph of the
// lowest citation, when a citation of another paragraph stays above it.
// citations: [{ start, paragraph_index }]. null when no such place exists.
export function insertTarget(citations) {
  const sorted = [...(citations ?? [])].sort((a, b) => a.start - b.start);
  const last = sorted.at(-1);
  if (!last || !sorted.some((c) => c.paragraph_index < last.paragraph_index)) return null;
  return last.paragraph_index;
}

// S3-36. before, after: [{ id, start }] of the citations before the paragraph
// was inserted and after the line numbers were read again; insertedAtLine:
// the line number of the inserted paragraph in the new PDF. A citation at or
// below that line moved by exactly one, one above it did not move.
export function lineShift(before, after, insertedAtLine, staleShown) {
  const by = new Map((after ?? []).map((c) => [c.id, c.start]));
  const moved = [];
  const kept = [];
  const wrong = [];
  for (const c of before ?? []) {
    const want = c.start >= insertedAtLine ? 1 : 0;
    const d = by.get(c.id) - c.start;
    (d !== want || !Number.isFinite(insertedAtLine) ? wrong : want ? moved : kept).push(c.id);
  }
  return {
    pass: wrong.length === 0 && moved.length > 0 && kept.length > 0 && staleShown === true,
    moved,
    kept,
    wrong,
    stale_shown: staleShown === true,
  };
}

const passed = (view) => /^All checks passed/.test(view?.summary ?? "");

// S3-37. before: the Response tab after a quoted sentence of `ref` was edited
// in Word and the checks ran; after: once the answer was saved again and the
// line numbers updated.
export function staleEvidence(before, after, ref) {
  const stale = before?.cards?.find((c) => c.ref === ref)?.stale ?? "";
  const stale_shown = stale.includes("Quote changed in the manuscript");
  const failed_before = (before?.checks ?? []).some((l) => l.startsWith(`FAIL ${ref} stale:`));
  const passed_after = passed(after);
  return {
    pass: stale_shown && failed_before && passed_after,
    stale_shown,
    failed_before,
    passed_after,
  };
}

// S3-38. failing: the Response tab after an export was asked for with a
// forbidden expression in an answer; exportedWhileFailing: a letter file was
// written or changed by that; after: the checks once the answer was put back.
export function ledgerEvidence(failing, exportedWhileFailing, after) {
  const ledger_fail = (failing?.checks ?? []).filter((l) => /^FAIL \S+ ledger: /.test(l)).length;
  const passed_after = passed(after);
  return {
    pass:
      ledger_fail > 0 &&
      failing?.export_label === "Export anyway" &&
      exportedWhileFailing === false &&
      passed_after,
    ledger_fail,
    exported_while_failing: exportedWhileFailing,
    passed_after,
  };
}

// S3-39. logLine: the daemon's "[response] draft <ref> lang=..." line of the
// draft; previewPrompt: the review prompt the Review tab previews; lang: the
// language set in the Response tab.
export function langEvidence(logLine, previewPrompt, lang) {
  const response_prompt_lang = / lang=(\S+)/.exec(String(logLine ?? ""))?.[1] ?? null;
  const rationale_prompt_lang =
    /Write the rationale in ([^:]+):/.exec(String(previewPrompt ?? ""))?.[1] ?? null;
  const korean_found = /Korean/.test(`${logLine ?? ""}\n${previewPrompt ?? ""}`);
  return {
    pass: response_prompt_lang === lang && rationale_prompt_lang === lang && !korean_found,
    response_prompt_lang,
    rationale_prompt_lang,
    korean_found,
  };
}

// S3-40. com: word-com.ps1 open-check on the first letter; modelParagraphs:
// the paragraphs of its model.json; firstPath, secondPath: the letter that is
// open in Word and the one an export wrote meanwhile.
export function openCheckEvidence(com, modelParagraphs, firstPath, secondPath) {
  const opened = com?.ok === true;
  const repaired = com?.repaired === true;
  const paragraphs_match = opened && com.paragraphs === modelParagraphs;
  return {
    pass: opened && !repaired && paragraphs_match && !!secondPath && secondPath !== firstPath,
    opened,
    repaired,
    paragraphs_match,
    first_path: firstPath,
    second_path: secondPath ?? null,
    paragraphs: com?.paragraphs ?? null,
    model_paragraphs: modelParagraphs,
  };
}

// A change handled by the author while ANOTHER tracked change is pending in
// its paragraph: it must show its own state, not Pending (the paragraph's
// revision mark total used to decide), and its rationale comment must be
// resolved. Shown on the first change of `ref`, a pure deletion that was
// rejected or a pure insertion that was accepted; `hand` is the text of the
// author's own tracked insertion in that paragraph, still pending in Word.
// revision: revision.json after the refresh; view: the pane; word:
// word-com.ps1 state.
export function siblingMatch(revision, view, word, { ref, hand }) {
  const ch = revision?.comments?.find((c) => c.ref === ref)?.changes?.[0];
  const badge = view?.comments?.find((c) => c.ref === ref)?.changes?.[0]?.badge;
  const want =
    ch?.kind === "delete"
      ? ["rejected", "Rejected"]
      : ch?.kind === "insert"
        ? ["accepted", "Accepted"]
        : null;
  const sibling = (word?.revisions ?? []).some((r) => r.type === 1 && has(r.text, hand.trim()));
  const text = ch ? norm(commentTextOf(ref, ch)) : "";
  const com = (word?.comments ?? []).find((c) => norm(c.text) === text);
  return {
    pass: !!want && sibling && ch.state === want[0] && badge === want[1] && com?.done === true,
    expect: `${ref} first change (a pure deletion rejected, or a pure insertion accepted) reads its own state with the author's tracked "${hand.trim()}" pending in the paragraph, and its rationale comment is resolved`,
    got: `kind ${ch?.kind ?? "-"}; state ${ch?.state ?? "-"}; badge ${badge ?? "-"}; sibling revision pending ${sibling}; comment ${com ? (com.done ? "done" : "open") : "not found"}`,
  };
}

// The document of make_tracked_special.py as the pane's snapshot shows it in
// Word: every tracked run read from the paragraph OOXML (trackedRuns) must be
// found in the paragraph text Word reports, with no revision mark left over
// (placedRuns of the daemon, the function the classifier uses). paragraphs:
// snapshot paragraphs; expected: how many paragraphs hold a tracked run.
export function trackedPlaced(paragraphs, expected) {
  const tracked = (paragraphs ?? []).filter(
    (p) => (p.tracked?.length ?? 0) > 0 || (p.w_ins ?? 0) + (p.w_del ?? 0) > 0,
  );
  const bad = [];
  for (const p of tracked) {
    const placed = placedRuns(p, normalizeParaText(p.text));
    if (
      !placed ||
      placed.short.ins ||
      placed.short.del ||
      placed.runs.length !== p.tracked.length
    ) {
      const codes = (s) => [...String(s)].map((c) => c.codePointAt(0).toString(16)).join(" ");
      bad.push(
        `paragraph ${p.index}: text [${codes(p.text)}] runs ${JSON.stringify((p.tracked ?? []).map((r) => [r.kind, codes(r.text)]))}`,
      );
    }
  }
  return {
    pass: tracked.length > 0 && tracked.length === expected && bad.length === 0,
    expect: `${expected} paragraphs with a tracked run, each run found in the paragraph text Word reports`,
    got: bad.length ? bad.join("; ") : `${tracked.length} paragraphs with tracked runs, all placed`,
  };
}

// Judgments of one pane step from the runner's record of it. Each returns
// { pass, expect, got } where expect quotes the acceptance procedure.
// ---- stage 4 (Tools tab) -----------------------------------------------------

// The findings scripts/e2e/make_aux_fixture.py plants (PLANTED there), as
// the rows of the tab word them: a statistics symbol and a species name not
// in italics, a formula digit not in subscript, one paragraph without the
// indent of the others, one empty table.
export const S4_PLANTED = { italic: 2, subscript: 1, indent: 1, empty_table: 1 };

const sameJson = (a, b) => JSON.stringify(a) === JSON.stringify(b);

// S4-33. translate, ask: the steps of the pane; before, after: the paragraphs
// of the document (the pane's snapshot) around them; log: what the daemon
// wrote meanwhile. A selection request starts no agent turn and calls no
// tool, and the document reads the same afterwards.
export function selectionEvidence({ translate, ask, before, after, log }) {
  const answers = [translate, ask].map((s) => String(s?.selection?.[1] ?? "").trim());
  const tool_calls = (
    String(log ?? "").match(
      /model called tool:|message → ensure loop|^\[agent\] turn complete/gm,
    ) ?? []
  ).length;
  const document_unchanged = Array.isArray(before) && before.length > 0 && sameJson(before, after);
  return {
    pass:
      answers.every(Boolean) &&
      !translate?.message_error &&
      !ask?.message_error &&
      document_unchanged &&
      tool_calls === 0,
    tool_calls,
    document_unchanged,
    answer_chars: answers.map((a) => a.length),
  };
}

// S4-30. extract, add: the steps of the pane; guideline: the text that was
// pasted; ledgerBefore, ledgerAfter: the ledger of revision.json around the
// add; pick: the RegExp source the kept candidate was chosen by. Every
// candidate cites a sentence of the guideline, and the one that was kept is
// the one rule added, with its sentence.
export function guidelineEvidence({ extract, add, guideline, ledgerBefore, ledgerAfter, pick }) {
  const candidates = extract?.rules ?? [];
  const flat = (s) => String(s ?? "").replace(/\s+/g, " ");
  const text = flat(guideline);
  const all_cited =
    candidates.length > 0 && candidates.every((c) => c.sentence && text.includes(c.sentence));
  const had = new Set((ledgerBefore ?? []).map((i) => i.id));
  const added = (ledgerAfter ?? [])
    .filter((i) => !had.has(i.id))
    .map((i) => ({ id: i.id, text: i.text, check: i.check, source: i.source }));
  const one = added.length === 1 ? added[0] : null;
  return {
    pass:
      all_cited &&
      !!one &&
      one.source?.kind === "guideline" &&
      text.includes(flat(one.source.sentence)) &&
      new RegExp(pick, "i").test(one.source.sentence) &&
      /added to the ledger/.test(add?.message ?? ""),
    candidates: candidates.length,
    all_cited,
    added,
    message: add?.message ?? null,
  };
}

// S4-31. check: the "Run checks" step; rule: the count rule that was added
// ({ id, check }). The planted abstract is over the limit of the rule.
export function countEvidence({ check, rule }) {
  const row = (check?.checks ?? []).find((t) => t.startsWith(`${rule?.id}: `)) ?? null;
  const m = /^\S+: (\d+) found, at most (\d+) allowed/.exec(row ?? "");
  return {
    pass:
      rule?.check?.type === "count" &&
      !!m &&
      Number(m[2]) === rule.check.max &&
      Number(m[1]) > Number(m[2]),
    rule_id: rule?.id ?? null,
    found: m ? Number(m[1]) : null,
    max: m ? Number(m[2]) : null,
    row,
  };
}

// S4-32. The format findings of the same step are exactly the planted ones,
// and the check changed nothing in the document (before, after: snapshots).
export function formatEvidence({ check, planted = S4_PLANTED, before, after }) {
  const rows = (check?.checks ?? []).filter((t) => /^Paragraph \d+: /.test(t));
  const kind = (t) =>
    /should be italic/.test(t)
      ? "italic"
      : /should be subscript/.test(t)
        ? "subscript"
        : /first-line indent/.test(t)
          ? "indent"
          : /empty table/.test(t)
            ? "empty_table"
            : "other";
  const kinds = {};
  for (const t of rows) kinds[kind(t)] = (kinds[kind(t)] ?? 0) + 1;
  const sorted = (o) => Object.entries(o).sort(([a], [b]) => a.localeCompare(b));
  const document_unchanged = Array.isArray(before) && before.length > 0 && sameJson(before, after);
  return {
    pass: sameJson(sorted(kinds), sorted(planted)) && document_unchanged,
    findings: rows.length,
    kinds,
    document_unchanged,
    rows,
  };
}

// S4-37. modes: one entry per mode: { mode, name (the file the button
// wrote), script_exit (the script on the same two files), equal (the two
// outputs byte for byte) }.
export function markedEvidence(modes) {
  const list = modes ?? [];
  return {
    pass:
      sameJson(list.map((m) => m.mode).sort(), ["color", "track_changes"]) &&
      list.every((m) => !!m.name && m.script_exit === 0 && m.equal === true),
    modes: list,
  };
}

// S4-38. step: the pane step; verifier: scripts/verify-anonymized.py run by
// the runner on the manuscript and the copy; exists: the copy is on disk.
export function anonymizeEvidence({ step, verifier, exists }) {
  const rows = step?.anonymized ?? [];
  return {
    pass:
      exists === true &&
      /^Written: \S+_anonymized\.docx\.$/.test(rows[0] ?? "") &&
      rows.some((t) => /^Verifier: total 0 left; text unchanged: yes/.test(t)) &&
      verifier?.total === 0 &&
      verifier?.body_equal === true,
    verifier: verifier
      ? {
          total: verifier.total,
          body_equal: verifier.body_equal,
          markup_equal: verifier.markup_equal,
          structure_ok: verifier.structure_ok,
        }
      : null,
    rows,
  };
}

const answered = (r, head) => ({
  pass: !r.message_error && head.test(r.selection?.[0] ?? "") && !!(r.selection?.[1] ?? "").trim(),
  expect: "the head line and a non-empty answer under the Selection buttons",
  got: r.message_error ? r.message : (r.selection ?? []).join(" | ").slice(0, 200),
});
const written = (r, rows) => ({
  pass: !r.message_error && /^Written: \S+\.docx/.test(rows?.[0] ?? ""),
  expect: "Written: <name>.docx ...",
  got: r.message_error ? r.message : (rows ?? []).join(" | "),
});

export const JUDGES = {
  "open-review": (r, ctx) => ({
    pass: r.file === `Review file: ${ctx.revisionPath}`,
    expect: `Review file: ${ctx.revisionPath}`,
    got: r.file,
  }),
  split: (r) => ({
    pass:
      /^Split by /.test(r.message) && r.draft_refs.filter((x) => /^R1\.\d$/.test(x)).length === 5,
    expect: "Split by regex: 5 comment(s) (R1.1-R1.5 after the non-reviewer rows go)",
    got: `${r.message} refs=${r.draft_refs.join(",")}`,
  }),
  "edit-draft": (r) => {
    const refs = r.comments.map((c) => c.ref).join(",");
    const r13 = r.comments.find((c) => c.ref === "R1.3")?.reqs_line ?? "";
    return {
      pass:
        refs === "R1.1,R1.2,R1.3,R1.4,R1.5" &&
        has(
          r13,
          "Requirements: q1 say how growth was followed; q2 describe the synthetic carbon source as a novel substrate",
        ),
      expect: "cards R1.1-R1.5; R1.3 Requirements: q1 ...; q2 ...",
      got: `${refs} | ${r13}`,
    };
  },
  "ledger-add": (r) => ({
    pass:
      r.ledger.length === 3 &&
      has(r.ledger[0], "L1 instruction") &&
      has(r.ledger[1], "L2 forbid") &&
      has(r.ledger[2], "L3 instruction"),
    expect: "L1 instruction, L2 forbid, L3 instruction",
    got: r.ledger.join(" | "),
  }),
  "check-all": (r) => {
    const n = (id) => r.violations.filter((v) => v.startsWith(`${id}:`)).length;
    const want = ["R1.3 q1 has no change", "R1.3 q2 has no change"];
    return {
      pass:
        n("L1") === 2 &&
        n("L2") === 1 &&
        n("L3") === 0 &&
        want.every((w) => r.decisions.some((d) => d.startsWith(w))),
      expect: "violations L1 2, L2 1, L3 0; waiting: R1.3 q1, R1.3 q2",
      got: `L1 ${n("L1")}, L2 ${n("L2")}, L3 ${n("L3")}; ${r.decisions.join(" | ")}`,
    };
  },
  preview: (r) => ({
    pass: ["[L1]", "[L2]", "[L3]"].every((t) => has(r.prompt, t)),
    expect: "prompt shows [L1], [L2], [L3]",
    got: ["[L1]", "[L2]", "[L3]"].filter((t) => has(r.prompt, t)).join(" "),
  }),
  propose: (r, ctx) => {
    const re = new RegExp(
      `^${ctx.ref.replace(".", "\\.")} running \\d+:\\d\\d\\. \\d+ of \\d+ comments not done\\.$`,
    );
    const ticks = r.progress_samples.filter((s) => re.test(s));
    return {
      pass: !r.message_error && ticks.length >= 2 && ctx.turnCompleted === true,
      expect: `${ctx.ref} running m:ss. n of 5 comments not done. (changing), turn complete in the log, no error`,
      got: `${ticks.slice(0, 2).join(" -> ")} | turn_complete=${ctx.turnCompleted} | ${r.message}`,
    };
  },
  refresh: (r) => {
    const m = /Changes (\d+), expected tool comments (\d+), found (\d+)/.exec(r.message);
    return {
      // The parts of a rewritten paragraph share one comment: every comment
      // that is expected is there, and there are not more comments than changes.
      pass: !!m && Number(m[1]) >= Number(m[2]) && m[2] === m[3] && !has(r.message, "MISMATCH"),
      expect: "Changes N, expected tool comments M, found M, M <= N. (no MISMATCH)",
      got: r.message,
    };
  },
  // A document in compatibility mode 14: Word refuses Comment.resolved. The
  // refresh still shows its counts, and one note for the handled changes.
  "refresh-mode14": (r) => {
    const m =
      /^Accepted (\d+), rejected (\d+), .*Changes (\d+), expected tool comments (\d+), found (\d+)/.exec(
        r.message,
      );
    const n = /(\d+) rationale comment\(s\) not marked resolved \(Word: /.exec(r.message);
    const handled = m ? Number(m[1]) + Number(m[2]) : 0;
    return {
      pass:
        !!m &&
        !!n &&
        handled > 0 &&
        Number(n[1]) === handled &&
        m[3] === m[4] &&
        m[4] === m[5] &&
        !has(r.message, "MISMATCH"),
      expect:
        "Accepted a, rejected r ... Changes N, expected tool comments N, found N. (a+r) rationale comment(s) not marked resolved (Word: ...)",
      got: r.message,
    };
  },
  // The badge part is what the procedure asks a person to confirm
  // (s2-list-matches); the outside-change part is L4.
  "refresh-after-user": (r) => {
    const b = badgesMatch(r);
    const at = r.external.findIndex((x) => has(x, "Lost: 0.82; 0.91"));
    return {
      pass: b.pass && has(r.message, "Outside changes:") && at >= 0,
      expect:
        "R1.1 Accepted, R1.2 Rejected, R1.3 first Edited by you; Outside changes: n; an outside item Lost: 0.82; 0.91",
      got: `${b.got}; outside item ${at + 1} of ${r.external.length}: ${r.external[at] ?? "-"}`,
    };
  },
  "clean-tool": (r) => ({
    pass: /Tool comments \d+ -> 0, your comments 1 -> 1 \(unchanged\)/.test(r.message),
    expect: "Tool comments N -> 0, your comments 1 -> 1 (unchanged)",
    got: r.message,
  }),
  "delete-all": (r) => ({
    pass:
      has(r.offer, "Back up and delete 2 comments (tool 0, yours 2)") &&
      has(r.message, "Backed up 2, deleted 2 of 2") &&
      !has(r.message, "MISMATCH"),
    expect: "Back up and delete 2 comments (tool 0, yours 2) then Backed up 2, deleted 2 of 2",
    got: `${r.offer} | ${r.message}`,
  }),
  // ---- stage 2.5 steps ----
  models: (r) => ({
    pass: /^Models your Claude account can use/.test(r.title ?? "") && r.options?.length > 0,
    expect: "the model menu filled from the account (tooltip: Models your Claude account can use)",
    got: `${r.title} | ${(r.options ?? []).map((o) => o.value).join(",")}`,
  }),
  "propose-stop": (r) => ({
    pass: r.stopped === true && r.agent_status === "Ready",
    expect: "Stop pressed while the turn ran; the pane is Ready again",
    got: r.stopped ? `stopped after ${r.seconds} s, status ${r.agent_status}` : String(r.reason),
  }),
  providers: (r) => {
    const e = providersEvidence(r);
    const ids = (k) =>
      Object.keys(e)
        .filter((id) => id !== "pass" && e[id][k])
        .join(",");
    return {
      pass: e.pass,
      expect: "the list loaded; Claude selectable for review turns and judgments",
      got: `judgments: ${ids("enabled")}; review turns: ${ids("turn_enabled")}`,
    };
  },
  "select-provider": (r, ctx) => ({
    pass: r.selected?.id === ctx.id && !r.refused && !r.message,
    expect: `${ctx.id} selected, no message`,
    got: `${r.selected?.id}${r.selected?.model ? `:${r.selected.model}` : ""}${r.refused ? ` | not selectable: ${r.refused}` : ""}${r.message ? ` | ${r.message}` : ""}`,
  }),
  "load-review-file": (r) => ({
    pass: /^Loaded .+\. Split into comments next\.$/.test(r.message) && r.chars > 0,
    expect: "Loaded <file>. Split into comments next. (text in the paste box)",
    got: `${r.message} chars=${r.chars}`,
  }),
  "export-ledger": (r, ctx) => {
    const m = /^Exported (\d+) ledger rule\(s\) to (.+)$/.exec(r.message ?? "");
    return {
      pass: !!m && (ctx.count == null || Number(m[1]) === ctx.count),
      expect: `Exported ${ctx.count ?? "n"} ledger rule(s) to <stem>.ledger.json`,
      got: r.message,
    };
  },
  "import-ledger": (r, ctx) => {
    const m = /^Added (\d+) ledger rule\(s\), skipped (\d+) already in the ledger\.$/.exec(
      r.message ?? "",
    );
    return {
      pass:
        !!m &&
        (ctx.added == null || Number(m[1]) === ctx.added) &&
        (ctx.skipped == null || Number(m[2]) === ctx.skipped),
      expect: `Added ${ctx.added ?? "a"} ledger rule(s), skipped ${ctx.skipped ?? "s"} already in the ledger.`,
      got: r.message,
    };
  },
  "suggest-ledger": (r) => {
    const picked = (r.candidates ?? []).filter((c) => c.checked).length;
    return {
      pass: picked > 0 && /^Added [1-9]\d* ledger rule\(s\)/.test(r.message ?? ""),
      expect: "candidates listed, the ticked ones added to the ledger",
      got: `${(r.candidates ?? []).length} candidate(s), ${picked} ticked | ${r.message}`,
    };
  },
  "propose-all": (r, ctx) => {
    const e = proposeAllEvidence({ message: r.message, log: ctx.log });
    return {
      pass: e.pass && !r.message_error,
      expect: "Propose all: n of n proposed. One review turn at a time.",
      got: `${r.message} | started ${e.started.join(",")} overlaps=${e.overlaps}`,
    };
  },
  // "Rewrite this paragraph": the turn ran and left changes of a rewrite
  // under its comment (the flag of a change row, review-view.js changeFlags).
  "rewrite-paragraph": (r, ctx) => {
    const changes = (r.comments ?? []).find((c) => c.ref === ctx.ref)?.changes ?? [];
    const n = changes.filter((ch) => (ch.flags ?? []).includes("Rewrite")).length;
    return {
      pass: !r.message_error && n > 0,
      expect: `changes flagged Rewrite under ${ctx.ref}, no error`,
      got: `${n} of ${changes.length} changes flagged Rewrite | ${r.message}`,
    };
  },
  "check-paragraphs": (r) => ({
    pass:
      !r.message_error &&
      /^(Checked \d+ revised paragraphs?( with .+?)?: (\d+ problems?|no problem found)\.|No paragraph holds two or more changes: nothing to check\.)/.test(
        r.message,
      ),
    expect: "Checked N revised paragraph(s)...: K problem(s) | no problem found.",
    got: `${r.message} | ${(r.consistency ?? []).length} line(s)`,
  }),
  "focus-refresh": (r) => ({
    pass: r.changed === true && /^Auto-refreshed /.test(r.after ?? ""),
    expect: "Auto-refreshed <time>: ... (no Refresh click)",
    got: r.after || "(the auto-refresh line did not change)",
  }),
  // ---- stage 4 steps (Tools tab) ----
  "open-tools": (r) => ({
    pass: /^Saved file: \S+\.docx\. Profile: /.test(r.doc ?? ""),
    expect: "Saved file: <name>.docx. Profile: <profile>. ...",
    got: r.doc || r.message,
  }),
  "tools-translate": (r) => answered(r, /^Translation into /),
  "tools-ask": (r) => answered(r, /^Answer /),
  "tools-extract": (r) => ({
    pass: !r.message_error && /^\d+ rule candidates\./.test(r.message ?? "") && r.rules.length > 0,
    expect: "n rule candidates. (each with the guideline sentence it cites)",
    got: `${r.message} | ${(r.rules ?? []).map((c) => c.kind).join(" | ")}`,
  }),
  "tools-add": (r) => ({
    pass: !r.message_error && r.kept === 1 && /^1 rule added to the ledger/.test(r.message ?? ""),
    expect: "1 rule added to the ledger. (one candidate kept ticked)",
    got: `kept ${r.kept} | ${r.message}`,
  }),
  "tools-check": (r) => ({
    pass: !r.message_error && /^\d+ findings\. Nothing was changed\.$/.test(r.message ?? ""),
    expect: "n findings. Nothing was changed.",
    got: `${r.message} | ${(r.checks ?? []).length} rows`,
  }),
  "tools-marked": (r) => written(r, r.marked),
  "tools-anonymize": (r) => written(r, r.anonymized),
  // ---- stage 3 steps (Response tab) ----
  "open-response": (r) => ({
    pass:
      !r.message_error &&
      /^Refreshed from the manuscript: \d+ comments?\.$/.test(r.message ?? "") &&
      (r.cards ?? []).length > 0,
    expect: "Refreshed from the manuscript: n comments. (one card per comment)",
    got: `${r.message} | cards ${(r.cards ?? []).map((c) => c.ref).join(",")} | profile ${r.profile}`,
  }),
  "load-letter": (r) => ({
    pass: /^Read the letter: /.test(r.message ?? "") && (r.rows ?? []).length > 0,
    expect: "Read the letter: n settings differ ... (proposal rows listed, nothing saved)",
    got: `${r.message} | ${(r.rows ?? []).join(" | ")}`,
  }),
  "apply-proposal": (r) => ({
    pass: !r.message_error && /^Applied to this paper: /.test(r.message ?? ""),
    expect: "Applied to this paper: n settings, m earlier answers kept for drafting.",
    got: r.message,
  }),
  "response-set-lang": (r) => ({
    pass: !r.message_error && /^Saved for this paper: profile /.test(r.message ?? ""),
    expect: "Saved for this paper: profile <name>.",
    got: r.message,
  }),
  // ctx.comments: the refs declined as a whole. No open request is left on any card.
  "response-decline": (r, ctx = {}) => {
    const cards = r.cards ?? [];
    const declined = (ctx.comments ?? []).every(
      (ref) => !!cards.find((c) => c.ref === ref)?.declined,
    );
    const open = cards.flatMap((c) => c.open_reqs ?? []);
    return {
      pass: !r.message_error && declined && open.length === 0,
      expect: `${(ctx.comments ?? []).join(", ")} declined with a reason; no request left open`,
      got: `${r.message} | declined ${cards
        .filter((c) => c.declined)
        .map((c) => c.ref)
        .join(",")} | open requests ${open.length}`,
    };
  },
  // The line may name what the draft ran on: "Drafted R1.1 with Claude · Opus 5.5."
  "response-draft": (r, ctx = {}) => ({
    pass:
      !r.message_error &&
      new RegExp(`^Drafted ${escapeRe(ctx.ref)}( with .+?)?\\.( |$)`).test(r.message ?? ""),
    expect: `Drafted ${ctx.ref}. (or: Drafted ${ctx.ref} with <provider and model>.)`,
    got: r.message,
  }),
  "response-draft-all": (r) => {
    const missing = (r.cards ?? []).filter((c) => !c.declined && !String(c.text ?? "").trim());
    return {
      pass:
        !r.message_error &&
        /^(Drafted \d+ responses?( with .+?)?\.( |$)|Nothing to draft)/.test(r.message ?? "") &&
        missing.length === 0,
      expect: "Drafted n responses. (every comment answered or declined)",
      got: `${r.message} | without an answer: ${missing.map((c) => c.ref).join(",") || "-"}`,
    };
  },
  "response-edit": (r) => ({
    pass: !r.message_error && /^Saved \S+\.$/.test(r.message ?? ""),
    expect: "Saved <ref>.",
    got: r.message,
  }),
  "response-refresh": (r) => ({
    pass: !r.message_error && /^Refreshed from the manuscript: /.test(r.message ?? ""),
    expect: "Refreshed from the manuscript: n comments.",
    got: r.message,
  }),
  "response-locate": (r) => {
    const m = /^Line numbers updated: (\d+)\/(\d+) quotes located\.$/.exec(r.message ?? "");
    return {
      pass: !r.message_error && !!m && m[1] === m[2] && Number(m[2]) > 0,
      expect: "Line numbers updated: n/n quotes located.",
      got: r.message,
    };
  },
  // No failure and no warning.
  "response-check": (r) => ({
    pass: r.summary === "All checks passed",
    expect: "All checks passed",
    got: `${r.summary} | ${(r.checks ?? []).join(" | ")}`,
  }),
  // After S3-34: no failure. A warning does not fail the step (an answer
  // that was edited or drafted again may differ from the one before it).
  "response-check-no-fail": (r) => ({
    pass: /^All checks passed/.test(r.summary ?? ""),
    expect: "All checks passed (warnings allowed)",
    got: `${r.summary} | ${(r.checks ?? []).join(" | ")}`,
  }),
  "response-export": (r) => ({
    pass: !r.message_error && /^Exported/.test(r.output ?? ""),
    expect: "Exported: <path>",
    got: r.output || r.message,
  }),
  "response-export-text": (r) => ({
    pass: !r.message_error && /^Exported/.test(r.output ?? ""),
    expect: "Exported: <stem>.response.R1.txt ...",
    got: r.output || r.message,
  }),
};

export function judge(name, record, ctx = {}) {
  const fn = JUDGES[name];
  if (!fn) return { pass: record.ok !== false, expect: "-", got: "-" };
  if (record.ok === false) return { pass: false, expect: "-", got: record.error };
  return fn(record, ctx);
}

// Fixed-width table of { step, pass, got } rows.
export function table(rows) {
  const w = Math.max(4, ...rows.map((r) => r.step.length));
  const line = (a, b, c) => `${a.padEnd(w)}  ${b.padEnd(7)}  ${c}`;
  return [
    line("step", "result", "detail"),
    ...rows.map((r) => line(r.step, r.result, String(r.got ?? "").slice(0, 160))),
  ].join("\n");
}
