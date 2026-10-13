// Pure helpers of the acceptance runner: the Word COM plan built from
// revision.json, the expected pane texts of each step (wording of the
// stage-2 acceptance procedure), and the path rewrite that runs a
// feature_list verify_command against this run's files.

import { formatRationaleComment, TOOL_COMMENT_PREFIX } from "../../daemon/review-tools.mjs";
import { UNSAVED_TEXT, UNSAVED_UNKNOWN } from "../../daemon/submission-check.mjs";
import { normalizeParaText, placedRuns } from "../../daemon/revision-classify.mjs";
import { rewriteSegments } from "../../daemon/word-diff.mjs";
import { paragraphHash } from "../../daemon/doc-diff.mjs";
import { narrowEdit } from "../../daemon/edit-facts.mjs";

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
  const at = escapeRe(ref);
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
// end of the turn; waiting: what the pane said 4 s and 20 s into the wait,
// and whether its progress line was in view where the click left the user
// (progress_in_view); progress: the progress lines a propose step saw.
// shot: what the user sees; shot_full: the whole tab; look_at: what was
// scrolled into view for the picture after the step (card, message, null);
// viewport: the size of the pane's window in CSS pixels and its pixel ratio.
const base = (p) => String(p).split(/[\\/]/).pop();
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
    shot: result?.shot ? shot : null,
    shot_full: result?.shot_full ? base(result.shot_full) : null,
    shot_error: result?.shot_error ?? null,
    look_at: result?.look_at ?? null,
    viewport: result?.viewport ?? null,
    waiting: (result?.waiting ?? []).map((w) => ({
      ...w,
      ...(w.shot ? { shot: base(w.shot) } : {}),
      ...(w.shot_full ? { shot_full: base(w.shot_full) } : {}),
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
  // A change of R1.4 recorded as a pure deletion.
  "S2-31": { trigger: /^\[review\] change recorded R1\.4-\d+ kind=delete\b/m },
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
      /^Split into /.test(String(split?.message ?? "")) &&
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
      /^Split into /.test(String(split?.message ?? "")) &&
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

// A plain word of the paragraph of `change` for the probe-tracked-at step:
// one of the words around the change (change.context, the left side first,
// nearest the change first), so not part of it, found once as a whole word
// in the paragraph as it read before the change, and not a word of `avoid`
// (the author's own insertion). null when there is none.
export function plainWordOf(change, avoid = "") {
  const words = (s) => String(s ?? "").match(/[A-Za-z]{3,}/g) ?? [];
  const before = String(change?.before ?? "");
  const once = (w) => (before.match(new RegExp(`\\b${w}\\b`, "g")) ?? []).length === 1;
  const shun = new Set(words(avoid).map((w) => w.toLowerCase()));
  const own = new Set(words(`${change?.find ?? ""} ${change?.text ?? ""}`));
  const near = [...words(change?.context?.left).reverse(), ...words(change?.context?.right)];
  return near.find((w) => once(w) && !own.has(w) && !shun.has(w.toLowerCase())) ?? null;
}

// The probe-tracked-at step (pane-driver.cjs): the pane's insertionHolds
// read a word inside the author's pending tracked insertion as inside it
// (true) and a plain word of the same paragraph as not (false). Any other
// answer, null (Word could not say) included, fails. hand: the author's
// insertion; insertionAround on the whole of it must read held with nothing
// left on either side and its text (result.whole): the condition on which
// the pane refuses a deletion of it (WholeAuthorInsertion).
export function trackedProbeMatch(result, hand) {
  const changes = (result?.changes ?? []).map((c) => `${c.type}:${JSON.stringify(c.text)}`);
  // insertionAround on "Checked", the start of the insertion "Checked: ":
  // held, nothing left of it, ": " right of it (ContainsStart and the
  // TrackedChange text, what in_insertion's remainders rest on).
  const a = result?.around;
  const aroundOk =
    a?.held === true &&
    a.left === "" &&
    typeof a.right === "string" &&
    normalizeParaText(a.right).trim() === ":";
  const w = result?.whole;
  const norm = (s) => normalizeParaText(String(s)).trim();
  const wholeOk =
    w?.held === true &&
    w.left === "" &&
    w.right === "" &&
    typeof w.text === "string" &&
    typeof hand === "string" &&
    norm(w.text) === norm(hand);
  const side = (s) => (typeof s === "string" ? JSON.stringify(s) : "-");
  return {
    pass: result?.inside === true && result?.plain === false && aroundOk && wholeOk,
    expect:
      'inside the author\'s tracked insertion true with nothing of it left and ":" right of the word, the whole insertion held with nothing left on either side, a plain word of its paragraph false',
    got:
      (result?.threw ?? result?.error)
        ? `threw ${result.threw ?? result.error}`
        : `inside ${result?.inside ?? "-"} (hits ${result?.inside_hits ?? "-"}); around held ${a?.held ?? "-"} left ${side(a?.left)} right ${side(a?.right)}; whole held ${w?.held ?? "-"} left ${side(w?.left)} right ${side(w?.right)} text ${side(w?.text)} (hits ${result?.whole_hits ?? "-"}); plain ${result?.plain ?? "-"} (hits ${result?.plain_hits ?? "-"}); changes ${changes.join(", ") || "-"}`,
  };
}

// The change of `ref` the pending-sibling case is demonstrated on: the
// FIRST pure deletion or insertion, or null when the ref has none (a
// replace carries no single author handling to read).
export function demoChange(revision, ref) {
  return (
    revision?.comments
      ?.find((c) => c.ref === ref)
      ?.changes?.find((c) => c.kind === "delete" || c.kind === "insert") ?? null
  );
}

// A change handled by the author while ANOTHER tracked change is pending in
// its paragraph: it must show its own state, not Pending (the paragraph's
// revision mark total used to decide), and its rationale comment must be
// resolved. Shown on the change demoChange picks: a pure deletion that was
// rejected or a pure insertion that was accepted; `hand` is the text of the
// author's own tracked insertion in that paragraph, still pending in Word.
// `shown` is false when `ref` has no such change and the case cannot be
// demonstrated.
// revision: revision.json after the refresh; view: the pane; word:
// word-com.ps1 state.
export function siblingMatch(revision, view, word, { ref, hand }) {
  const changes = revision?.comments?.find((c) => c.ref === ref)?.changes;
  const ch = demoChange(revision, ref);
  const expect = `${ref} first change (a pure deletion rejected, or a pure insertion accepted) reads its own state with the author's tracked "${hand.trim()}" pending in the paragraph, and its rationale comment is resolved`;
  if (!ch) {
    return {
      pass: false,
      shown: false,
      expect,
      got: `${ref} has no pure deletion or insertion: ${(changes ?? []).map((c) => c.kind).join(", ") || "-"}`,
    };
  }
  const badge = view?.comments?.find((c) => c.ref === ref)?.changes?.[changes.indexOf(ch)]?.badge;
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
    shown: true,
    expect,
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

// ---- the cite stage: tracked edits next to reference-manager fields ----
// The document is tests/fixtures/citation_manuscript.docx
// (make_citation_fixture.py); what is judged is what
// tools/citation_fields_check.py finds in the files Word saved, and for the
// edits of the run itself also what the tool answered: each case of
// make_citation_fixture.py CASES names the answer it expects (applied with
// no field touched; refused as FieldNotEditable; applied with a field whole
// in the deletion). A paragraph with fields must not be rewritten as a
// whole. What a model turn does is recorded; its fields are judged.

// The comments of tests/fixtures/cite_review.txt. R1.1 to R1.3 get a model
// turn each; the edits of the run itself are put on file under R1.4.
export const CITE_REFS = ["R1.1", "R1.2", "R1.3", "R1.4"];
export const CITE_TURN_REFS = ["R1.1", "R1.2", "R1.3"];
export const CITE_SEED_REF = "R1.4";

// One case of make_citation_fixture.py CASES as the edit the daemon stores
// for a blocked edit (FORBID_GUARD_EDIT has the same form).
export const citeEdit = (c, ref = CITE_SEED_REF) => ({
  comment_ref: ref,
  find: c.find,
  replace: c.replace,
  rationale: {
    what: `Acceptance run: ${c.what}.`,
    why: "Written by the acceptance run to see what a tracked edit does to a field; no model wrote it.",
    addresses: "the wording of the Introduction and the Methods",
  },
});

// The code of a refusal, from its message: the pane shows the message of a
// refused edit (review-pane.js override), not its code. The texts are those
// of review-tools.mjs applyTrackedEdit and tools-word.js toolTrackedEdit.
const REFUSAL_TEXTS = [
  [/is not rewritten as a whole/, "ParagraphNotPlain"],
  [/could not read this paragraph as Word shows it/, "FieldsUnread"],
  [/would change the text of a citation or another field/, "FieldNotEditable"],
  [/" not found$/, "NotFound"],
  [/^Paragraph not found: /, "HintNotFound"],
  [/" occurs \d+ times; /, "Ambiguous"],
  [
    /find (is not in the paragraph text|occurs more than once in the paragraph)/,
    "AmbiguousInParagraph",
  ],
  [
    /could not be told from the last status refresh|changed in Word after it was last read|changed since it was read|was not checked before the edit/,
    "StaleTarget",
  ],
  [/puts a number into the manuscript/, "NumberNeedsAuthor"],
  [/whole of a tracked insertion the author made/, "WholeAuthorInsertion"],
  [/forbidden rule\(s\) .+ cannot be checked/, "ForbidRuleBroken"],
  [/matches forbidden rule\(s\)/, "ForbidBlocked"],
  [/task pane is older than the daemon/, "TaskPaneOutdated"],
  [/could not be written to the review file first/, "AttemptNotRecorded"],
  [/Word did not answer|task pane did not answer/, "OutcomeUnknown"],
];
export const refusalCode = (message) =>
  REFUSAL_TEXTS.find(([re]) => re.test(String(message ?? "")))?.[1] ?? "Refused";

const revisionCount = (word) => (word?.ok === true ? (word.revisions ?? []).length : null);

// What became of one edit the run put on file and sent with "Allow
// exception" (pane step apply-seeded).
//   step      the answer of the pane step
//   log       the daemon log written since the step began
//   before    revision.json with the item on file, revision: after the step
//   word, wordBefore   word-com.ps1 state after and before
//   seedId    the id of the item
// answer: "applied" (the daemon recorded a change, or logged the approved
// exception), "refused" (the pane showed the refusal; code and message say
// which) or "unknown" (the step gave no answer that can be read).
export function citeOutcome({ step, log, before, revision, word, wordBefore, seedId, ref }) {
  const of = (rev) => (rev?.comments ?? []).find((c) => c.ref === ref)?.changes ?? [];
  const had = new Set(of(before).map((ch) => ch.id));
  const changes = of(revision)
    .filter((ch) => !had.has(ch.id))
    .map((ch) => ({
      id: ch.id,
      kind: ch.kind,
      find: ch.find,
      text: ch.text,
      narrowed: !!ch.narrowed_from,
      paragraph_index: ch.paragraph_index ?? null,
      word_comment_id: ch.word_comment_id ?? null,
      comment_missing: ch.comment_missing ?? null,
    }));
  const approved = String(log ?? "").includes(`[review] override ${seedId} approved`);
  const n = [revisionCount(wordBefore), revisionCount(word)];
  const lines = String(log ?? "")
    .split("\n")
    .filter((l) => l.startsWith("[review] "));
  const base = { changes, word_revisions: n, log: lines };
  if (changes.length || approved) return { answer: "applied", code: null, message: null, ...base };
  if (step?.ok !== false && step?.message_error === true) {
    return {
      answer: "refused",
      code: refusalCode(step.message),
      message: step.message,
      // A refusal must leave the document as it was.
      word_unchanged: n[0] !== null && n[0] === n[1],
      ...base,
    };
  }
  return { answer: "unknown", code: null, message: step?.error ?? step?.message ?? null, ...base };
}

// One line about what the checker found: the counts, then the first field
// that is wrong with the start of its paragraph, its kind and its markers.
export function fieldsLine(check) {
  if (!check || typeof check !== "object" || !check.after) {
    return `no checker result (${check?.error ?? "the checker did not run"})`;
  }
  const s = check.after.by_state ?? {};
  // A field deleted whole whose text was typed again as plain words passes,
  // and is told: after accept-all that citation is text.
  const flat = (check.flattened ?? []).map((f) => `${f.kind} ${f.shown}`);
  // The invariants of an edit, each with its count: no field split by a
  // tracked change, none left with its result gone (hollowed), no result
  // other than before.
  const hollow = check.hollowed ?? [];
  const other = check.result_changed ?? [];
  // w:fldData that differs and is not judged (Word rewrote it on its own
  // save of the fixture: run-acceptance mainCite).
  const data = check.fld_data?.differs && check.fld_data.judged === false;
  const counts =
    `fields ${check.before.fields} -> ${check.after.fields}` +
    ` (live ${s.live ?? 0}, in deletions ${s.deleted ?? 0}, in insertions ${s.inserted ?? 0}, split ${s.split ?? 0},` +
    ` hollowed ${hollow.length}, result changed ${other.length})` +
    (flat.length ? `; replaced by plain text: ${flat.join("; ")}` : "") +
    (data ? "; fldData differs (not judged)" : "");
  if (check.ok) return counts;
  const f =
    hollow[0] ?? other[0] ?? check.split?.[0] ?? check.missing?.[0] ?? check.added?.[0] ?? null;
  const u = check.unbalanced?.[0];
  const result = f && "was" in f ? ` result "${f.was}" -> "${f.accepted ?? ""}"` : "";
  const first = f
    ? ` first: ${f.kind} in "${f.text}" ${f.layout}${result}`
    : u
      ? ` first: ${u.what} in ${u.story} p${u.paragraph}`
      : "";
  return `${counts}; FAILED ${check.failures.join(", ")};${first}`;
}

// What the daemon logged about an edit refused for a field
// (review-tools.mjs): "why=<reason> by=<daemon|pane> ...".
const FIELD_REFUSED = /^\[review\] field edit refused \S+ (.*)$/;

// What a case of the run expects of the tool (make_citation_fixture.py
// CASES expect), in words.
const CITE_EXPECT = {
  applied: () => "applied, no field touched",
  refused: () => "refused FieldNotEditable",
  removed: (n) => `applied, ${n} field(s) whole in the deletion`,
};

// The row of one edit of the run. The fields are judged (the checker on the
// copy before the edit against the copy after it); a refusal that changed
// the document fails; and with `expect` the answer of the tool is judged
// too: "applied" (no field newly in a deletion), "refused" (as
// FieldNotEditable) or "removed" (applied, exactly `fields` fields newly
// whole in a deletion). The row says what the tool answered and why (the
// reason the daemon logged), whether the expectation was met, and the count
// of each invariant.
export function citeCaseRow({ name, what, expect = null, fields = 0, outcome, check }) {
  const why = (outcome.log ?? []).map((l) => FIELD_REFUSED.exec(l)?.[1]).find(Boolean);
  const tool =
    outcome.answer === "applied"
      ? `applied (${outcome.changes.map((c) => `${c.kind}${c.narrowed ? ", narrowed" : ""}${c.comment_missing ? `, no comment: ${c.comment_missing}` : ""}`).join("; ") || "exception approved"})`
      : outcome.answer === "refused"
        ? `refused ${outcome.code}${why ? ` [${why}]` : ""}${outcome.word_unchanged ? "" : ", but the tracked changes of the document differ"}`
        : `no readable answer (${outcome.message ?? "-"})`;
  const gone = (side) => check?.[side]?.by_state?.deleted ?? 0;
  const removed = check?.after ? gone("after") - gone("before") : null;
  const met =
    expect === "applied"
      ? outcome.answer === "applied" && removed === 0
      : expect === "refused"
        ? outcome.answer === "refused" && outcome.code === "FieldNotEditable"
        : expect === "removed"
          ? outcome.answer === "applied" && removed === fields
          : null;
  const expected =
    met === null
      ? ""
      : ` | expected ${CITE_EXPECT[expect](fields)}: ${met ? "met" : "NOT MET"} (fields newly in deletions: ${removed ?? "?"})`;
  const pass =
    check?.ok === true &&
    met !== false &&
    outcome.answer !== "unknown" &&
    (outcome.answer !== "refused" || outcome.word_unchanged === true);
  return {
    step: `cite: ${name} (${what})`,
    result: pass ? "PASS" : "FAIL",
    ...(met === null ? {} : { expect: CITE_EXPECT[expect](fields) }),
    got: `tool: ${tool}${expected} | ${fieldsLine(check)}`,
  };
}

// The call the daemon makes first for a paragraph rewritten as a whole
// (review-tools.mjs applyTrackedRewrite): the paragraph as the pane reads it
// and the new text are cut into the edits that differ (rewriteSegments), the
// last of them is sent first, with the rationale comment and plain_only, and
// with the hash of the paragraph it was cut from.
//   paragraphs  the pane's snapshot (index, id, text, reviewed_text)
//   rewrite     make_citation_fixture.py REWRITE
// Returns { args, paragraph, before, after, parts } or { error }.
export function citeRewriteCall(paragraphs, rewrite, ref = CITE_SEED_REF) {
  const read = (p) => normalizeParaText(p.reviewed_text ?? p.text);
  const hits = (paragraphs ?? []).filter((p) => read(p).includes(rewrite.paragraph_has));
  if (hits.length !== 1)
    return { error: `${hits.length} paragraphs hold "${rewrite.paragraph_has}"` };
  const p = hits[0];
  const before = read(p);
  let after = before;
  for (const [from, to] of rewrite.changes) {
    if (after.split(from).length !== 2) return { error: `"${from}" is not in the paragraph once` };
    after = after.replace(from, to);
  }
  const cut = rewriteSegments(before, after);
  if (cut.error || !cut.segments.length) return { error: cut.error ?? "nothing differs" };
  const part = cut.segments[cut.segments.length - 1];
  const narrow = narrowEdit(part.find, part.replace);
  const id = p.id ?? `p${p.index}`;
  return {
    paragraph: { index: p.index, id: p.id ?? null },
    before,
    after,
    parts: cut.segments.length,
    args: {
      find: part.find,
      replace: part.replace,
      paragraph_hint: id,
      ...(narrow ? { narrow } : {}),
      comment_text: formatRationaleComment(ref, {
        what: "Acceptance run: a paragraph with citation fields written anew.",
        why: "Sent by the acceptance run to see whether the rewrite is refused; no model wrote it.",
        addresses: "the wording of the Introduction and the Methods",
      }),
      plain_only: true,
      expected_before_hash: paragraphHash(before),
      expected_paragraph: { index: p.index, id: p.id ?? null },
    },
  };
}

// The rewrite of a paragraph with fields: it must be refused as
// ParagraphNotPlain with no tracked change added, and the fields must hold.
export function citeRewriteRow({ what, call, step, word, wordBefore, check }) {
  const r = step?.result ?? null;
  const n = [revisionCount(wordBefore), revisionCount(word)];
  const answer = call?.error
    ? `not sent (${call.error})`
    : step?.ok === false
      ? `no answer (${step.error})`
      : step?.threw
        ? `the tool threw: ${step.threw}`
        : r?.ok === false && r.applied !== true
          ? `refused ${r.code ?? refusalCode(r.message)}`
          : `applied (${JSON.stringify({ ok: r?.ok, code: r?.code ?? null, narrowed: r?.narrowed ?? null })})`;
  const refused = r?.ok === false && r.code === "ParagraphNotPlain" && r.applied !== true;
  const unchanged = n[0] !== null && n[0] === n[1];
  return {
    step: `cite: rewrite (${what})`,
    result: refused && unchanged && check?.ok === true ? "PASS" : "FAIL",
    expect: "refused ParagraphNotPlain, no tracked change added, the fields whole",
    got: `tool: ${answer}; tracked changes ${n[0]} -> ${n[1]} | ${fieldsLine(check)}`,
  };
}

// Word's own list of fields of a document (marked-diff-word.ps1 Shape):
// "code U+001F result" per field, in document order.
const wordFields = (job) =>
  (job?.shape?.codes ?? []).flatMap((s) => String(s).split("\u001e").filter(Boolean));

// The saved document with every change accepted and with every change
// rejected (two copies written by Word, marked-diff-word.ps1 save_as).
//   tracked, accepted, rejected   the checker's results, each against the
//       document as Word saved it before any edit (accepted: run with
//       --expect-removed <the fields that lay whole in a deletion>)
//   jobs   by id (baseline, tracked, accept, reject): what Word answered
// pass: Word opened every file as it is and left no revision; after
// reject-all the fields are those of the original; after accept-all they are
// the original without exactly the fields that lay in deletions, each of
// the others whole and live; and Word itself lists the same field codes
// and results after reject-all as in the original.
export function citeAcceptEvidence({ tracked, accepted, rejected, jobs = {} }) {
  const ids = (list) => (list ?? []).map((f) => f.id).sort();
  const deleted = (tracked?.after?.list ?? []).filter((f) => f.state === "deleted");
  const opened = ["baseline", "tracked", "accept", "reject"].map((id) => jobs[id]);
  const original = wordFields(jobs.baseline);
  const checks = {
    word_opened_all: opened.every((j) => j?.ok === true),
    no_revision_left: jobs.accept?.revisions_left === 0 && jobs.reject?.revisions_left === 0,
    rejected_equals_original:
      rejected?.ok === true && !rejected.missing?.length && !rejected.added?.length,
    rejected_all_live: (rejected?.after?.list ?? []).every((f) => f.state === "live"),
    accepted_whole: accepted?.ok === true,
    accepted_all_live: (accepted?.after?.list ?? []).every((f) => f.state === "live"),
    accepted_lacks_only_the_deleted:
      !!accepted?.missing && JSON.stringify(ids(accepted.missing)) === JSON.stringify(ids(deleted)),
    word_lists_the_original_after_reject:
      original.length > 0 && JSON.stringify(wordFields(jobs.reject)) === JSON.stringify(original),
  };
  return {
    pass: Object.values(checks).every(Boolean),
    checks,
    deleted_fields: deleted.map((f) => f.kind),
    removed_by_accept: (accepted?.missing ?? []).map((f) => ({ kind: f.kind, text: f.text })),
    word_field_counts: {
      original: original.length,
      rejected: wordFields(jobs.reject).length,
      accepted: wordFields(jobs.accept).length,
    },
    job_errors: opened.filter((j) => j && j.error).map((j) => `${j.id}: ${j.error}`),
    // Word changed the file while opening it (a field updated, a repair): told, not judged.
    changed_on_open: opened.filter((j) => j?.dirty_on_open === true).map((j) => j.id),
  };
}

// The reference-manager plugins of the Word of the run (word-com.ps1
// addins), by name. Reported, never judged.
export function pluginsEvidence(addins) {
  const names = [
    ...(addins?.com_addins ?? []).map((a) => `${a.prog_id} ${a.description}`),
    ...(addins?.templates ?? []).map((a) => `${a.name}`),
  ];
  const has = (re) => names.filter((n) => re.test(n));
  return {
    read: addins?.ok === true,
    zotero: has(/zotero/i),
    endnote: has(/endnote|cwyw/i),
    com_addins: (addins?.com_addins ?? []).length,
    templates: (addins?.templates ?? []).length,
    errors: addins?.errors ?? (addins?.error ? [addins.error] : []),
  };
}

// ---- stage parts ---------------------------------------------------------------
// A manuscript (paper.docx) and its supplementary Word file (SI.docx) in one
// revision: the link, the copied comment list, a turn in each file, the
// forbid guard in the supplementary file, where a request stands in the
// other file, two turns at once, and leaving. The documents, the review and
// which comment is of which kind are those of make_parts_fixture.py (--spec).
//
// Every judge below returns { pass, checks, got }. checks: one boolean per
// named check. got: for each failed check its name (which pane or file), what
// was expected and what was found, so a failed row can be read from the log
// alone; when all pass, one line of what was found.

// The rule a seeded edit is put on file with (cite uses the same trick): the
// ledger of the stage has one rule, L1, so "Allow exception" lifts nothing
// that exists and L1 has to refuse the edit.
export const PARTS_SEED_RULE = "L9";

// The edit the run sends to the supplementary file for the forbid guard, in
// the form the daemon stores a blocked edit (FORBID_GUARD_EDIT has it too).
export const partsGuardEdit = (guard) => ({
  comment_ref: guard.ref,
  find: guard.find,
  replace: guard.replace,
  rationale: {
    what: "Called the carbon source novel.",
    why: "Written by the acceptance run to show the forbid guard in a supplementary file; no model wrote it.",
    addresses: "the wording of the Supplementary Methods",
  },
});

const partsComment = (rev, ref) => (rev?.comments ?? []).find((c) => c.ref === ref) ?? null;
const changeIds = (rev, ref) => (partsComment(rev, ref)?.changes ?? []).map((ch) => ch.id);
const newChanges = (before, after, ref) => {
  const had = new Set(changeIds(before, ref));
  return (partsComment(after, ref)?.changes ?? []).filter((ch) => !had.has(ch.id));
};
const baseName = (p) =>
  String(p ?? "")
    .split(/[\\/]/)
    .pop();
const show = (x) => (typeof x === "string" ? x : JSON.stringify(x));
// checks: [{ name, ok, expect, found }]. ok: what was found when all pass.
function partsVerdict(checks, ok = "") {
  const failed = checks.filter((c) => !c.ok);
  return {
    pass: failed.length === 0,
    checks: Object.fromEntries(checks.map((c) => [c.name, !!c.ok])),
    got: failed.length
      ? failed
          .map((c) => `${c.name}: expected ${show(c.expect)}, found ${show(c.found)}`)
          .join(" | ")
      : ok,
  };
}
const wordRevisions = (word) => (word?.ok === true ? (word.revisions ?? []).length : null);
// Rationale comments of the tool for `ref` among the comments Word lists.
const toolComments = (word, ref) =>
  (word?.comments ?? []).filter((c) => String(c.text ?? "").startsWith(`[${ref}]`)).length;

// a. Before the join. view: the Review tab of the supplementary file as it
// opened. si, paper: the two review files (null: no such file yet).
export function partsBeforeJoin({ view, si, paper, names, comments }) {
  const want = `Join the revision of ${names.paper}, ${comments} comment${comments === 1 ? "" : "s"}`;
  return partsVerdict(
    [
      {
        name: `${names.si} pane: the join box is shown`,
        ok: view?.link_box?.shown === true,
        expect: true,
        found: view?.link_box?.shown ?? view?.error ?? null,
      },
      {
        name: `${names.si} pane: the candidate`,
        ok: (view?.link_box?.candidates ?? []).includes(want),
        expect: want,
        found: view?.link_box?.candidates ?? [],
      },
      {
        name: `${names.si} pane: no comment row before the join`,
        ok: (view?.comments ?? []).length === 0 && view?.ok !== false,
        expect: 0,
        found: (view?.comments ?? []).map((c) => c.ref),
      },
      {
        name: `review file of ${names.si}: no link (nothing joins by itself)`,
        ok: !si?.part_of,
        expect: "no part_of",
        found: si?.part_of ?? null,
      },
      {
        name: `review file of ${names.paper}: no linked file`,
        ok: (paper?.parts ?? []).length === 0,
        expect: "no parts",
        found: paper?.parts ?? null,
      },
    ],
    `join box with "${want}"; no link in either review file`,
  );
}

// b. After the join by the candidate. view: the supplementary file's pane
// after the join; mainView: the manuscript's pane after it waited, nothing
// pressed, for the line of linked files (wait-linked).
export function partsJoined({ view, mainView, si, paper, names }) {
  const mainRefs = (paper?.comments ?? []).map((c) => c.ref);
  const rows = (view?.comments ?? []).map((c) => c.ref);
  const ids = (rev) =>
    (rev?.comments ?? []).map((c) => [c.ref, (c.requirements ?? []).map((r) => [r.id, r.text])]);
  const note = `These rules are edited in ${names.paper}.`;
  const line = `Linked files: ${names.si}`;
  return partsVerdict(
    [
      {
        name: `${names.si} pane: the rows of the manuscript`,
        ok: mainRefs.length > 0 && JSON.stringify(rows) === JSON.stringify(mainRefs),
        expect: mainRefs,
        found: view?.error ?? rows,
      },
      {
        name: `${names.si} pane: the rules are shown read-only with their note`,
        ok:
          (view?.ledger ?? []).length === (paper?.ledger ?? []).length &&
          (view?.ledger ?? []).length > 0 &&
          view?.ledger_edit_buttons === 0 &&
          view?.ledger_note === note,
        expect: `${(paper?.ledger ?? []).length} rule(s), no Edit or Remove button, "${note}"`,
        found: `${(view?.ledger ?? []).length} rule(s), ${view?.ledger_edit_buttons} Edit/Remove button(s), "${view?.ledger_note ?? ""}"`,
      },
      {
        name: `${names.si} pane: Propose changes on every row`,
        ok: rows.length > 0 && (view?.comments ?? []).every((c) => c.can_propose === true),
        expect: "a Propose changes button in every card",
        found: (view?.comments ?? []).filter((c) => !c.can_propose).map((c) => c.ref),
      },
      {
        name: `review file of ${names.si}: part_of names ${names.paper}`,
        ok: baseName(si?.part_of?.manuscript) === names.paper,
        expect: names.paper,
        found: si?.part_of ?? null,
      },
      {
        name: `review file of ${names.paper}: parts lists ${names.si}`,
        ok: (paper?.parts ?? []).some((p) => baseName(p.doc) === names.si),
        expect: names.si,
        found: paper?.parts ?? null,
      },
      {
        name: "request ids are the same in both review files",
        ok: mainRefs.length > 0 && JSON.stringify(ids(si)) === JSON.stringify(ids(paper)),
        expect: ids(paper),
        found: ids(si),
      },
      {
        name: `${names.paper} pane: "${line}" without a reload`,
        ok: mainView?.matched === true && mainView?.linked === line,
        expect: line,
        found: mainView?.error ?? mainView?.linked ?? "",
      },
    ],
    `${rows.length} rows in ${names.si}; both review files hold the link; ${names.paper} shows "${line}"`,
  );
}

// c. The turn of the manuscript for the comment that concerns only the
// supplementary file: every request noted as belonging to another file, no
// change. before, after: the manuscript's review file. word, wordBefore:
// word-com.ps1 state of the manuscript.
export function partsMainOutside({ before, after, word, wordBefore, ref, names }) {
  const c = partsComment(after, ref);
  const wanted = (c?.requirements ?? []).map((r) => r.id);
  const noted = (c?.requirements_outside ?? []).map((n) => n.id);
  const added = newChanges(before, after, ref);
  const n = [wordRevisions(wordBefore), wordRevisions(word)];
  return partsVerdict(
    [
      {
        name: `review file of ${names.paper}: ${ref} requests noted as belonging to another file`,
        ok: wanted.length > 0 && wanted.every((id) => noted.includes(id)),
        expect: wanted,
        found: `outside ${show(noted)}, questions ${(c?.questions ?? []).length}, changes ${(c?.changes ?? []).length}`,
      },
      {
        name: `review file of ${names.paper}: no change recorded for ${ref}`,
        ok: added.length === 0,
        expect: 0,
        found: added.map((ch) => `${ch.id} ${ch.kind} "${ch.text ?? ""}"`),
      },
      {
        name: `${names.paper} in Word: no new tracked change`,
        ok: n[0] !== null && n[0] === n[1],
        expect: n[0],
        found: n[1] ?? word?.error ?? null,
      },
    ],
    `outside ${show(noted)}: ${(c?.requirements_outside ?? []).map((x) => x.why).join("; ")}`,
  );
}

// d. The turn of the supplementary file for the same comment: a tracked
// change with its rationale comment in SI.docx, recorded in its review file
// and not in the manuscript's, and the manuscript's text as it was.
// paperTexts: the paragraph texts of the manuscript before and after
// (the pane's snapshot step), or null where they could not be read.
export function partsSuppChange({
  siBefore,
  si,
  paperBefore,
  paper,
  word,
  wordBefore,
  paperTexts,
  ref,
  names,
}) {
  const added = newChanges(siBefore, si, ref);
  const inMain = newChanges(paperBefore, paper, ref);
  const n = [wordRevisions(wordBefore), wordRevisions(word)];
  const comments = toolComments(word, ref);
  const [textsBefore, textsAfter] = paperTexts ?? [null, null];
  const same =
    Array.isArray(textsBefore) &&
    Array.isArray(textsAfter) &&
    JSON.stringify(textsBefore) === JSON.stringify(textsAfter);
  const differs = same
    ? null
    : Array.isArray(textsBefore) && Array.isArray(textsAfter)
      ? (textsAfter.findIndex((t, i) => t !== textsBefore[i]) ?? -1)
      : "the text could not be read";
  return partsVerdict(
    [
      {
        name: `review file of ${names.si}: a change recorded for ${ref}`,
        ok: added.length >= 1,
        expect: "at least 1",
        found: added.length,
      },
      {
        name: `${names.si} in Word: a tracked change with a [${ref}] rationale comment`,
        ok:
          n[0] !== null &&
          n[1] > n[0] &&
          comments >= 1 &&
          added.every((ch) => ch.word_comment_id != null),
        expect: `more tracked changes than ${n[0]}, a comment that begins [${ref}], every change with its comment`,
        found: `${n[1] ?? word?.error} tracked change(s), ${comments} [${ref}] comment(s), without a comment: ${show(added.filter((ch) => ch.word_comment_id == null).map((ch) => ch.id))}`,
      },
      {
        name: `review file of ${names.paper}: no change recorded for ${ref}`,
        ok: inMain.length === 0,
        expect: 0,
        found: inMain.map((ch) => ch.id),
      },
      {
        name: `${names.paper} in Word: its text is as it was`,
        ok: same,
        expect: "the same paragraphs",
        found: differs === null ? "the same" : `differs at paragraph ${differs}`,
      },
    ],
    `${added.length} change(s) in ${names.si}: ${added.map((ch) => `${ch.id} "${ch.text ?? ""}"`).join("; ")}`,
  );
}

// f. The manuscript's pane, nothing pressed, after the supplementary file
// got its change: the row of the comment and the request the change names.
// view: the wait-state step of the manuscript's pane. si: the review file of
// the supplementary file (which requests its changes name).
export function partsCrossState({ view, si, ref, names }) {
  const card = (view?.comments ?? []).find((c) => c.ref === ref);
  const pending = (partsComment(si, ref)?.changes ?? []).filter((ch) =>
    ["pending", "unlocated"].includes(ch.state),
  );
  const named = [...new Set(pending.flatMap((ch) => ch.rationale?.requirement_ids ?? []))];
  const label = new RegExp(`^\\d+ changes? to review in ${escapeRe(names.si)}$`);
  const text = `Waiting for your decision in ${names.si}`;
  const rows = (card?.requests ?? []).filter((r) => named.includes(r.id));
  return partsVerdict(
    [
      {
        name: `review file of ${names.si}: a pending change of ${ref} names a request`,
        ok: named.length > 0,
        expect: "at least one request id in rationale.requirement_ids",
        found: pending.map(
          (ch) => `${ch.id} ${ch.state} ${show(ch.rationale?.requirement_ids ?? [])}`,
        ),
      },
      {
        name: `${names.paper} pane: the row of ${ref} says the change waits in ${names.si} (no Refresh pressed)`,
        ok: view?.matched === true && label.test(card?.state ?? ""),
        expect: `n change(s) to review in ${names.si}`,
        found: view?.error ?? card?.state ?? "no such row",
      },
      {
        name: `${names.paper} pane: the request shows "${text}"`,
        ok:
          rows.length > 0 && rows.every((r) => r.state === "waiting_elsewhere" && r.text === text),
        expect: named.map((id) => `${id}: ${text}`),
        found: (card?.requests ?? []).map((r) => `${r.id}: ${r.state} "${r.text}"`),
      },
    ],
    `${ref}: ${card?.state}; ${rows.map((r) => `${r.id} ${r.text}`).join("; ")}`,
  );
}

// g. A turn in each pane at the same time for the comment that concerns
// both files. log: the daemon log written while the two ran. steps: the
// answers of the two propose steps.
export function partsBoth({
  paperBefore,
  paper,
  siBefore,
  si,
  wordPaper,
  wordPaperBefore,
  wordSi,
  wordSiBefore,
  log,
  steps,
  ref,
  names,
}) {
  const text = String(log ?? "");
  const lines = text.split("\n");
  const at = (re) => lines.findIndex((l) => re.test(l));
  const esc = escapeRe(ref);
  // Both turns were queued before either ended.
  const prompts = lines
    .map((l, i) =>
      new RegExp(`^\\[review\\] prompt ${esc} `).test(l) && !/\(preview\)$/.test(l) ? i : -1,
    )
    .filter((i) => i >= 0);
  const firstEnd = at(new RegExp(`^\\[review\\] turn done ${esc}`));
  const overlap = prompts.length >= 2 && (firstEnd < 0 || prompts[1] < firstEnd);
  const file = (label, before, after, word, wordBefore) => {
    const added = newChanges(before, after, ref);
    const n = [wordRevisions(wordBefore), wordRevisions(word)];
    const had = toolComments(wordBefore, ref);
    const now = toolComments(word, ref);
    return [
      {
        name: `review file of ${label}: a change recorded for ${ref}`,
        ok: added.length >= 1,
        expect: "at least 1",
        found: added.length,
      },
      {
        name: `${label} in Word: its own tracked change(s) with their [${ref}] comments`,
        ok: n[0] !== null && n[1] > n[0] && now - had >= 1 && now - had <= added.length,
        expect: `more tracked changes than ${n[0]}; between 1 and ${added.length} new [${ref}] comment(s)`,
        found: `${n[1] ?? word?.error} tracked change(s), ${now - had} new [${ref}] comment(s)`,
      },
    ];
  };
  const bad = (re) => lines.filter((l) => re.test(l));
  const unknown = bad(/outcome unknown|OutcomeUnknown/);
  const crossed = bad(
    /tool_result .* from another pane|tool_result for unknown id|late result dropped/,
  );
  const added = (before, after) => newChanges(before, after, ref).map((ch) => ch.id);
  return partsVerdict(
    [
      {
        name: "both propose steps ended without an error",
        ok: steps?.paper?.ok !== false && steps?.si?.ok !== false,
        expect: "two turns that ended",
        found: `${names.paper}: ${steps?.paper?.error ?? "ended"}; ${names.si}: ${steps?.si?.error ?? "ended"}`,
      },
      {
        name: "the two turns ran at the same time",
        ok: overlap,
        expect: `two "[review] prompt ${ref}" lines before the first "[review] turn done ${ref}"`,
        found: `prompt lines at ${show(prompts)}, first turn done at ${firstEnd}`,
      },
      ...file(names.paper, paperBefore, paper, wordPaper, wordPaperBefore),
      ...file(names.si, siBefore, si, wordSi, wordSiBefore),
      {
        name: "no edit whose outcome is unknown",
        ok: unknown.length === 0,
        expect: "no such log line",
        found: unknown.slice(0, 3),
      },
      {
        name: "no tool result crossed between the panes",
        ok: crossed.length === 0,
        expect: "no such log line",
        found: crossed.slice(0, 3),
      },
    ],
    `${names.paper}: ${show(added(paperBefore, paper))}; ${names.si}: ${show(added(siBefore, si))}`,
  );
}

// h. The manuscript's comment list saved without the comment that has
// changes in the supplementary file: refused, the list as it was.
// step: the save-without step of the manuscript's pane.
export function partsKeep({ step, before, after, ref, names }) {
  const refs = (rev) => (rev?.comments ?? []).map((c) => c.ref);
  const want = new RegExp(
    `Comment ${escapeRe(ref)} has recorded change\\(s\\) in ${escapeRe(names.si)}; keep it`,
  );
  return partsVerdict(
    [
      {
        name: `${names.paper} pane: the save is refused with the reason`,
        ok: step?.ok !== false && step?.message_error === true && want.test(step?.message ?? ""),
        expect: `Comment ${ref} has recorded change(s) in ${names.si}; keep it`,
        found: step?.error ?? step?.message ?? "",
      },
      {
        name: `review file of ${names.paper}: the comment list is as it was`,
        ok:
          refs(after).includes(ref) && JSON.stringify(refs(after)) === JSON.stringify(refs(before)),
        expect: refs(before),
        found: refs(after),
      },
    ],
    step?.message ?? "",
  );
}

// j. After "Leave this revision" in the supplementary file. view: its pane
// after the step; siBefore: its review file before the step.
export function partsLeft({ view, siBefore, si, paper, names, comments }) {
  const want = `Join the revision of ${names.paper}, ${comments} comment${comments === 1 ? "" : "s"}`;
  const all = (rev) =>
    (rev?.comments ?? []).flatMap((c) => c.changes.map((ch) => `${c.ref}/${ch.id}`));
  const lost = all(siBefore).filter((id) => !all(si).includes(id));
  return partsVerdict(
    [
      {
        name: `${names.si} pane: the join box is shown again, with the manuscript`,
        ok: view?.link_box?.shown === true && (view?.link_box?.candidates ?? []).includes(want),
        expect: want,
        found:
          view?.error ??
          `shown ${view?.link_box?.shown}, ${show(view?.link_box?.candidates ?? [])}`,
      },
      {
        name: `review file of ${names.si}: the link is gone`,
        ok: !!si && !si.part_of,
        expect: "no part_of",
        found: si?.part_of ?? (si ? null : "no review file"),
      },
      {
        name: `review file of ${names.paper}: the link is gone`,
        ok: !!paper && (paper.parts ?? []).length === 0,
        expect: "no parts",
        found: paper?.parts ?? (paper ? null : "no review file"),
      },
      {
        name: `review file of ${names.si}: its own changes are kept`,
        ok: all(siBefore).length > 0 && lost.length === 0,
        expect: all(siBefore),
        found: all(si),
      },
    ],
    `${names.si} left; it keeps ${all(si).length} change(s)`,
  );
}

// i (not judged). One line of what Word said about unsaved changes and what
// the file on disk was, at each point the run read them.
// points: [{ at, saved: the doc-saved step, file: { mtime, bytes } | null }]
export function partsSavedLine(points) {
  return points
    .map((p) => {
      const s = p.saved ?? {};
      const saved =
        s.ok === false
          ? `not read (${s.error})`
          : `Document.saved=${show(s.raw)} (${s.raw_type}${s.error ? `, ${s.error}` : ""}), sent as ${show(s.saved)}`;
      const file = p.file ? `mtime ${p.file.mtime}, ${p.file.bytes} bytes` : "file not read";
      return `${p.at}: ${saved}; ${file}`;
    })
    .join(" | ");
}

// Which Agent SDK session each pane's turns ran on, from the line the daemon
// prints for every turn (daemon/sessions.mjs turnSessionLine, no debug flag):
//   [agent] turn doc=<path> session=<id> asked=<id or -> (started|resumed|continued)
// The line names the document by its path; the documents of the run are told
// apart by their file names here (the isolation check, foreignLogLines, is
// what reads the folder). The two panes must not share a session: no id a turn of one document ran
// on, or was asked to resume, may be an id of the other document (asked is
// read too, should the SDK give a resumed session a new id). A file of a
// revision with parts has a resume hint of its own (sessions.mjs), so with
// both files in one folder each pane starts or resumes its own.
// names: { paper, si }, the file names of the run.
export function partsSessions(log, names) {
  const rows = [
    ...String(log ?? "").matchAll(
      /^\[agent\] turn doc=(.*) session=(\S+) asked=(\S+) \((started|resumed|continued)\)$/gm,
    ),
  ].map((m) => ({
    doc: m[1].split(/[\\/]/).pop(),
    session: m[2],
    asked: m[3] === "-" ? null : m[3],
    how: m[4],
  }));
  const idsOf = (doc) =>
    new Set(rows.filter((r) => r.doc === doc).flatMap((r) => [r.session, r.asked].filter(Boolean)));
  const ran = (doc) => [...new Set(rows.filter((r) => r.doc === doc).map((r) => r.session))];
  const paper = idsOf(names.paper);
  const si = idsOf(names.si);
  const shared = [...paper].filter((id) => si.has(id));
  const short = (id) => String(id).slice(0, 8);
  const told = (doc) =>
    `${doc}: ${
      rows
        .filter((r) => r.doc === doc && r.how !== "continued")
        .map(
          (r) =>
            `${r.how} ${short(r.session)}${r.asked && r.asked !== r.session ? ` (asked ${short(r.asked)})` : ""}`,
        )
        .join(", ") || "no turn"
    }`;
  const line = rows.length
    ? `${told(names.paper)}; ${told(names.si)}`
    : "no turn line in the log (a daemon from before the line?)";
  return {
    rows,
    line,
    ...partsVerdict(
      [
        {
          name: `${names.paper}: a turn ran on a session`,
          ok: ran(names.paper).length > 0,
          expect: "at least one [agent] turn line",
          found: line,
        },
        {
          name: `${names.si}: a turn ran on a session`,
          ok: ran(names.si).length > 0,
          expect: "at least one [agent] turn line",
          found: line,
        },
        {
          name: "the two panes share no session",
          ok: shared.length === 0,
          expect: "no session id in common, run on or asked for",
          found: `${shared.map(short).join(", ")} in both: ${line}`,
        },
      ],
      line,
    ),
  };
}

// ---- k to m: the letter of the manuscript and the supplementary file --------
// (README "Supplementary Word documents", step A3)

// The ops of word-com.ps1 apply for the changes of one comment, all with one
// decision ("accept" or "reject"). revision: the review file of the document
// the plan is applied to.
export function partsPlan(revision, ref, op) {
  return buildComPlan(
    { comments: (revision?.comments ?? []).filter((c) => c.ref === ref) },
    { [ref]: op },
  );
}
const planApplied = (applied, plan) =>
  applied?.ok === true &&
  (plan?.ops ?? []).length > 0 &&
  (applied.results ?? []).length === plan.ops.length &&
  applied.results.every((r) => !r.error);
const appliedLine = (applied, plan) =>
  applied?.ok === undefined && !applied?.error
    ? "not tried"
    : (applied?.error ??
      `${(plan?.ops ?? []).length} op(s): ${
        (applied?.results ?? [])
          .map((r) => `${r.change_id} ${r.op}${r.error ? ` ERR ${r.error}` : ""}`)
          .join("; ") || "no result"
      }`);
const statesOf = (rev, ref) =>
  ((rev?.comments ?? []).find((c) => c.ref === ref)?.changes ?? []).map((ch) => ch.state);
const sameFile = (a, b) => !!a && !!b && a.mtime === b.mtime && a.bytes === b.bytes;
const fileLineOf = (f) => (f ? `${f.mtime}, ${f.bytes} bytes` : "not read");
const failRows = (view, re) => (view?.checks ?? []).filter((row) => re.test(row));
const cardOf = (view, ref) => (view?.cards ?? []).find((c) => c.ref === ref);
// The row of the check that says the record of a supplementary file cannot
// be confirmed (response-check.mjs partStateText: three reasons, one advice).
const partStateRow = (name) =>
  new RegExp(`^FAIL part_state: .*${escapeRe(name)}.* Open it, save, press Refresh\\.$`);

// k. The change of the supplementary file was accepted in Word, the file
// saved, refreshed and closed; then, in the manuscript, the answer was
// drafted, checked and the letter exported without force.
//   plan, applied: the ops and what word-com.ps1 apply answered
//   named: the link-label step; label: the name it gave the file
//   si, paper: the two review files after the steps; fileAfter: the file of
//   the supplement on disk after it was closed
//   check: the Response tab after "Run checks"; exported: after the export
//   (null: not pressed, a check failed); model: the letter model on disk
//   log: the daemon log of the row
export function partsCited(r) {
  const { ref, names, label } = r;
  const states = statesOf(r.si, ref);
  const state = r.si?.doc_state ?? null;
  const own = ((r.si?.comments ?? []).find((c) => c.ref === ref)?.response?.citations ?? []).filter(
    (c) => c.kind !== "removed",
  );
  const cited = (
    (r.paper?.comments ?? []).find((c) => c.ref === ref)?.response?.citations ?? []
  ).filter((c) => c.doc === names.si);
  const card = cardOf(r.check, ref);
  const failing = failRows(r.check, new RegExp(`^FAIL (${escapeRe(ref)} |part_state: )`));
  const rows = r.model?.paragraphs ?? [];
  const at = rows.findIndex(
    (p) => p.kind === "location" && p.ref === ref && p.text.startsWith(`${label} [Revised in `),
  );
  const quote = at >= 0 && rows[at + 1]?.kind === "quote" ? rows[at + 1].text : null;
  const link = (r.paper?.parts ?? []).find((p) => String(p.doc ?? "").endsWith(names.si));
  return partsVerdict(
    [
      {
        name: `${names.si} pane: the file is given its name in the letter`,
        ok: r.named?.ok !== false && link?.label === label,
        expect: `"${label}" in the link of ${names.paper}`,
        found:
          r.named?.error ??
          `message "${r.named?.message ?? ""}", link label ${show(link?.label ?? null)}`,
      },
      {
        name: `${names.si} in Word: the change(s) of ${ref} accepted`,
        ok: planApplied(r.applied, r.plan),
        expect: "every op applied",
        found: appliedLine(r.applied, r.plan),
      },
      {
        name: `${names.si}: saved through Word, then closed`,
        ok: r.saved?.ok === true && r.closed?.ok === true,
        expect: "save ok, close ok",
        found: `save ${r.saved?.ok === true ? "ok" : (r.saved?.error ?? "not ok")}, close ${r.closed?.ok === true ? "ok" : (r.closed?.error ?? "not ok")}`,
      },
      {
        name: `review file of ${names.si}: every change of ${ref} is accepted`,
        ok: states.length > 0 && states.every((s) => s === "accepted"),
        expect: "accepted",
        found: r.refreshed?.error ?? states,
      },
      {
        name: `review file of ${names.si}: the refresh was taken on a saved document, and the file on disk is still that file`,
        ok: state?.saved === true && sameFile(state.file, r.fileAfter),
        expect: `saved true, file ${fileLineOf(r.fileAfter)}`,
        found: state
          ? `saved ${show(state.saved)}, file ${fileLineOf(state.file)} (refreshed ${state.at})`
          : "no doc_state",
      },
      {
        name: `review file of ${names.si}: what the letter may cite of ${ref} is written there, without line numbers`,
        ok: own.length > 0 && own.every((c) => c.quote && !c.lines),
        expect: "at least one citation with a quote and no lines",
        found: `${own.length} citation(s)${own.some((c) => c.lines) ? ", with lines" : ""} | refresh: ${r.refreshed?.error ?? r.refreshed?.message ?? ""}`,
      },
      {
        name: `${names.paper} pane: an answer for ${ref} was drafted`,
        ok: r.drafted?.ok !== false && String(card?.text ?? "").trim() !== "",
        expect: "a text in the answer box",
        found: r.drafted?.error ?? `${String(card?.text ?? "").length} characters`,
      },
      {
        name: `${names.paper} pane: no failing check for ${ref}, none for the state of ${names.si}`,
        ok: r.check?.ok !== false && /^Checked: /.test(r.check?.message ?? "") && !failing.length,
        expect: "Checked: ..., no such FAIL row",
        found: r.check?.error ?? (failing.length ? failing : (r.check?.message ?? "")),
      },
      {
        name: `${names.paper} pane: the card of ${ref} cites ${label}`,
        ok: (card?.citations ?? []).some((c) => c.line.startsWith(`${label} [`)),
        expect: `a location line that starts with "${label} ["`,
        found: (card?.citations ?? []).map((c) => c.line),
      },
      {
        name: `${names.paper} pane: the letter is exported without force`,
        ok: /^Exported/.test(r.exported?.message ?? ""),
        expect: "Exported: ...",
        found: r.exported
          ? (r.exported.error ?? r.exported.message)
          : `not pressed, the button would force the export: ${show(failRows(r.check, /^FAIL /))}`,
      },
      {
        name: `letter: a location line of ${ref} names ${label}, with a quote of ${names.si} under it`,
        ok: quote !== null && cited.some((c) => c.quote === quote),
        expect: `"${label} [Revised in ...]", then a quote the review file of ${names.paper} records for ${names.si}`,
        found:
          at < 0
            ? rows.filter((p) => p.kind === "location" && p.ref === ref).map((p) => p.text)
            : `${rows[at].text} | ${show(quote)} | recorded: ${show(cited.map((c) => c.quote))}`,
      },
      {
        name: "the verifier passed on the exported letter",
        ok: /^\[response\] export path=.* verify=ok$/m.test(String(r.log ?? "")),
        expect: "[response] export path=... verify=ok",
        found:
          (String(r.log ?? "").match(/^\[response\] export .*$/gm) ?? []).join(" ; ") ||
          "no export line in the daemon log",
      },
    ],
    at >= 0 ? `${rows[at].text} | ${show(quote)}` : "",
  );
}

// l. The supplementary file was changed in Word and saved, and its pane did
// not refresh: the manuscript's check fails with part_state and the letter
// is not exported without force.
//   restored: the file as it was before the accept was put back (Word cannot
//   reject a change it has accepted); recorded: doc_state of the review file
//   after row k; recordedNow: the same after this row's save and close;
//   fileNow: the file on disk; opened: the Response tab before the export was
//   pressed; exported: after it; letters, export: [before, after] of the
//   letter files of the run and of response_export in the manuscript's file
export function partsGuard(r) {
  const { ref, names } = r;
  const rows = failRows(r.exported, partStateRow(names.si));
  const refreshedSince = r.recorded?.at !== r.recordedNow?.at;
  const same = (pair) => JSON.stringify(pair?.[0] ?? null) === JSON.stringify(pair?.[1] ?? null);
  return partsVerdict(
    [
      {
        name: `${names.si}: reopened as it was before the accept, the change(s) of ${ref} rejected in Word, saved, closed`,
        ok:
          r.restored?.ok === true &&
          r.reopened?.ok === true &&
          planApplied(r.applied, r.plan) &&
          r.saved?.ok === true &&
          r.closed?.ok === true,
        expect: "every step ok",
        found: `put back: ${r.restored?.ok === true ? "ok" : (r.restored?.error ?? "not ok")}; open: ${r.reopened?.ok === true ? "ok" : (r.reopened?.error ?? "not ok")}; reject: ${appliedLine(r.applied, r.plan)}; save: ${r.saved?.ok === true ? "ok" : (r.saved?.error ?? "not ok")}; close: ${r.closed?.ok === true ? "ok" : (r.closed?.error ?? "not ok")}`,
      },
      {
        name: `${names.si} on disk is no longer the file its review file records`,
        ok: !!r.fileNow && !!r.recordedNow?.file && !sameFile(r.recordedNow.file, r.fileNow),
        expect: `another time or size than ${fileLineOf(r.fileNow)}`,
        found: `${fileLineOf(r.recordedNow?.file)}${
          refreshedSince
            ? ` (the pane of ${names.si} refreshed by itself: doc_state.at ${r.recorded?.at} -> ${r.recordedNow?.at})`
            : ""
        }`,
      },
      {
        name: `${names.paper} pane: the export button did not offer to force`,
        ok: r.opened?.ok !== false && r.opened?.export_label === "Export letter (.docx)",
        expect: "Export letter (.docx)",
        found: r.opened?.error ?? r.opened?.export_label ?? "",
      },
      {
        name: `${names.paper} pane: the check fails with part_state for ${names.si}`,
        ok: rows.length === 1,
        expect: `FAIL part_state: ${names.si} changed since Sidewrite last read it. Open it, save, press Refresh.`,
        found: r.exported?.error ?? failRows(r.exported, /^FAIL /),
      },
      {
        name: `${names.paper} pane: the letter is not exported`,
        ok: /^Not exported: /.test(r.exported?.message ?? "") && same(r.letters) && same(r.export),
        expect:
          "Not exported: ..., the letter files and the record of the last export as they were",
        found: `${r.exported?.error ?? r.exported?.message ?? ""} | letter files ${same(r.letters) ? "unchanged" : show(r.letters?.[1])} | record ${same(r.export) ? "unchanged" : "changed"}`,
      },
      {
        name: "daemon log: the export was held",
        ok: /^\[response\] export held fail=[1-9]/m.test(String(r.log ?? "")),
        expect: "[response] export held fail=n",
        found:
          (String(r.log ?? "").match(/^\[response\] export .*$/gm) ?? []).join(" ; ") ||
          "no export line in the daemon log",
      },
    ],
    `${rows[0] ?? ""}${refreshedSince ? ` | the pane of ${names.si} refreshed by itself before the save (doc_state.at ${r.recorded?.at} -> ${r.recordedNow?.at})` : ""}`,
  );
}

// m. The supplementary file was opened again and refreshed: its record is
// confirmed, the change is rejected, so the request is open again and the
// answer that cited the change is stale.
//   reopened: word-com.ps1 open-edit; refreshed: the refresh in its pane;
//   si: its review file after it; fileNow: the file on disk;
//   check: the Response tab of the manuscript after "Run checks"
export function partsReopened(r) {
  const { ref, names } = r;
  const states = statesOf(r.si, ref);
  const state = r.si?.doc_state ?? null;
  const guard = failRows(r.check, partStateRow(names.si));
  const open = failRows(r.check, new RegExp(`^FAIL ${escapeRe(ref)} requirements: `));
  const card = cardOf(r.check, ref);
  return partsVerdict(
    [
      {
        name: `${names.si}: opened again and refreshed`,
        ok: r.reopened?.ok === true && r.refreshed?.ok !== false,
        expect: "open ok, refresh ok",
        found: `open: ${r.reopened?.ok === true ? "ok" : (r.reopened?.error ?? "not ok")}; refresh: ${r.refreshed?.error ?? r.refreshed?.message ?? ""}`,
      },
      {
        name: `review file of ${names.si}: every change of ${ref} is rejected`,
        ok: states.length > 0 && states.every((s) => s === "rejected"),
        expect: "rejected",
        found: states,
      },
      {
        name: `review file of ${names.si}: the refresh was taken on a saved document, and the file on disk is that file`,
        ok: state?.saved === true && sameFile(state.file, r.fileNow),
        expect: `saved true, file ${fileLineOf(r.fileNow)}`,
        found: state
          ? `saved ${show(state.saved)}, file ${fileLineOf(state.file)}`
          : "no doc_state",
      },
      {
        name: `${names.paper} pane: no failing check for the state of ${names.si}`,
        ok: r.check?.ok !== false && /^Checked: /.test(r.check?.message ?? "") && !guard.length,
        expect: "Checked: ..., no FAIL part_state row",
        found: r.check?.error ?? (guard.length ? guard : (r.check?.message ?? "")),
      },
      {
        name: `${names.paper} pane: the request of ${ref} is open again`,
        ok: open.length > 0 && (card?.open_reqs ?? []).length > 0,
        expect: `FAIL ${ref} requirements: ..., and the request shown as open on its card`,
        found: `${show(open)} | open on the card: ${show(card?.open_reqs ?? [])}`,
      },
      {
        name: `${names.paper} pane: the answer of ${ref} is marked stale`,
        ok: String(card?.stale ?? "") !== "",
        expect: "a stale line on the card (A cited change is no longer accepted)",
        found: card ? card.stale : "no card",
      },
    ],
    `${open[0] ?? ""} | ${card?.stale ?? ""}`,
  );
}

// A line of the daemon log that names a document: is every document it
// names one of this run's own files by its bare name ("SI.docx", as the link
// lines of the daemon write them)? names: the file names of the run.
export function onlyOwnNames(line, names) {
  const tokens = String(line ?? "").match(/[^\s"'()=]+\.docx\b/gi) ?? [];
  return tokens.length > 0 && tokens.every((t) => names.includes(t));
}

// The lines of the daemon log that name a document outside the run. own:
// the run folder as the log may write it (Windows and WSL path); names: the
// bare file names of the run (stage parts only, [] elsewhere). The hello of
// another Word's pane is not a tool call and is left out. A turn line
// ("[agent] turn doc=<path> ...") names its document by path, so one of
// another folder is foreign whatever its file name is.
// A line with its JSON-quoted values decoded ("D:\\run\\a.docx" becomes
// D:\run\a.docx), as figure-http.mjs writes the page loads. A quoted text
// that is not JSON stays as it is.
const unquoted = (l) =>
  l.replace(/"(?:[^"\\]|\\.)*"/g, (m) => {
    try {
      return String(JSON.parse(m));
    } catch {
      return m;
    }
  });
export function foreignLogLines(lines, own, names = []) {
  return (lines ?? [])
    .filter(
      (l) => /\.docx\b/i.test(l) && !own.some((p) => l.includes(p) || unquoted(l).includes(p)),
    )
    .filter((l) => !onlyOwnNames(l, names))
    .filter((l) => !/^\[bridge\] hello received|^\[diag\] hello /.test(l));
}

export const JUDGES = {
  "open-review": (r, ctx) => ({
    pass: r.file === `Review file: ${ctx.revisionPath}`,
    expect: `Review file: ${ctx.revisionPath}`,
    got: r.file,
  }),
  split: (r) => ({
    pass:
      /^Split into /.test(r.message) && r.draft_refs.filter((x) => /^R1\.\d$/.test(x)).length === 5,
    expect:
      "Split into n comments (found by the numbering in the text), R1.1-R1.5 after the non-reviewer rows go",
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
  // The cite stage: the four comments of tests/fixtures/cite_review.txt.
  "cite-split": (r) => ({
    pass:
      /^Split into /.test(r.message) &&
      CITE_REFS.every((ref) => r.draft_refs.includes(ref)) &&
      r.draft_refs.filter((x) => /^R\d+\.\d+$/.test(x)).length === CITE_REFS.length,
    expect: `Split into n comments, ${CITE_REFS.join(", ")} among them`,
    got: `${r.message} refs=${r.draft_refs.join(",")}`,
  }),
  "cite-draft": (r) => {
    const refs = r.comments.map((c) => c.ref).join(",");
    return {
      pass: refs === CITE_REFS.join(","),
      expect: `cards ${CITE_REFS.join(",")}`,
      got: refs,
    };
  },
  // The parts stage: the four comments of tests/fixtures/parts_review.txt,
  // the requests the run typed, and the one rule of parts_ledger.json.
  "parts-split": (r) => ({
    pass:
      /^Split into /.test(r.message) &&
      r.draft_refs.filter((x) => /^R\d+\.\d+$/.test(x)).join(",") === "R1.1,R1.2,R1.3,R1.4",
    expect: "Split into n comments, R1.1 to R1.4 among them",
    got: `${r.message} refs=${r.draft_refs.join(",")}`,
  }),
  "parts-draft": (r, ctx = {}) => {
    const refs = r.comments.map((c) => c.ref).join(",");
    const want = Object.entries(ctx.requirements ?? {}).map(([ref, lines]) => [ref, lines.length]);
    const got = r.comments.map((c) => [c.ref, (c.requests ?? []).length]);
    return {
      pass: refs === "R1.1,R1.2,R1.3,R1.4" && JSON.stringify(got) === JSON.stringify(want),
      expect: `cards R1.1 to R1.4 with ${JSON.stringify(want)} requests`,
      got: `${refs} | requests ${JSON.stringify(got)}`,
    };
  },
  "parts-ledger": (r) => ({
    pass: r.ledger.length === 1 && has(r.ledger[0], "L1 forbid"),
    expect: "L1 forbid",
    got: r.ledger.join(" | "),
  }),
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
    // A request no change answers waits as a "request_open" item of its comment.
    const want = [
      { ref: "R1.3", req: "q1" },
      { ref: "R1.3", req: "q2" },
    ];
    // The list holds an open request of a comment nobody has worked on yet, hidden:
    // a request waits for the author only once a turn ran for its comment (item 13).
    // At this step no turn has run, so both rows are in the list and not shown.
    const waiting = (r.decisions_raw ?? []).filter((d) => d.kind === "request_open");
    return {
      pass:
        n("L1") === 2 &&
        n("L2") === 1 &&
        n("L3") === 0 &&
        want.every((w) => waiting.some((d) => d.ref === w.ref && d.req === w.req)),
      expect: "violations L1 2, L2 1, L3 0; waiting: R1.3 q1, R1.3 q2",
      got: `L1 ${n("L1")}, L2 ${n("L2")}, L3 ${n("L3")}; waiting rows ${waiting.map((d) => `${d.ref} ${d.req}${d.shown === false ? " (hidden)" : ""}`).join(", ")}`,
    };
  },
  preview: (r) => ({
    pass: ["[L1]", "[L2]", "[L3]"].every((t) => has(r.prompt, t)),
    expect: "prompt shows [L1], [L2], [L3]",
    got: ["[L1]", "[L2]", "[L3]"].filter((t) => has(r.prompt, t)).join(" "),
  }),
  propose: (r, ctx) => {
    const re = new RegExp(
      `^${escapeRe(ctx.ref)} running \\d+:\\d\\d\\. \\d+ of \\d+ comments not done\\.$`,
    );
    const ticks = r.progress_samples.filter((s) => re.test(s));
    return {
      pass: !r.message_error && ticks.length >= 2 && ctx.turnCompleted === true,
      expect: `${ctx.ref} running m:ss. n of 5 comments not done. (changing), turn complete in the log, no error`,
      got: `${ticks.slice(0, 2).join(" -> ")} | turn_complete=${ctx.turnCompleted} | ${r.message}`,
    };
  },
  refresh: (r) => {
    // The numbers are under "Details"; the sentence says what to do next.
    const m = /Changes (\d+), expected tool comments (\d+), found (\d+)/.exec(r.message_details);
    return {
      // The parts of a rewritten paragraph share one comment: every comment
      // that is expected is there, and there are not more comments than changes.
      pass:
        !!m && Number(m[1]) >= Number(m[2]) && m[2] === m[3] && !has(r.message_details, "MISMATCH"),
      expect: "Details: Changes N, expected tool comments M, found M, M <= N. (no MISMATCH)",
      got: `${r.message} | ${r.message_details}`,
    };
  },
  // A document in compatibility mode 14: Word refuses Comment.resolved. The
  // refresh still shows its counts, and one note for the handled changes.
  "refresh-mode14": (r) => {
    const m =
      /^Accepted (\d+), rejected (\d+), .*Changes (\d+), expected tool comments (\d+), found (\d+)/.exec(
        r.message_details,
      );
    const n = /(\d+) rationale comment\(s\) not marked resolved \(Word: /.exec(r.message_details);
    const handled = m ? Number(m[1]) + Number(m[2]) : 0;
    return {
      pass:
        !!m &&
        !!n &&
        handled > 0 &&
        Number(n[1]) === handled &&
        m[3] === m[4] &&
        m[4] === m[5] &&
        !has(r.message_details, "MISMATCH"),
      expect:
        "Details: Accepted a, rejected r ... Changes N, expected tool comments N, found N. (a+r) rationale comment(s) not marked resolved (Word: ...)",
      got: `${r.message} | ${r.message_details}`,
    };
  },
  // The badge part is what the procedure asks a person to confirm
  // (s2-list-matches); the outside-change part is L4.
  "refresh-after-user": (r) => {
    const b = badgesMatch(r);
    const at = r.external.findIndex((x) => has(x, "The old text had 0.82; 0.91"));
    return {
      pass: b.pass && has(r.message_details, "Outside changes:") && at >= 0,
      expect:
        "R1.1 Accepted, R1.2 Rejected, R1.3 first Edited by you; Details: Outside changes: n; an outside item whose old text had 0.82; 0.91",
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

// ---- Real Word: "Check now", the figures button, the resident daemon ----
// Each returns judged rows { id, result: "PASS" | "FAIL" | "NOTE", expect,
// got } from what the run read itself: the rows the pane drew, Word over
// COM, revision.json, the hash of a file on the disk, the daemon log.

const rowOf = (id, pass, expect, got) => ({
  id,
  result: pass ? "PASS" : "FAIL",
  expect,
  got: String(got ?? ""),
});
const nameOfPath = (p) =>
  String(p ?? "")
    .split(/[\\/]/)
    .pop();

// The line aux-service.mjs writes for every submission_check job.
export function submissionLogLine(log) {
  const m = [
    ...String(log ?? "").matchAll(
      /^\[tools\] submission_check rows=(\d+) ready=(\d+) not_ready=(\d+) not_applicable=(\d+) saved=(\S+)$/gm,
    ),
  ].at(-1);
  if (!m) return null;
  return {
    rows: Number(m[1]),
    ready: Number(m[2]),
    not_ready: Number(m[3]),
    not_applicable: Number(m[4]),
    saved: m[5] === "true" ? true : m[5] === "false" ? false : null,
  };
}

// Did anything the letter model is built from change between two reads of
// revision.json, as far as the run can tell: the answers with their
// citations, a declined comment, the overrides of the profile.
export function letterInputsChanged(before, after) {
  const of = (rev) =>
    JSON.stringify({
      comments: (rev?.comments ?? []).map((c) => ({
        ref: c.ref,
        declined: c.declined ?? null,
        response: c.response ?? null,
      })),
      overrides: rev?.profile_overrides ?? null,
    });
  return of(before) !== of(after);
}

// The rows of one press of "Check now", judged.
//   step      the answer of the pane step submission-check: { ok, message,
//             message_error, rows: [{ id, state, text, file }], asof }
//   word      word-com.ps1 state: { saved, revisions, comments }
//   revision  revision.json read right after the press
//   letterSha sha256 of the file response_export.path names, as it is on the
//             disk now (null: not there)
//   letter    "ready": nothing the letter is built from changed since the
//             last export, as far as the run knows; "changed": it did
//   log       the daemon log written since the press
// Rows the run cannot tell by itself (answers, requests, format_counts, the
// rows of another file) are NOTE.
export function submissionEvidence({ step, word, revision, letterSha, letter = "ready", log }) {
  const rows = [];
  // The rows of the open file come before those of a supplementary file
  // with the same id (submission-check.mjs): the first of an id is judged.
  const shown = step?.rows ?? [];
  const byId = new Map();
  for (const r of shown) if (!byId.has(r.id)) byId.set(r.id, r);
  const describe = (r) => (r ? `${r.state}: ${r.text}` : "no such row");
  const saved = word?.saved;
  const line = submissionLogLine(log);

  rows.push(
    rowOf(
      "job",
      step?.ok !== false && !step?.message_error && shown.length > 0 && !!line,
      "rows drawn, no error, the daemon logged the job",
      step?.ok === false || step?.message_error
        ? `error: ${step?.error ?? step?.message}`
        : `${shown.length} rows; log: ${line ? `rows=${line.rows} saved=${line.saved}` : "no submission_check line"}`,
    ),
  );
  rows.push(
    rowOf(
      "saved",
      typeof saved === "boolean" && line?.saved === saved,
      `the pane sent what Word says over COM (Saved=${saved})`,
      `pane sent saved=${line ? line.saved : "-"}, COM Saved=${saved ?? "not read"}`,
    ),
  );

  const revisions = word?.revisions?.length ?? null;
  const comments = word?.comments ?? null;
  const tool = comments
    ? comments.filter((c) => TOOL_COMMENT_PREFIX.test(c.text ?? "")).length
    : null;
  const expectState = (id, ready, why) => {
    const r = byId.get(id);
    const want = ready ? "ready" : "not_ready";
    rows.push(rowOf(id, !!r && r.state === want, `${want} (${why})`, describe(r)));
  };
  if (revisions === null)
    rows.push(rowOf("changes_decided", false, "Word's revisions", "COM state not read"));
  else expectState("changes_decided", revisions === 0, `Word lists ${revisions} revisions`);
  if (tool === null)
    rows.push(rowOf("tool_comments", false, "Word's comments", "COM state not read"));
  else {
    expectState("tool_comments", tool === 0, `Word has ${tool} comments of the tool`);
    expectState(
      "file_clean",
      revisions === 0 && comments.length === 0,
      `Word lists ${revisions} revisions and ${comments.length} comments`,
    );
  }

  // The copies made from the saved file.
  const copy = (id, what, artifact) => {
    const r = byId.get(id);
    if (id === "anonymized_copy" && r?.state === "not_applicable")
      return rows.push(
        rowOf(
          id,
          /anonymized copy/.test(r.text),
          "not_applicable (the profile asks for none)",
          describe(r),
        ),
      );
    if (saved === false)
      return rows.push(
        rowOf(
          id,
          r?.state === "not_ready" && r.text === UNSAVED_TEXT,
          `not_ready: ${UNSAVED_TEXT}`,
          describe(r),
        ),
      );
    if (saved !== true)
      return rows.push(rowOf(id, false, "Word says whether the document is saved", describe(r)));
    if (!artifact?.path) {
      const head = `No ${what} is on record for this file.`;
      return rows.push(
        rowOf(
          id,
          r?.state === "not_ready" && r.text.startsWith(head) && !r.text.includes(UNSAVED_UNKNOWN),
          `not_ready: ${head} ...`,
          describe(r),
        ),
      );
    }
    rows.push(
      rowOf(
        id,
        !!r && r.text !== UNSAVED_TEXT && !r.text.includes(UNSAVED_UNKNOWN),
        "a row of the saved file (a copy is on record: its hashes are not judged here)",
        describe(r),
      ),
    );
  };
  copy("marked_copy", "marked copy", revision?.artifacts?.marked);
  copy("anonymized_copy", "anonymized copy", revision?.artifacts?.anonymized);

  // The exported letter: ready only when its record says no check failed,
  // the file on the disk is the file written, and nothing changed since.
  {
    const x = revision?.response_export;
    const r = byId.get("letter");
    const name = nameOfPath(x?.path);
    const sound =
      !!x && Number.isInteger(x.fail) && x.fail === 0 && !!x.docx_sha && letterSha === x.docx_sha;
    if (sound && letter === "ready")
      rows.push(
        rowOf(
          "letter",
          r?.state === "ready" && r.text.includes(name),
          `ready, naming ${name}`,
          describe(r),
        ),
      );
    else if (sound)
      rows.push(
        rowOf(
          "letter",
          r?.state === "not_ready" && /changed after .+ was exported/.test(r.text),
          `not_ready: ... changed after ${name} was exported ...`,
          describe(r),
        ),
      );
    else
      rows.push(
        rowOf(
          "letter",
          r?.state === "not_ready",
          `not_ready (record: ${x ? `fail=${x.fail}, file ${letterSha === x.docx_sha ? "as written" : "not as written"}` : "no export"})`,
          describe(r),
        ),
      );
  }

  // Figures: none are on record in this stage.
  {
    const r = byId.get("figures");
    const any = (revision?.comments ?? []).some((c) => (c.figure_changes ?? []).length > 0);
    if (any) rows.push({ id: "figures", result: "NOTE", expect: "-", got: describe(r) });
    else
      rows.push(
        rowOf(
          "figures",
          r?.state === "not_applicable",
          "not_applicable (no figure change on record)",
          describe(r),
        ),
      );
  }

  // The line under the rows.
  rows.push(
    rowOf(
      "as_of",
      saved === false
        ? /changes that are not saved/.test(step?.asof ?? "")
        : saved === true && /Press Check now again after a change\./.test(step?.asof ?? ""),
      saved === false
        ? "says the document has unsaved changes"
        : "Press Check now again after a change.",
      step?.asof ?? "",
    ),
  );

  const judged = new Set(rows.map((r) => r.id));
  for (const r of shown)
    if (!judged.has(r.id) || byId.get(r.id) !== r)
      rows.push({ id: r.id, result: "NOTE", expect: "-", got: describe(r) });
  return { pass: rows.every((r) => r.result !== "FAIL"), rows };
}

// The page loads the daemon logged (figure-http.mjs): { doc, ref, ua }.
export function figuresPageGets(log) {
  const out = [];
  for (const m of String(log ?? "").matchAll(
    /^\[figures\] page GET doc=(null|"(?:[^"\\]|\\.)*") ref=(null|"(?:[^"\\]|\\.)*") ua=("(?:[^"\\]|\\.)*")$/gm,
  )) {
    try {
      out.push({ doc: JSON.parse(m[1]), ref: JSON.parse(m[2]), ua: JSON.parse(m[3]) });
    } catch {
      /* not a line of figure-http.mjs */
    }
  }
  return out;
}

// The "Figures and other files" button, judged.
//   step  the answer of the pane step figures-open: { ok, href, spy:
//         { supported, wrapped, installed_on, calls, errors,
//         default_prevented }, pane_url_before, pane_url_after,
//         figures_pages_before, figures_pages_after }
//   gets  figuresPageGets of the log written since the click
// route, from default_prevented, which is what the pane's own click handler
// did (review-pane.js figuresLink): "api" (openInBrowser returned true: the
// API was there, called, and did not throw), "fallback" (the API is missing
// and the link itself was followed), "fallback-despite-api" (the API is
// there and the link was followed all the same: it threw or was not called)
// or "unknown" (the click was not seen). The wrapper, when it held, adds the
// address the API was called with.
export function figuresOpenEvidence({ step, gets }) {
  const rows = [];
  let doc = null;
  try {
    doc = new URL(step?.href ?? "").searchParams.get("doc");
  } catch {
    doc = null;
  }
  const spy = step?.spy ?? {};
  const calls = spy.calls ?? [];
  const prevented = spy.default_prevented;
  const route =
    prevented === true
      ? "api"
      : prevented === false
        ? spy.supported === false
          ? "fallback"
          : "fallback-despite-api"
        : "unknown";
  rows.push(
    rowOf(
      "clicked",
      step?.ok !== false && !!step?.href && !!doc,
      "the link of the pane, with the document in its address",
      step?.ok === false ? `error: ${step.error}` : step?.href,
    ),
  );
  rows.push(
    rowOf(
      "route",
      route === "fallback" ||
        (route === "api" &&
          // A wrapper Office would not let in ("wrap: ...") is no error of the call.
          !(spy.errors ?? []).some((e) => !String(e).startsWith("wrap:")) &&
          (!spy.wrapped || (calls.length === 1 && calls[0] === step?.href))),
      "the click handler prevented the link (openBrowserWindow called without an error; with the wrapper in place, once with the link), or the link fallback where the API is missing",
      `route=${route} default_prevented=${prevented ?? "-"} supported=${spy.supported ?? "-"} wrapped=${spy.wrapped ? (spy.installed_on ?? "yes") : "no (address not watched)"} calls=${JSON.stringify(calls)}${(spy.errors ?? []).length ? ` errors=${JSON.stringify(spy.errors)}` : ""}`,
    ),
  );
  const hit = gets.filter((g) => g.doc === doc);
  rows.push(
    rowOf(
      "served",
      hit.length > 0,
      `the daemon served /figures/ for ${doc}`,
      gets.length ? gets.map((g) => `doc=${g.doc} ref=${g.ref}`).join("; ") : "no page load logged",
    ),
  );
  const newInPane = (step?.figures_pages_after ?? []).filter(
    (u) => !(step?.figures_pages_before ?? []).includes(u),
  );
  rows.push(
    rowOf(
      "outside the pane",
      !!step?.pane_url_before &&
        step.pane_url_after === step.pane_url_before &&
        newInPane.length === 0,
      "the pane stayed on its page and no figures page opened inside Word's WebView2",
      `pane ${step?.pane_url_after === step?.pane_url_before ? "unchanged" : `moved to ${step?.pane_url_after}`}; figures pages in WebView2: ${newInPane.length ? newInPane.join(", ") : "none"}`,
    ),
  );
  if (hit[0]) rows.push({ id: "user agent", result: "NOTE", expect: "-", got: hit[0].ua });
  return { pass: rows.every((r) => r.result !== "FAIL"), route, rows };
}

// May the run close the browser window the button opened? Only when it is
// surely the run's own: a process that was not there before the click, of a
// browser no process of which ran before it (otherwise the page may have
// opened as a tab among the user's), the only such process with a window, and
// its window shows the figures page. browser-windows.ps1 checks the title
// again and counts the tabs before it closes anything.
//   before, after: [{ pid, name, title, has_window }] (browser-windows.ps1 list)
export function browserCloseDecision({ before, after, title }) {
  if (!Array.isArray(before) || !Array.isArray(after))
    return { action: "leave", reason: "the browser processes could not be listed" };
  const old = new Set(before.map((p) => p.pid));
  const fresh = after.filter((p) => !old.has(p.pid) && p.has_window);
  const ours = fresh.filter((p) => String(p.title ?? "").includes(title));
  if (ours.length === 0)
    return {
      action: "leave",
      reason: `no new browser window shows "${title}" (the page opened as a tab of a browser that was already running, or under another title)`,
    };
  if (ours.length > 1)
    return { action: "leave", reason: `${ours.length} new browser windows show "${title}"` };
  const p = ours[0];
  if (before.some((b) => b.name === p.name))
    return {
      action: "leave",
      reason: `a ${p.name} process was running before the click, so its window may hold tabs of the user`,
    };
  if (fresh.filter((q) => q.name === p.name).length > 1)
    return { action: "leave", reason: `more than one new ${p.name} window` };
  return { action: "close", pid: p.pid, name: p.name };
}

// ---- --resident: the installed daemon, not stopped and not replaced ----

// The file the unit writes the daemon's output to (`systemctl --user cat`),
// or null when it writes to the journal.
export function unitLogFile(unitText) {
  const m = [...String(unitText ?? "").matchAll(/^StandardOutput=(?:append|file):(.+)$/gm)].at(-1);
  return m ? m[1].trim() : null;
}

// The processes that listen on 127.0.0.1:<port>, from `ss -ltnp`.
export function listenerPids(ssText, port) {
  const pids = new Set();
  for (const l of String(ssText ?? "").split("\n"))
    if (l.includes(`127.0.0.1:${port} `))
      for (const m of l.matchAll(/pid=(\d+)/g)) pids.add(Number(m[1]));
  return [...pids];
}

// before, after: { active, main_pid, ws: [pids], http: [pids], uptime_s }
// read before the Word of the run started and after it was closed.
// elapsed_s: seconds between the two reads.
export function residentEvidence({ before, after, elapsed_s }) {
  const serves = (s) =>
    !!s?.active && s.main_pid > 0 && s.ws?.includes(s.main_pid) && s.http?.includes(s.main_pid);
  const say = (s) =>
    s
      ? `active=${s.active} MainPID=${s.main_pid} 47833:${JSON.stringify(s.ws)} 47834:${JSON.stringify(s.http)} uptime=${s.uptime_s ?? "-"}s`
      : "not read";
  const rows = [
    rowOf(
      "serving",
      serves(before),
      "the unit is active and its MainPID holds both ports",
      say(before),
    ),
    rowOf(
      "not restarted",
      !!after &&
        after.main_pid === before?.main_pid &&
        Number.isFinite(after.uptime_s) &&
        Number.isFinite(before?.uptime_s) &&
        after.uptime_s >= before.uptime_s + Math.max(0, elapsed_s - 5),
      `the same MainPID, uptime grew by about ${elapsed_s}s`,
      `before: ${say(before)} | after: ${say(after)}`,
    ),
    rowOf(
      "still serving",
      serves(after),
      "active, both ports held by MainPID after the run",
      say(after),
    ),
  ];
  return { pass: rows.every((r) => r.result === "PASS"), rows };
}
