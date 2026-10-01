// Pure helpers of the acceptance runner: the Word COM plan built from
// revision.json, the expected pane texts of each step (wording of the
// stage-2 acceptance procedure), and the path rewrite that runs a
// feature_list verify_command against this run's files.

import { formatRationaleComment } from "../../daemon/review-tools.mjs";

// The author's handling of the proposed changes in the acceptance procedure:
// R1.1 accepted, R1.2 rejected, R1.3 first change accepted then one inserted
// word edited, everything else accepted.
export const DEFAULT_DECISIONS = {
  "R1.1": "accept",
  "R1.2": "reject",
  "R1.3": ["accept_edit", "accept"],
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
        comment_text: formatRationaleComment(c.ref, ch.rationale, { lost: ch.lost_facts ?? [] }),
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
  const map = { [`/tmp/${prefix}-`]: `${noSlash(runDir)}/ok/${prefix}-` };
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
  // A review turn run by a CLI provider. None starts when the provider is
  // "judgments only" (the CLI will not call the office tools unattended).
  "S25-33": { trigger: /^\[cli-turn\] start provider=codex /m },
  "S25-34": { trigger: /^\[cli-turn\] start provider=codex /m },
  "S25-35": { trigger: /^\[cli-turn\] start provider=antigravity /m },
  "S25-36": { trigger: /^\[cli-turn\] start provider=/m },
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
      const text = formatRationaleComment(c.ref, ch.rationale, { lost: ch.lost_facts ?? [] });
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
  const known = { s2: "S2-18..S2-34", s25: "S25-30..S25-41" };
  if (!known[stage])
    throw new Error(`no default feature range for stage ${stage}: give --features`);
  return known[stage];
}

// The name of the manuscript copy in the run folder. The stage 2.5 feature
// list reads paper.decisions.jsonl and paper.ledger.json next to it.
export const docStem = (stage) => (stage === "s2" ? "review_manuscript" : "paper");

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

// What the log says about the review turns of one CLI provider.
function cliTurnLog(provider, log) {
  const text = String(log ?? "");
  const p = escapeRe(provider);
  return {
    started: new RegExp(`^\\[cli-turn\\] start provider=${p} `, "m").test(text),
    done: new RegExp(`^\\[cli-turn\\] done provider=${p} code=0 `, "m").test(text),
    endpoint_calls: (text.match(/^\[mcp-http\] POST route=/gm) ?? []).length,
  };
}
const selectedFor = (select, provider) => select?.selected?.id === provider && !select?.refused;
const changesOf = (revision, ref) =>
  (revision?.comments ?? []).find((c) => c.ref === ref)?.changes ?? [];

// S25-33 (Codex) and S25-35 (Antigravity): one review turn of the provider on
// `ref`. log: the daemon log lines of that turn; revision: the file after it.
export function cliTurnEvidence({ provider, ref, select, log, revision }) {
  const selected = selectedFor(select, provider);
  const l = cliTurnLog(provider, selected ? log : "");
  const changes = changesOf(revision, ref);
  return {
    pass:
      selected &&
      l.started &&
      l.done &&
      l.endpoint_calls > 0 &&
      changes.length > 0 &&
      changes.every((c) => c.word_comment_id != null),
    provider,
    ref,
    ...l,
    changes: selected ? changes.length : 0,
    ...(selected ? {} : { reason: refusal(select) }),
  };
}

// S25-34: a CLI turn on a comment whose request asks for a forbidden word.
// document_changes: recorded changes of that comment whose new text holds the
// word (must be 0: the edit is blocked, as in a Claude turn). The tool filter
// of a CLI turn (no writer but office_tracked_edit) is tested in
// tests/mcp-http.test.mjs and tests/review-turn-parity.test.mjs.
export function guardEvidence({ provider, ref, forbidden, select, log, revision }) {
  const selected = selectedFor(select, provider);
  const l = cliTurnLog(provider, selected ? log : "");
  const re = new RegExp(forbidden, "i");
  const leaked = changesOf(revision, ref).filter((c) => re.test(String(c.text ?? "")));
  const document_changes = selected && l.done ? leaked.length : null;
  return {
    pass: document_changes === 0,
    document_changes,
    blocked: new RegExp(`^\\[review\\] blocked ${escapeRe(ref)} by=`, "m").test(String(log ?? "")),
    provider,
    ref,
    ...l,
    ...(selected ? {} : { reason: refusal(select) }),
  };
}

// S25-36. step: the `propose-stop` pane step; processes: CLI processes of
// the provider still running after it.
export function stopEvidence({ select, step, log, processes }) {
  const provider = select?.selected?.id;
  const selected = !!provider && provider !== "claude" && !select?.refused;
  const started = selected && cliTurnLog(provider, log).started;
  return {
    pass:
      selected &&
      started &&
      step?.stopped === true &&
      step?.agent_status === "Ready" &&
      processes === 0,
    stopped: step?.stopped === true,
    agent_status: step?.agent_status ?? null,
    started,
    processes,
    ...(selected ? {} : { reason: refusal(select) }),
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
  };
}

// Judgments of one pane step from the runner's record of it. Each returns
// { pass, expect, got } where expect quotes the acceptance procedure.
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
      pass: !!m && m[1] === m[2] && m[2] === m[3] && !has(r.message, "MISMATCH"),
      expect: "Changes N, expected tool comments N, found N. (no MISMATCH)",
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
  "focus-refresh": (r) => ({
    pass: r.changed === true && /^Auto-refreshed /.test(r.after ?? ""),
    expect: "Auto-refreshed <time>: ... (no Refresh click)",
    got: r.after || "(the auto-refresh line did not change)",
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
