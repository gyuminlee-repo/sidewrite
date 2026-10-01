#!/usr/bin/env node
// Stage-2 review acceptance run without hands: daemon, a separate Word
// process with a CDP port on its add-in pane, the Review tab steps, the
// author's Word steps over COM, the feature_list verify commands, cleanup.
// Runs from WSL; Word, its pane and COM are on Windows. See README.md.

import { spawn, execFileSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { parseArgs } from "node:util";
import {
  badgesMatch,
  buildComPlan,
  cleanupDecision,
  exitDecision,
  judge,
  minimalFind,
  norm,
  resolvedMatch,
  siblingMatch,
  rewriteCommand,
  table,
} from "./judge.mjs";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "../..");
const WS_PORT = 47833;
const HTTP_PORT = 47834;
const OUTSIDE_SENTENCE =
  "Across the replicates the plateau optical density ranged from 0.82 to 0.91.";
// The author's own tracked insertion next to the R1.4 change (siblingMatch).
const HAND_INSERT = "Checked: ";
const R13_REQUIREMENTS = [
  "say how growth was followed",
  "describe the synthetic carbon source as a novel substrate",
];

const { values: opt } = parseArgs({
  options: {
    "base-dir": { type: "string", default: process.env.E2E_BASE_DIR },
    "run-dir": { type: "string" },
    fixture: { type: "string", default: path.join(REPO, "tests/fixtures/review_manuscript.docx") },
    review: { type: "string", default: path.join(REPO, "tests/fixtures/fake_review.txt") },
    ledger: { type: "string", default: path.join(REPO, "tests/fixtures/review_ledger.json") },
    manifest: { type: "string", default: path.join(REPO, "manifests/word.xml") },
    log: { type: "string", default: "/tmp/daemon-s2-e2e.log" },
    "cdp-port": { type: "string", default: process.env.E2E_CDP_PORT ?? "9333" },
    "win-node": {
      type: "string",
      default: process.env.E2E_WIN_NODE ?? "C:\\Program Files\\nodejs\\node.exe",
    },
    "playwright-core": { type: "string", default: process.env.E2E_PLAYWRIGHT_CORE },
    "feature-list": { type: "string", default: process.env.E2E_FEATURE_LIST },
    "plan-dir": { type: "string", default: process.env.E2E_PLAN_DIR },
    "plan-repo": { type: "string", default: process.env.E2E_PLAN_REPO },
    "plan-log": { type: "string", default: "/tmp/daemon-s2.log" },
    features: { type: "string", default: "S2-18..S2-34" },
    // Feature ids that depend on an event the run cannot force (judge.mjs
    // CONDITIONS). Repeat the flag or separate ids with commas.
    conditional: { type: "string", multiple: true, default: [] },
    "skip-knob": { type: "boolean", default: false },
    // Leave out the second Word run on a compatibility mode 14 copy.
    "skip-mode14": { type: "boolean", default: false },
    // Leave Word and the daemon running when a step fails, to look at them.
    "keep-on-fail": { type: "boolean", default: false },
  },
});

const die = (m) => {
  console.error(`run-acceptance: ${m}`);
  process.exit(2);
};
if (!opt["base-dir"] && !opt["run-dir"])
  die("--base-dir (a folder on a Windows drive, WSL path) is required");
if (!opt["playwright-core"])
  die("--playwright-core (Windows path of a playwright-core package) is required");

const stamp = new Date().toISOString().replace(/[-:]/g, "").replace("T", "-").slice(0, 15);
const RUN = path.resolve(opt["run-dir"] ?? path.join(opt["base-dir"], `e2e-${stamp}`));
const TOOLS = path.join(RUN, "tools");
const OK_DIR = path.join(RUN, "ok");
const DOC = path.join(RUN, "review_manuscript.docx");
const REVISION = path.join(RUN, "review_manuscript.revision.json");
const LOG = opt.log;
const KNOB_LOG = LOG.replace(/\.log$/, "") + "-knob.log";
const MODE14_LOG = LOG.replace(/\.log$/, "") + "-mode14.log";
const MODE14_DIR = path.join(RUN, "mode14");
const DOC14 = path.join(MODE14_DIR, "review_manuscript.docx");
const winPath = (p) => execFileSync("wslpath", ["-w", p], { encoding: "utf8" }).trim();
const unixPath = (p) => execFileSync("wslpath", ["-u", p], { encoding: "utf8" }).trim();

const record = {
  run_dir: RUN,
  started_at: new Date().toISOString(),
  steps: [],
  features: [],
  word: {},
  cleanup: {},
};
const rows = [];
const save = () => fs.writeFileSync(path.join(RUN, "result.json"), JSON.stringify(record, null, 2));

function note(step, result, extra = {}) {
  const j = judge(extra.judge ?? step, result, extra.ctx ?? {});
  const row = {
    step: extra.label ?? step,
    result: j.pass ? "PASS" : "FAIL",
    expect: j.expect,
    got: j.got,
  };
  rows.push(row);
  record.steps.push({ ...row, raw: result });
  save();
  console.log(`${row.result}  ${row.step}  ${String(row.got).slice(0, 200)}`);
  return result;
}

// ---- Windows calls -------------------------------------------------------
function powershell(script, args, timeoutMs = 300_000) {
  const out = execFileSync(
    "powershell.exe",
    [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      winPath(path.join(TOOLS, script)),
      ...args,
    ],
    { cwd: RUN, encoding: "utf8", timeout: timeoutMs, maxBuffer: 64 << 20 },
  );
  const line = out.replace(/\r/g, "").trim().split("\n").filter(Boolean).pop();
  return JSON.parse(line);
}
function com(action, extra = [], doc = DOC) {
  return powershell("word-com.ps1", ["-Doc", winPath(doc), "-Action", action, ...extra]);
}
const psList = (cmd) =>
  execFileSync("powershell.exe", ["-NoProfile", "-Command", cmd], { encoding: "utf8" })
    .replace(/\r/g, "")
    .trim()
    .split(",")
    .filter(Boolean)
    .map(Number);
const wordPids = () =>
  psList("(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id) -join ','");

let paneTag = "";
let pane = (step, args = {}) => {
  const argsFile = path.join(RUN, `args-${paneTag}${step}.json`);
  fs.writeFileSync(argsFile, JSON.stringify(args));
  let out;
  try {
    out = execFileSync(
      unixPath(opt["win-node"]),
      [winPath(path.join(TOOLS, "pane-driver.cjs")), step, winPath(argsFile)],
      {
        cwd: RUN,
        encoding: "utf8",
        timeout: 30 * 60_000,
        maxBuffer: 64 << 20,
        env: {
          ...process.env,
          PLAYWRIGHT_CORE: opt["playwright-core"],
          E2E_CDP_PORT: opt["cdp-port"],
          WSLENV: [process.env.WSLENV, "PLAYWRIGHT_CORE", "E2E_CDP_PORT"].filter(Boolean).join(":"),
        },
      },
    );
  } catch (e) {
    out = e.stdout || JSON.stringify({ ok: false, error: e.message });
  }
  const line = String(out).replace(/\r/g, "").trim().split("\n").filter(Boolean).pop();
  return JSON.parse(line);
};

// ---- daemon ----------------------------------------------------------------
const listening = (port) => {
  try {
    return execFileSync("ss", ["-ltn"], { encoding: "utf8" }).includes(`127.0.0.1:${port} `);
  } catch {
    return false;
  }
};
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
async function until(fn, ms, what) {
  const end = Date.now() + ms;
  while (Date.now() < end) {
    if (await fn()) return;
    await sleep(500);
  }
  throw new Error(`timed out waiting for ${what}`);
}
const readLog = (f) => (fs.existsSync(f) ? fs.readFileSync(f, "utf8") : "");
const count = (f, re) => (readLog(f).match(re) ?? []).length;

let daemon = null;
let daemonLog = LOG;
async function startDaemon(logFile, extraEnv = {}) {
  const fd = fs.openSync(logFile, "w");
  daemon = spawn(process.execPath, ["daemon/index.mjs"], {
    cwd: REPO,
    env: { ...process.env, CC_OFFICE_DEBUG: "1", ...extraEnv },
    stdio: ["ignore", fd, fd],
  });
  daemonLog = logFile;
  await until(
    () => readLog(logFile).includes("HTTP server listening") && listening(WS_PORT),
    30_000,
    "daemon",
  );
}
async function stopDaemon() {
  if (!daemon || daemon.exitCode !== null) return;
  daemon.kill("SIGTERM");
  await until(
    () => daemon.exitCode !== null || daemon.signalCode !== null,
    10_000,
    "daemon exit",
  ).catch(() => daemon.kill("SIGKILL"));
  await until(() => !listening(WS_PORT) && !listening(HTTP_PORT), 10_000, "ports free");
}
// Logs this run wrote. A -knob.log of an earlier run is removed at the start,
// so a verify command or a condition check never reads it.
const runLogs = () => [...new Set([LOG, daemonLog])].map(readLog).join("\n");

// A Word of its own for `doc`, with the CDP port on its pane.
function launchWord(doc, userData) {
  const launch = powershell("word-launch.ps1", [
    "-DocPath",
    winPath(doc),
    "-Port",
    opt["cdp-port"],
    "-UserDataFolder",
    winPath(userData),
  ]);
  if (launch.exited || record.word.preexisting.includes(launch.pid))
    throw new Error(`Word launch failed: ${JSON.stringify(launch)}`);
  return launch;
}
const embedTaskpane = (src, out) =>
  execFileSync(
    "python3",
    [path.join(HERE, "embed_taskpane.py"), "--manifest", opt.manifest, "--src", src, "--out", out],
    { stdio: "inherit" },
  );

// ---- verify commands -------------------------------------------------------
function featureIds(spec) {
  const [a, b] = spec.split("..").map((s) => Number(s.replace(/^S2-/, "")));
  const out = [];
  for (let i = a; i <= (b ?? a); i++) out.push(`S2-${String(i).padStart(2, "0")}`);
  return out;
}
function runFeatures() {
  if (!opt["feature-list"]) return;
  const list = JSON.parse(fs.readFileSync(opt["feature-list"], "utf8"));
  const features = Array.isArray(list) ? list : list.features;
  const map = {};
  if (opt["plan-dir"]) map[opt["plan-dir"].replace(/\/$/, "")] = RUN;
  if (opt["plan-repo"]) map[`cd ${opt["plan-repo"].replace(/\/$/, "")} `] = `cd ${REPO} `;
  map[opt["plan-log"]] = LOG;
  map[opt["plan-log"].replace(/\.log$/, "") + "-knob.log"] = KNOB_LOG;
  map["/tmp/s2-"] = path.join(OK_DIR, "s2-");
  for (const id of featureIds(opt.features)) {
    const f = features.find((x) => x.id === id);
    if (!f?.verify_command) continue;
    const cmd = rewriteCommand(f.verify_command, map);
    let code = 0;
    try {
      execFileSync("bash", ["-c", cmd], { stdio: "ignore", timeout: 300_000 });
    } catch (e) {
      code = e.status ?? 1;
    }
    record.features.push({ id, exit: code, command: cmd });
    save();
  }
}

// Marker files the procedure asks a person to create, written only from the
// machine evidence saved next to them.
function marker(name, evidence) {
  fs.mkdirSync(OK_DIR, { recursive: true });
  fs.writeFileSync(path.join(OK_DIR, `${name}.evidence.json`), JSON.stringify(evidence, null, 2));
  if (evidence.pass) fs.writeFileSync(path.join(OK_DIR, `${name}.ok`), "");
}

// ---- the run ---------------------------------------------------------------
async function propose(ref, label) {
  const before = count(daemonLog, /^\[agent\] turn complete/gm);
  const r = pane("propose", { ref });
  await until(
    () => count(daemonLog, /^\[agent\] turn complete/gm) > before,
    60_000,
    "turn complete",
  ).catch(() => {});
  const turnCompleted = count(daemonLog, /^\[agent\] turn complete/gm) > before;
  return note("propose", r, { label: label ?? `propose ${ref}`, ctx: { ref, turnCompleted } });
}

async function main() {
  fs.mkdirSync(TOOLS, { recursive: true });
  for (const f of ["pane-driver.cjs", "word-com.ps1", "word-launch.ps1"])
    fs.copyFileSync(path.join(HERE, f), path.join(TOOLS, f));
  if (listening(WS_PORT) || listening(HTTP_PORT))
    die(`port ${WS_PORT} or ${HTTP_PORT} is in use; stop that daemon first`);
  fs.rmSync(KNOB_LOG, { force: true });

  record.word.preexisting = wordPids();

  embedTaskpane(opt.fixture, DOC);
  await startDaemon(LOG);

  record.word.launch = launchWord(DOC, path.join(RUN, "webview2"));
  save();

  note("open-review", pane("open-review"), { ctx: { revisionPath: REVISION } });
  note("split", pane("split", { review_file: winPath(opt.review) }));
  note("edit-draft", pane("edit-draft", { requirements: { "R1.3": R13_REQUIREMENTS } }));
  note("ledger-add", pane("ledger-add", { ledger_file: winPath(opt.ledger) }));
  note("check-all", pane("check-all"));
  note("preview", pane("preview", { ref: "R1.3" }), { label: "preview R1.3" });
  for (const ref of ["R1.1", "R1.2", "R1.3", "R1.4", "R1.5"]) await propose(ref);

  // Knob: the model may avoid "novel" once it sees the forbidden list; then
  // the block is shown with the list left out of the prompt (procedure 8).
  if (!opt["skip-knob"] && !/^\[review\] blocked R1\.3 by=L2$/m.test(readLog(LOG))) {
    await stopDaemon();
    await startDaemon(KNOB_LOG, { CC_OFFICE_REVIEW_PROMPT_FORBID: "0" });
    const win = winPath(DOC);
    await until(() => readLog(KNOB_LOG).includes(`active doc: ${win}`), 120_000, "pane reconnect");
    await sleep(3000);
    await propose("R1.3", "propose R1.3 (forbid list off)");
    record.knob = { blocked: /^\[review\] blocked R1\.3 by=L2$/m.test(readLog(KNOB_LOG)) };
  }
  note("refresh", pane("refresh"), { label: "refresh before Word" });

  // The author in Word (procedure 10).
  const revision = JSON.parse(fs.readFileSync(REVISION, "utf8"));
  const plan = buildComPlan(revision);
  const planFile = path.join(RUN, "com-plan.json");
  fs.writeFileSync(planFile, JSON.stringify(plan, null, 2));
  record.word.state_before = com("state");
  // A tracked insertion of the author in the paragraph of the R1.4 change,
  // left pending: R1.4 must then show its own state, not the paragraph's.
  const r14 = plan.ops.find((o) => o.ref === "R1.4");
  const hand = r14
    ? com("hand-insert", [
        "-PlanJson",
        winPath(planFile),
        "-ChangeId",
        r14.change_id,
        "-Text",
        HAND_INSERT,
      ])
    : { ok: false, error: "no R1.4 change with a rationale comment" };
  record.word.hand_insert = hand;
  rows.push({
    step: "word: tracked insertion next to R1.4",
    result: hand.ok ? "PASS" : "FAIL",
    got: hand.ok ? `revisions in the paragraph: ${hand.revisions_in_paragraph}` : hand.error,
  });
  const applied = com("apply", ["-PlanJson", winPath(planFile)]);
  record.word.apply = applied;
  rows.push({
    step: "word: accept/reject/edit",
    result: applied.ok ? "PASS" : "FAIL",
    got: applied.results
      .map(
        (x) =>
          `${x.change_id} ${x.op}${x.error ? ` ERR ${x.error}` : ""}${x.edited ? ` ${x.edited.from}->${x.edited.to}` : ""}`,
      )
      .join("; "),
  });
  if (!applied.ok && opt["keep-on-fail"]) throw new Error("Word COM apply failed");
  const del = com("delete-text", ["-Text", OUTSIDE_SENTENCE]);
  record.word.delete_text = del;
  rows.push({
    step: "word: delete outside sentence",
    result: del.ok ? "PASS" : "FAIL",
    got: del.removed ?? del.error,
  });
  save();

  const after = note("refresh-after-user", pane("refresh"), { label: "refresh after Word" });
  marker("s2-list-matches", { pass: after.ok !== false && badgesMatch(after).pass, view: after });

  // C1: the rationale comments of the accepted and rejected changes are now
  // resolved, as Word itself reports them (Comment.Done).
  record.word.state_after_refresh = com("state");
  const resolved = resolvedMatch(
    JSON.parse(fs.readFileSync(REVISION, "utf8")),
    record.word.state_after_refresh,
    after.message,
  );
  record.word.resolved = resolved;
  rows.push({
    step: "word: rationale comments resolved",
    result: resolved.pass ? "PASS" : "FAIL",
    got: resolved.got,
  });
  const sibling = siblingMatch(
    JSON.parse(fs.readFileSync(REVISION, "utf8")),
    after,
    record.word.state_after_refresh,
    { ref: "R1.4", hand: HAND_INSERT },
  );
  record.word.sibling = sibling;
  rows.push({
    step: "word: R1.4 next to a pending revision",
    result: sibling.pass ? "PASS" : "FAIL",
    got: sibling.got,
  });
  save();

  note("clean-tool", pane("clean-tool"));
  note("refresh", pane("refresh"), { label: "refresh after cleanup", judge: "none" });
  const added = com("add-comment", ["-Text", "Author note added for the delete-all check."]);
  record.word.add_comment = added;
  rows.push({
    step: "word: add author comment",
    result: added.ok ? "PASS" : "FAIL",
    got: `comments=${added.comments}`,
  });
  note("delete-all", pane("delete-all"));
  note("refresh", pane("refresh"), { label: "refresh after delete-all", judge: "none" });

  // Evidence the procedure leaves to a person, judged from the files.
  const final = JSON.parse(fs.readFileSync(REVISION, "utf8"));
  const r15 = final.comments.find((c) => c.ref === "R1.5");
  const r15Turn = record.steps.find((s) => s.step === "propose R1.5");
  marker("s2-r15-declined", {
    pass: r15?.changes.length === 0,
    changes: r15?.changes.length,
    assistant: r15Turn?.raw?.assistant,
  });
  const all = final.comments.flatMap((c) => c.changes);
  marker("s2-narrow-minimal", {
    pass: all.length > 0 && all.every(minimalFind),
    changes: all.map((c) => ({
      id: c.id,
      kind: c.kind,
      find: c.find,
      text: c.text,
      minimal: minimalFind(c),
    })),
  });
}

// Close the Word this run started (pid, with `doc` open in it). It is stopped
// by force only when quit ran and the process stayed (cleanupDecision); a
// refused quit leaves Word and its documents alone and fails the run.
function quitWord(pid, doc, key = "quit") {
  if (!pid) return;
  let quit;
  try {
    quit = com("quit", ["-AllowDir", winPath(RUN), "-ExpectPid", String(pid)], doc);
  } catch (e) {
    quit = { ok: false, error: e.message };
  }
  record.cleanup[key] = quit;
  let alive = true;
  try {
    alive = wordPids().includes(pid);
  } catch {}
  const d = cleanupDecision({
    pid,
    preexisting: record.word.preexisting ?? [],
    quit,
    runDir: winPath(RUN),
    alive,
  });
  if (d.action === "kill") {
    try {
      execFileSync("powershell.exe", [
        "-NoProfile",
        "-Command",
        `Stop-Process -Id ${Number(pid)} -ErrorAction SilentlyContinue`,
      ]);
      (record.cleanup.stopped_pids ??= []).push(pid);
    } catch {}
  }
  if (d.action === "leave") {
    (record.cleanup.left_running ??= []).push({ pid, reason: d.reason });
    console.error(`run-acceptance: ${d.reason}`);
  }
  if (d.fail) rows.push({ step: `cleanup: Word ${pid}`, result: "FAIL", got: d.reason });
}

// A copy of the manuscript in compatibility mode 14 (Word 2010), where Word
// refuses Comment.resolved: one comment, one proposed change, accepted in
// Word. The refresh must still show its counts, with the note that the
// rationale comment could not be marked resolved, and the comment stays open.
// Runs after the main Word is gone (its CDP port is free), in a Word and a
// daemon log of its own.
async function mode14() {
  const label = (x) => `mode-14: ${x}`;
  const src = path.join(MODE14_DIR, "source.docx");
  fs.mkdirSync(MODE14_DIR, { recursive: true });
  execFileSync(
    "python3",
    [
      path.join(REPO, "scripts/make-test-manuscript.py"),
      "--variant",
      "review",
      "--compat",
      "14",
      src,
    ],
    { stdio: "ignore" },
  );
  embedTaskpane(src, DOC14);
  await stopDaemon();
  await startDaemon(MODE14_LOG);
  // The WebView2 of the Word that just quit may hold the CDP port a moment.
  const cdpBusy = () =>
    psList(
      `[int][bool](Get-NetTCPConnection -LocalPort ${Number(opt["cdp-port"])} -State Listen -ErrorAction SilentlyContinue)`,
    )[0] === 1;
  await until(() => !cdpBusy(), 60_000, "CDP port free");
  paneTag = "mode14-";
  record.word.launch_mode14 = launchWord(DOC14, path.join(RUN, "webview2-mode14"));
  save();
  const revisionPath = path.join(MODE14_DIR, "review_manuscript.revision.json");
  note("open-review", pane("open-review"), { label: label("open-review"), ctx: { revisionPath } });
  note("split", pane("split", { review_file: winPath(opt.review) }), { label: label("split") });
  note("edit-draft", pane("edit-draft", { keep_ref: "^R1\\.1$" }), {
    label: label("keep R1.1"),
    judge: "none",
  });
  await propose("R1.1", label("propose R1.1"));
  note("refresh", pane("refresh"), { label: label("refresh before Word") });
  const plan = buildComPlan(JSON.parse(fs.readFileSync(revisionPath, "utf8")), {});
  const planFile = path.join(MODE14_DIR, "com-plan.json");
  fs.writeFileSync(planFile, JSON.stringify(plan, null, 2));
  const applied = com("apply", ["-PlanJson", winPath(planFile)], DOC14);
  record.word.apply_mode14 = applied;
  rows.push({
    step: label("word: accept"),
    result: applied.ok && plan.ops.length > 0 ? "PASS" : "FAIL",
    got: [applied.results ?? []]
      .flat()
      .map((x) => `${x.change_id} ${x.op}${x.error ? ` ERR ${x.error}` : ""}`)
      .join("; "),
  });
  note("refresh-mode14", pane("refresh"), { label: label("refresh after Word") });
  const state = com("state", [], DOC14);
  record.word.state_mode14 = state;
  const texts = new Set(plan.ops.map((o) => norm(o.comment_text)));
  const tool = state.comments.filter((c) => texts.has(c.text));
  rows.push({
    step: label("word: rationale comments stay open"),
    result: tool.length > 0 && tool.every((c) => c.done !== true) ? "PASS" : "FAIL",
    got: tool.map((c) => `done=${c.done}`).join("; ") || "no rationale comment found",
  });
  save();
}

// After the main run and its verify commands: close the Word of the main
// run, then the mode-14 check, then the daemon.
async function cleanup() {
  const pid = record.word.launch?.pid;
  quitWord(pid, DOC);
  if (!failed && !opt["skip-mode14"]) {
    try {
      if (record.cleanup.left_running?.length || wordPids().includes(pid))
        throw new Error("the Word of the main run is still up");
      await mode14();
    } catch (e) {
      record.mode14_error = String(e?.stack ?? e);
      rows.push({ step: "mode-14: run", result: "FAIL", got: String(e?.message ?? e) });
    }
    quitWord(record.word.launch_mode14?.pid, DOC14, "quit_mode14");
  }
  await stopDaemon().catch((e) => (record.cleanup.daemon_error = e.message));
  try {
    const left = wordPids();
    const own = [pid, record.word.launch_mode14?.pid].filter(Boolean);
    record.cleanup.word_after = left;
    record.cleanup.preexisting_alive = (record.word.preexisting ?? []).every((p) =>
      left.includes(p),
    );
    record.cleanup.own_word_gone = own.length ? own.every((p) => !left.includes(p)) : null;
  } catch (e) {
    record.cleanup.word_after_error = e.message;
  }
}

let failed = null;
try {
  await main();
} catch (e) {
  failed = e;
  record.error = String(e?.stack ?? e);
  console.error(record.error);
}
// Another Word may have a pane connected to this daemon. Apart from its
// hello, no log line may name a document outside the run folder.
function isolationCheck() {
  const own = [winPath(RUN), RUN];
  const foreign = [LOG, KNOB_LOG]
    .flatMap((f) => readLog(f).split("\n"))
    .filter((l) => /\.docx\b/i.test(l) && !own.some((p) => l.includes(p)))
    .filter((l) => !/^\[bridge\] hello received|^\[diag\] hello /.test(l));
  record.isolation = { foreign_lines: foreign.slice(0, 20), count: foreign.length };
  rows.push({
    step: "no tool call on other documents",
    result: foreign.length ? "FAIL" : "PASS",
    got: foreign.length ? foreign[0] : "only hello lines name other documents",
  });
}

// Verify commands first: S2-18 checks that the daemon still listens.
isolationCheck();
try {
  runFeatures();
} catch (e) {
  record.features_error = String(e?.message ?? e);
}
// Before the mode-14 check starts a daemon with a log of its own.
const mainLogs = runLogs();
if (failed && opt["keep-on-fail"]) {
  record.cleanup.kept = { word_pid: record.word.launch?.pid, daemon_pid: daemon?.pid };
  daemon?.unref();
} else {
  await cleanup();
}
const decision = exitDecision({
  failed: !!failed,
  rows,
  features: record.features,
  featuresError: record.features_error,
  conditional: opt.conditional.flatMap((x) => x.split(",")).filter(Boolean),
  logs: mainLogs,
});
record.features = decision.features;
record.exit = { code: decision.code, reasons: decision.reasons };
record.finished_at = new Date().toISOString();
save();
console.log("\n" + table(rows));
if (record.features.length) {
  console.log("\nfeature  exit  status");
  for (const f of record.features)
    console.log(`${f.id.padEnd(7)}  ${String(f.exit).padEnd(4)}  ${f.status}`);
}
if (record.features_error) console.log(`\nfeatures error: ${record.features_error}`);
console.log(`\ncleanup: ${JSON.stringify(record.cleanup)}`);
console.log(`result: ${path.join(RUN, "result.json")}`);
for (const r of decision.reasons) console.error(`run-acceptance: ${r}`);
process.exit(decision.code);
