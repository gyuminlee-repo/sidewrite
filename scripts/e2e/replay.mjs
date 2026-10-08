#!/usr/bin/env node
// Replay of a real revision: a COPY of a manuscript is opened in a Word of
// its own, the review is pasted into the Review tab, every comment gets its
// "Propose changes" turn, and what the tool did is written out for a
// comparison with what the authors did (replay-compare.py). No hands.
//
//   node scripts/e2e/replay.mjs --manuscript <docx> --review <txt> \
//     --base-dir <folder on a Windows drive, WSL path> \
//     --playwright-core '<Windows path of a playwright-core package>' \
//     [--model opus] [--writing-model <alias>] [--judgment-model <alias>] \
//     [--only R1.1,R2.3] [--instructions <txt>] [--forbid <txt>] \
//     [--draft-responses] [--turn-minutes 10] [--repair] [--rewrite-stacked]
//     [--follow-through] [--answers <json>]
//
// Before the run, a Word of its own is asked whether it opens the copy the
// run will use (the add-in tagged on) as it is. A file Word wants to repair
// ends the run ("run again with --repair"); with --repair Word opens the
// manuscript with Open and Repair and saves a repaired copy in the run
// folder, the run's copy is made from that one and asked about again
// (summary.txt says so). The three models (Edits, Writing, Judgments) are chosen in Setup; the
// pane's own line says what ran, and a model other than the one asked for
// fails the run when the line gives no reason for it (a document that stays
// on another model because the chosen one declines it is no failure).
// With --rewrite-stacked, after the paragraph check every paragraph that
// holds three or more changes gets its "Rewrite this paragraph" turn.
// With --follow-through, after the turns of the comments and before the
// paragraph check every comment whose card offers "Check other places" gets
// that turn: the places its changes did not touch (the heading above, the
// legends of the figures cited, the methods, the rest of the section).
// With --answers <json> ({ "<ref>": "<answer text>", ... }), after the turns
// of the comments and before the follow-through turns every comment of the
// file that has open questions gets them answered as the author would on
// its card (each with the answer of its ref), and then its turn once more.
// The file is only read; an answer is written to the review file by the
// pane and to no report.
//
// The manuscript and the review are only read. Everything is written to a
// new folder <base-dir>/replay-<UTC time>/:
//   manuscript.docx              the copy, saved with its tracked changes
//                                and rationale comments
//   manuscript.revision.json     the record of the tool
//   manuscript.decisions.jsonl   the receipts of the model's judgments
//   changes.json                 per comment: requests and their state,
//                                changes (before, after, rationale, state,
//                                rewrite, words from the review), questions,
//                                outside notes, named terms and where they
//                                stand, how the turn ended, its time and model
//   paragraph_check.json         "Check revised paragraphs" at the end of
//                                the run: paragraphs with two or more
//                                changes and what the model found in them
//   responses.json               with --draft-responses: answers and quotes
//   replies.json                 the answer of the model in each turn
//                                (of a comment, after answers, follow-through,
//                                rewrite), for reading the run afterwards
//   summary.txt                  counts and times
//   result.json, daemon.log      every step as the pane showed it; the log
// The text of the manuscript is in these files and nowhere else: nothing of
// it is printed.
//
// Runs from WSL; Word, its pane and COM are on Windows, as run-acceptance.mjs
// (the same Windows scripts, the same rules: only the Word this run started
// is closed, the resident daemon is stopped for the run and started again).
// A turn that fails or runs out of time does not end the run. One that ends
// on the usage limit of the account does: no further turn is started, the
// steps that need the model are passed over, the copy is saved and the
// reports are written (summary.txt: "stopped: usage limit reached at ...").
// Exit code: 0 when the flow ran to its end, the copy was saved and the
// pane's line names no other model than the one asked for without a reason,
// else 1; 2 for bad arguments; 3 for a usage stop with the copy saved and
// cleanup complete. A save or cleanup failure still returns 1.

import { spawn, execFileSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { parseArgs } from "node:util";
import { cleanupDecision } from "./judge.mjs";
import { pauseResident } from "./resident-unit.mjs";
import { execWithInteropRetry } from "./wsl-interop.mjs";
import { consistencyLines } from "../../taskpane/shared/review-view.js";
import { followState } from "../../taskpane/shared/follow-places.js";
import {
  changesReport,
  extractBatchesFromLog,
  keepRefPattern,
  ledgerFromLines,
  modelFailures,
  answeredOf,
  limitStop,
  replayExitCode,
  answerPlan,
  parseAnswers,
  replyRow,
  parseOnly,
  responsesReport,
  restoredFromLog,
  staysFromLog,
  summaryText,
  newWords,
  sweepSelection,
  mergeAttempts,
  retryTurn,
  turnFromLog,
  turnResult,
} from "./replay-lib.mjs";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "../..");
const WS_PORT = 47833;
const HTTP_PORT = 47834;

const { values: opt } = parseArgs({
  options: {
    manuscript: { type: "string" },
    review: { type: "string" },
    "base-dir": { type: "string", default: process.env.E2E_BASE_DIR },
    "playwright-core": { type: "string", default: process.env.E2E_PLAYWRIGHT_CORE },
    // The Claude models of the three kinds of work (Setup, Providers): Edits
    // (the review turns), Writing (response drafts), Judgments (split,
    // requirements, checks). Writing and Judgments not given: the pane's own
    // choice stays.
    model: { type: "string", default: "opus" },
    "writing-model": { type: "string" },
    "judgment-model": { type: "string" },
    // Let Word repair a file it does not open as it is (Open and Repair), and
    // run on the repaired copy.
    repair: { type: "boolean", default: false },
    only: { type: "string" },
    instructions: { type: "string" },
    forbid: { type: "string" },
    "draft-responses": { type: "boolean", default: false },
    // After the paragraph check: a rewrite turn for every paragraph that
    // holds three or more changes ("Rewrite this paragraph").
    "rewrite-stacked": { type: "boolean", default: false },
    // After the turns of the comments: a second turn for every comment whose
    // changes left other places to check ("Check other places").
    "follow-through": { type: "boolean", default: false },
    // What the author answers when asked: a JSON file { "<ref>": "<text>" }.
    // After the turns of the comments their open questions get these
    // answers, and each such comment its turn again.
    answers: { type: "string" },
    // The longest one review turn may take.
    "turn-minutes": { type: "string", default: "10" },
    manifest: { type: "string", default: path.join(REPO, "manifests/word.xml") },
    "cdp-port": { type: "string", default: process.env.E2E_CDP_PORT ?? "9333" },
    "win-node": {
      type: "string",
      default: process.env.E2E_WIN_NODE ?? "C:\\Program Files\\nodejs\\node.exe",
    },
    // Leave Word and the daemon running when the run fails, to look at them.
    "keep-on-fail": { type: "boolean", default: false },
  },
});

const die = (m) => {
  console.error(`replay: ${m}`);
  process.exit(2);
};
for (const name of ["manuscript", "review", "base-dir", "playwright-core"])
  if (!opt[name]) die(`--${name} is required`);
for (const name of ["manuscript", "review", "instructions", "forbid", "answers"])
  if (opt[name] && !fs.existsSync(opt[name])) die(`--${name}: no such file`);
if (!/\.docx$/i.test(opt.manuscript)) die("--manuscript takes a .docx file");
const TURN_MS = Number(opt["turn-minutes"]) * 60_000;
if (!(TURN_MS > 0)) die("--turn-minutes takes a number of minutes");
let ONLY;
// The refs that have an answer in the file of --answers. The answers are
// read here only to check the file; the pane step reads it for itself.
let ANSWERS = null;
try {
  ONLY = parseOnly(opt.only);
  if (opt.answers) {
    let json;
    try {
      json = JSON.parse(fs.readFileSync(opt.answers, "utf8"));
    } catch {
      throw new Error("--answers: the file is not JSON");
    }
    ANSWERS = Object.fromEntries(Object.keys(parseAnswers(json)).map((ref) => [ref, true]));
  }
} catch (e) {
  die(e.message);
}

const stamp = new Date().toISOString().replace(/[-:]/g, "").replace("T", "-").slice(0, 15);
const RUN = path.resolve(opt["base-dir"], `replay-${stamp}`);
const TOOLS = path.join(RUN, "tools");
const DOC = path.join(RUN, "manuscript.docx");
// The manuscript as it was given, and what Word made of it with --repair.
const SOURCE = path.join(RUN, "source.docx");
const REPAIRED = path.join(RUN, "source_repaired.docx");
const REVISION = path.join(RUN, "manuscript.revision.json");
const LOG = path.join(RUN, "daemon.log");
const winPath = (p) => execFileSync("wslpath", ["-w", p], { encoding: "utf8" }).trim();
const unixPath = (p) => execFileSync("wslpath", ["-u", p], { encoding: "utf8" }).trim();
const started = Date.now();

const record = {
  run_dir: RUN,
  started_at: new Date().toISOString(),
  steps: [],
  word: {},
  cleanup: {},
};
const save = () => fs.writeFileSync(path.join(RUN, "result.json"), JSON.stringify(record, null, 2));
// One line per step on the screen: its name and whether it worked. What the
// pane showed (it may quote the manuscript) goes to result.json only.
function note(step, result) {
  record.steps.push({ step, ok: result?.ok !== false, raw: result });
  save();
  console.log(`${result?.ok === false ? "FAIL" : "ok  "}  ${step}`);
  return result;
}

// ---- Windows calls (as run-acceptance.mjs) ---------------------------------
// Every start of a Windows program goes through winExec: it is tried once
// more when WSL itself could not start it (wsl-interop.mjs).
const winExec = (file, argv, options, what = path.basename(file)) =>
  execWithInteropRetry(() => execFileSync(file, argv, options), { what });
function powershell(script, args, timeoutMs = 300_000) {
  const out = winExec(
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
    `powershell.exe ${script} ${args.join(" ")}`,
  );
  return JSON.parse(out.replace(/\r/g, "").trim().split("\n").filter(Boolean).pop());
}
const psList = (cmd) =>
  winExec("powershell.exe", ["-NoProfile", "-Command", cmd], { encoding: "utf8" }, "powershell.exe")
    .replace(/\r/g, "")
    .trim()
    .split(",")
    .filter(Boolean)
    .map(Number);
const wordPids = () =>
  psList("(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id) -join ','");
// Every WINWORD running now: { pid, started (UTC, ISO; null when Windows does
// not give it), title }.
const wordProcesses = () =>
  JSON.parse(
    winExec(
      "powershell.exe",
      [
        "-NoProfile",
        "-Command",
        "ConvertTo-Json -Compress -InputObject @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $s = $null; try { $s = $_.StartTime.ToUniversalTime().ToString('o') } catch { $s = $null }; [ordered]@{ pid = $_.Id; started = $s; title = [string]$_.MainWindowTitle } })",
      ],
      { encoding: "utf8" },
      "powershell.exe",
    ).trim() || "[]",
  );

let stepNo = 0;
function pane(step, args = {}, timeoutMs = 30 * 60_000) {
  const argsFile = path.join(RUN, `args-${String(++stepNo).padStart(3, "0")}-${step}.json`);
  fs.writeFileSync(argsFile, JSON.stringify(args));
  let out;
  try {
    out = winExec(
      unixPath(opt["win-node"]),
      [winPath(path.join(TOOLS, "replay-driver.cjs")), step, winPath(argsFile)],
      {
        cwd: RUN,
        encoding: "utf8",
        timeout: timeoutMs,
        maxBuffer: 64 << 20,
        env: {
          ...process.env,
          PLAYWRIGHT_CORE: opt["playwright-core"],
          E2E_CDP_PORT: opt["cdp-port"],
          E2E_TURN_TIMEOUT_MS: String(TURN_MS),
          WSLENV: [process.env.WSLENV, "PLAYWRIGHT_CORE", "E2E_CDP_PORT", "E2E_TURN_TIMEOUT_MS"]
            .filter(Boolean)
            .join(":"),
        },
      },
      `replay-driver ${step}`,
    );
  } catch (e) {
    // The driver prints its own failure as a JSON line before it exits.
    out = e.stdout || JSON.stringify({ ok: false, error: e.message });
  }
  const line = String(out).replace(/\r/g, "").trim().split("\n").filter(Boolean).pop();
  try {
    return JSON.parse(line);
  } catch (e) {
    return { ok: false, error: `the pane step gave no JSON line (${e.message})` };
  }
}

// ---- daemon (as run-acceptance.mjs) ----------------------------------------
const listening = (port) =>
  execFileSync("ss", ["-ltn"], { encoding: "utf8" }).includes(`127.0.0.1:${port} `);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
async function until(fn, ms, what) {
  const end = Date.now() + ms;
  while (Date.now() < end) {
    if (await fn()) return;
    await sleep(500);
  }
  throw new Error(`timed out waiting for ${what}`);
}
const readLog = () => (fs.existsSync(LOG) ? fs.readFileSync(LOG, "utf8") : "");

let daemon = null;
// The resident daemon of this machine holds both ports: it is stopped for
// the run and started again when this process ends, whichever way.
let resident = { paused: false, resume: () => {} };
process.on("exit", () => resident.resume());
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
for (const [signal, code] of [
  ["SIGINT", 130],
  ["SIGTERM", 143],
]) {
  process.on(signal, () => {
    // The run's own daemon first, until its ports are free: a resident unit
    // started while they are held fails.
    stopDaemon()
      .catch((e) => console.error(`replay: ${e.message}`))
      .finally(() => process.exit(code));
  });
}
async function startDaemon() {
  const fd = fs.openSync(LOG, "w");
  daemon = spawn(process.execPath, ["daemon/index.mjs"], {
    cwd: REPO,
    env: { ...process.env, CC_OFFICE_DEBUG: "1" },
    stdio: ["ignore", fd, fd],
  });
  await until(
    () => readLog().includes("HTTP server listening") && listening(WS_PORT),
    30_000,
    "daemon",
  );
}

// Quit through COM with current ownership checks. A failed or refused quit
// leaves Word running; an old PID does not authorize process termination.
function quitWord(pid) {
  if (!pid) return;
  let quit;
  try {
    quit = powershell("word-com.ps1", [
      "-Doc",
      winPath(DOC),
      "-Action",
      "quit",
      "-AllowDir",
      winPath(RUN),
      "-ExpectPid",
      String(pid),
    ]);
  } catch (e) {
    quit = { ok: false, error: e.message };
  }
  record.cleanup.quit = quit;
  // Not known: taken as still running, which never leads to more than the
  // decision below allows.
  let alive = true;
  try {
    alive = wordPids().includes(pid);
  } catch (e) {
    record.cleanup.pid_check_error = e.message;
  }
  const d = cleanupDecision({
    pid,
    preexisting: record.word.preexisting ?? [],
    quit,
    runDir: winPath(RUN),
    alive,
  });
  if (d.action === "leave" || (d.action === "kill" && alive)) {
    record.cleanup.left_running = { pid, reason: d.reason };
    record.cleanup.failed = true;
    console.error(`replay: ${d.reason}`);
  } else {
    record.cleanup.failed ||= d.fail;
  }
}

// A Word step of the run, with the WINWORD processes listed right before and
// after it. One that came up and is not the step's own Word (own(result)) is
// recorded in record.word.strays with the step for reporting only.
function withStrays(step, run, own = () => null) {
  const before = wordProcesses();
  let result;
  try {
    result = run();
  } finally {
    const found = newWords(before, wordProcesses(), own(result)).map((p) => ({ ...p, step }));
    (record.word.strays ??= []).push(...found);
    if (result && typeof result === "object") result.strays = found;
  }
  return result;
}

// Report remaining Word processes without closing any unproven stray.
function sweepWord() {
  const preexisting = record.word.preexisting;
  // Without the list of the run's start nothing is provably the run's.
  if (!Array.isArray(preexisting)) return;
  const keep = [record.cleanup.kept?.word_pid, record.cleanup.left_running?.pid].filter(Boolean);
  let sel;
  try {
    sel = sweepSelection({
      processes: wordProcesses(),
      strays: record.word.strays ?? [],
      preexisting,
      keep,
    });
  } catch (e) {
    record.cleanup.sweep = { error: e.message };
    console.error(`replay: the Word processes could not be listed: ${e.message}`);
    return;
  }
  record.cleanup.sweep = { closed: [], skipped: sel.skip };
  for (const s of sel.skip.filter((x) => x.reason.startsWith("left open")))
    console.log(`sweep: Word ${s.pid} ("${s.title}") ${s.reason}`);
}

// One run of replay-open.ps1 in a Word of its own. Cleanup failures remain
// visible and never trigger a process kill. file: what is opened; repairTo: the file
// is opened with Open and Repair and saved under that name.
let checks = 0;
function wordOpens(file, repairTo = null) {
  const anchor = path.join(RUN, "anchor.docx");
  const pidFile = path.join(RUN, `open-check-${++checks}-pid.json`);
  if (!fs.existsSync(anchor))
    fs.copyFileSync(path.join(REPO, "tests/fixtures/spike_manuscript.docx"), anchor);
  const ownPid = (r) =>
    r?.pid ??
    (fs.existsSync(pidFile)
      ? JSON.parse(fs.readFileSync(pidFile, "utf8").replace(/^\uFEFF/, "")).pid
      : null);
  // A Word that came up during the check and is not the one it started is
  // recorded in r.strays, with the check it came from.
  const r = withStrays(
    `open check ${checks}`,
    () => {
      try {
        return powershell(
          "replay-open.ps1",
          [
            ...["-File", winPath(file), "-Anchor", winPath(anchor)],
            ...["-AllowDir", winPath(RUN), "-PidFile", winPath(pidFile)],
            ...(repairTo ? ["-RepairTo", winPath(repairTo)] : []),
          ],
          5 * 60_000,
        );
      } catch (e) {
        return { ok: false, error: e.message };
      }
    },
    ownPid,
  );
  const pid = ownPid(r);
  if (pid && wordPids().includes(pid)) {
    r.left_running = {
      pid,
      reason: r.quit?.refused ?? r.quit?.error ?? "Word still running after open check",
    };
    record.cleanup.left_running = r.left_running;
    record.cleanup.failed = true;
    console.error(`replay: Word ${pid} left running: ${r.left_running.reason}`);
  }
  (record.word.open_checks ??= []).push({ file: path.basename(file), repair: !!repairTo, ...r });
  save();
  return r;
}

const embed = (src) => {
  fs.rmSync(DOC, { force: true });
  execFileSync(
    "python3",
    [path.join(HERE, "embed_taskpane.py"), "--manifest", opt.manifest, "--src", src, "--out", DOC],
    { stdio: "inherit" },
  );
};

// The copy the pane opens (the add-in tagged on, embed_taskpane.py), made
// from the manuscript, and the check that Word opens exactly that file as it
// is. Returns how the manuscript came to be opened, for summary.txt.
function prepareCopy() {
  const why = (r) => r.repair_error ?? r.open_error ?? r.error ?? "no answer";
  embed(SOURCE);
  const first = wordOpens(DOC);
  if (first.opens) return "opened as it is";
  if (first.ok === false) throw new Error(`the open check did not run: ${why(first)}`);
  if (!opt.repair) throw new Error("Word needs to repair this file; run again with --repair");
  const fixed = wordOpens(SOURCE, REPAIRED);
  if (!fixed.repaired || !fs.existsSync(REPAIRED))
    throw new Error(`Word could not repair the manuscript: ${why(fixed)}`);
  embed(REPAIRED);
  const second = wordOpens(DOC);
  if (!second.opens) throw new Error(`Word does not open the repaired copy either: ${why(second)}`);
  return "Word could not open the file as it is; the run used the copy Word repaired (Open and Repair)";
}

// ---- the run ---------------------------------------------------------------
async function setup() {
  fs.mkdirSync(TOOLS, { recursive: true });
  for (const f of [
    "pane-driver.cjs",
    "replay-driver.cjs",
    "word-com.ps1",
    "word-launch.ps1",
    "word-own.ps1",
    "word-rot.ps1",
    "word-back.ps1",
    "replay-open.ps1",
    "replay-save.ps1",
  ])
    fs.copyFileSync(path.join(HERE, f), path.join(TOOLS, f));
  record.word.preexisting = wordPids();
  // The manuscript is read once, into the run folder; Word only ever sees
  // copies.
  fs.copyFileSync(opt.manuscript, SOURCE);
  // The resident daemon first: the pane of the copy that the open check
  // opens must not find a daemon to talk to.
  resident = pauseResident();
  if (resident.paused)
    await until(() => !listening(WS_PORT) && !listening(HTTP_PORT), 10_000, "ports free");
  if (listening(WS_PORT) || listening(HTTP_PORT))
    throw new Error(`port ${WS_PORT} or ${HTTP_PORT} is in use; stop that daemon first`);
  record.source = prepareCopy();
  save();
  console.log(`ok    open check: ${record.source}`);
  await startDaemon();
  const launch = withStrays(
    "word launch",
    () =>
      powershell("word-launch.ps1", [
        "-DocPath",
        winPath(DOC),
        "-Port",
        opt["cdp-port"],
        "-UserDataFolder",
        winPath(path.join(RUN, "webview2")),
      ]),
    (r) => r?.pid ?? null,
  );
  if (launch.exited || record.word.preexisting.includes(launch.pid))
    throw new Error(`Word launch failed: ${JSON.stringify(launch)}`);
  record.word.launch = launch;
  save();
}

// A step the run cannot go on without.
function must(step, result) {
  note(step, result);
  if (result?.ok === false) throw new Error(`${step}: ${result.error}`);
  if (result?.refused) throw new Error(`${step}: ${result.refused}`);
  return result;
}

const turns = {};
let saved = false;
let drafted = null;
// What the pane says each kind of work ran on, at the end of the run.
let modelsLine = "";
// The rewrite turns of --rewrite-stacked, and the last turn that ran.
const rewriteTurns = [];
// The turns of --follow-through.
const followTurns = [];
// --answers: the questions answered per comment, and its turn after them.
const answered = {};
const answerTurns = {};
// The answer of the model in each turn (replies.json).
const replies = [];
let lastTurn = null;
// The turn that ended on the usage limit (replay-lib.mjs limitStop): no turn
// is started after it.
let stopped = null;
// A model other than the one asked for, with no reason in the pane's line
// (replay-lib.mjs modelFailures).
const otherModel = () => modelFailures(modelsLine, lastTurn).length > 0;

// One review turn through the pane: the step, its part of the daemon log,
// and what the pane says the edits ran on right after it. kind: what turn
// it is, for replies.json (replay-lib.mjs REPLY_KINDS).
function runTurn(ref, label, kind, step) {
  const at = readLog().length;
  const t0 = Date.now();
  const done = note(label, step());
  replies.push(replyRow(ref, kind, done));
  if (done.ok === false) note(`stop-turn ${ref}`, pane("stop-turn", { wait_ms: 120_000 }));
  const log = readLog();
  lastTurn = turnResult(ref, {
    step: done,
    log: log.slice(at),
    seconds: done.seconds ?? Math.round((Date.now() - t0) / 1000),
    model: turnFromLog(log, ref).model,
    line: pane("models-line").line ?? null,
  });
  stopped ??= limitStop(ref, kind, lastTurn);
  if (stopped) {
    record.stopped = stopped;
    save();
  }
  return lastTurn;
}

// The turn of a comment, or its follow-through turn: one that ended on an
// API error is run once more after a minute (replay-lib.mjs retryTurn). The
// entry keeps how each try ended (mergeAttempts).
async function proposeTurn(ref, label, kind, step) {
  const attempts = [runTurn(ref, label, kind, step)];
  while (retryTurn(attempts)) {
    await sleep(60_000);
    attempts.push(runTurn(ref, `${label} (again after an API error)`, kind, step));
  }
  return mergeAttempts(attempts);
}

async function main() {
  must("open-review", pane("open-review"));
  // The three kinds of work, chosen in Setup (Providers).
  must("set-model (Edits)", pane("set-model", { model: opt.model }));
  for (const [role, name] of [
    ["writing", "writing-model"],
    ["oneshot", "judgment-model"],
  ]) {
    if (!opt[name]) continue;
    const r = must(
      `select-provider (${role})`,
      pane("select-provider", { role, id: "claude", model: opt[name] }),
    );
    if (r.message) throw new Error(`select-provider (${role}): ${r.message}`);
  }
  must("open-review", pane("open-review"));

  must("split", pane("split", { review_file: winPath(path.resolve(opt.review)) }));
  const draft = must("edit-draft", pane("edit-draft", { keep_ref: keepRefPattern(ONLY) }));
  const refs = draft.kept ?? [];
  const absent = ONLY.filter((r) => !refs.includes(r));
  if (absent.length) throw new Error(`the split gave no comment ${absent.join(", ")}`);
  if (!refs.length) throw new Error("the split gave no comment");
  record.refs = refs;
  record.refs_removed = draft.removed ?? [];

  const ledger = ledgerFromLines({
    instructions: opt.instructions ? fs.readFileSync(opt.instructions, "utf8") : "",
    forbid: opt.forbid ? fs.readFileSync(opt.forbid, "utf8") : "",
  });
  if (ledger.items.length) {
    const file = path.join(RUN, "ledger.json");
    fs.writeFileSync(file, JSON.stringify(ledger, null, 2));
    must("ledger-add", pane("ledger-add", { ledger_file: winPath(file) }));
  }

  // One turn per comment, in order. A turn that fails or runs out of time is
  // written down and the next comment goes on.
  for (const ref of refs) {
    if (stopped) break;
    turns[ref] = await proposeTurn(ref, `propose ${ref}`, "ordinary", () =>
      pane("propose", { ref }, TURN_MS + 3 * 60_000),
    );
    record.turns = turns;
    save();
    console.log(`      ${ref}: ${turns[ref].outcome}, ${turns[ref].seconds} s`);
  }
  note("refresh", pane("refresh"));
  if (ANSWERS && !stopped) {
    // The comments of the file that now have open questions, by the review
    // file of that refresh: their questions are answered on the card, then
    // the comment gets its turn again, with the answers in its prompt. A ref
    // of the file without an open question is named and passed over.
    const plan = answerPlan(readRevision(), ANSWERS, refs);
    record.answers = { due: plan.due, skipped: plan.skipped };
    save();
    for (const s of plan.skipped) console.log(`      answers ${s.ref}: skipped, ${s.reason}`);
    for (const { ref, open } of plan.due) {
      if (stopped) break;
      note(
        `answer-questions ${ref}`,
        pane("answer-questions", { ref, answers_file: winPath(path.resolve(opt.answers)) }),
      );
      // What the review file now holds, whatever the step said: no turn
      // after a step that saved nothing.
      answered[ref] = answeredOf(readRevision(), ref, open);
      if (!answered[ref]) continue;
      answerTurns[ref] = await proposeTurn(ref, `propose ${ref} (after answers)`, "answer", () =>
        pane("propose", { ref }, TURN_MS + 3 * 60_000),
      );
      record.answer_turns = answerTurns;
      save();
      console.log(
        `      ${ref} after ${answered[ref]} answers: ${answerTurns[ref].outcome}, ${answerTurns[ref].seconds} s`,
      );
    }
    if (Object.keys(answerTurns).length) note("refresh", pane("refresh"));
  }
  if (opt["follow-through"] && !stopped) {
    // The comments whose card now offers the check, by the rule of the card
    // (follow-places.js followState) on the review file of that refresh.
    const rev = readRevision();
    const due = refs.filter(
      (ref) =>
        followState(
          (rev?.comments ?? []).find((c) => c.ref === ref),
          rev?.snapshot,
        ).offer,
    );
    for (const ref of due) {
      if (stopped) break;
      const turn = await proposeTurn(ref, `follow-through ${ref}`, "follow-through", () =>
        pane("follow-through", { ref }, TURN_MS + 3 * 60_000),
      );
      followTurns.push({ ref, ...turn });
      record.follow_turns = followTurns;
      save();
      console.log(`      follow-through ${ref}: ${turn.outcome}, ${turn.seconds} s`);
    }
    if (due.length) note("refresh", pane("refresh"));
  }
  // The paragraphs that now hold two or more changes, read once more by the
  // model of the edits. Its result is in revision.json (consistency).
  if (!stopped) note("check-paragraphs", pane("check-paragraphs", {}, 25 * 60_000));
  if (opt["rewrite-stacked"] && !stopped) {
    // The buttons stand in the order of these lines (review-pane.js draws
    // them from the same list), and a rewrite does not change the list: the
    // check is not run again.
    const stacked = consistencyLines(readRevision()?.consistency).filter((l) => l.rewrite_ref);
    for (const [nth, line] of stacked.entries()) {
      if (stopped) break;
      const ref = line.rewrite_ref;
      const turn = runTurn(ref, `rewrite ${ref} p${line.paragraph_index}`, "rewrite", () =>
        pane("rewrite-paragraph", { nth }, TURN_MS + 3 * 60_000),
      );
      rewriteTurns.push({ ref, paragraph_index: line.paragraph_index, ...turn });
      record.rewrite_turns = rewriteTurns;
      save();
      console.log(
        `      rewrite ${ref} p${line.paragraph_index}: ${turn.outcome}, ${turn.seconds} s`,
      );
    }
    if (stacked.length) note("refresh", pane("refresh"));
  }

  if (opt["draft-responses"] && !stopped) {
    const opened = note("open-response", pane("open-response"));
    if (opened.ok !== false) {
      const r = note(
        "response-draft-all",
        pane("response-draft-all", {}, TURN_MS * refs.length + 3 * 60_000),
      );
      drafted = r.ok === false ? 0 : (r.cards ?? []).filter((c) => c.text.trim()).length;
    }
  }

  modelsLine = note("models-line", pane("models-line")).line ?? "";
  const headings = note("headings", pane("headings")).headings ?? [];
  // Our copy: it is saved as it is, tracked changes and comments in it.
  const done = note(
    "save copy",
    withStrays("save copy", () =>
      powershell("replay-save.ps1", ["-Doc", winPath(DOC), "-AllowDir", winPath(RUN)]),
    ),
  );
  saved = done.ok === true;
  writeReports(headings);
}

const readRevision = () =>
  fs.existsSync(REVISION) ? JSON.parse(fs.readFileSync(REVISION, "utf8")) : null;

function writeReports(headings = []) {
  const revision = readRevision();
  if (!revision) return;
  const report = changesReport(revision, headings, turns, rewriteTurns, followTurns, {
    answered,
    turns: answerTurns,
  });
  const log = readLog();
  const out = (name, text) => fs.writeFileSync(path.join(RUN, name), text);
  out("changes.json", JSON.stringify(report, null, 2));
  out("replies.json", JSON.stringify(replies, null, 2));
  out(
    "paragraph_check.json",
    JSON.stringify(
      revision.consistency ?? { checked_at: null, note: "the check left no result" },
      null,
      2,
    ),
  );
  if (opt["draft-responses"])
    out("responses.json", JSON.stringify(responsesReport(revision), null, 2));
  out(
    "summary.txt",
    summaryText(report, {
      seconds: Math.round((Date.now() - started) / 1000),
      saved,
      models: {
        asked: {
          edits: opt.model,
          writing: opt["writing-model"] ?? null,
          judgments: opt["judgment-model"] ?? null,
        },
        line: modelsLine,
      },
      source: record.source ?? null,
      paragraphCheck: revision.consistency ?? null,
      extraction: extractBatchesFromLog(log),
      stays: staysFromLog(log),
      restored: restoredFromLog(log),
      lastTurn,
      stopped,
      responses: opt["draft-responses"] ? (drafted ?? 0) : null,
    }),
  );
}

let failed = null;
try {
  await setup();
  await main();
} catch (e) {
  failed = e;
  record.error = String(e?.stack ?? e);
  console.error(`replay: ${e?.message ?? e}`);
  // What was recorded before the failure is still written out.
  try {
    writeReports();
  } catch (x) {
    record.report_error = String(x?.message ?? x);
    console.error(`replay: the reports could not be written: ${record.report_error}`);
  }
}

if (failed && opt["keep-on-fail"]) {
  record.cleanup.kept = { word_pid: record.word.launch?.pid, daemon_pid: daemon?.pid };
  daemon?.unref();
} else {
  quitWord(record.word.launch?.pid);
  await stopDaemon().catch((e) => (record.cleanup.daemon_error = e.message));
}
sweepWord();
record.finished_at = new Date().toISOString();
if (fs.existsSync(RUN)) save();
const summary = path.join(RUN, "summary.txt");
if (fs.existsSync(summary)) console.log(`\n${fs.readFileSync(summary, "utf8")}`);
console.log(`run folder: ${RUN}`);
process.exit(
  replayExitCode({ failed, stopped, saved, otherModel: otherModel(), cleanup: record.cleanup }),
);
