#!/usr/bin/env node
// Review acceptance run without hands: daemon, a separate Word process with
// a CDP port on its add-in pane, the Review tab steps, the author's Word
// steps over COM, the feature_list verify commands, cleanup. Five flows:
// stage 2 (main), stage 2.5 (mainS25), stage 3 (mainS3), stage 4 (mainS4,
// the Tools tab) and stage cite (mainCite, reference-manager fields), chosen
// by --stage or --features.
// Runs from WSL; Word, its pane and COM are on Windows. See README.md.

import { spawn, execFileSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import {
  addBlocked,
  dropBlocked,
  loadRevision,
  saveRevision,
} from "../../daemon/revision-store.mjs";
import { parseArgs } from "node:util";
import {
  anonymizeEvidence,
  badgesMatch,
  buildComPlan,
  CITE_SEED_REF,
  CITE_TURN_REFS,
  citeAcceptEvidence,
  citeCaseRow,
  citeEdit,
  citeOutcome,
  citeRewriteCall,
  citeRewriteRow,
  cleanupDecision,
  cleanupEvidence,
  countEvidence,
  defaultFeatures,
  docStem,
  exitDecision,
  featureIds,
  exportEvidence,
  fieldsLine,
  focusEvidence,
  formatEvidence,
  guidelineEvidence,
  insertTarget,
  FORBID_GUARD_EDIT,
  FORBID_GUARD_SEED_RULE,
  forbidBlockEvidence,
  shotName,
  stepTiming,
  JUDGES,
  judge,
  langEvidence,
  ledgerEvidence,
  lineShift,
  markedEvidence,
  minimalFind,
  modelsEvidence,
  norm,
  ollamaSplitEvidence,
  openCheckEvidence,
  openEvidence,
  pluginsEvidence,
  proposalEvidence,
  proposeAllEvidence,
  providersEvidence,
  resolvedMatch,
  restartSplitEvidence,
  rewriteVerify,
  S3_DECISIONS,
  secondReader,
  selectionEvidence,
  siblingMatch,
  stageOf,
  staleEvidence,
  staleMarkers,
  suggestEvidence,
  trackedPlaced,
  table,
} from "./judge.mjs";
import { pauseResident } from "./resident-unit.mjs";
import { refusalsFromLog, rewritesFromLog } from "./replay-lib.mjs";

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
    // Default: tests/fixtures/review_manuscript.docx; stage 4:
    // tests/fixtures/aux_manuscript.docx (scripts/e2e/make_aux_fixture.py).
    fixture: { type: "string" },
    // Stage 4: the guideline text that is pasted, the earlier version the
    // marked copy is compared with, the language the selection is translated
    // into.
    guideline: { type: "string", default: path.join(REPO, "tests/fixtures/fake_guideline.txt") },
    baseline: { type: "string", default: path.join(REPO, "tests/fixtures/aux_manuscript_v1.docx") },
    "translate-lang": { type: "string", default: "German" },
    // Default: tests/fixtures/fake_review.txt; stage cite: cite_review.txt.
    review: { type: "string" },
    ledger: { type: "string", default: path.join(REPO, "tests/fixtures/review_ledger.json") },
    manifest: { type: "string", default: path.join(REPO, "manifests/word.xml") },
    // Default /tmp/daemon-<stage>-e2e.log.
    log: { type: "string" },
    "cdp-port": { type: "string", default: process.env.E2E_CDP_PORT ?? "9333" },
    "win-node": {
      type: "string",
      default: process.env.E2E_WIN_NODE ?? "C:\\Program Files\\nodejs\\node.exe",
    },
    "playwright-core": { type: "string", default: process.env.E2E_PLAYWRIGHT_CORE },
    "feature-list": { type: "string", default: process.env.E2E_FEATURE_LIST },
    "plan-dir": { type: "string", default: process.env.E2E_PLAN_DIR },
    "plan-repo": { type: "string", default: process.env.E2E_PLAN_REPO },
    // Default /tmp/daemon-<stage>.log; --plan-dir defaults to /tmp/<stage>-run.
    "plan-log": { type: "string" },
    // Default: the range of the stage (S2-18..S2-34, S25-30..S25-41,
    // S3-30..S3-40).
    features: { type: "string" },
    // Features whose verify command runs after the cleanup (Word closed, the
    // daemon stopped). Default for stage 3: S3-41.
    "after-cleanup": { type: "string" },
    // The provider of the judgments (split, ledger suggestions, and in stage 3
    // the drafts of the letter). Stage 2.5: default ollama. Stage 3: when not
    // given, the choice of the pane stays as it is.
    "judgment-provider": { type: "string" },
    // Stage 3: an existing response letter whose colours and frames are read.
    letter: { type: "string", default: path.join(REPO, "tests/fixtures/prior_response.docx") },
    // Stage 2.5: a review text with a rule the ledger suggestions must find.
    "rules-review": {
      type: "string",
      default: path.join(REPO, "tests/fixtures/fake_review_rules.txt"),
    },
    // Prefix of the feature ids and of the marker files (s2, s25, ...).
    // Default: the stage of --features.
    stage: { type: "string" },
    // Provider for the review turns, chosen in the pane's Providers section
    // after the Review tab opened. Only claude is offered for them.
    provider: { type: "string" },
    // Feature ids that depend on an event the run cannot force (judge.mjs
    // CONDITIONS). Repeat the flag or separate ids with commas.
    conditional: { type: "string", multiple: true, default: [] },
    "skip-knob": { type: "boolean", default: false },
    // A PNG of the task pane after every pane step, the tabs before anything
    // is pasted, the Word window once and the pane without a daemon, in
    // <run folder>/shots, with shots/timings.json (how long each step took
    // and what the pane showed meanwhile).
    screenshots: { type: "boolean", default: false },
    // Leave out the second Word run on a compatibility mode 14 copy.
    "skip-mode14": { type: "boolean", default: false },
    // Leave out the Word run on the document of tracked special characters.
    "skip-special": { type: "boolean", default: false },
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

let STAGE;
try {
  STAGE = (opt.stage ?? stageOf(opt.features ?? defaultFeatures("s2"))).toLowerCase();
  // Stage cite has no feature list: its rows and markers are its result.
  if (STAGE !== "cite") {
    if (!/^s\d+$/.test(STAGE))
      throw new Error(`--stage takes s2, s25, s3, s4 or cite, not ${opt.stage}`);
    opt.features ??= defaultFeatures(STAGE);
    featureIds(opt.features);
  }
} catch (e) {
  die(e.message);
}
if (!["s2", "s25", "s3", "s4", "cite"].includes(STAGE))
  die(`there is no flow for stage ${STAGE} (s2, s25, s3, s4, cite)`);
opt.fixture ??= path.join(
  REPO,
  "tests/fixtures",
  { s4: "aux_manuscript.docx", cite: "citation_manuscript.docx" }[STAGE] ??
    "review_manuscript.docx",
);
opt.review ??= path.join(
  REPO,
  "tests/fixtures",
  STAGE === "cite" ? "cite_review.txt" : "fake_review.txt",
);
let AFTER_CLEANUP = [];
try {
  const spec = opt["after-cleanup"] ?? (STAGE === "s3" ? "S3-41" : null);
  if (spec) AFTER_CLEANUP = featureIds(spec);
} catch (e) {
  die(`--after-cleanup: ${e.message}`);
}
const PLAN_DIR = opt["plan-dir"] ?? `/tmp/${STAGE}-run`;
const PLAN_LOG = opt["plan-log"] ?? `/tmp/daemon-${STAGE}.log`;

const stamp = new Date().toISOString().replace(/[-:]/g, "").replace("T", "-").slice(0, 15);
const RUN = path.resolve(opt["run-dir"] ?? path.join(opt["base-dir"], `e2e-${stamp}`));
const TOOLS = path.join(RUN, "tools");
const OK_DIR = path.join(RUN, "ok");
// The copy of the manuscript: review_manuscript.docx (stage 2), paper.docx
// (stage 2.5, whose feature list reads paper.decisions.jsonl and
// paper.ledger.json next to it).
const STEM = docStem(STAGE);
const DOC = path.join(RUN, `${STEM}.docx`);
const REVISION = path.join(RUN, `${STEM}.revision.json`);
// The manuscript as Word rendered it for the line numbers (stage 3).
const RENDER_PDF = path.join(RUN, `${STEM}.render.pdf`);
const LOG = opt.log ?? `/tmp/daemon-${STAGE}-e2e.log`;
const KNOB_LOG = LOG.replace(/\.log$/, "") + "-knob.log";
const MODE14_LOG = LOG.replace(/\.log$/, "") + "-mode14.log";
const MODE14_DIR = path.join(RUN, "mode14");
const DOC14 = path.join(MODE14_DIR, "review_manuscript.docx");
// Stage cite: the copies of the document Word saved, and the two Word wrote
// with every change accepted and rejected.
const CITE_DIR = path.join(RUN, "cite");
const SPECIAL_DIR = path.join(RUN, "special");
const DOC_SPECIAL = path.join(SPECIAL_DIR, "tracked_special.docx");
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

// --screenshots: the PNG of each pane step and the record of its wait.
const SHOTS = path.join(RUN, "shots");
const timings = [];
let shotNo = 0;
let paneTag = "";
let pane = (step, args = {}) => {
  const argsFile = path.join(RUN, `args-${paneTag}${step}.json`);
  fs.writeFileSync(argsFile, JSON.stringify(args));
  const shot = opt.screenshots
    ? path.join(SHOTS, shotName(++shotNo, `${paneTag}${step}`, args))
    : null;
  if (shot) fs.mkdirSync(SHOTS, { recursive: true });
  const t0 = Date.now();
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
          ...(shot ? { E2E_SHOT: winPath(shot) } : {}),
          WSLENV: [process.env.WSLENV, "PLAYWRIGHT_CORE", "E2E_CDP_PORT", shot && "E2E_SHOT"]
            .filter(Boolean)
            .join(":"),
        },
      },
    );
  } catch (e) {
    out = e.stdout || JSON.stringify({ ok: false, error: e.message });
  }
  const line = String(out).replace(/\r/g, "").trim().split("\n").filter(Boolean).pop();
  const result = JSON.parse(line);
  if (shot) {
    timings.push(
      stepTiming({
        n: shotNo,
        step: `${paneTag}${step}`,
        args,
        shot: path.basename(shot),
        seconds: Math.round((Date.now() - t0) / 100) / 10,
        result,
      }),
    );
    fs.writeFileSync(path.join(SHOTS, "timings.json"), JSON.stringify(timings, null, 2));
  }
  return result;
};
// --screenshots only: a tab as it is, nothing pressed. Never a step of the
// run: what it answers is not judged.
const showTab = (tab) => (opt.screenshots ? pane("show-tab", { tab }) : null);
// --screenshots only: the whole Word window of this run (pane beside the
// document), drawn by the window itself (PrintWindow), so nothing of another
// window is in the picture.
function wordWindowShot(name) {
  if (!opt.screenshots) return;
  const file = path.join(SHOTS, name);
  let got;
  try {
    got = powershell("word-shot.ps1", [
      "-ProcessId",
      String(record.word.launch.pid),
      "-Out",
      winPath(file),
    ]);
  } catch (e) {
    got = { ok: false, error: String(e?.message ?? e).split("\n")[0] };
  }
  // A picture whose pane came out black is not kept (word-shot.ps1).
  record.word_window_shot = { file: got.ok ? name : null, ...got };
  save();
}

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
// The resident daemon of this machine (scripts/install), stopped in setup()
// while the run has its own on the same ports. It is started again when this
// process ends, whichever way: the end of the run, process.exit, an error
// nobody caught, Ctrl-C.
let resident = { paused: false, resume: () => {} };
process.on("exit", () => resident.resume());
for (const [signal, code] of [
  ["SIGINT", 130],
  ["SIGTERM", 143],
]) {
  process.on(signal, () => {
    // The run's own daemon first, and until its ports are free: a resident
    // unit started while they are held fails and, after a few tries, stays
    // down.
    stopDaemon()
      .catch((e) => console.error(`run-acceptance: ${e.message}`))
      .finally(() => process.exit(code));
  });
}
// append: keep the log and write on (a restart in the middle of a run).
async function startDaemon(logFile, extraEnv = {}, { append = false } = {}) {
  const from = append ? readLog(logFile).length : 0;
  const fd = fs.openSync(logFile, append ? "a" : "w");
  daemon = spawn(process.execPath, ["daemon/index.mjs"], {
    cwd: REPO,
    env: { ...process.env, CC_OFFICE_DEBUG: "1", ...extraEnv },
    stdio: ["ignore", fd, fd],
  });
  daemonLog = logFile;
  await until(
    () => readLog(logFile).slice(from).includes("HTTP server listening") && listening(WS_PORT),
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
function runFeatures(ids = featureIds(opt.features)) {
  if (!opt["feature-list"]) return;
  const list = JSON.parse(fs.readFileSync(opt["feature-list"], "utf8"));
  const features = Array.isArray(list) ? list : list.features;
  for (const id of ids) {
    const f = features.find((x) => x.id === id);
    if (!f?.verify_command) continue;
    const cmd = rewriteVerify(f.verify_command, {
      runDir: RUN,
      prefix: STAGE,
      planDir: PLAN_DIR,
      planRepo: opt["plan-repo"],
      repo: REPO,
      planLog: PLAN_LOG,
      log: LOG,
      knobLog: KNOB_LOG,
    });
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
// S2-23 without a model: an edit that puts "novel" into the paragraph of R1.3
// is sent to the guarded write of the daemon (review-tools.mjs
// applyTrackedEdit) and must be refused by L2 before Word. The only way from
// the pane to that function without a turn is "Allow exception" of a blocked
// edit (review_override), which sends the stored edit once more and lifts
// only the rules the item names. So an item is put on file for R1.3 that
// names L1, not L2 (the store's own addBlocked, written while no request is
// running), the pane presses its button, and L2 has to refuse it: the same
// check and the same log line as for an edit of a turn. The item is taken
// off the file again; the blocked edit the daemon recorded itself stays.
async function forbidGuard() {
  // A refresh first: what the last turn left to write is written before the
  // run writes the file itself.
  note("refresh", pane("refresh"), { label: "refresh (before the forbid guard item)" });
  const seeded = addBlocked(await loadRevision(REVISION, winPath(DOC)), "R1.3", {
    args: FORBID_GUARD_EDIT,
    blocked_by: [FORBID_GUARD_SEED_RULE],
  });
  await saveRevision(REVISION, seeded.revision);
  const seedId = seeded.blocked.id;
  // The pane reads the file again; the item has to be on the card of R1.3
  // before its button is pressed.
  const shown = note("refresh", pane("refresh"), { label: "refresh (forbid guard item on file)" });
  const onCard = (shown.comments ?? [])
    .find((c) => c.ref === "R1.3")
    ?.blocked.some((b) => b.startsWith(`Blocked by ${FORBID_GUARD_SEED_RULE}:`));
  const wordBefore = com("state");
  const at = readLog(daemonLog).length;
  const step = onCard
    ? pane("allow-exception", { ref: "R1.3", blocked_by: FORBID_GUARD_SEED_RULE })
    : { ok: false, error: "the seeded item is not on the card of R1.3 after the refresh" };
  const evidence = forbidBlockEvidence({
    step,
    log: readLog(daemonLog).slice(at),
    revision: JSON.parse(fs.readFileSync(REVISION, "utf8")),
    before: seeded.revision,
    word: com("state"),
    wordBefore,
    seedId,
  });
  marker("s2-forbid-block", evidence);
  record.forbid_guard = { seed: seedId, step, evidence };
  rows.push({
    step: "forbid guard: an edit with novel sent to R1.3",
    result: evidence.pass ? "PASS" : "FAIL",
    got: `${Object.entries(evidence.checks)
      .map(([k, v]) => `${k} ${v}`)
      .join("; ")} | ${evidence.message ?? "-"}`,
  });
  const cleaned = dropBlocked(await loadRevision(REVISION, winPath(DOC)), "R1.3", seedId);
  if (cleaned) await saveRevision(REVISION, cleaned);
  save();
}

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

// Both flows: the run folder, the daemon, a Word of its own on the copy.
async function setup() {
  fs.mkdirSync(TOOLS, { recursive: true });
  for (const f of ["pane-driver.cjs", "word-com.ps1", "word-launch.ps1", "word-shot.ps1"])
    fs.copyFileSync(path.join(HERE, f), path.join(TOOLS, f));
  // Stage cite: a Word of its own for accept-all and reject-all.
  if (STAGE === "cite")
    for (const f of ["word-own.ps1", "marked-diff-word.ps1"])
      fs.copyFileSync(path.join(HERE, f), path.join(TOOLS, f));
  // The installed daemon holds both ports: it makes way for the run.
  resident = pauseResident();
  if (resident.paused)
    await until(() => !listening(WS_PORT) && !listening(HTTP_PORT), 10_000, "ports free");
  if (listening(WS_PORT) || listening(HTTP_PORT))
    die(`port ${WS_PORT} or ${HTTP_PORT} is in use; stop that daemon first`);
  fs.rmSync(KNOB_LOG, { force: true });
  // A reused run folder (--run-dir): markers and evidence of an earlier run
  // are removed, so only what this run judges can satisfy a verify command.
  if (fs.existsSync(OK_DIR)) {
    record.stale_markers_removed = staleMarkers(fs.readdirSync(OK_DIR), STAGE);
    for (const n of record.stale_markers_removed) fs.unlinkSync(path.join(OK_DIR, n));
  }

  record.word.preexisting = wordPids();

  embedTaskpane(opt.fixture, DOC);
  // Stage 4: the journal profile whose format rules find what the fixture
  // plants, next to the manuscript, and named as the default of this daemon.
  const env = {};
  if (STAGE === "s4") {
    const profile = JSON.parse(
      execFileSync("python3", [path.join(HERE, "make_aux_fixture.py"), "--profile"], {
        encoding: "utf8",
      }),
    );
    fs.mkdirSync(path.join(RUN, "journal_profiles"), { recursive: true });
    fs.writeFileSync(
      path.join(RUN, "journal_profiles", `${profile.name}.json`),
      JSON.stringify(profile, null, 2),
    );
    env.CC_OFFICE_PROFILE = profile.name;
  }
  await startDaemon(LOG, env);

  record.word.launch = launchWord(DOC, path.join(RUN, "webview2"));
  save();
}

// Stage 2 (procedure of the stage-2 plan, S2-18..S2-34).
async function main() {
  // --screenshots: the pane as it opens on a document nothing was done with,
  // and every tab before anything is pasted or pressed.
  for (const tab of ["chat", "review", "response", "tools", "library", "setup"]) showTab(tab);
  note("open-review", pane("open-review"), { ctx: { revisionPath: REVISION } });
  // A provider that cannot run review turns is refused here with its reason
  // (the step fails); the turns below then run with Claude.
  if (opt.provider) {
    note("select-provider", pane("select-provider", { role: "turn", id: opt.provider }), {
      label: `select-provider turn=${opt.provider}`,
      ctx: { id: opt.provider },
    });
    pane("open-review");
  }
  note("split", pane("split", { review_file: winPath(opt.review) }));
  note("edit-draft", pane("edit-draft", { requirements: { "R1.3": R13_REQUIREMENTS } }));
  note("ledger-add", pane("ledger-add", { ledger_file: winPath(opt.ledger) }));
  note("check-all", pane("check-all"));
  note("preview", pane("preview", { ref: "R1.3" }), { label: "preview R1.3" });
  for (const ref of ["R1.1", "R1.2", "R1.3", "R1.4", "R1.5"]) {
    await propose(ref);
    // --screenshots: once a first tracked change is in, the Word window.
    if (ref === "R1.1") wordWindowShot("word-window-after-R1.1.png");
  }
  if (opt.screenshots) {
    // After the turns: the chat, the three model rows of Setup, and what
    // "Check revised paragraphs" answers. None of it is judged.
    showTab("chat");
    showTab("setup");
    showTab("review");
    pane("check-paragraphs");
  }

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
  await forbidGuard();
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
    `${after.message} ${after.message_details ?? ""}`,
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

// Stage 2.5 (S25-30..S25-41). Every feature gets its evidence file under
// ok/, and its .ok marker only when the evidence passes; nothing here is a
// person's judgment. Review turns are run with Claude: Codex is used for
// judgments only and no step chooses it for a review turn.
async function mainS25() {
  const revisionNow = () => JSON.parse(fs.readFileSync(REVISION, "utf8"));
  const logAt = () => readLog(LOG).length;
  const logFrom = (at) => readLog(LOG).slice(at);
  note("open-review", pane("open-review"), { ctx: { revisionPath: REVISION } });

  // S25-30: the model menu holds what the account lists.
  const models = note("models", pane("models"));
  marker("s25-models", modelsEvidence({ step: models, log: readLog(LOG) }));

  // S25-31: which providers can be chosen, and why the others cannot.
  marker("s25-providers", providersEvidence(note("providers", pane("providers"))));

  // S25-32: judgments by Ollama, the review text from a file, then the split.
  const judgeBy = opt["judgment-provider"] ?? "ollama";
  const picked = note(
    "select-provider",
    pane("select-provider", { role: "oneshot", id: judgeBy }),
    {
      label: `select-provider oneshot=${judgeBy}`,
      ctx: { id: judgeBy },
    },
  );
  pane("open-review");
  note("load-review-file", pane("load-review-file", { path: winPath(opt.review) }));
  const split = note("split", pane("split"));
  marker(
    "s25-ollama-split",
    ollamaSplitEvidence({
      select: picked,
      split,
      decisions: readLog(path.join(RUN, `${STEM}.decisions.jsonl`)),
    }),
  );

  // The stored judgment provider across a daemon restart: the daemon comes
  // back with no choice for this pane, the pane reconnects and Split is
  // clicked at once. Its judgments must run on the stored provider; a
  // judgment by Claude here would be text sent where the user did not choose.
  await stopDaemon();
  const restartAt = logAt();
  const decisionsFile = path.join(RUN, `${STEM}.decisions.jsonl`);
  const receiptsAt = readLog(decisionsFile).length;
  await startDaemon(LOG, {}, { append: true });
  await until(
    () => logFrom(restartAt).includes(`active doc: ${winPath(DOC)}`),
    120_000,
    "pane reconnect",
  );
  const again = note("split", pane("split"), { label: "split right after a daemon restart" });
  const restart = restartSplitEvidence({
    provider: judgeBy,
    split: again,
    log: logFrom(restartAt),
    decisions: readLog(decisionsFile).slice(receiptsAt),
  });
  marker("s25-restart-split", restart);
  rows.push({
    step: "no judgment by another provider after the restart",
    result: restart.pass ? "PASS" : "FAIL",
    got: `restored=${restart.restored} judgments=${JSON.stringify(restart.judgments)} receipts=${restart.receipt_models.join(",")}`,
  });
  note("edit-draft", pane("edit-draft", { requirements: { "R1.3": R13_REQUIREMENTS } }));

  // S25-39: the ledger from a file, out to <stem>.ledger.json, and in again.
  note("import-ledger", pane("import-ledger", { path: winPath(opt.ledger) }), {
    label: "import-ledger (fixture)",
    ctx: { added: 3, skipped: 0 },
  });
  note("export-ledger", pane("export-ledger"), { ctx: { count: 3 } });
  const exported = path.join(RUN, `${STEM}.ledger.json`);
  if (fs.existsSync(exported)) {
    note("import-ledger", pane("import-ledger", { path: winPath(exported) }), {
      label: "import-ledger (the exported file again)",
      ctx: { added: 0, skipped: 3 },
    });
  } else {
    rows.push({
      step: "import-ledger (the exported file again)",
      result: "FAIL",
      got: `no ${exported}`,
    });
  }

  // S25-40: a rule the review text asks for, found and added on a tick.
  const rule = "cell-free extract";
  const suggested = note(
    "suggest-ledger",
    pane("suggest-ledger", { review_file: winPath(opt["rules-review"]), pick: [rule] }),
  );
  marker("s25-ledger-suggest", suggestEvidence({ step: suggested, pick: rule, before: 3 }));

  // S25-37: every comment that is not done, one turn after another.
  const at = logAt();
  const all = pane("propose-all");
  note("propose-all", all, { ctx: { log: logFrom(at) } });
  marker("s25-propose-all", proposeAllEvidence({ message: all.message, log: logFrom(at) }));

  // S25-38: a hand edit in Word, then the pane gets the focus back.
  const del = com("delete-text", ["-Text", OUTSIDE_SENTENCE]);
  record.word.delete_text = del;
  rows.push({
    step: "word: delete outside sentence",
    result: del.ok ? "PASS" : "FAIL",
    got: del.removed ?? del.error,
  });
  marker("s25-focus-refresh", focusEvidence(note("focus-refresh", pane("focus-refresh"))));
}

// Stage 3 (S3-30..S3-41): the response letter. The review part is the stage 2
// procedure up to the author's decisions in Word, with S3_DECISIONS and
// without the steps that leave something undecided (the author's own pending
// insertion, the knob run): a pending tracked change is a warning of the
// letter check, and S3-34 asks for none. Every feature gets its evidence file
// under ok/, and its .ok marker only when the evidence passes.
async function mainS3() {
  const revisionNow = () => JSON.parse(fs.readFileSync(REVISION, "utf8"));
  const logAt = () => readLog(LOG).length;
  const logFrom = (at) => readLog(LOG).slice(at);
  const row = (step, pass, got, skip = false) =>
    rows.push({ step, result: pass ? "PASS" : skip ? "SKIP" : "FAIL", got });
  // A step that only reads, or computes again what a second run computes the
  // same, is run once more when it fails: one slow answer of Word or of the
  // daemon is not a failed run. A step that writes or asks a model is not.
  const paneTry = (step, args = {}) => {
    const first = pane(step, args);
    if (first.ok !== false && !first.message_error) return first;
    (record.retries ??= []).push({ step, first });
    save();
    return pane(step, args);
  };
  // The citations with a line range, as revision.json holds them.
  const cits = () =>
    revisionNow().comments.flatMap((c) =>
      c.declined
        ? []
        : (c.response?.citations ?? []).flatMap((x, i) =>
            x.kind === "revised" && x.lines
              ? [
                  {
                    id: `${c.ref}#${i}`,
                    ref: c.ref,
                    quote: x.quote,
                    paragraph_index: x.paragraph_index,
                    ...x.lines,
                  },
                ]
              : [],
          ),
    );
  // The second reader of the rendered PDF (poppler), independent of PyMuPDF.
  const layout = () =>
    execFileSync("pdftotext", ["-layout", RENDER_PDF, "-"], {
      encoding: "utf8",
      maxBuffer: 64 << 20,
    });
  const letters = () =>
    fs
      .readdirSync(RUN)
      .filter((n) => n.startsWith(`${STEM}.response`) && n.endsWith(".docx"))
      .sort()
      .map((n) => `${n}:${fs.statSync(path.join(RUN, n)).mtimeMs}`);

  // ---- the review, as stage 2 ----
  note("open-review", pane("open-review"), { ctx: { revisionPath: REVISION } });
  if (opt.provider) {
    note("select-provider", pane("select-provider", { role: "turn", id: opt.provider }), {
      label: `select-provider turn=${opt.provider}`,
      ctx: { id: opt.provider },
    });
    pane("open-review");
  }
  // The provider of the judgments also writes the drafts of the letter.
  // Without --judgment-provider the pane's own choice stays (Claude).
  if (opt["judgment-provider"]) {
    const id = opt["judgment-provider"];
    note("select-provider", pane("select-provider", { role: "oneshot", id }), {
      label: `select-provider oneshot=${id}`,
      ctx: { id },
    });
    pane("open-review");
  }
  note("split", pane("split", { review_file: winPath(opt.review) }));
  note("edit-draft", pane("edit-draft", { requirements: { "R1.3": R13_REQUIREMENTS } }));
  note("ledger-add", pane("ledger-add", { ledger_file: winPath(opt.ledger) }));
  for (const ref of ["R1.1", "R1.2", "R1.3", "R1.4", "R1.5"]) await propose(ref);
  note("refresh", pane("refresh"), { label: "refresh before Word" });

  const plan = buildComPlan(revisionNow(), S3_DECISIONS);
  const planFile = path.join(RUN, "com-plan.json");
  fs.writeFileSync(planFile, JSON.stringify(plan, null, 2));
  const applied = com("apply", ["-PlanJson", winPath(planFile)]);
  record.word.apply = applied;
  row(
    "word: accept/reject/edit",
    applied.ok,
    (applied.results ?? [])
      .map((x) => `${x.change_id} ${x.op}${x.error ? ` ERR ${x.error}` : ""}`)
      .join("; "),
  );
  if (!applied.ok && opt["keep-on-fail"]) throw new Error("Word COM apply failed");
  const afterWord = note("refresh", pane("refresh"), {
    label: "refresh after Word",
    judge: "none",
  });
  const badges =
    afterWord.ok === false ? { pass: false, got: afterWord.error } : badgesMatch(afterWord);
  row("word decisions shown on the Review tab", badges.pass, badges.got);

  // ---- S3-30: the Response tab lists every comment as the reviewer wrote it ----
  const opened = note("open-response", paneTry("open-response"));
  marker("s3-open", openEvidence(opened, revisionNow()));

  // ---- S3-31: an existing letter is read, proposed, then applied ----
  const letter = note("load-letter", pane("load-letter", { path: winPath(opt.letter) }));
  const beforeApply = revisionNow().profile_overrides ?? {};
  note("apply-proposal", pane("apply-proposal"));
  marker(
    "s3-proposal",
    proposalEvidence({
      rows: letter.rows,
      before: beforeApply,
      after: revisionNow().profile_overrides,
    }),
  );

  // ---- S3-32: declines with reasons, then a draft for every other comment ----
  note(
    "response-decline",
    pane("response-decline", {
      comments: [{ ref: "R1.5", reason: "No antifoam agent was used in this study." }],
      open_requests: true,
      request_reason: "This wording was ruled out in the previous round of review.",
    }),
    { ctx: { comments: ["R1.5"] } },
  );
  note("response-draft-all", pane("response-draft-all"));
  const r14 = revisionNow().comments.find((c) => c.ref === "R1.4");
  const r14Kinds = (r14?.response?.citations ?? []).map((x) => x.kind);
  row(
    "R1.4 cites the removal",
    r14Kinds.includes("removed"),
    `citations: ${r14Kinds.join(",") || "-"}`,
    true,
  );

  // ---- S3-33: line ranges, and a second reader of the same PDF ----
  note("response-locate", paneTry("response-locate"));
  marker("s3-lines", secondReader(layout(), cits()));

  // ---- S3-34: the checks pass ----
  const checked = note("response-check", paneTry("response-check"));
  marker("s3-check", {
    pass: JUDGES["response-check"](checked).pass,
    summary: checked.summary,
    checks: checked.checks,
  });

  // ---- S3-35: the letter, in the colours of the profile ----
  const exported = note("response-export", pane("response-export"));
  const docx = path.join(RUN, `${STEM}.response.docx`);
  let read = null;
  try {
    read = JSON.parse(
      execFileSync("python3", [path.join(REPO, "tools/response_read.py"), "--docx", docx], {
        encoding: "utf8",
        maxBuffer: 64 << 20,
      }),
    );
  } catch (e) {
    record.response_read_error = String(e?.message ?? e);
  }
  marker("s3-export", exportEvidence(exported, read, docx, fs.existsSync(docx)));

  // ---- S3-36: one more line in the manuscript ----
  const before = cits();
  const target = insertTarget(before);
  const INSERTED = "One short line added for the line shift test.";
  if (target == null) {
    marker("s3-line-shift", {
      pass: false,
      error: "no two cited paragraphs to insert between",
      before,
    });
    row("word: insert a paragraph", false, "no two cited paragraphs to insert between");
  } else {
    // Snapshot paragraph n is paragraph n + 1 of Word.
    const inserted = com("insert-paragraph", ["-Text", INSERTED, "-Paragraph", String(target + 1)]);
    record.word.insert = inserted;
    row(
      "word: insert a paragraph",
      inserted.ok,
      inserted.ok ? `before: ${inserted.next_text}` : inserted.error,
    );
    // The Review tab's refresh marks the answers stale; its button is only
    // there while that tab is shown.
    pane("open-review");
    note("refresh", pane("refresh"), { label: "refresh after the insertion", judge: "none" });
    const staleView = note("open-response", paneTry("open-response"), {
      label: "open-response after the insertion",
    });
    note("response-locate", paneTry("response-locate"), {
      label: "response-locate after the insertion",
    });
    const text = layout();
    const at = Number(/^\s{0,3}(\d+)\s+One short line added/m.exec(text)?.[1]);
    const shift = lineShift(
      before,
      cits(),
      at,
      (staleView.cards ?? []).some((c) => /Line numbers are older/.test(c.stale)),
    );
    const second = secondReader(text, cits());
    marker("s3-line-shift", { ...shift, inserted_at: at, second, pass: shift.pass && second.pass });
  }

  // ---- S3-37: a quoted sentence edited in Word ----
  const first = [...cits()].sort((a, b) => a.start - b.start)[0];
  const words = String(first?.quote ?? "").split(/\s+/);
  if (words.length < 5) {
    marker("s3-stale", { pass: false, error: "no cited sentence of five words or more" });
    row("word: edit a quoted sentence", false, "no cited sentence of five words or more");
  } else {
    // One word more after the first three words of the quoted sentence.
    const from = words.slice(0, 4).join(" ");
    const to = [...words.slice(0, 3), "indeed", words[3]].join(" ");
    const replaced = com("replace-text", ["-Text", from, "-With", to]);
    record.word.replace = replaced;
    row(
      "word: edit a quoted sentence",
      replaced.ok,
      replaced.ok ? `${from} -> ${to}` : replaced.error,
    );
    const at = logAt();
    pane("open-review");
    note("refresh", pane("refresh"), { label: "refresh after the edit", judge: "none" });
    paneTry("open-response");
    const failing = note("response-check", pane("response-check"), {
      label: "response-check with a stale answer",
      judge: "none",
    });
    // The author reads the answer and keeps it: saving it again clears the mark.
    note("response-edit", pane("response-edit", { ref: first.ref }), {
      label: `response-edit ${first.ref} (kept)`,
    });
    note("response-locate", paneTry("response-locate"), {
      label: "response-locate after the edit",
    });
    const passing = note("response-check", paneTry("response-check"), {
      label: "response-check after the answer was saved again",
      judge: "response-check-no-fail",
    });
    marker("s3-stale", {
      ...staleEvidence(failing, passing, first.ref),
      ref: first.ref,
      log_stale: (new RegExp(
        `^\\[response\\] stale ${first.ref.replace(".", "\\.")} .*$`,
        "m",
      ).exec(logFrom(at)) ?? [null])[0],
      checks_before: failing.checks,
    });
  }

  // ---- S3-38: a forbidden expression of the response scope ----
  const ledgerFile = path.join(RUN, "s3-response-ledger.json");
  fs.writeFileSync(
    ledgerFile,
    JSON.stringify({
      items: [
        {
          kind: "forbid",
          text: "Do not write unfortunately in the response letter",
          scope: "response",
          check: { type: "regex", pattern: "\\bunfortunately\\b", flags: "i", expect: "absent" },
        },
      ],
    }),
  );
  pane("open-review");
  note("ledger-add", pane("ledger-add", { ledger_file: winPath(ledgerFile) }), {
    label: "ledger-add (response scope)",
    judge: "none",
  });
  paneTry("open-response");
  const ref = first?.ref ?? "R1.1";
  const edited = note("response-edit", pane("response-edit", { ref, prepend: "Unfortunately, " }), {
    label: `response-edit ${ref} (forbidden word)`,
  });
  const lettersBefore = letters();
  // Not checked first: the button still reads "Export letter (.docx)", the
  // daemon finds the failure and holds the letter back.
  const held = note("response-export", pane("response-export"), {
    label: "response-export while a check fails",
    judge: "none",
  });
  const exportedWhileFailing = JSON.stringify(letters()) !== JSON.stringify(lettersBefore);
  note("response-edit", pane("response-edit", { ref, text: edited.previous_text }), {
    label: `response-edit ${ref} (put back)`,
  });
  const afterLedger = note("response-check", paneTry("response-check"), {
    label: "response-check after the word was removed",
    judge: "response-check-no-fail",
  });
  marker("s3-ledger", {
    ...ledgerEvidence(held, exportedWhileFailing, afterLedger),
    checks: held.checks,
    message: held.message,
  });

  // ---- S3-40: plain text; the letter opens in Word; an export meanwhile ----
  note("response-export-text", pane("response-export", { format: "text" }));
  const modelFile = path.join(RUN, `${STEM}.response.model.json`);
  const modelParagraphs = fs.existsSync(modelFile)
    ? JSON.parse(fs.readFileSync(modelFile, "utf8")).paragraphs.length
    : null;
  const openCheck = fs.existsSync(docx)
    ? com("open-check", ["-Other", winPath(docx)])
    : { ok: false, error: `no ${docx}` };
  record.word.open_check = openCheck;
  row(
    "word: the letter opens without repair",
    openCheck.ok === true,
    openCheck.ok ? `paragraphs ${openCheck.paragraphs}` : openCheck.error,
  );
  const again = note("response-export", pane("response-export"), {
    label: "response-export with the first letter open in Word",
  });
  const secondPath =
    /^Exported under a new name: (.+?) \(/.exec(again.output ?? "")?.[1] ??
    /^Exported: (.+)$/.exec(again.output ?? "")?.[1] ??
    null;
  if (openCheck.ok) record.word.close_other = com("close-other", ["-Other", winPath(docx)]);
  marker("s3-open-check", openCheckEvidence(openCheck, modelParagraphs, docx, secondPath));

  // ---- S3-39: the languages of the profile reach both prompts ----
  note(
    "response-set-lang",
    pane("response-set-lang", { lang: "German", rationale_lang: "German" }),
  );
  const at = logAt();
  note("response-draft", pane("response-draft", { ref: "R1.1" }), {
    label: "response-draft R1.1 (German)",
    ctx: { ref: "R1.1" },
  });
  const draftLine = (logFrom(at).match(/^\[response\] draft R1\.1 .*$/gm) ?? []).at(-1);
  pane("open-review");
  const preview = note("preview", pane("preview", { ref: "R1.1" }), {
    label: "preview R1.1 (German)",
    judge: "none",
  });
  marker("s3-lang", {
    ...langEvidence(draftLine, preview.prompt, "German"),
    log_line: draftLine ?? null,
  });
  paneTry("open-response");
  note("response-set-lang", pane("response-set-lang", { lang: "", rationale_lang: "" }), {
    label: "response-set-lang (the profile's own again)",
  });
}

// Stage 4 (S4-30..S4-33, S4-37, S4-38): the Tools tab on the fake manuscript
// of make_aux_fixture.py. Nothing here writes to the document: the pane's
// snapshot is compared around the selection requests and around the checks.
// Every feature gets its evidence file under ok/, and its .ok marker only
// when the evidence passes.
async function mainS4() {
  const ledgerNow = () =>
    fs.existsSync(REVISION) ? (JSON.parse(fs.readFileSync(REVISION, "utf8")).ledger ?? []) : [];
  const logAt = () => readLog(LOG).length;
  const logFrom = (at) => readLog(LOG).slice(at);
  const paragraphs = () => pane("snapshot").paragraphs ?? [];
  const python = (args) => {
    try {
      return {
        exit: 0,
        out: execFileSync("python3", args, { encoding: "utf8", maxBuffer: 64 << 20 }),
      };
    } catch (e) {
      return { exit: e.status ?? 1, out: String(e.stdout ?? "") };
    }
  };

  note("open-tools", pane("open-tools"));
  if (opt["judgment-provider"]) {
    const id = opt["judgment-provider"];
    note("select-provider", pane("select-provider", { role: "oneshot", id }), {
      label: `select-provider oneshot=${id}`,
      ctx: { id },
    });
    pane("open-tools");
  }

  // ---- S4-33: translate the selection and ask about it; the document stays ----
  const before = paragraphs();
  const target = before.find((p) => /model organism/.test(p.text))?.index ?? 0;
  const at = logAt();
  const translate = note(
    "tools-translate",
    pane("tools-translate", { paragraph: target, lang: opt["translate-lang"] }),
  );
  const ask = note(
    "tools-ask",
    pane("tools-ask", { paragraph: target, question: "Which organism is this paragraph about?" }),
  );
  const afterSelection = paragraphs();
  marker(
    "s4-selection",
    selectionEvidence({ translate, ask, before, after: afterSelection, log: logFrom(at) }),
  );

  // ---- S4-30: the guideline gives cited candidates; one is adopted ----
  const pick = "abstract must not exceed";
  const ledgerBefore = ledgerNow();
  const extract = note(
    "tools-extract",
    pane("tools-extract", { guideline_file: winPath(opt.guideline) }),
  );
  const add = note("tools-add", pane("tools-add", { keep: pick }));
  const adopted = guidelineEvidence({
    extract,
    add,
    guideline: fs.readFileSync(opt.guideline, "utf8"),
    ledgerBefore,
    ledgerAfter: ledgerNow(),
    pick,
  });
  marker("s4-profile", adopted);

  // ---- S4-31, S4-32: the checks find what the fixture plants, and only report ----
  const check = note("tools-check", pane("tools-check"));
  marker("s4-count", countEvidence({ check, rule: adopted.added[0] }));
  marker("s4-format", formatEvidence({ check, before: afterSelection, after: paragraphs() }));

  // ---- S4-37: the button writes what the script writes, in both modes ----
  const modes = [];
  for (const mode of ["track_changes", "color"]) {
    const step = note("tools-marked", pane("tools-marked", { path: winPath(opt.baseline), mode }), {
      label: `tools-marked ${mode}`,
    });
    const name = /^Written: (\S+\.docx)/.exec(step.marked?.[0] ?? "")?.[1] ?? null;
    const byScript = path.join(RUN, `script_marked_${mode}.docx`);
    const script = python([
      path.join(REPO, "tools/build_marked_diff.py"),
      "--baseline",
      opt.baseline,
      "--clean",
      DOC,
      "--out",
      byScript,
      "--mode",
      mode,
    ]);
    const equal =
      !!name &&
      script.exit === 0 &&
      fs.existsSync(path.join(RUN, name)) &&
      fs.readFileSync(path.join(RUN, name)).equals(fs.readFileSync(byScript));
    modes.push({ mode, name, script_exit: script.exit, equal });
  }
  marker("s4-marked", markedEvidence(modes));

  // ---- S4-38: the anonymized copy, checked again by the runner ----
  const anonymized = note("tools-anonymize", pane("tools-anonymize"));
  const copy = path.join(RUN, `${STEM}_anonymized.docx`);
  let verifier = null;
  if (fs.existsSync(copy)) {
    const run = python([
      path.join(REPO, "scripts/verify-anonymized.py"),
      "--before",
      DOC,
      "--after",
      copy,
      "--json",
    ]);
    try {
      verifier = JSON.parse(run.out.trim().split("\n").pop());
    } catch (e) {
      record.verifier_error = String(e?.message ?? e);
    }
  }
  marker(
    "s4-anonymize",
    anonymizeEvidence({ step: anonymized, verifier, exists: fs.existsSync(copy) }),
  );
}

// ---- stage cite ------------------------------------------------------------
// Do tracked edits leave the fields of Zotero and EndNote whole? The document
// is the fake manuscript of make_citation_fixture.py. What a tool answers
// (applied, refused and why) is recorded; what is judged is what
// tools/citation_fields_check.py finds in the files Word saved: after every
// edit of the run, after every model turn, and at the end with the changes
// tracked, all accepted and all rejected.

const firstLine = (e) => String(e?.message ?? e).split("\n")[0];
// A row of the stage that is no pane step: in the table of the run and,
// like the steps, in result.json.
const citeRow = (row) => {
  rows.push(row);
  record.steps.push({ ...row });
  console.log(`${row.result}  ${row.step}  ${String(row.got ?? "").slice(0, 400)}`);
  save();
};

// The checker on two files. Its lines (the paragraph, the kind of the field
// and its markers before and after) go to the screen, its result into
// result.json under cite.checks[name].
function fieldsCheck(name, before, after, extra = []) {
  let exit = 0;
  let out = "";
  try {
    out = execFileSync(
      "python3",
      [
        path.join(REPO, "tools/citation_fields_check.py"),
        ...["--before", before, "--after", after, "--detail", ...extra],
      ],
      { encoding: "utf8", maxBuffer: 64 << 20 },
    );
  } catch (e) {
    exit = e.status ?? 1;
    out = String(e.stdout ?? "");
  }
  const lines = out.trim().split("\n").filter(Boolean);
  let result;
  try {
    result = JSON.parse(lines.pop());
  } catch (e) {
    result = { ok: false, failures: ["no-result"], error: firstLine(e) };
  }
  for (const l of lines) console.log(`      ${name}: ${l}`);
  record.cite.checks[name] = { before, after, exit, lines, result };
  save();
  return result;
}

// The document as Word holds it now: saved in place, then copied to
// cite/<name>.docx. null when it could not be saved.
function citeSave(name) {
  let saved;
  try {
    saved = com("save", ["-AllowDir", winPath(RUN)]);
  } catch (e) {
    saved = { ok: false, error: firstLine(e) };
  }
  record.cite.saves[name] = saved;
  if (!saved.ok) return null;
  const copy = path.join(CITE_DIR, `${name}.docx`);
  try {
    fs.copyFileSync(DOC, copy);
  } catch (e) {
    // The row of this check then says that there is no file to check.
    record.cite.saves[name] = { ...saved, ok: false, error: `copy: ${firstLine(e)}` };
    return null;
  }
  return copy;
}

// One edit of the run, through the guarded write of the daemon as
// forbidGuard() sends one: the edit is put on file as a blocked item of
// R1.4 (the store's own addBlocked, while no request runs) that names a rule
// the ledger does not have, and "Allow exception" sends it to
// applyTrackedEdit. No rule refuses it, so it goes to Word as an edit of a
// turn does. prev: the last saved copy; returns the copy saved after it.
async function citeCase(c, prev) {
  const ref = CITE_SEED_REF;
  const entry = { name: c.name, what: c.what, find: c.find, replace: c.replace };
  let outcome;
  try {
    // A refresh first: what is left to write is written before the run
    // writes the file itself.
    note("refresh", pane("refresh"), { label: `refresh (before ${c.name})`, judge: "none" });
    const seeded = addBlocked(await loadRevision(REVISION, winPath(DOC)), ref, {
      args: citeEdit(c),
      blocked_by: [FORBID_GUARD_SEED_RULE],
    });
    await saveRevision(REVISION, seeded.revision);
    const seedId = seeded.blocked.id;
    const shown = pane("refresh");
    const onCard = (shown.comments ?? [])
      .find((x) => x.ref === ref)
      ?.blocked.some((b) => b.startsWith(`Blocked by ${FORBID_GUARD_SEED_RULE}:`));
    const wordBefore = com("state");
    const at = readLog(daemonLog).length;
    const step = onCard
      ? pane("apply-seeded", { ref, blocked_by: FORBID_GUARD_SEED_RULE })
      : { ok: false, error: `the item is not on the card of ${ref} after the refresh` };
    outcome = citeOutcome({
      step,
      log: readLog(daemonLog).slice(at),
      before: seeded.revision,
      revision: JSON.parse(fs.readFileSync(REVISION, "utf8")),
      word: com("state"),
      wordBefore,
      seedId,
      ref,
    });
    // A refused item is back on file: the run takes its own item off.
    const cleaned = dropBlocked(await loadRevision(REVISION, winPath(DOC)), ref, seedId);
    if (cleaned) await saveRevision(REVISION, cleaned);
    Object.assign(entry, { seed: seedId, step });
  } catch (e) {
    entry.error = String(e?.stack ?? e);
    outcome = {
      answer: "unknown",
      message: firstLine(e),
      changes: [],
      word_revisions: [null, null],
      log: [],
    };
  }
  const copy = citeSave(`case-${c.name}`);
  const check = copy ? fieldsCheck(`case-${c.name}`, prev, copy) : null;
  record.cite.cases.push({ ...entry, outcome, check_ok: check?.ok ?? null });
  citeRow(citeCaseRow({ name: c.name, what: c.what, outcome, check }));
  return copy ?? prev;
}

// office_tracked_rewrite on a paragraph with fields. No step of the pane
// reaches that tool without a model, so the call it makes first is built
// here with the daemon's own functions (citeRewriteCall) and handed to the
// pane's office_tracked_edit, which is where a paragraph that is not plain
// text is refused. The daemon's half of the rewrite runs only when a model
// turn below chooses the tool.
function citeRewrite(rw, prev) {
  let call = null;
  let step = null;
  let wordBefore = null;
  let word = null;
  try {
    note("refresh", pane("refresh"), { label: "refresh (before the rewrite)", judge: "none" });
    const snap = pane("snapshot");
    call = snap.ok === false ? { error: snap.error } : citeRewriteCall(snap.paragraphs, rw);
    wordBefore = com("state");
    if (!call.error) step = pane("tracked-edit-direct", { args: call.args });
    word = com("state");
  } catch (e) {
    call = { ...(call ?? {}), error: firstLine(e) };
  }
  const copy = citeSave("case-rewrite");
  const check = copy ? fieldsCheck("case-rewrite", prev, copy) : null;
  record.cite.rewrite = { call, step, check_ok: check?.ok ?? null };
  citeRow(citeRewriteRow({ what: rw.what, call, step, word, wordBefore, check }));
  return copy ?? prev;
}

async function mainCite() {
  fs.mkdirSync(CITE_DIR, { recursive: true });
  record.cite = { checks: {}, saves: {}, cases: [], turns: [] };
  const spec = JSON.parse(
    execFileSync("python3", [path.join(HERE, "make_citation_fixture.py"), "--cases"], {
      encoding: "utf8",
    }),
  );

  note("open-review", pane("open-review"), { ctx: { revisionPath: REVISION } });
  if (opt.provider) {
    note("select-provider", pane("select-provider", { role: "turn", id: opt.provider }), {
      label: `select-provider turn=${opt.provider}`,
      ctx: { id: opt.provider },
    });
    pane("open-review");
  }

  // The plugins of this Word, read only. Nothing is started or refreshed.
  let addins;
  try {
    addins = com("addins");
  } catch (e) {
    addins = { ok: false, error: firstLine(e) };
  }
  const plugins = pluginsEvidence(addins);
  record.cite.plugins = { ...plugins, raw: addins };
  citeRow({
    step: "cite: the Zotero and EndNote plugins of this Word (not judged)",
    result: "NOTE",
    got: plugins.read
      ? `Zotero: ${plugins.zotero.join("; ") || "not found"} | EndNote: ${plugins.endnote.join("; ") || "not found"} | ${plugins.com_addins} COM add-in(s), ${plugins.templates} template(s)`
      : `could not be read (${plugins.errors.join("; ")})`,
  });

  // The document in Word's own form before any edit: every later file is
  // compared with this one, so what Word itself changes when it saves the
  // fixture is not taken for the work of a tool.
  const baseline = citeSave("baseline");
  if (!baseline) throw new Error("the document could not be saved before the edits");
  const own = fieldsCheck("word-save", opt.fixture, baseline);
  citeRow({
    step: "cite: Word's own save keeps the fields of the fixture (not judged)",
    result: own.ok ? "PASS" : "NOTE",
    got: fieldsLine(own),
  });

  note("split", pane("split", { review_file: winPath(opt.review) }), { judge: "cite-split" });
  note("edit-draft", pane("edit-draft"), { judge: "cite-draft" });

  // a. The edits of the run, one at a time, each saved and checked.
  let prev = baseline;
  for (const c of spec.cases) prev = await citeCase(c, prev);
  prev = citeRewrite(spec.rewrite, prev);

  // b. One model turn per comment.
  note("refresh", pane("refresh"), { label: "refresh (before the turns)", judge: "none" });
  for (const ref of CITE_TURN_REFS) {
    const at = readLog(daemonLog).length;
    await propose(ref);
    const log = readLog(daemonLog).slice(at);
    const turn = {
      ref,
      changes: (log.match(/^\[review\] change recorded /gm) ?? []).length,
      rewrites: rewritesFromLog(log, ref),
      refused: refusalsFromLog(log, ref),
      stale: (log.match(/^\[review\] stale target /gm) ?? []).length,
    };
    const copy = citeSave(`turn-${ref}`);
    const check = copy ? fieldsCheck(`turn-${ref}`, prev, copy) : null;
    record.cite.turns.push({ ...turn, check_ok: check?.ok ?? null });
    citeRow({
      step: `cite: fields after the turn of ${ref}`,
      result: check?.ok === true ? "PASS" : "FAIL",
      got: `turn: ${turn.changes} change(s), rewrites ${JSON.stringify(turn.rewrites)}, refused ${JSON.stringify(turn.refused)}, stale ${turn.stale} | ${fieldsLine(check)}`,
    });
    if (copy) prev = copy;
  }

  // c. The saved document with its changes tracked, against the baseline.
  note("refresh", pane("refresh"), { label: "refresh before the save" });
  const trackedCopy = citeSave("tracked");
  const tracked = trackedCopy ? fieldsCheck("tracked", baseline, trackedCopy) : null;
  marker("cite-tracked", { pass: tracked?.ok === true, check: tracked });
  citeRow({
    step: "cite: saved with the changes tracked, every field whole",
    result: tracked?.ok === true ? "PASS" : "FAIL",
    got: fieldsLine(tracked),
  });
  const changed = tracked?.text_changed ?? [];
  citeRow({
    step: "cite: citations whose text differs from what their field records (not judged)",
    result: "NOTE",
    got: changed.length
      ? changed.map((f) => `${f.kind} "${f.recorded}" -> "${f.accepted}"`).join("; ")
      : "none",
  });
  // Whole by the rules, and still lost as a citation once accepted.
  const flat = tracked?.flattened ?? [];
  citeRow({
    step: "cite: citations an edit replaced by plain text (not judged)",
    result: "NOTE",
    got: flat.length ? flat.map((f) => `${f.kind} ${f.shown} in "${f.text}"`).join("; ") : "none",
  });
}

// The saved document of stage cite in a Word of its own (word-own.ps1, the
// Word judge of the marked copies): opened as it is, once with
// Revisions.AcceptAll() and once with Revisions.RejectAll(), each result
// written as a file, and the checker run on both. After the Word of the main
// run is gone and the daemon is stopped, so the pane of these copies reaches
// no daemon.
function citeJudge() {
  const baseline = path.join(CITE_DIR, "baseline.docx");
  const tracked = path.join(CITE_DIR, "tracked.docx");
  const accepted = path.join(CITE_DIR, "accepted.docx");
  const rejected = path.join(CITE_DIR, "rejected.docx");
  const anchor = path.join(CITE_DIR, "anchor.docx");
  const jobsFile = path.join(CITE_DIR, "jobs.json");
  const outFile = path.join(CITE_DIR, "word.json");
  const pidFile = path.join(CITE_DIR, "word-pid.json");
  for (const f of [accepted, rejected, outFile, pidFile]) fs.rmSync(f, { force: true });
  fs.copyFileSync(opt.fixture, anchor);
  fs.writeFileSync(
    jobsFile,
    JSON.stringify([
      { id: "baseline", path: winPath(baseline), action: "none" },
      { id: "tracked", path: winPath(tracked), action: "none" },
      { id: "accept", path: winPath(tracked), action: "accept", save_as: winPath(accepted) },
      { id: "reject", path: winPath(tracked), action: "reject", save_as: winPath(rejected) },
    ]),
  );
  const readJson = (f) =>
    fs.existsSync(f) ? JSON.parse(fs.readFileSync(f, "utf8").replace(/^﻿/, "")) : null;
  let done = null;
  try {
    done = powershell(
      "marked-diff-word.ps1",
      [
        ...["-Jobs", winPath(jobsFile), "-Out", winPath(outFile), "-Anchor", winPath(anchor)],
        ...["-AllowDir", winPath(RUN), "-PidFile", winPath(pidFile)],
      ],
      15 * 60_000,
    );
  } catch (e) {
    record.cite.word_error = firstLine(e);
  }
  if (done && !done.ok) record.cite.word_error = done.error;
  // Its Word: stopped by force only when its quit ran and the process stayed.
  const pid = done?.pid ?? readJson(pidFile)?.pid;
  if (pid) {
    record.word.launch_cite = { pid };
    record.cleanup.quit_cite = done?.quit ?? null;
    let alive = true;
    try {
      alive = wordPids().includes(pid);
    } catch (e) {
      // Not known: taken as alive, so cleanupDecision does not call it gone.
      record.cleanup.cite_pids_error = firstLine(e);
    }
    const d = cleanupDecision({
      pid,
      preexisting: record.word.preexisting ?? [],
      quit: done?.quit,
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
      } catch (e) {
        // cleanup() reports a Word of the run that is still there.
        record.cleanup.cite_stop_error = firstLine(e);
      }
    }
    if (d.action === "leave") {
      (record.cleanup.left_running ??= []).push({ pid, reason: d.reason });
      console.error(`run-acceptance: ${d.reason}`);
    }
    if (d.fail) rows.push({ step: `cleanup: Word ${pid}`, result: "FAIL", got: d.reason });
  }

  const jobs = Object.fromEntries((readJson(outFile)?.jobs ?? []).map((j) => [j.id, j]));
  const trackedCheck = record.cite.checks.tracked?.result ?? null;
  const inDeletions = (trackedCheck?.after?.list ?? []).filter((f) => f.state === "deleted");
  const rej = fs.existsSync(rejected) ? fieldsCheck("rejected", baseline, rejected) : null;
  const acc = fs.existsSync(accepted)
    ? fieldsCheck("accepted", baseline, accepted, ["--expect-removed", String(inDeletions.length)])
    : null;
  const e = citeAcceptEvidence({ tracked: trackedCheck, accepted: acc, rejected: rej, jobs });
  marker("cite-accept-reject", { ...e, word_error: record.cite.word_error ?? null });
  record.cite.accept_reject = e;
  const k = e.checks;
  citeRow({
    step: "cite: Word opens the saved file as it is, accepts and rejects every change",
    result: k.word_opened_all && k.no_revision_left ? "PASS" : "FAIL",
    got:
      [record.cite.word_error, ...e.job_errors].filter(Boolean).join("; ") ||
      `revisions left: accept ${jobs.accept?.revisions_left}, reject ${jobs.reject?.revisions_left}; changed on open: ${e.changed_on_open.join(",") || "none"}`,
  });
  citeRow({
    step: "cite: after reject-all the fields are the original ones",
    result: k.rejected_equals_original && k.rejected_all_live ? "PASS" : "FAIL",
    got: fieldsLine(rej),
  });
  citeRow({
    step: "cite: after accept-all only the fields that lay in deletions are gone, the others whole",
    result:
      k.accepted_whole && k.accepted_all_live && k.accepted_lacks_only_the_deleted
        ? "PASS"
        : "FAIL",
    got: `in deletions: ${e.deleted_fields.join(", ") || "none"}; gone after accept: ${
      e.removed_by_accept.map((f) => `${f.kind} in "${f.text}"`).join("; ") || "none"
    } | ${fieldsLine(acc)}`,
  });
  citeRow({
    step: "cite: Word lists the original fields (code and result) after reject-all",
    result: k.word_lists_the_original_after_reject ? "PASS" : "FAIL",
    got: `fields Word lists: original ${e.word_field_counts.original}, after reject-all ${e.word_field_counts.rejected}, after accept-all ${e.word_field_counts.accepted}`,
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

// A document whose tracked insertions and deletions hold the run elements
// that are text without being w:t (no-break hyphen, soft hyphen, tab, line
// break, symbol). Opened in a Word of its own; the pane's snapshot must find
// every tracked run in the paragraph text Word reports. No model turn.
async function special() {
  const src = path.join(SPECIAL_DIR, "source.docx");
  fs.mkdirSync(SPECIAL_DIR, { recursive: true });
  const built = execFileSync(
    "python3",
    [path.join(HERE, "make_tracked_special.py"), src, "--count"],
    { encoding: "utf8" },
  );
  const expected = Number(built.trim().split("\n").pop());
  embedTaskpane(src, DOC_SPECIAL);
  const cdpBusy = () =>
    psList(
      `[int][bool](Get-NetTCPConnection -LocalPort ${Number(opt["cdp-port"])} -State Listen -ErrorAction SilentlyContinue)`,
    )[0] === 1;
  await until(() => !cdpBusy(), 60_000, "CDP port free");
  paneTag = "special-";
  record.word.launch_special = launchWord(DOC_SPECIAL, path.join(RUN, "webview2-special"));
  save();
  const snap = pane("snapshot");
  record.word.special_snapshot = snap;
  const j =
    snap.ok === false ? { pass: false, got: snap.error } : trackedPlaced(snap.paragraphs, expected);
  rows.push({
    step: "special: tracked runs found in Word's paragraph text",
    result: j.pass ? "PASS" : "FAIL",
    got: j.got,
  });
  save();
}

// After the main run and its verify commands: close the Word of the main
// run, then the mode-14 check, the special character check, then the daemon.
async function cleanup() {
  const pid = record.word.launch?.pid;
  quitWord(pid, DOC);
  if (!failed && STAGE === "s2" && !opt["skip-mode14"]) {
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
  if (!failed && ["s2", "s25"].includes(STAGE) && !opt["skip-special"]) {
    try {
      if (record.cleanup.left_running?.length) throw new Error("a Word of this run is still up");
      await special();
    } catch (e) {
      record.special_error = String(e?.stack ?? e);
      rows.push({ step: "special: run", result: "FAIL", got: String(e?.message ?? e) });
    }
    quitWord(record.word.launch_special?.pid, DOC_SPECIAL, "quit_special");
  }
  await stopDaemon().catch((e) => (record.cleanup.daemon_error = e.message));
  // Stage cite: accept-all and reject-all on the saved document, once the
  // main Word is gone and no daemon listens.
  if (STAGE === "cite" && record.cite?.checks.tracked) {
    try {
      if (record.cleanup.left_running?.length || wordPids().includes(pid))
        throw new Error("the Word of the main run is still up");
      citeJudge();
    } catch (e) {
      record.cite.judge_error = String(e?.stack ?? e);
      rows.push({
        step: "cite: accept-all and reject-all in Word",
        result: "FAIL",
        got: firstLine(e),
      });
    }
  }
  try {
    const left = wordPids();
    const own = [
      pid,
      record.word.launch_mode14?.pid,
      record.word.launch_special?.pid,
      record.word.launch_cite?.pid,
    ].filter(Boolean);
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
  await setup();
  await { s2: main, s25: mainS25, s3: mainS3, s4: mainS4, cite: mainCite }[STAGE]();
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

isolationCheck();
// Before the mode-14 check starts a daemon with a log of its own.
const mainLogs = runLogs();
function verify(ids) {
  // Stage cite has no feature list and no feature range.
  if (STAGE === "cite") return;
  try {
    runFeatures(ids);
  } catch (e) {
    record.features_error = String(e?.message ?? e);
  }
}
async function tidy() {
  if (failed && opt["keep-on-fail"]) {
    record.cleanup.kept = { word_pid: record.word.launch?.pid, daemon_pid: daemon?.pid };
    daemon?.unref();
  } else {
    await cleanup();
  }
}
if (STAGE === "s2") {
  // Verify commands first: S2-18 checks that the daemon still listens.
  verify();
  if (opt.screenshots && !failed) {
    // What a user sees when the daemon does not answer. After the verify
    // commands, before the cleanup; nothing of it is judged.
    await stopDaemon().catch((e) => (record.screenshots_error = e.message));
    showTab("chat");
    showTab("review");
    pane("refresh-offline");
    // The daemon again, for the Word runs of the cleanup that use it.
    await startDaemon(daemonLog, {}, { append: true }).catch(
      (e) => (record.screenshots_error = e.message),
    );
  }
  await tidy();
} else {
  // Stage 3: its features before the cleanup, then the ones of
  // --after-cleanup (S3-41: Word and the daemon are gone). Stage 2.5: the
  // cleanup first, S25-41 is in its range.
  if (AFTER_CLEANUP.length)
    verify(featureIds(opt.features).filter((id) => !AFTER_CLEANUP.includes(id)));
  await tidy();
  marker(
    `${STAGE}-cleanup`,
    cleanupEvidence({
      cleanup: record.cleanup,
      listening: [WS_PORT, HTTP_PORT].filter(listening).length,
    }),
  );
  verify(AFTER_CLEANUP.length ? AFTER_CLEANUP : undefined);
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
