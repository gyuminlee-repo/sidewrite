#!/usr/bin/env node
// Word as the judge of the marked-diff tool.
//
// The tool checks its own output with its own model of "accept all" and
// "reject all". This run asks Word instead. For a fixed selection of pairs
// from tests/marked_diff_corpus.py it writes the baseline and the clean
// document, runs tools/build_marked_diff.py in both modes, and then, in a
// Word process of its own:
//   - opens the track-changes output, Revisions.AcceptAll(), reads the
//     structure, closes without saving; the same with RejectAll();
//   - opens the clean and the baseline document and reads their structure;
//   - opens the colour output and counts its paragraphs and tables.
// A pair passes when Word opens every file without an error, accept-all is
// the clean document, reject-all is the baseline, no revision is left, and
// the colour output has the paragraphs and tables of the track-changes one.
//
// Runs from WSL; Word is on Windows. No daemon and no port. Everything is
// written under <base-dir>/mdw-<time>/ (a new folder every run; nothing is
// deleted). Exit code 0 only when every pair passed and the Word of this run
// is gone. See README.md.

import { execFileSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { parseArgs } from "node:util";
import { cleanupDecision } from "./judge.mjs";
import { PAIRS, exitCode, jobsFor, judgePair, renderRows } from "./marked-diff-word-lib.mjs";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "../..");

const { values: opt } = parseArgs({
  options: {
    "base-dir": { type: "string", default: process.env.E2E_BASE_DIR },
    "run-dir": { type: "string" },
    python: { type: "string", default: "python3" },
    // Only the pairs whose name has one of these texts (comma separated).
    only: { type: "string" },
    "timeout-min": { type: "string", default: "30" },
    // Write the pairs, the tool's output and jobs.json, and stop before Word.
    "prepare-only": { type: "boolean", default: false },
  },
});
const die = (m) => {
  console.error(`marked-diff-word-check: ${m}`);
  process.exit(2);
};
if (!opt["base-dir"] && !opt["run-dir"]) {
  die("--base-dir (a folder on a Windows drive, WSL path) is required");
}
const stamp = new Date().toISOString().replace(/[-:]/g, "").replace("T", "-").slice(0, 15);
const RUN = path.resolve(opt["run-dir"] ?? path.join(opt["base-dir"], `mdw-${stamp}`));
if (fs.existsSync(RUN) && fs.readdirSync(RUN).length) die(`${RUN} is not empty; give a new folder`);

const winPath = (p) => execFileSync("wslpath", ["-w", p], { encoding: "utf8" }).trim();
const wanted = (opt.only ?? "")
  .split(",")
  .map((s) => s.trim())
  .filter(Boolean);
const pairs = PAIRS.filter((p) => !wanted.length || wanted.some((w) => p.name.includes(w)));
if (!pairs.length) die("no pair matches --only");

const record = {
  run_dir: RUN,
  started_at: new Date().toISOString(),
  pairs: [],
  word: {},
  cleanup: {},
};
const save = () => fs.writeFileSync(path.join(RUN, "result.json"), JSON.stringify(record, null, 2));
const lastLine = (text) => String(text).replace(/\r/g, "").trim().split("\n").filter(Boolean).pop();

// The tool in one mode: its exit code and the last line it printed.
function runTool(dir, mode, out) {
  const args = [
    path.join(REPO, "tools/build_marked_diff.py"),
    ...["--baseline", path.join(dir, "base.docx"), "--clean", path.join(dir, "clean.docx")],
    ...["--out", out, "--mode", mode, "--json"],
  ];
  try {
    const stdout = execFileSync(opt.python, args, {
      encoding: "utf8",
      stdio: ["ignore", "pipe", "pipe"],
    });
    return { code: 0, last: lastLine(stdout) };
  } catch (e) {
    return { code: e.status ?? 1, last: lastLine(e.stderr || e.stdout || e.message) };
  }
}

// The ids of the Word processes that run now; null (and a note in the
// record) when Windows could not be asked.
function wordPids(when) {
  try {
    const out = execFileSync(
      "powershell.exe",
      [
        "-NoProfile",
        "-Command",
        "(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id) -join ','",
      ],
      { encoding: "utf8" },
    );
    return out.replace(/\r/g, "").trim().split(",").filter(Boolean).map(Number);
  } catch (e) {
    record.cleanup[`pids_error_${when}`] = String(e?.message ?? e).slice(0, 500);
    return null;
  }
}

fs.mkdirSync(path.join(RUN, "tools"), { recursive: true });
save();

// 1. The pairs and the tool's output for each.
const namesFile = path.join(RUN, "names.json");
fs.writeFileSync(namesFile, JSON.stringify(pairs.map((p) => p.name)));
const made = JSON.parse(
  lastLine(
    execFileSync(
      opt.python,
      [
        path.join(HERE, "marked_diff_pairs.py"),
        ...["--out", path.join(RUN, "pairs"), "--names", namesFile],
      ],
      { encoding: "utf8" },
    ),
  ),
);
const jobs = [];
made.forEach((pair, i) => {
  const files = {
    clean: path.join(pair.dir, "clean.docx"),
    base: path.join(pair.dir, "base.docx"),
    track: path.join(pair.dir, "marked_track.docx"),
    color: path.join(pair.dir, "marked_color.docx"),
  };
  const tool = {
    track: runTool(pair.dir, "track_changes", files.track),
    color: runTool(pair.dir, "color", files.color),
  };
  const win = Object.fromEntries(
    Object.entries(files).map(([k, f]) => [k, fs.existsSync(f) ? winPath(f) : null]),
  );
  jobs.push(...jobsFor(i + 1, win));
  record.pairs.push({
    index: i + 1,
    name: pair.name,
    covers: pairs[i].covers,
    dir: pair.dir,
    tool,
  });
});
save();

// 2. Word.
const anchor = path.join(RUN, "anchor.docx");
fs.copyFileSync(path.join(made[0].dir, "clean.docx"), anchor);
const ps1 = path.join(RUN, "tools", "marked-diff-word.ps1");
fs.copyFileSync(path.join(HERE, "marked-diff-word.ps1"), ps1);
fs.copyFileSync(path.join(HERE, "word-own.ps1"), path.join(RUN, "tools", "word-own.ps1"));
const jobsFile = path.join(RUN, "jobs.json");
const outFile = path.join(RUN, "word.json");
const pidFile = path.join(RUN, "word-pid.json");
fs.writeFileSync(jobsFile, JSON.stringify(jobs));
if (opt["prepare-only"]) {
  for (const x of record.pairs) {
    console.log(`${x.index} ${x.name}: track ${x.tool.track.code}, colour ${x.tool.color.code}`);
  }
  console.log(
    `${record.pairs.length} pairs, ${jobs.length} Word jobs in ${RUN}; Word was not started`,
  );
  const written = record.pairs.every((x) => x.tool.track.code === 0 && x.tool.color.code === 0);
  process.exit(written ? 0 : 1);
}
const preexisting = wordPids("before") ?? [];
let done = null;
try {
  const out = execFileSync(
    "powershell.exe",
    [
      ...["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", winPath(ps1)],
      ...["-Jobs", winPath(jobsFile), "-Out", winPath(outFile), "-Anchor", winPath(anchor)],
      ...["-AllowDir", winPath(RUN), "-PidFile", winPath(pidFile)],
    ],
    {
      cwd: RUN,
      encoding: "utf8",
      timeout: Number(opt["timeout-min"]) * 60_000,
      maxBuffer: 64 << 20,
    },
  );
  done = JSON.parse(lastLine(out));
} catch (e) {
  record.word.error = String(e?.message ?? e).slice(0, 2000);
}
record.word.run = done;
const readJson = (f) =>
  fs.existsSync(f) ? JSON.parse(fs.readFileSync(f, "utf8").replace(/^\uFEFF/, "")) : null;
const started = readJson(pidFile);
if (done && !done.ok) record.word.error = done.error;

// 3. Cleanup: the Word of this run is stopped by force only when its quit
// ran and the process stayed; otherwise it is left alone (cleanupDecision).
const pid = done?.pid ?? started?.pid;
let cleanupFail = false;
if (pid) {
  const before = started?.preexisting ?? preexisting;
  const now = wordPids("after_quit");
  const d = cleanupDecision({
    pid,
    preexisting: before,
    quit: done?.quit,
    runDir: winPath(RUN),
    alive: now === null || now.includes(pid),
  });
  Object.assign(record.cleanup, { pid, quit: done?.quit ?? null, decision: d });
  if (d.action === "kill") {
    try {
      execFileSync("powershell.exe", [
        "-NoProfile",
        "-Command",
        `Stop-Process -Id ${Number(pid)} -ErrorAction SilentlyContinue`,
      ]);
      record.cleanup.stopped = true;
    } catch (e) {
      record.cleanup.stop_error = String(e?.message ?? e).slice(0, 500);
    }
  }
  if (d.action === "leave") console.error(`marked-diff-word-check: ${d.reason}`);
  cleanupFail = d.fail;
  const left = wordPids("at_end");
  if (left) {
    record.cleanup.preexisting_alive = before.every((p) => left.includes(p));
    record.cleanup.own_word_gone = !left.includes(pid);
    if (!record.cleanup.own_word_gone) cleanupFail = true;
  }
}

// 4. The verdicts.
const byId = new Map((readJson(outFile)?.jobs ?? []).map((j) => [j.id, j]));
const rows = record.pairs.map((pair) => {
  const job = (role) => byId.get(`${pair.index}:${role}`);
  const verdict = judgePair({
    tool: { track: pair.tool.track.code, color: pair.tool.color.code },
    accept: job("accept"),
    reject: job("reject"),
    clean: job("clean"),
    base: job("base"),
    color: job("color"),
  });
  pair.result = verdict.pass ? "PASS" : "FAIL";
  pair.failures = verdict.failures;
  return {
    step: `${String(pair.index).padStart(2, "0")} ${pair.name}`,
    result: pair.result,
    got: verdict.first,
  };
});
const code = exitCode({ rows, wordError: record.word.error, cleanupFail });
record.exit = {
  code,
  passed: rows.filter((r) => r.result === "PASS").length,
  pairs: rows.length,
};
record.finished_at = new Date().toISOString();
save();

console.log(renderRows(rows));
if (record.word.error) console.log(`\nWord run: ${record.word.error}`);
console.log(`\n${record.exit.passed} of ${rows.length} pairs passed`);
console.log(`cleanup: ${JSON.stringify(record.cleanup)}`);
console.log(`result: ${path.join(RUN, "result.json")}`);
process.exit(code);
