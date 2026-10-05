#!/usr/bin/env node
// How does Word write a tracked change of an equation?
//
// The marked-diff tool refuses an equation that is new, gone or changed,
// because it does not know the notation. This probe lets Word answer: in a
// Word process of its own, word-math-probe.ps1 makes a document with
// equations, then a copy in which equations are deleted, inserted and
// changed with Track Changes on. The word/document.xml of both files is
// written to the run folder and summarised here: which element wraps which
// around the equations (w:ins, w:del, m:oMath, m:r, w:rPr, m:ctrlPr, the
// paragraph marks).
//
// Runs from WSL; Word is on Windows. No daemon and no port. Everything is
// written under <base-dir>/wmp-<time>/ (a new folder every run; nothing is
// deleted). Exit code 0 when Word made both files, every change ran, and
// the Word of this run is gone; 1 otherwise. See README.md.

import { execFileSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { parseArgs } from "node:util";
import { cleanupDecision } from "./judge.mjs";
import { markRelations, mathSummary } from "./word-math-probe-lib.mjs";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "../..");

const { values: opt } = parseArgs({
  options: {
    "base-dir": { type: "string", default: process.env.E2E_BASE_DIR },
    python: { type: "string", default: "python3" },
    "timeout-min": { type: "string", default: "10" },
  },
});
if (!opt["base-dir"]) {
  console.error("word-math-probe: --base-dir (a folder on a Windows drive, WSL path) is required");
  process.exit(2);
}
const stamp = new Date().toISOString().replace(/[-:]/g, "").replace("T", "-").slice(0, 15);
const RUN = path.resolve(opt["base-dir"], `wmp-${stamp}`);
fs.mkdirSync(path.join(RUN, "tools"), { recursive: true });

const winPath = (p) => execFileSync("wslpath", ["-w", p], { encoding: "utf8" }).trim();
const lastLine = (text) => String(text).replace(/\r/g, "").trim().split("\n").filter(Boolean).pop();
const record = { run_dir: RUN, started_at: new Date().toISOString(), word: {}, cleanup: {} };
const save = () => fs.writeFileSync(path.join(RUN, "result.json"), JSON.stringify(record, null, 2));

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

// 1. Word.
for (const name of ["word-math-probe.ps1", "word-own.ps1", "word-rot.ps1", "word-back.ps1"]) {
  fs.copyFileSync(path.join(HERE, name), path.join(RUN, "tools", name));
}
const anchor = path.join(RUN, "anchor.docx");
fs.copyFileSync(path.join(REPO, "tests/fixtures/spike_manuscript.docx"), anchor);
const files = { base: path.join(RUN, "base.docx"), revised: path.join(RUN, "revised.docx") };
const outFile = path.join(RUN, "word.json");
const pidFile = path.join(RUN, "word-pid.json");
const preexisting = wordPids("before") ?? [];
let done = null;
try {
  const out = execFileSync(
    "powershell.exe",
    [
      ...["-NoProfile", "-ExecutionPolicy", "Bypass"],
      ...["-File", winPath(path.join(RUN, "tools", "word-math-probe.ps1"))],
      ...["-Base", winPath(RUN) + "\\base.docx", "-Revised", winPath(RUN) + "\\revised.docx"],
      ...["-Out", winPath(RUN) + "\\word.json", "-Anchor", winPath(anchor)],
      ...["-AllowDir", winPath(RUN), "-PidFile", winPath(RUN) + "\\word-pid.json"],
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
  record.word.error = String(e?.stderr || e?.message || e).slice(0, 4000);
}
record.word.run = done;
if (done && !done.ok) record.word.error = done.error;
const readJson = (f) =>
  fs.existsSync(f) ? JSON.parse(fs.readFileSync(f, "utf8").replace(/^﻿/, "")) : null;
const started = readJson(pidFile);
if (!done) record.word.partial = readJson(outFile);

// 2. Cleanup, as marked-diff-word-check.mjs does: the Word of this run is
// stopped by force only when its quit ran and the process stayed; otherwise
// it is left alone (cleanupDecision).
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
  if (d.action === "leave") console.error(`word-math-probe: ${d.reason}`);
  cleanupFail = d.fail;
  const left = wordPids("at_end");
  if (left) {
    record.cleanup.preexisting_alive = before.every((p) => left.includes(p));
    record.cleanup.own_word_gone = !left.includes(pid);
    if (!record.cleanup.own_word_gone) cleanupFail = true;
  }
}

// 3. The XML of both files, and what it shows.
const UNZIP =
  "import sys, zipfile; sys.stdout.buffer.write(zipfile.ZipFile(sys.argv[1]).read('word/document.xml'))";
const xmlFiles = {};
for (const [name, docx] of Object.entries(files)) {
  if (!fs.existsSync(docx)) continue;
  try {
    const xml = execFileSync(opt.python, ["-c", UNZIP, docx], { maxBuffer: 256 << 20 });
    xmlFiles[name] = path.join(RUN, `${name}.document.xml`);
    fs.writeFileSync(xmlFiles[name], xml);
  } catch (e) {
    record.word[`unzip_error_${name}`] = String(e?.message ?? e).slice(0, 500);
  }
}
const steps = done?.steps ?? [];
const ok =
  Boolean(done?.ok) &&
  steps.length === 5 &&
  steps.every((s) => s.ok) &&
  Boolean(xmlFiles.base && xmlFiles.revised) &&
  !cleanupFail;
record.steps = steps;
record.info = done?.info ?? null;
record.xml = xmlFiles;
record.exit = { code: ok ? 0 : 1 };
record.finished_at = new Date().toISOString();
save();

for (const s of steps)
  console.log(`${s.ok ? "ok  " : "FAIL"} ${s.name}${s.ok ? "" : `: ${s.error}`}`);
if (done?.info) console.log(`revised document: ${JSON.stringify(done.info)}`);
if (record.word.error) console.log(`Word run: ${record.word.error}`);
for (const [name, file] of Object.entries(xmlFiles)) {
  const xml = fs.readFileSync(file, "utf8");
  console.log(`\n== ${name}: ${file}`);
  console.log(mathSummary(xml).join("\n"));
  const marks = markRelations(xml);
  if (marks.length) console.log(`-- where the change marks stand\n${marks.join("\n")}`);
}
console.log(`\ncleanup: ${JSON.stringify(record.cleanup)}`);
console.log(`result: ${path.join(RUN, "result.json")}`);
process.exit(ok ? 0 : 1);
