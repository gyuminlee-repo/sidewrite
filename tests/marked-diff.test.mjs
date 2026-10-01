import { test } from "node:test";
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { copyFile, mkdtemp, readdir } from "node:fs/promises";
import { existsSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { markedOutPath, runMarkedDiff, MARKED_NOTICE, NEEDS_LXML } from "../daemon/marked-diff.mjs";
import { SAVE_FIRST } from "../daemon/doc-tools.mjs";
import { DISK_NOTICE } from "../daemon/doc-bytes.mjs";
import { missing } from "./fixtures/py-packages.mjs";

const NOW = new Date(2031, 0, 2, 3, 4, 5);
const SUMMARY = { written: true, checks: 13, failures: 0, mode: "color", paragraphs: { n_mod: 2 } };

function fakePy(answer) {
  const calls = [];
  const runPy = async (script, args) => {
    calls.push({ script, args });
    return answer;
  };
  return { runPy, calls };
}
const result = (code, json, stderr = "") => ({
  ok: code === 0,
  code,
  json,
  stderr,
  error: code === 0 ? null : stderr.trim().split("\n").pop() || "failed",
});
const run = (more = {}) =>
  runMarkedDiff({
    baseline: "/tmp/up/base.docx",
    clean: "/tmp/up/current.docx",
    docPath: "/work/paper.docx",
    root: "/repo",
    ...more,
  });

test("the marked copy is named after the document; a taken name gets a time stamp", () => {
  assert.equal(markedOutPath("/work/paper.docx"), "/work/paper_marked.docx");
  const taken = (p) => p === "/work/paper_marked.docx";
  assert.equal(
    markedOutPath("/work/paper.docx", { exists: taken, now: NOW }),
    "/work/paper_marked_20310102-030405.docx",
  );
  assert.equal(markedOutPath(null), null);
});

test("exit 0: the path, the counts and the notice about pending changes", async () => {
  const { runPy, calls } = fakePy(result(0, SUMMARY));
  const r = await run({ runPy });
  assert.deepEqual(r, {
    ok: true,
    path: "/work/paper_marked.docx",
    report: SUMMARY,
    notice: MARKED_NOTICE,
  });
  assert.equal(calls.length, 1);
  assert.equal(calls[0].script, join("/repo", "tools", "build_marked_diff.py"));
  assert.deepEqual(calls[0].args, [
    "--baseline",
    "/tmp/up/base.docx",
    "--clean",
    "/tmp/up/current.docx",
    "--out",
    "/work/paper_marked.docx",
    "--json",
  ]);
  assert.match(MARKED_NOTICE, /pending tracked changes .* accepted/i);
});

test("colours, mode and author come from the profile", async () => {
  const { runPy, calls } = fakePy(result(0, SUMMARY));
  const profile = {
    colors: { marked_insert: "0000ff", marked_delete: "00AA00" },
    marked_mode: "track_changes",
    marked_author: "Reviewer copy",
  };
  await run({ runPy, profile });
  assert.deepEqual(calls[0].args.slice(7), [
    "--insert-rgb",
    "0000FF",
    "--delete-rgb",
    "00AA00",
    "--mode",
    "track_changes",
    "--author",
    "Reviewer copy",
  ]);
  const one = fakePy(result(0, SUMMARY));
  await run({ runPy: one.runPy, profile: { colors: { marked_insert: "112233" } } });
  assert.deepEqual(one.calls[0].args.slice(7), ["--insert-rgb", "112233"]);
});

test("a profile value the tool would not take is refused before it runs", async () => {
  const { runPy, calls } = fakePy(result(0, SUMMARY));
  for (const profile of [
    { colors: { marked_insert: "red" } },
    { colors: { marked_delete: "12345" } },
    { marked_mode: "highlight" },
    { colors: { marked_insert: "--out" } },
  ]) {
    const r = await run({ runPy, profile });
    assert.equal(r.ok, false, JSON.stringify(profile));
  }
  assert.equal(calls.length, 0);
});

test("no saved document, or an output that is an input: refused before the tool runs", async () => {
  const { runPy, calls } = fakePy(result(0, SUMMARY));
  assert.deepEqual(await run({ runPy, docPath: null }), { ok: false, error: SAVE_FIRST });
  for (const out of ["/tmp/up/base.docx", "/tmp/up/./current.docx"]) {
    const r = await run({ runPy, out });
    assert.deepEqual([r.ok, /same file/.test(r.error)], [false, true]);
  }
  assert.equal((await run({ runPy, baseline: null })).ok, false);
  assert.equal(calls.length, 0);
});

test("an existing marked copy is not overwritten", async () => {
  const { runPy, calls } = fakePy(result(0, SUMMARY));
  const r = await run({ runPy, exists: (p) => p === "/work/paper_marked.docx", now: NOW });
  assert.equal(r.path, "/work/paper_marked_20310102-030405.docx");
  assert.equal(calls[0].args[5], r.path);
});

test("exit 1: the tool's own check failed and nothing was written", async () => {
  const failed = { ...SUMMARY, written: false, failures: 2 };
  const r = await run({ runPy: fakePy(result(1, failed)).runPy });
  assert.equal(r.ok, false);
  assert.equal(r.report, failed);
  assert.equal(r.path, undefined);
  assert.match(r.error, /2 of 13 checks failed/);
  assert.match(r.error, /nothing was written/);
});

test("exit 2 and a missing interpreter: the reason as the tool gave it", async () => {
  const refused = await run({
    runPy: fakePy(
      result(
        2,
        undefined,
        "the clean document already has 1 run(s) in the insert colour FF0000; choose another colour with --insert-rgb\n",
      ),
    ).runPy,
  });
  assert.deepEqual([refused.ok, /insert colour FF0000/.test(refused.error)], [false, true]);
  const none = await run({
    runPy: async () => ({ ok: false, code: null, error: "python3 not found" }),
  });
  assert.deepEqual(none, { ok: false, error: "python3 not found" });
});

test("a missing lxml is named with what to install; nothing is installed", async () => {
  const trace =
    "Traceback (most recent call last):\n  File \"tools/build_marked_diff.py\", line 43, in <module>\n    from lxml import etree\nModuleNotFoundError: No module named 'lxml'\n";
  const r = await run({ runPy: fakePy(result(1, undefined, trace)).runPy });
  assert.deepEqual(r, { ok: false, error: NEEDS_LXML });
  assert.match(NEEDS_LXML, /pip install lxml/);
});

test("when the current document was read from disk the notice says so", async () => {
  const r = await run({ runPy: fakePy(result(0, SUMMARY)).runPy, source: "disk" });
  assert.equal(r.notice, `${MARKED_NOTICE} ${DISK_NOTICE}`);
});

// ---- the real tool, through a runner shaped like the daemon's
const ROOT = fileURLToPath(new URL("..", import.meta.url));
const realPy = (script, args) =>
  new Promise((resolve) => {
    execFile("python3", [script, ...args], { encoding: "utf8" }, (err, stdout, stderr) => {
      let json;
      try {
        json = JSON.parse(String(stdout).trim().split("\n").pop());
      } catch {
        json = undefined;
      }
      const code = err ? (typeof err.code === "number" ? err.code : null) : 0;
      const last = String(stderr).trim().split("\n").pop();
      resolve({ ok: code === 0, code, json, stderr, error: err ? last || err.message : null });
    });
  });

test(
  "with the real tool: a marked copy next to the document, and a second one beside it",
  { skip: missing("lxml") },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), "marked-"));
    const fixtures = join(ROOT, "tests", "fixtures");
    const baseline = join(dir, "base.docx");
    const clean = join(dir, "paper.docx");
    await copyFile(join(fixtures, "spike_manuscript.docx"), baseline);
    await copyFile(join(fixtures, "review_manuscript.docx"), clean);
    const first = await runMarkedDiff({
      baseline,
      clean,
      docPath: clean,
      runPy: realPy,
      exists: existsSync,
      root: ROOT,
    });
    assert.equal(first.ok, true, first.error);
    assert.equal(first.path, join(dir, "paper_marked.docx"));
    assert.deepEqual([first.report.written, first.report.failures], [true, 0]);
    const second = await runMarkedDiff({
      baseline,
      clean,
      docPath: clean,
      runPy: realPy,
      exists: existsSync,
      root: ROOT,
    });
    assert.equal(second.ok, true, second.error);
    assert.notEqual(second.path, first.path);
    assert.equal((await readdir(dir)).length, 4);
  },
);
