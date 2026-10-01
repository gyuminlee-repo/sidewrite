import { test } from "node:test";
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { copyFile, mkdtemp, readdir, readFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { anonymizedPath, runAnonymize, ANONYMIZE_NOTICE } from "../daemon/anonymize.mjs";
import { pyJson, siblingPath, SAVE_FIRST } from "../daemon/doc-tools.mjs";

const NOW = new Date(2031, 0, 2, 3, 4, 5);
const CLEAN = {
  total: 0,
  body_equal: true,
  markup_equal: true,
  structure_ok: true,
  unreadable_parts: [],
  not_cleared: ["hidden_text"],
  categories: {
    names: { status: "cleared", count: 0, default: "clear" },
    hidden_text: { status: "not cleared", count: 1, default: "report" },
  },
};
const LEFT = {
  ...CLEAN,
  total: 3,
  not_cleared: ["image_meta", "hidden_text"],
  categories: {
    ...CLEAN.categories,
    image_meta: { status: "not cleared", count: 3, default: "clear" },
  },
};

// A runPy that answers from a script of results and records the calls.
function fakePy(...answers) {
  const calls = [];
  const runPy = async (script, args) => {
    calls.push({ script, args });
    return answers[calls.length - 1];
  };
  return { runPy, calls };
}
const done = (json) => ({ ok: true, code: 0, json, stderr: "", error: null });
const exit = (code, json, stderr = "") => ({ ok: false, code, json, stderr, error: stderr || "x" });

test("the copy is named after the document; a taken name gets a time stamp, never an overwrite", () => {
  assert.equal(anonymizedPath("/work/paper.docx"), "/work/paper_anonymized.docx");
  assert.equal(siblingPath("/work/my paper.v2.docx", "marked"), "/work/my paper.v2_marked.docx");
  const taken = (p) => p === "/work/paper_anonymized.docx";
  assert.equal(
    anonymizedPath("/work/paper.docx", { exists: taken, now: NOW }),
    "/work/paper_anonymized_20310102-030405.docx",
  );
  assert.equal(anonymizedPath(null), null);
  assert.equal(anonymizedPath(""), null);
});

test("the tool runs first, then the verifier, with the same options; the result is the verifier's report", async () => {
  const { runPy, calls } = fakePy(done({ output: "x", changed: { a: 1 } }), done(CLEAN));
  const r = await runAnonymize({
    src: "/work/paper.docx",
    author: "Reviewer copy",
    initials: "RC",
    keep: ["rsids", "dates"],
    also: ["title"],
    dates: "remove",
    runPy,
    root: "/repo",
  });
  assert.deepEqual(r, {
    ok: true,
    path: "/work/paper_anonymized.docx",
    report: CLEAN,
    changed: { a: 1 },
    notice: ANONYMIZE_NOTICE,
  });
  assert.equal(calls.length, 2);
  assert.equal(calls[0].script, join("/repo", "tools", "anonymize_docx.py"));
  assert.deepEqual(calls[0].args, [
    "--in", "/work/paper.docx", "--out", "/work/paper_anonymized.docx",
    "--author", "Reviewer copy", "--initials", "RC",
    "--keep", "rsids,dates", "--also", "title", "--dates", "remove",
  ]); // prettier-ignore
  assert.equal(calls[1].script, join("/repo", "scripts", "verify-anonymized.py"));
  assert.deepEqual(calls[1].args, [
    "--before", "/work/paper.docx", "--after", "/work/paper_anonymized.docx", "--json",
    "--author", "Reviewer copy", "--initials", "RC",
    "--keep", "rsids,dates", "--also", "title",
  ]); // prettier-ignore
  assert.match(ANONYMIZE_NOTICE, /text itself is not changed/);
});

test("defaults: Anonymous, A, nothing kept, and no option the tool does not need", async () => {
  const { runPy, calls } = fakePy(done({ changed: {} }), done(CLEAN));
  await runAnonymize({ src: "/work/paper.docx", runPy, root: "/repo" });
  assert.deepEqual(calls[0].args.slice(4), ["--author", "Anonymous", "--initials", "A"]);
});

test("no saved document, or an output that is the input: refused before anything runs", async () => {
  const { runPy, calls } = fakePy();
  assert.deepEqual(await runAnonymize({ src: null, runPy }), { ok: false, error: SAVE_FIRST });
  for (const out of ["/work/paper.docx", "/work/./paper.docx", "/work/sub/../paper.docx"]) {
    const r = await runAnonymize({ src: "/work/paper.docx", out, runPy });
    assert.deepEqual([r.ok, /same file/.test(r.error)], [false, true], out);
  }
  const same = await runAnonymize({
    src: "/work/paper.docx",
    out: "/work/link.docx",
    runPy,
    sameFile: async () => true,
  });
  assert.equal(same.ok, false);
  assert.equal(calls.length, 0);
});

test("a tool failure is shown as the tool said it and the verifier does not run", async () => {
  const refused = fakePy(exit(2, undefined, "not a docx (not a zip file): /work/paper.docx"));
  const r = await runAnonymize({ src: "/work/paper.docx", runPy: refused.runPy, root: "/repo" });
  assert.deepEqual(r, { ok: false, error: "not a docx (not a zip file): /work/paper.docx" });
  assert.equal(refused.calls.length, 1);
  const noPython = fakePy({ ok: false, code: null, error: "python3 not found" });
  const n = await runAnonymize({ src: "/work/paper.docx", runPy: noPython.runPy, root: "/repo" });
  assert.deepEqual(n, { ok: false, error: "python3 not found" });
});

test("when the verifier finds something left, the copy is set aside as INCOMPLETE and the result names it", async () => {
  const moved = [];
  const { runPy } = fakePy(done({ changed: {} }), exit(1, LEFT));
  const r = await runAnonymize({
    src: "/work/paper.docx",
    runPy,
    root: "/repo",
    move: async (a, b) => moved.push([a, b]),
  });
  assert.equal(r.ok, false);
  assert.equal(r.path, "/work/paper_anonymized.INCOMPLETE.docx");
  assert.deepEqual(moved, [
    ["/work/paper_anonymized.docx", "/work/paper_anonymized.INCOMPLETE.docx"],
  ]);
  assert.equal(r.report, LEFT);
  assert.match(r.error, /image_meta/);
  assert.ok(!/hidden_text/.test(r.error)); // reported only: not a reason to fail
  const changed = { ...CLEAN, body_equal: false, structure_ok: false };
  const c = await runAnonymize({
    src: "/work/paper.docx",
    runPy: fakePy(done({ changed: {} }), exit(1, changed)).runPy,
    root: "/repo",
    move: async () => {},
  });
  assert.match(c.error, /text changed/);
  assert.match(c.error, /missing part/);
});

test("a verifier that could not run leaves no file passing as checked", async () => {
  const moved = [];
  const { runPy } = fakePy(done({ changed: {} }), exit(2, undefined, "cannot read"));
  const r = await runAnonymize({
    src: "/work/paper.docx",
    runPy,
    root: "/repo",
    move: async (a, b) => moved.push([a, b]),
  });
  assert.deepEqual(
    [r.ok, r.path, /could not be checked/.test(r.error)],
    [false, "/work/paper_anonymized.INCOMPLETE.docx", true],
  );
  assert.equal(moved.length, 1);
});

test("the runner's result is read whether it carries parsed JSON or raw stdout", () => {
  assert.deepEqual(pyJson({ json: { a: 1 } }), { a: 1 });
  assert.deepEqual(pyJson({ stdout: 'PASS x\n{"a": 2}\n' }), { a: 2 });
  assert.deepEqual(pyJson({ stdout: '{\n "a": 3\n}\n' }), { a: 3 });
  assert.equal(pyJson({ stdout: "no json" }), undefined);
  assert.equal(pyJson({}), undefined);
});

// ---- the real tools, through a runner shaped like the daemon's
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

test("with the real tools: a copy next to the source, verified, and the source untouched", async () => {
  const dir = await mkdtemp(join(tmpdir(), "anon-"));
  const src = join(dir, "paper.docx");
  await copyFile(join(ROOT, "tests", "fixtures", "review_manuscript.docx"), src);
  const before = await readFile(src);
  const r = await runAnonymize({ src, runPy: realPy, exists: existsSync, root: ROOT });
  assert.equal(r.ok, true, r.error);
  assert.equal(r.path, join(dir, "paper_anonymized.docx"));
  assert.equal(r.report.total, 0);
  assert.equal(r.report.categories.names.status, "cleared");
  assert.ok((await readFile(src)).equals(before));
  // A second run does not overwrite the first copy.
  const again = await runAnonymize({ src, runPy: realPy, exists: existsSync, root: ROOT });
  assert.equal(again.ok, true, again.error);
  assert.notEqual(again.path, r.path);
  assert.equal((await readdir(dir)).length, 3);
  // Kept on purpose: the verifier is told, and the result still passes.
  const kept = await runAnonymize({
    src,
    out: join(dir, "kept.docx"),
    keep: ["names", "comment_authors"],
    runPy: realPy,
    root: ROOT,
  });
  assert.equal(kept.ok, true, kept.error);
  assert.equal(kept.report.categories.names.status, "kept by option");
});
