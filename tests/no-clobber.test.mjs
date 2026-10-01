// No existing file is ever overwritten by the tools that write a copy of the
// manuscript: not by the name that is picked, and not by the write itself.

import { test } from "node:test";
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { copyFile, mkdtemp, readFile, readdir, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { runAnonymize } from "../daemon/anonymize.mjs";
import { runMarkedDiff } from "../daemon/marked-diff.mjs";
import { siblingPath, moveExclusive } from "../daemon/doc-tools.mjs";
import { createUploads } from "../daemon/uploads.mjs";
import { toChunks } from "../daemon/doc-bytes.mjs";
import { missing } from "./fixtures/py-packages.mjs";

const ROOT = fileURLToPath(new URL("..", import.meta.url));
const FIXTURES = join(ROOT, "tests", "fixtures");
const NOW = new Date(2031, 0, 2, 3, 4, 5);
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
const newDir = () => mkdtemp(join(tmpdir(), "noclobber-"));
const contents = async (dir) =>
  Object.fromEntries(
    await Promise.all(
      (await readdir(dir)).sort().map(async (n) => [n, await readFile(join(dir, n))]),
    ),
  );

test("a taken name gets a time stamp, and a taken time stamp gets a counter", () => {
  const taken = new Set(["/w/p_marked.docx"]);
  const exists = (p) => taken.has(p);
  const first = siblingPath("/w/p.docx", "marked", { exists, now: NOW });
  assert.equal(first, "/w/p_marked_20310102-030405.docx");
  taken.add(first);
  const second = siblingPath("/w/p.docx", "marked", { exists, now: NOW });
  assert.equal(second, "/w/p_marked_20310102-030405-2.docx");
  taken.add(second);
  assert.equal(
    siblingPath("/w/p.docx", "marked", { exists, now: NOW }),
    "/w/p_marked_20310102-030405-3.docx",
  );
});

test("anonymized copies made in the same second all stay, each as it was written", async () => {
  const dir = await newDir();
  const src = join(dir, "paper.docx");
  await copyFile(join(FIXTURES, "review_manuscript.docx"), src);
  const run = (author) =>
    runAnonymize({ src, author, runPy: realPy, exists: existsSync, now: NOW, root: ROOT });
  const made = [];
  for (const author of ["First", "Second", "Third"]) {
    const r = await run(author);
    assert.equal(r.ok, true, r.error);
    made.push([r.path, await readFile(r.path)]);
  }
  assert.equal(new Set(made.map(([p]) => p)).size, 3);
  for (const [path, bytes] of made) assert.ok((await readFile(path)).equals(bytes), path);
  assert.equal((await readdir(dir)).length, 4);
});

test("marked copies made in the same second all stay", { skip: missing("lxml") }, async () => {
  const dir = await newDir();
  const clean = join(dir, "paper.docx");
  await copyFile(join(FIXTURES, "review_manuscript.docx"), clean);
  const baseline = join(FIXTURES, "spike_manuscript.docx");
  const paths = [];
  for (const marked_insert of ["FF0000", "0000FF", "00AA00"]) {
    const r = await runMarkedDiff({
      baseline,
      clean,
      docPath: clean,
      profile: { colors: { marked_insert } },
      runPy: realPy,
      exists: existsSync,
      now: NOW,
      root: ROOT,
    });
    assert.equal(r.ok, true, r.error);
    paths.push(r.path);
  }
  assert.equal(new Set(paths).size, 3);
  assert.equal((await readdir(dir)).length, 4);
});

test(
  "a name that was free when it was picked but is taken when the tool writes is not overwritten",
  { skip: missing("lxml") },
  async () => {
    // The check and the write are two steps: another process can create the file in between.
    const dir = await newDir();
    const src = join(dir, "paper.docx");
    await copyFile(join(FIXTURES, "review_manuscript.docx"), src);
    const theirs = Buffer.from("somebody else's file, written between the check and the write");
    for (const [name, run] of [
      [
        "paper_anonymized.docx",
        (exists) => runAnonymize({ src, runPy: realPy, exists, now: NOW, root: ROOT }),
      ],
      [
        "paper_marked.docx",
        (exists) =>
          runMarkedDiff({
            baseline: join(FIXTURES, "spike_manuscript.docx"),
            clean: src,
            docPath: src,
            runPy: realPy,
            exists,
            now: NOW,
            root: ROOT,
          }),
      ],
    ]) {
      const target = join(dir, name);
      await writeFile(target, theirs);
      // The name check is told the file is not there the first time it asks.
      let asked = 0;
      const r = await run((p) => (p === target && asked++ === 0 ? false : existsSync(p)));
      assert.ok((await readFile(target)).equals(theirs), `${name} was overwritten`);
      assert.equal(r.ok, true, r.error);
      assert.notEqual(r.path, target);
    }
  },
);

test(
  "the Python tools refuse an existing output unless they are told to overwrite",
  { skip: missing("lxml") },
  async () => {
    const dir = await newDir();
    const src = join(dir, "paper.docx");
    await copyFile(join(FIXTURES, "review_manuscript.docx"), src);
    const old = Buffer.from("an earlier copy");
    const cases = [
      [join(ROOT, "tools", "anonymize_docx.py"), ["--in", src]],
      [
        join(ROOT, "tools", "build_marked_diff.py"),
        ["--baseline", join(FIXTURES, "spike_manuscript.docx"), "--clean", src],
      ],
    ];
    for (const [tool, args] of cases) {
      const out = join(dir, "out.docx");
      await writeFile(out, old);
      const refused = await realPy(tool, [...args, "--out", out]);
      assert.equal(refused.code, 2, tool);
      assert.match(refused.stderr, /already exists/);
      assert.ok((await readFile(out)).equals(old));
      const forced = await realPy(tool, [...args, "--out", out, "--overwrite"]);
      assert.equal(forced.code, 0, forced.stderr);
      assert.ok(!(await readFile(out)).equals(old));
      assert.deepEqual((await readdir(dir)).sort(), ["out.docx", "paper.docx"]);
    }
    // The marked-diff report file is not overwritten either.
    const report = join(dir, "report.json");
    await writeFile(report, old);
    const [tool, args] = cases[1];
    const r = await realPy(tool, [...args, "--out", join(dir, "new.docx"), "--report", report]);
    assert.equal(r.code, 2);
    assert.ok((await readFile(report)).equals(old));
    assert.equal(existsSync(join(dir, "new.docx")), false);
  },
);

test("a copy that failed its check is set aside without replacing an earlier set-aside copy", async () => {
  const dir = await newDir();
  const a = join(dir, "a.docx");
  const b = join(dir, "b.docx");
  await writeFile(a, "first");
  await writeFile(b, "second");
  await assert.rejects(moveExclusive(a, b), /EEXIST|exists/);
  assert.deepEqual(
    Object.fromEntries(Object.entries(await contents(dir)).map(([k, v]) => [k, String(v)])),
    { "a.docx": "first", "b.docx": "second" },
  );
  await moveExclusive(a, join(dir, "c.docx"));
  assert.deepEqual((await readdir(dir)).sort(), ["b.docx", "c.docx"]);

  // runAnonymize: the verifier fails twice; both copies stay, under different INCOMPLETE names.
  const src = join(dir, "paper.docx");
  await copyFile(join(FIXTURES, "review_manuscript.docx"), src);
  const failing = async (script, args) => {
    const r = await realPy(script, args);
    return script.endsWith("verify-anonymized.py") ? { ...r, ok: false, code: 1 } : r;
  };
  const one = await runAnonymize({ src, runPy: failing, exists: existsSync, now: NOW, root: ROOT });
  const two = await runAnonymize({ src, runPy: failing, exists: existsSync, now: NOW, root: ROOT });
  assert.deepEqual([one.ok, two.ok], [false, false]);
  assert.match(one.path, /INCOMPLETE/);
  assert.match(two.path, /INCOMPLETE/);
  assert.notEqual(one.path, two.path);
  assert.ok(existsSync(one.path) && existsSync(two.path));
});

test("an upload does not replace a file that is already in the upload folder", async () => {
  const dir = await newDir();
  await writeFile(join(dir, "base.docx"), "already here");
  const uploads = createUploads({ dir });
  const [only] = toChunks(Buffer.from("new bytes"), 64);
  const r = await uploads.accept({ upload_id: "base", ...only });
  assert.deepEqual([r.ok, /already/.test(r.error)], [false, true]);
  assert.equal(String(await readFile(join(dir, "base.docx"))), "already here");
  assert.equal(uploads.pathOf("base"), null);
});
