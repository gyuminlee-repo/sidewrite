// The "Make marked copy" button (daemon/aux-service.mjs marked_build): what
// the button writes is what the script writes, and a refusal of the tool
// reaches the pane word for word.
import { test } from "node:test";
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { existsSync } from "node:fs";
import { copyFile, mkdtemp, readFile, readdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { createAuxService, AUX_JOBS, MARKED_MODES } from "../daemon/aux-service.mjs";
import { MARKED_NOTICE } from "../daemon/marked-diff.mjs";
import { DISK_NOTICE } from "../daemon/doc-bytes.mjs";

const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..");
const FIXTURES = join(ROOT, "tests", "fixtures");
const ENV = { CC_OFFICE_WSL_PATHS: "0", CC_OFFICE_CONFIG_DIR: "/nonexistent-config" };

async function setup({ runPy, saved = true } = {}) {
  const dir = await mkdtemp(join(tmpdir(), "mbuild-"));
  const doc = join(dir, "paper.docx");
  await copyFile(join(FIXTURES, "review_manuscript.docx"), doc);
  const events = [];
  const service = createAuxService({
    getContext: () => ({ activeDoc: saved ? doc : null }),
    sendEvent: (key, event) => events.push(event),
    log: () => {},
    env: ENV,
    ...(runPy && { runPy }),
  });
  async function job(msg) {
    let first;
    await service.handlers.marked_build(
      { request_id: "q", ...msg },
      (r) => (first = r),
      "k",
      "word",
    );
    assert.equal(first.accepted, true);
    const mine = () => events.find((e) => e.job_id === first.job_id);
    for (let i = 0; i < 3000 && !mine(); i++) await new Promise((r) => setTimeout(r, 10));
    return mine().result;
  }
  const baseline_base64 = (await readFile(join(FIXTURES, "spike_manuscript.docx"))).toString(
    "base64",
  );
  return { dir, doc, job, baseline_base64, done: () => rm(dir, { recursive: true }) };
}

// The script, run as the README says.
const script = (args) =>
  new Promise((resolve) => {
    execFile(
      process.env.CC_OFFICE_PYTHON || "python3",
      [join(ROOT, "tools", "build_marked_diff.py"), ...args],
      { maxBuffer: 64 << 20 },
      (err) => resolve(err ? err.code : 0),
    );
  });

test("marked_build is a job and takes the two modes of the tool", () => {
  assert.ok(AUX_JOBS.has("marked_build"));
  assert.deepEqual(MARKED_MODES, ["track_changes", "color"]);
});

for (const mode of MARKED_MODES) {
  test(`${mode}: the button writes the file the script writes, next to the manuscript`, async () => {
    const t = await setup();
    try {
      const r = await t.job({ mode, baseline_base64: t.baseline_base64 });
      assert.equal(r.ok, true, r.error);
      assert.equal(r.path, join(t.dir, "paper_marked.docx"));
      assert.equal(r.name, "paper_marked.docx");
      assert.equal(r.mode, mode);
      assert.equal(r.report.mode, mode);
      assert.deepEqual([r.report.written, r.report.failures], [true, 0]);
      assert.equal(r.notice, `${MARKED_NOTICE} ${DISK_NOTICE}`);
      // The manuscript is not touched and the baseline copy is gone.
      assert.deepEqual((await readdir(t.dir)).sort(), ["paper.docx", "paper_marked.docx"]);
      const byScript = join(t.dir, "by_script.docx");
      const code = await script([
        "--baseline",
        join(FIXTURES, "spike_manuscript.docx"),
        "--clean",
        t.doc,
        "--out",
        byScript,
        "--mode",
        mode,
      ]);
      assert.equal(code, 0);
      assert.ok((await readFile(r.path)).equals(await readFile(byScript)));
    } finally {
      await t.done();
    }
  });
}

test("a refusal of the tool (exit 2) is the error of the job, with its paragraph number", async () => {
  const REFUSAL =
    'paragraph 12 ("The cells were"): a text direction mark (w:dir) was added; it is not compared';
  const seen = [];
  const t = await setup({
    runPy: async (script, args) => {
      seen.push({ script, args, baselineThere: existsSync(args[1]) });
      return { ok: false, code: 2, stderr: `reading\n${REFUSAL}\n`, error: REFUSAL };
    },
  });
  try {
    const r = await t.job({ mode: "color", baseline_base64: t.baseline_base64 });
    assert.deepEqual(r, { ok: false, error: REFUSAL });
    assert.match(seen[0].script, /tools[\\/]build_marked_diff\.py$/);
    const a = seen[0].args;
    assert.deepEqual(
      [a[0], a[2], a[3], a[4], a[5], a[6], a.slice(-2)],
      [
        "--baseline",
        "--clean",
        t.doc,
        "--out",
        join(t.dir, "paper_marked.docx"),
        "--json",
        ["--mode", "color"],
      ],
    );
    // The earlier version was a temporary file of the job, removed after it.
    assert.equal(seen[0].baselineThere, true);
    assert.equal(existsSync(a[1]), false);
    assert.deepEqual(await readdir(t.dir), ["paper.docx"]);
  } finally {
    await t.done();
  }
});

test("no earlier version, an unknown mode or an unsaved document: an error and the tool is not run", async () => {
  let ran = 0;
  const runPy = async () => (ran++, { ok: true, code: 0, json: { written: true } });
  let t = await setup({ runPy });
  try {
    assert.match((await t.job({ mode: "color" })).error, /Choose the earlier version/);
    assert.match(
      (await t.job({ mode: "redline", baseline_base64: t.baseline_base64 })).error,
      /Choose how the changes are shown/,
    );
  } finally {
    await t.done();
  }
  t = await setup({ runPy, saved: false });
  try {
    const r = await t.job({ mode: "color", baseline_base64: t.baseline_base64 });
    assert.match(r.error, /Save the document to a local folder first/);
  } finally {
    await t.done();
  }
  assert.equal(ran, 0);
});
