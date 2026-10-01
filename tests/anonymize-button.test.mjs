// The "Make anonymized copy" button (daemon/aux-service.mjs anonymize): the
// copy is written from the saved file and the verifier's report comes back.
import { test } from "node:test";
import assert from "node:assert/strict";
import { copyFile, mkdtemp, readFile, readdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { createAuxService, AUX_JOBS } from "../daemon/aux-service.mjs";
import { ANONYMIZE_NOTICE } from "../daemon/anonymize.mjs";

const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..");
const SOURCE = join(ROOT, "tests", "fixtures", "review_manuscript.docx");
const ENV = { CC_OFFICE_WSL_PATHS: "0", CC_OFFICE_CONFIG_DIR: "/nonexistent-config" };

async function setup({ runPy, saved = true } = {}) {
  const dir = await mkdtemp(join(tmpdir(), "anon-"));
  const doc = join(dir, "paper.docx");
  await copyFile(SOURCE, doc);
  const events = [];
  const service = createAuxService({
    getContext: () => ({ activeDoc: saved ? doc : null }),
    sendEvent: (key, event) => events.push(event),
    log: () => {},
    env: ENV,
    ...(runPy && { runPy }),
  });
  async function job() {
    let first;
    await service.handlers.anonymize({ request_id: "q" }, (r) => (first = r), "k", "word");
    assert.equal(first.accepted, true);
    const mine = () => events.find((e) => e.job_id === first.job_id);
    for (let i = 0; i < 3000 && !mine(); i++) await new Promise((r) => setTimeout(r, 10));
    return mine().result;
  }
  return { dir, doc, job, done: () => rm(dir, { recursive: true }) };
}

test("with the real tools: the copy is next to the manuscript and the verifier counts 0", async () => {
  assert.ok(AUX_JOBS.has("anonymize"));
  const t = await setup();
  try {
    const r = await t.job();
    assert.equal(r.ok, true, r.error);
    assert.equal(r.path, join(t.dir, "paper_anonymized.docx"));
    assert.equal(r.name, "paper_anonymized.docx");
    assert.equal(r.report.total, 0);
    assert.deepEqual(r.report.not_cleared, []);
    assert.deepEqual(
      [r.report.body_equal, r.report.markup_equal, r.report.structure_ok],
      [true, true, true],
    );
    assert.equal(r.notice, ANONYMIZE_NOTICE);
    // The manuscript itself is as it was.
    assert.ok((await readFile(t.doc)).equals(await readFile(SOURCE)));
    assert.deepEqual((await readdir(t.dir)).sort(), ["paper.docx", "paper_anonymized.docx"]);
  } finally {
    await t.done();
  }
});

test("a copy the verifier does not pass: ok is false, the report and the set-aside name come with it", async () => {
  const report = {
    total: 2,
    body_equal: true,
    markup_equal: true,
    structure_ok: true,
    unreadable_parts: [],
    categories: { names: { status: "not cleared", count: 2, default: "clear", what: "creator" } },
  };
  const t = await setup({
    runPy: async (script, args) => {
      if (/anonymize_docx\.py$/.test(script)) {
        await copyFile(args[1], args[3]);
        return { ok: true, code: 0, json: { changed: {} } };
      }
      return { ok: false, code: 1, json: report };
    },
  });
  try {
    const r = await t.job();
    assert.equal(r.ok, false);
    assert.equal(r.name, "paper_anonymized.INCOMPLETE.docx");
    assert.equal(r.report.total, 2);
    assert.equal(r.error, "The copy is incomplete: still present: names (2).");
    assert.deepEqual((await readdir(t.dir)).sort(), [
      "paper.docx",
      "paper_anonymized.INCOMPLETE.docx",
    ]);
  } finally {
    await t.done();
  }
});

test("a tool that fails, or a document that is not saved, is an error with its reason", async () => {
  let t = await setup({
    runPy: async () => ({ ok: false, code: 2, stderr: "cannot read the input: bad zip\n" }),
  });
  try {
    assert.deepEqual(await t.job(), { ok: false, error: "cannot read the input: bad zip" });
  } finally {
    await t.done();
  }
  t = await setup({ saved: false });
  try {
    assert.match((await t.job()).error, /Save the document to a local folder first/);
  } finally {
    await t.done();
  }
});
