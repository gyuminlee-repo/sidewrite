// Row 7 of the submission checklist: the exported letter against the record
// as it is now, by hash (daemon/submission-check.mjs letter row,
// response-model.mjs letterModelSha, response-service.mjs response_export).
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile, rm, writeFile } from "node:fs/promises";
import { submissionRows, submissionPaths } from "../daemon/submission-check.mjs";
import { buildLetterModel, letterModelSha } from "../daemon/response-model.mjs";
import { RevisionSchema, saveRevision } from "../daemon/revision-store.mjs";
import { FILE, liveOf, recordOf, profileOf, rowOf } from "./helpers/submission-fixture.mjs";
import { world, ACCEPTED, REJECTED } from "./helpers/submission-world.mjs";

const LETTER = "/m/paper_response.docx";
const DOCX_SHA = "d".repeat(64);
const profile = profileOf();
const live = liveOf();

// The record right after an export that passed: the hashes of what was
// written. then: what happened to the record afterwards.
function exported({ then = () => {}, stored = {} } = {}) {
  const revision = recordOf();
  revision.response_export = {
    path: LETTER,
    exported_at: "2026-03-01T10:00:00.000Z",
    fail: 0,
    warn: 0,
    model_sha: letterModelSha(buildLetterModel(revision, profile, { gloss: false })),
    docx_sha: DOCX_SHA,
    ...stored,
  };
  then(revision);
  return revision;
}
const letter = (revision, files = { [LETTER]: DOCX_SHA }) =>
  rowOf(submissionRows({ revision, profile, live, files, file: FILE }), "letter");

test("the letter model hash: equal for two key orders, and a stored value for a fixed letter", () => {
  const a = {
    paragraphs: [
      { kind: "reviewer_head", text: "Reviewer #1", color: null, ref: null, reviewer: 1 },
      { kind: "comment_head", text: "Comment 1", color: null, ref: "R1.1", reviewer: 1 },
      { kind: "comment", text: "Please state the unit.", color: null, ref: "R1.1", reviewer: 1 },
      {
        kind: "response",
        text: "We now state the unit (30 °C).",
        color: "C00000",
        ref: "R1.1",
        reviewer: 1,
      },
      {
        kind: "location",
        text: "[Revised in 2. Materials and Methods, line 9]",
        color: "0070C0",
        ref: "R1.1",
        reviewer: 1,
        removed: false,
      },
      {
        kind: "quote",
        text: "The incubation temperature was held at 30 °C.",
        color: "0070C0",
        ref: "R1.1",
        reviewer: 1,
      },
    ],
  };
  // The same letter with the keys of every paragraph in another order, a key
  // of a later version added and the keys the hash does not read left out.
  const b = {
    paragraphs: a.paragraphs.map((p) => ({
      later_key: "x",
      ref: p.ref,
      color: p.color,
      text: p.text,
      kind: p.kind,
    })),
  };
  // sha256 of the JSON list of [kind, text, color, ref], computed outside
  // this code (Python hashlib over json.dumps with compact separators).
  const STORED = "b718e945ed01160dfb38dd0cac77ef782397f33ae1476d32f20946189d40076f";
  assert.equal(letterModelSha(a), STORED);
  assert.equal(letterModelSha(b), STORED);
  // Each of the four values counts, and so does the order of the paragraphs.
  for (const [key, value] of [
    ["kind", "gloss"],
    ["text", "We now state the unit."],
    ["color", "000000"],
    ["ref", "R1.2"],
  ]) {
    const c = structuredClone(a);
    c.paragraphs[3][key] = value;
    assert.notEqual(letterModelSha(c), STORED, key);
  }
  const swapped = structuredClone(a);
  swapped.paragraphs.reverse();
  assert.notEqual(letterModelSha(swapped), STORED);
  assert.equal(
    letterModelSha({ paragraphs: [] }),
    "4f53cda18c2baa0c0354bb5f9a3ecbe5ed12ab4d8e11ba873c2f11161202b945",
  );
});

test("row 7: ready when the model built now hashes as the exported one and the file is the file written", () => {
  const r = letter(exported());
  assert.equal(r.state, "ready");
  assert.equal(
    r.text,
    "The exported letter says what the answers and the accepted changes say now: paper_response.docx.",
  );
  assert.equal(r.go, null);
  assert.deepEqual(submissionPaths(exported()), [LETTER]);
});

test("row 7: no export yet is not ready and leads to the export button", () => {
  const r = letter(recordOf());
  assert.equal(r.state, "not_ready");
  assert.equal(r.text, "The response letter (.docx) has not been exported yet.");
  assert.deepEqual(r.go, { tab: "response", target: "response-export" });
});

test("row 7: an answer edited after the export is not ready", () => {
  const r = letter(
    exported({ then: (rev) => (rev.comments[0].response.text = "We state the unit now.") }),
  );
  assert.equal(r.state, "not_ready");
  assert.equal(
    r.text,
    "The answers, the accepted changes or the line numbers changed after paper_response.docx was exported. Export the letter again.",
  );
});

test("row 7: a citation that left the letter (its change was rejected) is not ready", () => {
  const r = letter(exported({ then: (rev) => (rev.comments[0].response.citations = []) }));
  assert.equal(r.state, "not_ready");
  assert.match(r.text, /changed after paper_response\.docx was exported/);
});

test("row 7: new line numbers after the export are not ready", () => {
  const r = letter(
    exported({
      then: (rev) =>
        (rev.comments[0].response.citations[0].lines = {
          start: 12,
          end: 12,
          start_page: 1,
          end_page: 1,
          match_count: 1,
        }),
    }),
  );
  assert.equal(r.state, "not_ready");
});

test("row 7: a letter file that is gone, or is no longer the file written, is not ready", () => {
  const gone = letter(exported(), { [LETTER]: null });
  assert.equal(gone.state, "not_ready");
  assert.equal(gone.text, "The exported letter is gone: paper_response.docx. Export it again.");
  // A path nobody hashed counts as gone: never ready on a missing value.
  assert.equal(letter(exported(), {}).state, "not_ready");
  const edited = letter(exported(), { [LETTER]: "e".repeat(64) });
  assert.equal(edited.state, "not_ready");
  assert.equal(
    edited.text,
    "paper_response.docx was changed after it was exported, so it is no longer the letter the tool wrote. Export it again.",
  );
});

test("row 7: an export that was forced while checks failed is not ready", () => {
  const r = letter(exported({ stored: { fail: 2 } }));
  assert.equal(r.state, "not_ready");
  assert.equal(
    r.text,
    "paper_response.docx was exported while 2 checks failed. Fix them in the Response tab, then export the letter again.",
  );
});

test("row 7: a record without the hashes (an older daemon wrote or saved it) reads export again", () => {
  for (const stored of [
    { model_sha: null, docx_sha: null },
    { model_sha: null },
    { docx_sha: null },
  ]) {
    const r = letter(exported({ stored }));
    assert.equal(r.state, "not_ready");
    assert.match(r.text, /^Export the letter again: paper_response\.docx was written before/);
  }
});

// ---- through the services: an export, then what happens afterwards ---------

test("an export records the two hashes, and the row is ready right after it", async () => {
  const t = await world();
  try {
    await t.answered();
    const out = await t.ask("response_export", t.withDoc(ACCEPTED));
    assert.equal(out.exported, true, JSON.stringify(out.checks));
    const x = (await t.rev()).response_export;
    assert.match(x.model_sha, /^[0-9a-f]{64}$/);
    assert.match(x.docx_sha, /^[0-9a-f]{64}$/);
    // The model that was written is the model that was hashed.
    const written = JSON.parse(await readFile(out.paths.model, "utf8"));
    assert.equal(x.model_sha, letterModelSha(written));
    const rows = await t.check(ACCEPTED);
    assert.equal(rows.letter.state, "ready", rows.letter.text);
  } finally {
    await t.cleanup();
  }
});

test("a change rejected in Word after the export makes the letter not ready, with no export or refresh in between", async () => {
  const t = await world();
  try {
    await t.answered();
    await t.ask("response_export", t.withDoc(ACCEPTED));
    assert.equal((await t.check(ACCEPTED)).letter.state, "ready");
    const rows = await t.check(REJECTED);
    assert.equal(rows.letter.state, "not_ready");
    assert.match(rows.letter.text, /changed after paper\.response\.docx was exported/);
    assert.deepEqual(rows.letter.go, { tab: "response", target: "response-export" });
  } finally {
    await t.cleanup();
  }
});

test("an answer saved after the export, a letter edited by hand and a letter deleted each make the row not ready", async () => {
  const t = await world();
  try {
    await t.answered();
    const out = await t.ask("response_export", t.withDoc(ACCEPTED));
    await writeFile(out.path, "edited in Word");
    assert.match((await t.check(ACCEPTED)).letter.text, /was changed after it was exported/);
    await rm(out.path);
    assert.match((await t.check(ACCEPTED)).letter.text, /^The exported letter is gone/);

    const again = await t.ask("response_export", { ...t.withDoc(ACCEPTED), overwrite: true });
    assert.equal(again.exported, true);
    assert.equal((await t.check(ACCEPTED)).letter.state, "ready");
    const saved = await t.ask("response_save", {
      ref: "R1.1",
      text: "The unit is now stated.",
      ...t.withDoc(ACCEPTED),
    });
    assert.equal(saved.ok, true, saved.error);
    const rows = await t.check(ACCEPTED);
    assert.equal(rows.letter.state, "not_ready");
    assert.match(rows.letter.text, /changed after .* was exported/);
  } finally {
    await t.cleanup();
  }
});

test("a forced export with a failing check is recorded and is not ready", async () => {
  const t = await world();
  try {
    // No answer saved: the `answered` check fails.
    const held = await t.ask("response_export", t.withDoc(ACCEPTED));
    assert.equal(held.exported, false);
    const forced = await t.ask("response_export", { ...t.withDoc(ACCEPTED), force: true });
    assert.equal(forced.exported, true);
    const rows = await t.check(ACCEPTED);
    assert.equal(rows.letter.state, "not_ready");
    assert.match(rows.letter.text, /was exported while \d+ checks? failed/);
  } finally {
    await t.cleanup();
  }
});

// ---- the stored record: old files load, an older daemon drops the hashes ---

// The schema as it was before the checklist: response_export without the two
// hashes. (A copy of the earlier shape of that one member; everything else
// is the schema of today.)
const BEFORE = RevisionSchema.extend({
  response_export: RevisionSchema.shape.response_export
    .unwrap()
    .unwrap()
    .omit({ model_sha: true, docx_sha: true })
    .nullable()
    .default(null),
});

test("a sidecar from before the hashes loads and saves unchanged, and the row reads export again", async () => {
  const t = await world();
  try {
    await t.answered();
    await t.ask("response_export", t.withDoc(ACCEPTED));
    // The file as the earlier daemon wrote it.
    const old = JSON.parse(await t.raw());
    delete old.response_export.model_sha;
    delete old.response_export.docx_sha;
    const text = JSON.stringify(old, null, 2) + "\n";
    await writeFile(t.revPath, text);

    const loaded = await t.rev();
    assert.equal(loaded.version, 1);
    assert.equal(loaded.response_export.model_sha, undefined);
    const rows = await t.check(ACCEPTED);
    assert.equal(rows.letter.state, "not_ready");
    assert.match(rows.letter.text, /^Export the letter again/);

    // Saved again by this daemon: the file is byte for byte what it was.
    await saveRevision(t.revPath, loaded, new Date(old.updated_at));
    assert.equal(await t.raw(), text);
  } finally {
    await t.cleanup();
  }
});

test("a sidecar with the hashes read by the earlier schema loads, the hashes drop, and the row is then not ready", async () => {
  const t = await world();
  try {
    await t.answered();
    await t.ask("response_export", t.withDoc(ACCEPTED));
    assert.equal((await t.check(ACCEPTED)).letter.state, "ready");
    // An older daemon loads the file and saves it: what it does not know is gone.
    const viaOld = BEFORE.parse(JSON.parse(await t.raw()));
    assert.equal("model_sha" in viaOld.response_export, false);
    assert.equal(viaOld.response_export.fail, 0);
    await writeFile(t.revPath, JSON.stringify(viaOld, null, 2) + "\n");
    const rows = await t.check(ACCEPTED);
    assert.equal(rows.letter.state, "not_ready");
    assert.match(rows.letter.text, /^Export the letter again/);
  } finally {
    await t.cleanup();
  }
});
