// Rows 8 and 9 of the submission checklist: the marked copy and the
// anonymized copy against their record (daemon/submission-check.mjs,
// aux-service.mjs marked_build and anonymize, revision.artifacts).
import { test } from "node:test";
import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFile, readdir, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { submissionRows, submissionPaths, MANUSCRIPT } from "../daemon/submission-check.mjs";
import { RevisionSchema, saveRevision, setArtifact } from "../daemon/revision-store.mjs";
import { withDefaults } from "../daemon/journal-profile.mjs";
import { FILE, liveOf, recordOf, profileOf, rowOf } from "./helpers/submission-fixture.mjs";
import { world, ACCEPTED } from "./helpers/submission-world.mjs";

const sha = (text) => createHash("sha256").update(text).digest("hex");
const SOURCE = "a".repeat(64);
const COPY = "b".repeat(64);
const artifact = (path, extra = {}) => ({
  path,
  sha256: COPY,
  at: "2026-03-01T10:00:00.000Z",
  source_sha256: SOURCE,
  ...extra,
});
const live = liveOf();
const rows = ({ artifacts = null, files = {}, profile = profileOf() } = {}) =>
  submissionRows({
    revision: recordOf((r) => (r.artifacts = artifacts)),
    profile,
    live,
    files,
    file: FILE,
  });
const MARKED = { marked: artifact("paper_marked.docx", { mode: "track_changes" }) };
const ASKS = profileOf({ submission: { anonymized_copy: true } });

test("row 8: ready while the copy is the file that was made and the manuscript the file it was made from", () => {
  const r = rowOf(
    rows({ artifacts: MARKED, files: { "paper_marked.docx": COPY, [MANUSCRIPT]: SOURCE } }),
    "marked_copy",
  );
  assert.equal(r.state, "ready");
  assert.equal(
    r.text,
    "The marked copy was made from the manuscript as it is saved now: paper_marked.docx.",
  );
  assert.deepEqual(submissionPaths(recordOf((rev) => (rev.artifacts = MARKED))), [
    "paper_marked.docx",
  ]);
});

test("row 8: not ready after the manuscript bytes change; the sentence puts the cleanup first", () => {
  const r = rowOf(
    rows({ artifacts: MARKED, files: { "paper_marked.docx": COPY, [MANUSCRIPT]: "c".repeat(64) } }),
    "marked_copy",
  );
  assert.equal(r.state, "not_ready");
  assert.equal(
    r.text,
    "The manuscript was saved again after paper_marked.docx was made (removing the tool's comments counts too). Finish the cleanup above, save, then make the marked copy again.",
  );
  assert.deepEqual(r.go, { tab: "tools", target: "tools-marked-build" });
  // A manuscript that could not be hashed is not "the same".
  assert.equal(
    rowOf(rows({ artifacts: MARKED, files: { "paper_marked.docx": COPY } }), "marked_copy").state,
    "not_ready",
  );
});

test("row 8: no record, a copy that is gone and a copy that was changed are each not ready", () => {
  const none = rowOf(rows(), "marked_copy");
  assert.equal(none.state, "not_ready");
  assert.equal(
    none.text,
    "No marked copy is on record for this file. Finish the cleanup above and save the manuscript first (removing comments changes the saved file), then make the copy.",
  );
  const gone = rowOf(rows({ artifacts: MARKED, files: { [MANUSCRIPT]: SOURCE } }), "marked_copy");
  assert.equal(gone.state, "not_ready");
  assert.equal(gone.text, "The marked copy is gone: paper_marked.docx. Make it again.");
  const edited = rowOf(
    rows({ artifacts: MARKED, files: { "paper_marked.docx": SOURCE, [MANUSCRIPT]: SOURCE } }),
    "marked_copy",
  );
  assert.equal(edited.state, "not_ready");
  assert.match(edited.text, /^paper_marked\.docx was changed after it was made/);
});

const UNSAVED =
  "This document has changes that are not saved. Save it, then press Check now again.";
const NOT_CHECKED = " Whether the document has unsaved changes could not be checked.";
const COPIES = {
  marked: artifact("paper_marked.docx", { mode: "track_changes" }),
  anonymized: artifact("paper_anonymized.docx"),
};
const COPY_FILES = {
  "paper_marked.docx": COPY,
  "paper_anonymized.docx": COPY,
  [MANUSCRIPT]: SOURCE,
};
const copyRows = (saved) =>
  submissionRows({
    revision: recordOf((r) => (r.artifacts = COPIES)),
    profile: ASKS,
    live: { ...live, saved },
    files: COPY_FILES,
    file: FILE,
  });

test("rows 8 and 9: a document with unsaved changes is not ready although the copies match the saved file", () => {
  // Saved: both ready (the copies were made from the file on the disk).
  for (const id of ["marked_copy", "anonymized_copy"])
    assert.equal(rowOf(copyRows(true), id).state, "ready", id);
  // The author accepted changes in Word and did not save: the file on the
  // disk still hashes as it did, and the copy is old all the same.
  for (const id of ["marked_copy", "anonymized_copy"]) {
    const r = rowOf(copyRows(false), id);
    assert.equal(r.state, "not_ready", id);
    assert.equal(r.text, UNSAVED);
    assert.equal(r.go, null);
  }
  // The other rows do not depend on it.
  const rest = (rows) => rows.filter((r) => !/_copy$/.test(r.id));
  assert.deepEqual(rest(copyRows(false)), rest(copyRows(true)));
  // An anonymized copy the journal does not ask for stays not needed.
  const plain = submissionRows({
    revision: recordOf((r) => (r.artifacts = COPIES)),
    profile: profileOf(),
    live: { ...live, saved: false },
    files: COPY_FILES,
    file: FILE,
  });
  assert.equal(rowOf(plain, "anonymized_copy").state, "not_applicable");
});

test("rows 8 and 9: when Word did not say whether the document is saved, the rows are as before and say so", () => {
  for (const saved of [null, undefined, "true"]) {
    const marked = rowOf(copyRows(saved), "marked_copy");
    assert.equal(marked.state, "ready");
    assert.equal(
      marked.text,
      `The marked copy was made from the manuscript as it is saved now: paper_marked.docx.${NOT_CHECKED}`,
    );
    assert.equal(rowOf(copyRows(saved), "anonymized_copy").text.endsWith(NOT_CHECKED), true);
  }
  // A row that is not ready for another reason keeps its sentence and gets the note too.
  const gone = submissionRows({
    revision: recordOf((r) => (r.artifacts = COPIES)),
    profile: ASKS,
    live: { ...live, saved: null },
    files: { [MANUSCRIPT]: SOURCE },
    file: FILE,
  });
  assert.equal(
    rowOf(gone, "marked_copy").text,
    `The marked copy is gone: paper_marked.docx. Make it again.${NOT_CHECKED}`,
  );
});

test("Check now with unsaved changes: the service hands the flag on and answers it", async () => {
  const t = await world();
  try {
    assert.equal((await build(t)).ok, true);
    const saved = await t.ask("submission_check", { ...t.withDoc(ACCEPTED), saved: true });
    assert.equal(saved.saved, true);
    assert.equal(saved.rows.find((r) => r.id === "marked_copy").state, "ready");
    const unsaved = await t.ask("submission_check", { ...t.withDoc(ACCEPTED), saved: false });
    assert.equal(unsaved.saved, false);
    const row = unsaved.rows.find((r) => r.id === "marked_copy");
    assert.equal(row.state, "not_ready");
    assert.equal(row.text, UNSAVED);
    // A pane that sends nothing (an older one): not known, never taken for saved.
    const { saved: _, ...older } = t.withDoc(ACCEPTED);
    const old = await t.ask("submission_check", older);
    assert.equal(old.saved, null);
    assert.equal(old.rows.find((r) => r.id === "marked_copy").text.endsWith(NOT_CHECKED), true);
  } finally {
    await t.cleanup();
  }
});

test("the rows of the cleanup stand before the marked copy", () => {
  const ids = rows().map((r) => r.id);
  for (const id of ["changes_decided", "tool_comments", "file_clean"])
    assert.ok(ids.indexOf(id) < ids.indexOf("marked_copy"), id);
});

test("row 9: not applicable when the profile says null or false", () => {
  assert.equal(withDefaults({}).submission.anonymized_copy, null);
  const unsaid = rowOf(rows(), "anonymized_copy");
  assert.equal(unsaid.state, "not_applicable");
  assert.equal(unsaid.text, "The journal profile does not say that an anonymized copy is needed.");
  const no = rowOf(
    rows({ profile: profileOf({ submission: { anonymized_copy: false } }) }),
    "anonymized_copy",
  );
  assert.equal(no.state, "not_applicable");
  assert.equal(no.text, "The journal profile asks for no anonymized copy.");
  // A record of a copy does not make the row applicable.
  const made = { anonymized: artifact("paper_anonymized.docx") };
  assert.equal(rowOf(rows({ artifacts: made }), "anonymized_copy").state, "not_applicable");
});

test("row 9: the profile asks for one: not ready without a record, ready with a copy that matches", () => {
  const none = rowOf(rows({ profile: ASKS }), "anonymized_copy");
  assert.equal(none.state, "not_ready");
  assert.match(none.text, /^No anonymized copy is on record for this file\./);
  assert.deepEqual(none.go, { tab: "tools", target: "tools-anonymize" });

  const artifacts = { anonymized: artifact("paper_anonymized.docx") };
  const ok = rowOf(
    rows({
      profile: ASKS,
      artifacts,
      files: { "paper_anonymized.docx": COPY, [MANUSCRIPT]: SOURCE },
    }),
    "anonymized_copy",
  );
  assert.equal(ok.state, "ready");
  const stale = rowOf(
    rows({
      profile: ASKS,
      artifacts,
      files: { "paper_anonymized.docx": COPY, [MANUSCRIPT]: COPY },
    }),
    "anonymized_copy",
  );
  assert.equal(stale.state, "not_ready");
});

// ---- through the service: the copy is made, then things change -------------

const BASELINE = Buffer.from("earlier version").toString("base64");
const build = (t) => t.ask("marked_build", { mode: "track_changes", baseline_base64: BASELINE });

test("a marked copy that was made is put on record with both hashes, and the row is ready", async () => {
  const t = await world();
  try {
    assert.equal((await t.check()).marked_copy.state, "not_ready");
    const r = await build(t);
    assert.equal(r.ok, true, r.error);
    const copy = join(t.dir, "paper_marked.docx");
    assert.equal(r.path, copy);
    const { marked } = (await t.rev()).artifacts;
    assert.deepEqual(
      { ...marked, at: "t" },
      {
        path: "paper_marked.docx",
        sha256: sha(await readFile(copy)),
        at: "t",
        source_sha256: sha(await readFile(t.doc)),
        mode: "track_changes",
      },
    );
    const rows = await t.check();
    assert.equal(rows.marked_copy.state, "ready", rows.marked_copy.text);
  } finally {
    await t.cleanup();
  }
});

test("the manuscript saved again after the copy was made: the row is not ready until the copy is made again", async () => {
  const t = await world();
  try {
    await build(t);
    assert.equal((await t.check()).marked_copy.state, "ready");
    // Word saves the manuscript after the tool's comments were removed.
    await writeFile(t.doc, "manuscript as saved, comments removed");
    const stale = (await t.check()).marked_copy;
    assert.equal(stale.state, "not_ready");
    assert.match(stale.text, /^The manuscript was saved again after paper_marked\.docx was made/);
    const again = await build(t);
    assert.equal(again.ok, true, again.error);
    const fresh = (await t.check()).marked_copy;
    assert.equal(fresh.state, "ready", fresh.text);
    assert.notEqual(again.path, join(t.dir, "paper_marked.docx"), "the earlier copy is kept");
  } finally {
    await t.cleanup();
  }
});

test("a marked copy edited or deleted afterwards is not ready", async () => {
  const t = await world();
  try {
    const r = await build(t);
    await writeFile(r.path, "edited by hand");
    assert.match((await t.check()).marked_copy.text, /was changed after it was made/);
    const { rm } = await import("node:fs/promises");
    await rm(r.path);
    assert.match((await t.check()).marked_copy.text, /^The marked copy is gone/);
  } finally {
    await t.cleanup();
  }
});

test("a build the tool refuses records nothing", async () => {
  const t = await world({
    tools: {
      "build_marked_diff.py": async () => ({
        ok: false,
        code: 1,
        json: { written: false, checks: 3, failures: 1 },
      }),
    },
  });
  try {
    const r = await build(t);
    assert.equal(r.ok, false);
    assert.match(r.error, /did not pass its own check/);
    assert.equal((await t.rev()).artifacts ?? null, null);
    assert.equal((await t.check()).marked_copy.state, "not_ready");
  } finally {
    await t.cleanup();
  }
});

test("a manuscript that changed while the copy was being made is not put on record", async () => {
  let doc;
  const t = await world({
    tools: {
      "build_marked_diff.py": async (args, arg) => {
        await writeFile(arg("--out"), "marked copy of the earlier save");
        await writeFile(doc, "saved again while the tool ran");
        return { ok: true, code: 0, json: { written: true, checks: 3, failures: 0 } };
      },
    },
  });
  doc = t.doc;
  try {
    const r = await build(t);
    assert.equal(r.ok, true, r.error);
    assert.equal((await t.rev()).artifacts ?? null, null);
    assert.equal((await t.check()).marked_copy.state, "not_ready");
  } finally {
    await t.cleanup();
  }
});

const ASK_OVERRIDE = { profile_overrides: { submission: { anonymized_copy: true } } };

test("an anonymized copy the verifier passed is put on record, and the row is ready when the profile asks for one", async () => {
  const t = await world({ revision: ASK_OVERRIDE });
  try {
    assert.equal((await t.check()).anonymized_copy.state, "not_ready");
    const r = await t.ask("anonymize");
    assert.equal(r.ok, true, r.error);
    const { anonymized } = (await t.rev()).artifacts;
    assert.equal(anonymized.path, "paper_anonymized.docx");
    assert.equal(anonymized.sha256, sha(await readFile(r.path)));
    assert.equal(anonymized.source_sha256, sha(await readFile(t.doc)));
    assert.equal("mode" in anonymized, false);
    assert.equal((await t.check()).anonymized_copy.state, "ready");
    await writeFile(t.doc, "saved again");
    assert.equal((await t.check()).anonymized_copy.state, "not_ready");
  } finally {
    await t.cleanup();
  }
});

test("an anonymized copy the verifier does not pass is set aside as INCOMPLETE and records nothing", async () => {
  const report = {
    total: 2,
    body_equal: true,
    markup_equal: true,
    structure_ok: true,
    unreadable_parts: [],
    categories: { names: { status: "not cleared", count: 2, default: "clear", what: "creator" } },
  };
  const t = await world({
    revision: ASK_OVERRIDE,
    tools: { "verify-anonymized.py": async () => ({ ok: false, code: 1, json: report }) },
  });
  try {
    const r = await t.ask("anonymize");
    assert.equal(r.ok, false);
    assert.match(r.error, /incomplete/);
    assert.ok((await readdir(t.dir)).some((f) => f.includes(".INCOMPLETE.")));
    assert.equal((await t.rev()).artifacts ?? null, null);
    const row = (await t.check()).anonymized_copy;
    assert.equal(row.state, "not_ready");
    assert.match(row.text, /^No anonymized copy is on record/);
  } finally {
    await t.cleanup();
  }
});

test("a paper without the profile key: the anonymized row is not applicable, with or without a copy", async () => {
  const t = await world();
  try {
    assert.equal((await t.check()).anonymized_copy.state, "not_applicable");
    await t.ask("anonymize");
    assert.equal((await t.check()).anonymized_copy.state, "not_applicable");
  } finally {
    await t.cleanup();
  }
});

// ---- the stored record ------------------------------------------------------

test("setArtifact records one copy, keeps the other, and leaves no empty member behind", () => {
  const base = recordOf((r) => delete r.artifacts);
  const a = artifact("paper_marked.docx", { mode: "color" });
  const one = setArtifact(base, "marked", a);
  assert.deepEqual(one.artifacts, { marked: a });
  const two = setArtifact(one, "anonymized", artifact("paper_anonymized.docx"));
  assert.deepEqual(Object.keys(two.artifacts).sort(), ["anonymized", "marked"]);
  assert.deepEqual(setArtifact(two, "marked", null).artifacts, {
    anonymized: artifact("paper_anonymized.docx"),
  });
  assert.equal("artifacts" in setArtifact(one, "marked", null), false);
  assert.equal("artifacts" in base, false, "the input is not changed");
  assert.throws(() => setArtifact(base, "clean", a), /Unknown copy/);
});

test("the fixture sidecar from before the checklist loads, and saves byte for byte", async () => {
  const fixture = new URL("./fixtures/revision_before_figures.json", import.meta.url);
  const raw = JSON.parse(await readFile(fixture, "utf8"));
  assert.equal("artifacts" in raw, false, "the fixture is the old format");
  const t = await world();
  try {
    await writeFile(t.revPath, await readFile(fixture, "utf8"));
    const loaded = await t.rev();
    assert.equal(loaded.version, 1);
    assert.equal("artifacts" in loaded, false);
    // Saved by this daemon: the checklist adds no key to the file. (The keys
    // of the figure lane are added, as before: tests/figure-store.test.mjs.)
    const saved = await saveRevision(t.revPath, loaded, new Date(raw.updated_at));
    const text = await t.raw();
    assert.equal(text.includes("artifacts"), false);
    assert.equal(text.includes("model_sha"), false);
    assert.equal(text.includes("docx_sha"), false);
    for (const c of saved.comments) {
      delete c.figure_changes;
      delete c.figure_seq;
      if (c.response) delete c.response.figure_ids;
    }
    delete saved.figure_lane;
    assert.deepEqual(saved, raw);
    // Loaded and saved once more: byte for byte the same file.
    await saveRevision(t.revPath, await t.rev(), new Date(raw.updated_at));
    assert.equal(await t.raw(), text);
  } finally {
    await t.cleanup();
  }
});

// The schema as it was before the checklist: no `artifacts`. (Everything else
// is the schema of today.)
const BEFORE = RevisionSchema.omit({ artifacts: true });

test("a sidecar with a recorded copy read by the earlier schema loads, the record drops, and the row is then not ready", async () => {
  const t = await world();
  try {
    await build(t);
    assert.equal((await t.check()).marked_copy.state, "ready");
    // An older daemon loads the file and saves it: what it does not know is gone.
    const viaOld = BEFORE.parse(JSON.parse(await t.raw()));
    assert.equal("artifacts" in viaOld, false);
    assert.equal(viaOld.comments.length, 1);
    await writeFile(t.revPath, JSON.stringify(viaOld, null, 2) + "\n");
    const row = (await t.check(ACCEPTED)).marked_copy;
    assert.equal(row.state, "not_ready");
    assert.match(row.text, /^No marked copy is on record/);
  } finally {
    await t.cleanup();
  }
});
