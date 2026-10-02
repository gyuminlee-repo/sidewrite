// Row 10 of the submission checklist: the adopted figures and other files
// against the copies in the submission folder (daemon/submission-check.mjs
// figures row; the record is the figure lane's, revision-store.mjs).
import { test } from "node:test";
import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdir, rm, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { submissionRows, submissionPaths } from "../daemon/submission-check.mjs";
import { FILE, liveOf, recordOf, profileOf, rowOf } from "./helpers/submission-fixture.mjs";
import { world, commentsOf } from "./helpers/submission-world.mjs";

const sha = (text) => createHash("sha256").update(text).digest("hex");
const HASH = (c) => c.repeat(64);
const AT = "2026-02-01T00:00:00.000Z";
const fileRef = (path, sha256) => ({ path, sha256, bytes: 10, mtime: AT });

// A figure change as the figure lane records it. copy: the adopted copy.
const change = (id, figure, state, copy = null) => ({
  id,
  figure,
  kind: "revised",
  requirement_ids: [],
  summary: "The control curve was added.",
  before: { source: fileRef("figures/fig2_before.svg", HASH("a")) },
  after: { source: fileRef("figures/fig2_after.svg", copy?.sha256 ?? HASH("b")) },
  state,
  adopted: state === "adopted" ? { ...copy, at: AT } : null,
  created_at: AT,
  decided_at: state === "proposed" ? null : AT,
});
const COPY = { path: "figures/submission/fig2_after.svg", sha256: HASH("b") };
const TABLE = { path: "figures/submission/table_s3.xlsx", sha256: HASH("c") };

const live = liveOf();
const figures = (changes, files = {}) =>
  rowOf(
    submissionRows({
      revision: recordOf((r) => (r.comments[0].figure_changes = changes)),
      profile: profileOf(),
      live,
      files,
      file: FILE,
    }),
    "figures",
  );

test("row 10: not applicable without a figure or file change, and with rejected ones only", () => {
  const none = figures([]);
  assert.equal(none.state, "not_applicable");
  assert.equal(none.text, "No figure or file change is on record for this paper.");
  assert.equal(none.go, null);
  assert.equal(figures([change("R1.1-f1", "Figure 2", "rejected")]).state, "not_applicable");
});

test("row 10: ready when every adopted copy is in the submission folder with its recorded hash", () => {
  const changes = [
    change("R1.1-f1", "Figure 2", "adopted", COPY),
    change("R1.1-f2", "Table S3", "adopted", TABLE),
    change("R1.1-f3", "Figure 5", "rejected"),
  ];
  const r = figures(changes, { [COPY.path]: COPY.sha256, [TABLE.path]: TABLE.sha256 });
  assert.equal(r.state, "ready");
  assert.equal(r.text, "2 adopted figures or files are in the submission folder, unchanged.");
  assert.deepEqual(submissionPaths(recordOf((rev) => (rev.comments[0].figure_changes = changes))), [
    COPY.path,
    TABLE.path,
  ]);
  const one = figures([changes[0]], { [COPY.path]: COPY.sha256 });
  assert.equal(one.text, "1 adopted figure or file is in the submission folder, unchanged.");
});

test("row 10: an adopted copy edited on disk is not ready, names the file and leads to the figures page", () => {
  const r = figures([change("R1.1-f1", "Figure 2", "adopted", COPY)], {
    [COPY.path]: HASH("e"),
  });
  assert.equal(r.state, "not_ready");
  assert.equal(
    r.text,
    "The copy of Figure 2 was changed after it was adopted: fig2_after.svg. Settle it on the figures page; a copy that is missing or changed is adopted again there.",
  );
  assert.deepEqual(r.go, { figures: true, ref: "R1.1" });
});

test("row 10: an adopted copy that is gone is not ready and names the file", () => {
  for (const files of [{ [COPY.path]: null }, {}]) {
    const r = figures([change("R1.1-f1", "Figure 2", "adopted", COPY)], files);
    assert.equal(r.state, "not_ready");
    assert.match(
      r.text,
      /^The copy of Figure 2 is missing from the submission folder: fig2_after\.svg\./,
    );
  }
});

test("row 10: a proposed change is not ready and is named, also next to adopted ones that are fine", () => {
  const r = figures(
    [change("R1.1-f1", "Figure 2", "adopted", COPY), change("R1.1-f2", "Table S3", "proposed")],
    { [COPY.path]: COPY.sha256 },
  );
  assert.equal(r.state, "not_ready");
  assert.match(r.text, /^Table S3 is not decided yet \(adopt or reject it\)\. Settle it on/);
  assert.deepEqual(r.go, { figures: true, ref: "R1.1" });
});

// ---- through the service: the files on the disk -----------------------------

test("an adopted copy overwritten by hand in the submission folder makes the row not ready", async () => {
  const body = "<svg>after</svg>\n";
  const copy = { path: "figures/submission/fig2_after.svg", sha256: sha(body) };
  const t = await world({
    seed: commentsOf({
      figure_changes: [change("R1.1-f1", "Figure 2", "adopted", copy)],
      figure_seq: 1,
    }),
  });
  try {
    await mkdir(join(t.dir, "figures", "submission"), { recursive: true });
    const file = join(t.dir, "figures", "submission", "fig2_after.svg");
    await writeFile(file, body);
    const ok = (await t.check()).figures;
    assert.equal(ok.state, "ready", ok.text);

    await writeFile(file, "<svg>edited after it was adopted</svg>\n");
    const edited = (await t.check()).figures;
    assert.equal(edited.state, "not_ready");
    assert.match(
      edited.text,
      /^The copy of Figure 2 was changed after it was adopted: fig2_after\.svg/,
    );

    await rm(file);
    assert.match((await t.check()).figures.text, /is missing from the submission folder/);
  } finally {
    await t.cleanup();
  }
});

test("a proposed figure change on record makes the row not ready; none at all is not applicable", async () => {
  const t = await world({
    seed: commentsOf({
      figure_changes: [change("R1.1-f1", "Figure 2", "proposed")],
      figure_seq: 1,
    }),
  });
  const plain = await world();
  try {
    const row = (await t.check()).figures;
    assert.equal(row.state, "not_ready");
    assert.match(row.text, /^Figure 2 is not decided yet/);
    assert.equal((await plain.check()).figures.state, "not_applicable");
  } finally {
    await t.cleanup();
    await plain.cleanup();
  }
});

test("the checklist ships ten rows, in the order the author works through them", async () => {
  const t = await world();
  try {
    assert.deepEqual(Object.keys(await t.check()), [
      "answers",
      "requests",
      "changes_decided",
      "tool_comments",
      "file_clean",
      "format_counts",
      "letter",
      "marked_copy",
      "anonymized_copy",
      "figures",
    ]);
  } finally {
    await t.cleanup();
  }
});
