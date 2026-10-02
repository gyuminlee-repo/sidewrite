// "Check now" through the services (daemon/aux-service.mjs submission_check):
// the rows are computed from the saved record, the files on disk and the
// document as the pane read it, on an invented paper.
import { test } from "node:test";
import assert from "node:assert/strict";
import { AUX_JOBS } from "../daemon/aux-service.mjs";
import {
  world,
  ACCEPTED,
  PENDING,
  REJECTED,
  snapshotOf,
  AFTER,
} from "./helpers/submission-world.mjs";

test("submission_check is a job, answers the rows of the open file and saves nothing", async () => {
  assert.equal(AUX_JOBS.has("submission_check"), true);
  const t = await world();
  try {
    await t.answered();
    const before = await t.raw();
    const r = await t.ask("submission_check", t.withDoc(ACCEPTED));
    assert.equal(r.ok, true, r.error);
    assert.equal(r.doc_name, "paper.docx");
    assert.deepEqual(r.rows.map((row) => row.id).slice(0, 6), [
      "answers",
      "requests",
      "changes_decided",
      "tool_comments",
      "file_clean",
      "format_counts",
    ]);
    for (const row of r.rows) assert.equal(row.file, "paper.docx");
    assert.equal(await t.raw(), before, "the record on disk is as it was");
  } finally {
    await t.cleanup();
  }
});

test("the rows follow the document as it is now, not the record as it was saved", async () => {
  const t = await world();
  try {
    // Saved while the change was still tracked: the request had no accepted change.
    await t.answered(PENDING);
    const tracked = await t.check(PENDING);
    assert.equal(tracked.changes_decided.state, "not_ready");
    assert.equal(tracked.requests.state, "not_ready");
    assert.equal(tracked.file_clean.state, "not_ready");
    // Accepted in Word since then; no refresh was saved in between.
    const accepted = await t.check(ACCEPTED);
    assert.equal(accepted.changes_decided.state, "ready");
    assert.equal(accepted.requests.state, "ready");
    // Rejected in Word: the request is open again.
    const rejected = await t.check(REJECTED);
    assert.equal(rejected.changes_decided.state, "ready");
    assert.equal(rejected.requests.state, "not_ready");
  } finally {
    await t.cleanup();
  }
});

test("a comment of the tool in the document: rows 4 and 5 are not ready until it is gone", async () => {
  const t = await world();
  try {
    await t.answered();
    const tool = [{ id: "7", content: "[R1.1] Unit written / asked / q1", resolved: false }];
    const before = await t.check(snapshotOf({ text: AFTER }, tool));
    assert.equal(before.tool_comments.state, "not_ready");
    assert.equal(before.file_clean.state, "not_ready");
    const after = await t.check(ACCEPTED);
    assert.equal(after.tool_comments.state, "ready");
    assert.equal(after.file_clean.state, "ready");
  } finally {
    await t.cleanup();
  }
});

test("a request without a document read is refused", async () => {
  const t = await world();
  try {
    const r = await t.ask("submission_check", {});
    assert.equal(r.ok, false);
    assert.match(r.error, /no document snapshot/);
  } finally {
    await t.cleanup();
  }
});
