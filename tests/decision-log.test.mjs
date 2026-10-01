import { test } from "node:test";
import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { decisionRecord, appendDecisions, redactQuotes } from "../daemon/decision-log.mjs";
import { decisionsPathFor } from "../daemon/revision-store.mjs";

test("decisionsPathFor sits next to the manuscript", () => {
  assert.equal(decisionsPathFor("/a/b/paper.docx"), join("/a/b", "paper.decisions.jsonl"));
});

test("a record keeps a sha256 fingerprint of its input, never the input", () => {
  const r = decisionRecord({
    kind: "comment_split",
    input: "the review text",
    model: "haiku",
    output: { refs: ["R1.1"] },
    now: new Date("2026-09-30T00:00:00Z"),
  });
  assert.deepEqual(r, {
    ts: "2026-09-30T00:00:00.000Z",
    kind: "comment_split",
    input_fingerprint: createHash("sha256").update("the review text").digest("hex"),
    model: "haiku",
    output: { refs: ["R1.1"] },
    expected: null,
  });
});

test("redactQuotes cuts a warning at its first quoted review text", () => {
  assert.equal(
    redactQuotes('2 sentence(s) of the review are in no comment: "Secret one.", "Two."'),
    "2 sentence(s) of the review are in no comment: [quote omitted]",
  );
  assert.equal(redactQuotes("duplicate ref R1.2: renumber it"), "duplicate ref R1.2: renumber it");
});

test("appends from overlapping calls stay whole lines", async () => {
  const dir = await mkdtemp(join(tmpdir(), "decisions-"));
  const path = join(dir, "paper.decisions.jsonl");
  try {
    const rec = (i) => decisionRecord({ kind: "k", input: String(i), model: null, output: { i } });
    await Promise.all(Array.from({ length: 20 }, (_, i) => appendDecisions(path, [rec(i)])));
    const lines = (await readFile(path, "utf8")).trim().split("\n");
    assert.equal(lines.length, 20);
    assert.deepEqual(
      lines.map((l) => JSON.parse(l).output.i).sort((a, b) => a - b),
      Array.from({ length: 20 }, (_, i) => i),
    );
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
