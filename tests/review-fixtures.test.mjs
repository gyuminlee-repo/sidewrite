// The acceptance-test fixtures must split and validate the way the plan
// says: 5 comments R1.1-R1.5 by regex (after E.1, the disclaimer line
// before the reviewer head, which the split keeps), 3 valid ledger items L1-L3, and the
// planted sentences carry the facts the lesson checks (L1, L4) must report.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { splitReviewRegex } from "../daemon/review-split.mjs";
import { emptyRevision, addLedgerItem } from "../daemon/revision-store.mjs";
import { forbidHits } from "../daemon/ledger-check.mjs";
import { lostFacts, unsourcedTerms } from "../daemon/edit-facts.mjs";

const read = (p) => readFile(new URL(p, import.meta.url), "utf8");

test("fake_review.txt splits into R1.1-R1.5 without an LLM", async () => {
  const c = splitReviewRegex(await read("./fixtures/fake_review.txt"));
  assert.deepEqual(
    c.map((x) => x.ref),
    ["E.1", "R1.1", "R1.2", "R1.3", "R1.4", "R1.5"],
  );
  assert.match(c[0].text, /^Fake review for software testing\./);
  assert.match(c[3].text, /novel substrate/);
  assert.match(c[4].text, /flasks came from/);
  assert.match(c[5].text, /antifoam agent/);
});

// Same strings as PLANTS in scripts/make-test-manuscript.py.
const FLASK =
  " The flasks came from a local glassware supplier (catalogue 4411) and each held 250 mL of medium.";
const RANGE = " Across the replicates the plateau optical density ranged from 0.82 to 0.91.";

test("planted sentences: deleting them loses the facts the acceptance test greps for", async () => {
  const gen = await read("../scripts/make-test-manuscript.py");
  assert.ok(gen.includes(FLASK) && gen.includes(RANGE), "PLANTS changed: update this test");
  assert.deepEqual(lostFacts(FLASK, "", { sentenceStart: true }), ["catalogue 4411", "250 mL"]);
  assert.deepEqual(lostFacts(RANGE, "", { sentenceStart: true }), ["0.82", "0.91"]);
  // R1.5: a named agent and a concentration are in neither the manuscript nor the comment.
  const comment = splitReviewRegex(await read("./fixtures/fake_review.txt")).find(
    (x) => x.ref === "R1.5",
  ).text;
  // "Antifoam" itself is in the comment, so only the number and the product code count.
  assert.deepEqual(unsourcedTerms("with 0.01 % Antifoam 204", [gen, comment]), ["0.01 %", "204"]);
});

test("review_ledger.json items validate and get L1-L3; L2 blocks 'novel'", async () => {
  const { items } = JSON.parse(await read("./fixtures/review_ledger.json"));
  let rev = emptyRevision("/m/review_manuscript.docx");
  for (const item of items) rev = addLedgerItem(rev, item).revision;
  assert.deepEqual(
    rev.ledger.map((i) => [i.id, i.kind]),
    [
      ["L1", "instruction"],
      ["L2", "forbid"],
      ["L3", "instruction"],
    ],
  );
  assert.deepEqual(forbidHits(rev.ledger, "a novel substrate"), ["L2"]);
});
