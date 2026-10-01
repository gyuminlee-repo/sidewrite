// A count rule and the guideline sentence a rule was taken from, in the store
// and in the ledger check of the Review tab.
import { test } from "node:test";
import assert from "node:assert/strict";
import { addLedgerItem, emptyRevision } from "../daemon/revision-store.mjs";
import { checkLedgerIncremental } from "../daemon/ledger-check.mjs";
import { COUNT_TARGETS } from "../daemon/count-check.mjs";
import { withDefaults } from "../daemon/journal-profile.mjs";
import { ledgerStatusText } from "../taskpane/shared/review-view.js";

const base = () => emptyRevision({ doc_path: "/tmp/paper.docx" });
const COUNT = {
  kind: "instruction",
  text: "The abstract has at most 250 words.",
  scope: "manuscript",
  check: { type: "count", target: "section_words", section: "abstract", max: 250 },
  source: { kind: "guideline", sentence: "Abstracts must not exceed 250 words." },
};

test("the store keeps a count check and the guideline sentence of a rule", () => {
  const { item } = addLedgerItem(base(), COUNT);
  assert.deepEqual(item.check, COUNT.check);
  assert.deepEqual(item.source, COUNT.source);
  // A rule the author typed has no source.
  const typed = addLedgerItem(base(), { kind: "forbid", text: "x", check: { type: "llm" } });
  assert.equal(typed.item.source, null);
});

test("every count target the checker knows is a count check the store takes", () => {
  for (const target of COUNT_TARGETS) {
    const check = { type: "count", target };
    assert.deepEqual(addLedgerItem(base(), { ...COUNT, check }).item.check, check);
  }
  assert.throws(() =>
    addLedgerItem(base(), { ...COUNT, check: { type: "count", target: "pages" } }),
  );
});

test("the ledger check of the Review tab skips a count rule and says where it is checked", async () => {
  const { item } = addLedgerItem(base(), COUNT);
  const [out] = await checkLedgerIncremental([item], [{ index: 0, text: "One two." }]);
  assert.equal(out.status, "skipped");
  assert.equal(out.skipped, "count");
  assert.deepEqual(out.violations, []);
  assert.equal(ledgerStatusText({ ...item, status: "skipped" }), "Checked in the Tools tab");
});

test("the review_scope pack is on unless a profile switches it off", () => {
  assert.deepEqual(withDefaults({}).policy_packs, { review_scope: true });
  assert.deepEqual(withDefaults({ policy_packs: { review_scope: false } }).policy_packs, {
    review_scope: false,
  });
});
