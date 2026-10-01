// P15: the instruction ledger as a JSON file (export, import). Store side;
// the message handlers are in review-service.test.mjs.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import {
  ledgerExportDoc,
  importLedgerItems,
  emptyRevision,
  addLedgerItem,
} from "../daemon/revision-store.mjs";

const fx = JSON.parse(
  readFileSync(new URL("./fixtures/review_ledger.json", import.meta.url), "utf8"),
);

test("import: fixture shape, duplicates skipped, all-or-nothing on a bad row", () => {
  const r0 = emptyRevision("/x/paper.docx");
  const a = importLedgerItems(r0, fx.items, { mode: "append" });
  assert.deepEqual(
    a.revision.ledger.map((i) => i.id),
    ["L1", "L2", "L3"],
  );
  assert.equal(a.added.length, 3);
  assert.equal(r0.ledger.length, 0, "the revision passed in is not changed");
  const b = importLedgerItems(a.revision, fx.items, { mode: "append" });
  assert.deepEqual([b.added.length, b.skipped.length], [0, 3]);
  assert.throws(() => importLedgerItems(r0, [fx.items[0], { kind: "nope", text: "" }]), /item 2/);
  assert.throws(() => importLedgerItems(r0, { items: [] }), /items array/);
  assert.throws(() => importLedgerItems(r0, fx.items, { mode: "merge" }), /mode/);
  const c = importLedgerItems(a.revision, [fx.items[1]], { mode: "replace" });
  assert.deepEqual(
    c.revision.ledger.map((i) => i.text),
    [fx.items[1].text],
  );
});

test("import: the same rule twice in one file is added once", () => {
  const r = importLedgerItems(emptyRevision("/x/p.docx"), [fx.items[0], fx.items[0]]);
  assert.deepEqual([r.added.length, r.skipped.length], [1, 1]);
});

test("export keeps rules, drops ids and check results, re-imports", () => {
  let rev = emptyRevision("/x/p.docx");
  rev = addLedgerItem(rev, {
    kind: "forbid",
    text: "no novel",
    check: { type: "regex", pattern: "novel", flags: "i", expect: "absent" },
  }).revision;
  const doc = ledgerExportDoc(rev);
  assert.equal(doc.format, "manuscript-pane-ledger");
  assert.equal(doc.version, 1);
  assert.deepEqual(Object.keys(doc.items[0]).sort(), ["check", "kind", "scope", "text"]);
  assert.equal(importLedgerItems(emptyRevision("/y.docx"), doc.items).added.length, 1);
  // Importing an export into the revision it came from adds nothing.
  assert.equal(importLedgerItems(rev, doc.items).skipped.length, 1);
});
