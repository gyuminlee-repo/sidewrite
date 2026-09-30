// L4 outside changes and the L10 paragraph key, with known answers.
import { test } from "node:test";
import assert from "node:assert/strict";
import {
  paragraphHash,
  snapshotTexts,
  alignParagraphs,
  outsideChanges,
} from "../daemon/doc-diff.mjs";

const P = [
  "1. Introduction",
  "Cultures were grown at 30 °C for 48 h. Growth was followed at 600 nm.",
  "A cell-free extract was prepared by sonication.",
  "Enzyme activity was measured over five minutes.",
];

test("paragraphHash ignores comment anchors and spacing, not words", () => {
  assert.match(paragraphHash("abc"), /^[0-9a-f]{8}$/);
  assert.equal(paragraphHash("a  b\u0005 c "), paragraphHash("a b c"));
  assert.notEqual(paragraphHash("30 °C"), paragraphHash("37 °C"));
});

test("snapshotTexts reads reviewed_text in index order", () => {
  const snap = {
    paragraphs: [
      { index: 1, text: "old new", reviewed_text: "new" },
      { index: 0, text: "Title" },
    ],
  };
  assert.deepEqual(snapshotTexts(snap), ["Title", "new"]);
});

test("alignParagraphs: an inserted paragraph does not shift the others into changes", () => {
  const next = [P[0], "Inserted by hand.", P[1], P[2], P[3]];
  const a = alignParagraphs(P, next);
  assert.deepEqual(
    a.filter((x) => !x.same),
    [{ old: null, new: 1 }],
  );
});

// The recorded tool change in P[2], as review_refresh joins it to its row.
const LYSATE = {
  paragraph_index: 2,
  kind: "replace",
  find: "cell-free extract",
  text: "clarified lysate",
  context: { left: "A ", right: " was prepared by" },
};

// Positive: each must be reported.
test("L4 positive: sentence deleted outside the tool loses its facts", () => {
  const next = [P[0], "Growth was followed at 600 nm.", P[2], P[3]];
  const out = outsideChanges(P, next, []);
  assert.equal(out.length, 1);
  assert.equal(out[0].kind, "modified");
  assert.equal(out[0].paragraph_index, 1);
  assert.deepEqual(out[0].lost_facts, ["30 °C", "48 h"]);
});

test("L4 positive: deleted and inserted paragraphs, shifted indexes", () => {
  const next = [P[0], P[1], P[3], "Added by Codex (v1.2.0)."];
  const out = outsideChanges(P, next, []);
  assert.deepEqual(
    out.map((x) => [x.kind, x.old_index, x.paragraph_index]),
    [
      ["deleted", 2, null],
      ["inserted", null, 3],
    ],
  );
});

test("L4 positive: a paragraph whose recorded change is user_edited is reported", () => {
  const next = [P[0], P[1], "A clarified lysate was prepared on ice.", P[3]];
  const rows = [{ ...LYSATE, state: "user_edited" }];
  assert.equal(outsideChanges(P, next, rows).length, 1);
});

test("L4 positive: the paragraph of a pending tool change is explained, a hand-edited one is not", () => {
  const next = [
    P[0],
    P[1],
    "A clarified lysate was prepared by sonication.",
    "Enzyme activity was measured.",
  ];
  const rows = [{ ...LYSATE, state: "pending" }];
  const out = outsideChanges(P, next, rows);
  assert.deepEqual(
    out.map((x) => x.paragraph_index),
    [3],
  );
});

test("L4 positive: changes that lose facts come first", () => {
  const next = [P[0], "Growth was followed.", P[2], "Enzyme activity was measured over 5 min."];
  const out = outsideChanges(P, next, []);
  assert.deepEqual(
    out.map((x) => x.paragraph_index),
    [1, 3],
  );
  assert.deepEqual(out[0].lost_facts, ["30 °C", "48 h", "600 nm"]);
  assert.deepEqual(out[1].lost_facts, []);
});

// Negative: none may be reported.
test("L4 negative: unchanged, tool-explained, spacing-only", () => {
  assert.deepEqual(outsideChanges(P, [...P], []), []);
  const edited = [P[0], P[1], "A clarified lysate was prepared by sonication.", P[3]];
  for (const state of ["pending", "accepted"]) {
    assert.deepEqual(outsideChanges(P, edited, [{ ...LYSATE, state }]), [], state);
  }
  // Rejected after the last refresh saw it pending: the reviewed text goes
  // back from the new words to the old ones.
  assert.deepEqual(outsideChanges(edited, P, [{ ...LYSATE, state: "rejected" }]), []);
  const spaced = snapshotTexts({
    paragraphs: P.map((t, index) => ({ index, text: ` ${t.replace(/ /g, "  ")}\u0005` })),
  });
  assert.deepEqual(outsideChanges(P, spaced, []), []);
});

test("L4 negative: rows at another paragraph do not explain this one", () => {
  const next = [P[0], P[1], P[2], "Enzyme activity was measured."];
  const out = outsideChanges(P, next, [{ ...LYSATE, state: "accepted" }]);
  assert.equal(out.length, 1);
  assert.equal(out[0].paragraph_index, 3);
});

// A tool change in a paragraph explains only its own range. A later hand
// edit in the same paragraph is still an outside change, and its lost facts
// are the ones the hand edit dropped, not the tool edit.
const GROWTH = "Cells grew rapidly and were harvested after 24 h.";
const SLOWLY = {
  paragraph_index: 0,
  kind: "replace",
  find: "rapidly",
  text: "slowly",
  context: { left: "Cells grew ", right: " and were harvested" },
};

test("L4 positive: a hand edit next to an accepted tool change is reported", () => {
  const now = ["Cells grew slowly and were harvested after 48 h."];
  // The last refresh saw the tool edit already (reviewed view) or not yet.
  for (const old of [[GROWTH], ["Cells grew slowly and were harvested after 24 h."]]) {
    for (const state of ["accepted", "pending"]) {
      const out = outsideChanges(old, now, [{ ...SLOWLY, state }]);
      assert.equal(out.length, 1, `${old[0]} ${state}`);
      assert.equal(out[0].kind, "modified");
      assert.deepEqual(out[0].lost_facts, ["24 h"]);
      assert.equal(out[0].before, "Cells grew slowly and were harvested after 24 h.");
    }
  }
});

test("L4 negative: a tool change alone is explained in every state", () => {
  const slow = ["Cells grew slowly and were harvested after 24 h."];
  for (const state of ["accepted", "pending"]) {
    assert.deepEqual(outsideChanges([GROWTH], slow, [{ ...SLOWLY, state }]), [], state);
    assert.deepEqual(outsideChanges(slow, slow, [{ ...SLOWLY, state }]), [], state);
  }
  assert.deepEqual(outsideChanges(slow, [GROWTH], [{ ...SLOWLY, state: "rejected" }]), []);
  // Two tool changes in one paragraph, no hand edit.
  const HOURS = {
    paragraph_index: 0,
    kind: "replace",
    find: "24 h",
    text: "one day",
    context: { left: "were harvested after ", right: "" },
  };
  const both = ["Cells grew slowly and were harvested after one day."];
  const rows = [
    { ...SLOWLY, state: "accepted" },
    { ...HOURS, state: "pending" },
  ];
  assert.deepEqual(outsideChanges([GROWTH], both, rows), []);
});

test("L4 positive: a rejected tool change does not hide a hand edit", () => {
  const out = outsideChanges(
    [GROWTH],
    ["Cells grew rapidly and were harvested after 48 h."],
    [{ ...SLOWLY, state: "rejected" }],
  );
  assert.deepEqual(
    out.map((x) => x.lost_facts),
    [["24 h"]],
  );
});
