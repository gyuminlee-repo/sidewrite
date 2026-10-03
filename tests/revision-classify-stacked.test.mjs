// Two changes of one turn stacked on the same stretch of a paragraph, and
// nobody has decided anything: every change must read pending. Shapes of a
// replay (2026-10-03) where one such change read accepted and another
// rejected, which made the status check report a mismatch and the paragraph
// check not run. Invented text (bridge works, transport planning).
import { test } from "node:test";
import assert from "node:assert/strict";
import { classifyChange, changeContext, normalizeParaText } from "../daemon/revision-classify.mjs";
import { invariantReport } from "../taskpane/shared/review-view.js";

// A snapshot paragraph whose tracked runs are given by their text; start is
// where the text stands in the normalized paragraph text (as the task pane
// reports it).
function para(text, runs) {
  const t = normalizeParaText(text);
  const tracked = runs.map(([kind, s]) => {
    const start = t.indexOf(s.trim());
    assert.notEqual(start, -1, `run text not in paragraph: ${s}`);
    return { kind, text: s.trim(), start, marks: 1 };
  });
  return {
    index: 4,
    id: "P4",
    text,
    w_ins: runs.filter(([k]) => k === "ins").length,
    w_del: runs.filter(([k]) => k === "del").length,
    tracked,
  };
}
const snap = (p, comments = []) => ({ paragraphs: [p], comments });

// An insertion, then a sibling that rewrote the words right after it. The
// insertion's own text holds its left context word followed by the first
// word of its old right context, so "L + R" is found inside it.
test("an insertion whose right context a sibling rewrote stays pending", () => {
  const BEFORE = "Ridership rose on the east line (Table 2). This shows a gain for buses.";
  const I = " Fares stayed flat (Table 2). This held for rail.";
  const ins = {
    id: "R2.1-1",
    kind: "insert",
    paragraph_index: 4,
    paragraph_id: "P4",
    word_comment_id: "11",
    find: "line (Table 2).",
    text: I,
    before: BEFORE,
    after: BEFORE.replace("(Table 2).", `(Table 2).${I}`),
    context: changeContext(BEFORE, "line (Table 2).", { kind: "insert" }),
  };
  assert.deepEqual(ins.context, { left: "line (Table 2).", right: " This shows a" });
  // Word shows a pending replace as inserted then deleted text.
  const I2 = "Bus use grew";
  const D2 = "This shows a gain for buses";
  const pending = `Ridership rose on the east line (Table 2).${I} ${I2}${D2}.`;
  const p = para(pending, [
    ["ins", I],
    ["ins", I2],
    ["del", D2],
  ]);
  const comments = [
    { id: "11", paragraph_index: 4 },
    { id: "12", paragraph_index: 4 },
  ];
  assert.equal(classifyChange(ins, snap(p, comments)).state, "pending");
  // Control: the insertion rejected, its text and comment gone, the
  // sibling decided as well: rejected as before.
  const rejected = para(
    "Ridership rose on the east line (Table 2). This shows a gain for buses.",
    [],
  );
  assert.equal(classifyChange(ins, snap(rejected)).state, "rejected");
});

// A replace whose new text a later change of the same turn edited: the
// later change's range lies inside the pending insertion, so Word drops its
// deleted text and leaves no deletion to show.
test("a replace inside a sibling's pending insertion stays pending", () => {
  const S1_BEFORE = "Crews checked the north span daily. The old bearings were replaced.";
  const I1 = "New elastomeric bearings with steel plates were fitted";
  const D1 = "The old bearings were replaced";
  const S2_BEFORE = `Crews checked the north span daily. ${I1}${D1}.`;
  const s2 = {
    id: "R2.2-2",
    kind: "replace",
    paragraph_index: 4,
    paragraph_id: "P4",
    word_comment_id: "22",
    find: "plates ",
    text: "shims ",
    before: S2_BEFORE,
    after: S2_BEFORE.replace("plates ", "shims "),
    context: changeContext(S2_BEFORE, "plates "),
  };
  assert.deepEqual(s2.context, { left: "bearings with steel ", right: "were fittedThe old" });
  const I1now = I1.replace("plates", "shims");
  const pending = `Crews checked the north span daily. ${I1now}${D1}.`;
  const p = para(pending, [
    ["ins", I1now],
    ["del", D1],
  ]);
  assert.equal(classifyChange(s2, snap(p, [{ id: "22", paragraph_index: 4 }])).state, "pending");
  // Control: everything accepted, nothing tracked: accepted as before.
  const done = para(`Crews checked the north span daily. ${I1now}.`, []);
  assert.equal(classifyChange(s2, snap(done)).state, "accepted");
  // Unchanged without the per-run list (an older task pane): the text
  // shape decides.
  const { tracked, ...noRuns } = p;
  assert.ok(tracked.length);
  assert.equal(classifyChange(s2, snap(noRuns)).state, "accepted");
});

// The status check after such a run: nobody decided anything, a rewrite's
// parts share one comment, and one change was read rejected although its
// comment is still in the document.
const tool = (id, ref) => ({ id, content: `[${ref}] what / why / addresses` });
const RUN = {
  comments: [
    {
      ref: "R2.1",
      changes: [
        { id: "R2.1-1", kind: "insert", state: "rejected", word_comment_id: "11" },
        { id: "R2.1-2", kind: "replace", state: "pending", word_comment_id: "12" },
      ],
    },
    {
      ref: "R2.2",
      changes: [
        { id: "R2.2-1", kind: "replace", state: "pending", word_comment_id: "21" },
        { id: "R2.2-2", kind: "replace", state: "accepted", word_comment_id: "22" },
      ],
    },
    {
      ref: "R2.3",
      changes: [1, 2, 3].map((n) => ({
        id: `R2.3-${n}`,
        kind: "replace",
        state: "pending",
        word_comment_id: "31",
        rewrite_id: "W1",
      })),
    },
  ],
};
const IN_WORD = [
  tool("11", "R2.1"),
  tool("12", "R2.1"),
  tool("21", "R2.2"),
  tool("22", "R2.2"),
  tool("31", "R2.3"),
];

test("status check: a run nobody decided does not report a mismatch", () => {
  const r = invariantReport(RUN, IN_WORD);
  assert.deepEqual(
    [
      r.ok,
      r.changes,
      r.expected_comments,
      r.tool_comments,
      r.missing_comment,
      r.unrecorded_comment_ids,
    ],
    [true, 7, 5, 5, [], []],
  );
});

test("status check: a missing or an unrecorded tool comment still trips it", () => {
  const missing = invariantReport(
    RUN,
    IN_WORD.filter((c) => c.id !== "21"),
  );
  assert.deepEqual([missing.ok, missing.missing_comment], [false, ["R2.2-1"]]);
  const stray = invariantReport(RUN, [...IN_WORD, tool("99", "R2.9")]);
  assert.deepEqual([stray.ok, stray.unrecorded_comment_ids], [false, ["99"]]);
  // A rejected insertion whose comment Word removed expects none.
  const gone = invariantReport(
    RUN,
    IN_WORD.filter((c) => c.id !== "11"),
  );
  assert.deepEqual([gone.ok, gone.expected_comments, gone.tool_comments], [true, 4, 4]);
});
