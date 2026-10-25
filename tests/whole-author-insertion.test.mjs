// A tool deletion that would take the whole of a pending tracked insertion
// the author made: Word drops such words outright and keeps no mark, so the
// task pane refuses it (tools-word.js WholeAuthorInsertion; pane tests in
// comment-anchor-dropped.test.mjs). Here the daemon side: what it sends
// the task pane (own_insertions: the texts its recorded changes put into
// the paragraph), and how the refusal is handled: logged as a guard
// refusal, nothing recorded, the model told to ask with the paragraph id.
// Invented text.
import { test } from "node:test";
import assert from "node:assert/strict";
import {
  applyTrackedEdit,
  applyTrackedRewrite,
  ownInsertions,
  AUTHOR_INSERTION_CAP,
} from "../daemon/review-tools.mjs";
import { PANE_CAPS } from "../taskpane/shared/tools-word.js";
import { refusalsFromLog } from "../scripts/e2e/replay-lib.mjs";
import { WHOLE_AUTHOR_INSERTION } from "../taskpane/shared/review-view.js";

const DOC = [
  "Title",
  "Crews checked the north span daily. New bearings were fitted.",
  "The road stayed open.",
];
const IDS = ["U0", "U1", "U2"];
const R = { what: "w", why: "y", addresses: "a" };
const TURN = { ref: "R2.1" };
const edit = (find, replace) => ({ comment_ref: "R2.1", find, replace, rationale: R });

function deps(result, ctx = {}) {
  const d = {
    calls: [],
    edits: [],
    notApplied: [],
    lines: [],
    log: (s) => d.lines.push(s),
    getLedger: async () => [],
    getReviewContext: async () => ({ paragraphs: DOC, paragraph_ids: IDS, ...ctx }),
    paneChecksTarget: () => true,
    paneCan: () => true,
    onAttempt: async () => ({ id: "R2.1-b1" }),
    call: async (name, args) => {
      d.calls.push(args);
      return typeof result === "function" ? result(args) : result;
    },
    onEdit: async ({ change }) => (d.edits.push(change), { id: "R2.1-1" }),
    onNotApplied: async (x) => d.notApplied.push(x.attempt_id),
    onStaleTarget: async () => {},
  };
  return d;
}
const REFUSAL = {
  ok: false,
  code: "WholeAuthorInsertion",
  applied: false,
  paragraph_index: 1,
  paragraph_id: "U1",
  reason: "whole",
  message: WHOLE_AUTHOR_INSERTION,
};

test("ownInsertions: the texts recorded in that paragraph, by id; by index only while its text still matches", () => {
  const now = DOC[1];
  const inserted = [
    { paragraph_index: 1, paragraph_id: "U1", text: "New bearings were fitted." },
    { paragraph_index: 1, paragraph_id: "U1", text: "New bearings were fitted." },
    { paragraph_index: 2, paragraph_id: "U2", text: "stayed" },
    // Moved: the id decides, not the index it had then.
    { paragraph_index: 2, paragraph_id: "U1", text: "daily" },
    // No id: counted only when the paragraph still reads as it did after
    // that change (spacing aside).
    { paragraph_index: 1, paragraph_id: null, text: "north", after: `${now}  ` },
    { paragraph_index: 1, paragraph_id: null, text: "south", after: "Crews rested." },
    { paragraph_index: 1, paragraph_id: null, text: "east" },
  ];
  assert.deepEqual(ownInsertions(inserted, 1, IDS, now), [
    "New bearings were fitted.",
    "daily",
    "north",
  ]);
  // No snapshot ids: a record with an id cannot be placed and is not counted.
  assert.deepEqual(ownInsertions(inserted, 1, null, now), ["north"]);
  assert.deepEqual(ownInsertions(inserted, null, IDS, now), []);
  assert.deepEqual(ownInsertions(undefined, 1, IDS, now), []);
});

test("ownInsertions: one entry per text, none for a text found twice in the paragraph", () => {
  // The pane lets each entry cover one piece (tools-word.js insertionAround
  // covered); a text the tool put in twice is not told apart, so neither is
  // sent and the cut is refused as the author's.
  const now = "Crews fitted bolts bolts here.";
  const twice = { paragraph_index: 1, paragraph_id: "U1", text: "bolts ", after: now };
  assert.deepEqual(ownInsertions([twice, twice], 1, IDS, now), []);
  const once = { ...twice, text: "fitted", after: now };
  assert.deepEqual(ownInsertions([once, once], 1, IDS, now), ["fitted"]);
});

test("daemon: a deletion carries the texts the tool put into its paragraph; none goes out when there are none", async () => {
  const inserted = [
    { paragraph_index: 1, paragraph_id: "U1", text: "New bearings were fitted." },
    { paragraph_index: 2, paragraph_id: "U2", text: "stayed" },
  ];
  let d = deps(REFUSAL, { inserted });
  await applyTrackedEdit(d, edit("New bearings were fitted.", ""), TURN);
  assert.deepEqual(d.calls[0].own_insertions, ["New bearings were fitted."]);
  d = deps(REFUSAL);
  await applyTrackedEdit(d, edit("New bearings were fitted.", ""), TURN);
  assert.equal("own_insertions" in d.calls[0], false);
});

test("daemon: WholeAuthorInsertion is a guard refusal: logged, nothing recorded, the model asks", async () => {
  const d = deps(REFUSAL);
  const r = await applyTrackedEdit(d, edit("New bearings were fitted.", ""), TURN);
  assert.deepEqual(
    [r.ok, r.code, r.applied, r.paragraph_index],
    [false, "WholeAuthorInsertion", false, 1],
  );
  assert.ok(r.message.startsWith(WHOLE_AUTHOR_INSERTION));
  assert.match(r.message, /office_ask_author/);
  assert.match(r.message, /Paragraph: U1 \(give it as paragraph_id\)\./);
  assert.deepEqual([d.edits.length, d.notApplied], [0, ["R2.1-b1"]]);
  // No manuscript text in the log.
  assert.deepEqual(d.lines, ["[review] whole author insertion R2.1 paragraph=1 why=whole"]);
  assert.deepEqual(refusalsFromLog(d.lines.join("\n"), "R2.1"), { WholeAuthorInsertion: 1 });
});

test("daemon: the part of a rewrite that would take the whole insertion stops the rewrite there", async () => {
  const d = deps(REFUSAL, {
    inserted: [{ paragraph_index: 1, paragraph_id: "U1", text: "New bearings were fitted." }],
  });
  const r = await applyTrackedRewrite(
    d,
    {
      comment_ref: "R2.1",
      paragraph_id: "U1",
      text: "Crews checked the north span daily.",
      what: "w",
      why: "y",
      addresses: "a",
    },
    TURN,
  );
  assert.deepEqual([r.ok, r.code, r.applied_parts], [false, "WholeAuthorInsertion", 0]);
  assert.match(r.message, /office_ask_author/);
  assert.deepEqual(d.calls[0].own_insertions, ["New bearings were fitted."]);
  assert.equal(d.edits.length, 0);
  assert.ok(
    d.lines.includes("[review] whole author insertion R2.1 paragraph=1 why=whole"),
    d.lines.join("\n"),
  );
});

test("daemon: a replace inside the author's insertion is recorded with the flag and remainders", async () => {
  const applied = (extra) => ({
    ok: true,
    word_comment_id: 7,
    paragraph_index: 1,
    paragraph_id: "U1",
    before: DOC[1],
    after: "after",
    target_checked: true,
    ...extra,
  });
  const flag = { in_insertion: true, insertion_left: "New ", insertion_right: " were fitted." };
  let d = deps(applied(flag));
  let r = await applyTrackedEdit(d, edit("bearings", "pads"), TURN);
  assert.equal(r.ok, true, r.message);
  const e = d.edits[0];
  assert.deepEqual(
    [e.kind, e.in_insertion, e.insertion_left, e.insertion_right],
    ["replace", true, "New ", " were fitted."],
  );
  // Plain text: the flag alone, false.
  d = deps(applied({ in_insertion: false }));
  r = await applyTrackedEdit(d, edit("bearings", "pads"), TURN);
  assert.deepEqual([d.edits[0].in_insertion, "insertion_left" in d.edits[0]], [false, false]);
  // An insertion never carries it.
  d = deps(applied(flag));
  r = await applyTrackedEdit(
    d,
    { comment_ref: "R2.1", find: "bearings", insert_after: " and pads", rationale: R },
    TURN,
  );
  assert.equal(r.ok, true, r.message);
  assert.equal("in_insertion" in d.edits[0], false);
});

test("daemon: an insertion made earlier in the same turn counts as the tool's own before its record lands", async () => {
  // The first edit is applied but its record has not reached the review
  // file (the context still has none); the turn keeps what it put in.
  const turn = { ref: "R2.1" };
  const d = deps((args) =>
    args.replace === ""
      ? REFUSAL
      : {
          ok: true,
          word_comment_id: 7,
          paragraph_index: 1,
          paragraph_id: "U1",
          before: DOC[1],
          after: DOC[1],
          target_checked: true,
        },
  );
  // The paragraph as the daemon reads it after the edit; the list of
  // recorded insertions (ctx.inserted) still lacks it.
  const after = DOC[1].replace("New bearings", "Spare pads");
  let paragraphs = DOC;
  d.getReviewContext = async () => ({ paragraphs, paragraph_ids: IDS });
  const first = await applyTrackedEdit(d, edit("New bearings", "Spare pads"), turn);
  assert.equal(first.ok, true, first.message);
  paragraphs = [DOC[0], after, DOC[2]];
  await applyTrackedEdit(d, edit("Spare pads", ""), turn);
  assert.deepEqual(d.calls[1].own_insertions, ["Spare pads"]);
  // Another turn does not inherit it.
  const other = deps(REFUSAL);
  await applyTrackedEdit(other, edit("Spare pads", ""), { ref: "R2.1" });
  assert.equal("own_insertions" in other.calls[0], false);
});

test("daemon: a task pane without the insertion guard is sent no deletion or replace", async () => {
  const d = deps(REFUSAL);
  d.paneCan = (cap) => cap !== AUTHOR_INSERTION_CAP;
  for (const [find, replace] of [
    ["New bearings were fitted.", ""],
    ["bearings", "pads"],
  ]) {
    const r = await applyTrackedEdit(d, edit(find, replace), { ref: "R2.1" });
    assert.deepEqual([r.ok, r.code], [false, "TaskPaneOutdated"], find);
    assert.match(r.message, /reload the task pane/);
  }
  assert.equal(d.calls.length, 0);
  assert.equal(
    d.lines.filter((l) => l === `[review] refused R2.1: task pane without ${AUTHOR_INSERTION_CAP}`)
      .length,
    2,
  );
  // An insertion removes no words: it still goes out.
  d.call = async (name, args) => (
    d.calls.push(args),
    { ok: false, code: "NotFound", message: "x" }
  );
  await applyTrackedEdit(
    d,
    { comment_ref: "R2.1", find: "bearings", insert_after: " and pads", rationale: R },
    { ref: "R2.1" },
  );
  assert.equal(d.calls.length, 1);
  assert.ok(PANE_CAPS.includes(AUTHOR_INSERTION_CAP));
});

test("daemon: the refusal names the paragraph by its snapshot id, or by none", async () => {
  const d = deps(REFUSAL, { paragraph_ids: null });
  const r = await applyTrackedEdit(d, edit("New bearings were fitted.", ""), { ref: "R2.1" });
  assert.equal(r.code, "WholeAuthorInsertion");
  assert.doesNotMatch(r.message, /Paragraph: /);
  assert.doesNotMatch(r.message, /\bp1\b/);
});
