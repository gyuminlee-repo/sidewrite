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
import { applyTrackedEdit, applyTrackedRewrite, ownInsertions } from "../daemon/review-tools.mjs";
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
  message: WHOLE_AUTHOR_INSERTION,
};

test("ownInsertions: the texts recorded in that paragraph, by id, else by index", () => {
  const inserted = [
    { paragraph_index: 1, paragraph_id: "U1", text: "New bearings were fitted." },
    { paragraph_index: 1, paragraph_id: "U1", text: "New bearings were fitted." },
    { paragraph_index: 2, paragraph_id: "U2", text: "stayed" },
    // Moved: the id decides, not the index it had then.
    { paragraph_index: 2, paragraph_id: "U1", text: "daily" },
    { paragraph_index: 1, paragraph_id: null, text: "north" },
  ];
  assert.deepEqual(ownInsertions(inserted, 1, IDS), [
    "New bearings were fitted.",
    "daily",
    "north",
  ]);
  assert.deepEqual(ownInsertions(inserted, 1, null), ["New bearings were fitted.", "north"]);
  assert.deepEqual(ownInsertions(inserted, null, IDS), []);
  assert.deepEqual(ownInsertions(undefined, 1, IDS), []);
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
  assert.deepEqual(d.lines, ["[review] whole author insertion R2.1 paragraph=1"]);
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
    d.lines.includes("[review] whole author insertion R2.1 paragraph=1"),
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
