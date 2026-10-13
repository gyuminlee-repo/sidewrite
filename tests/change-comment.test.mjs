// A change Word applied without its rationale comment (comment_missing) gets
// the comment later, from the Review tab: "Add rationale comment" on the
// change. The daemon gives the text the comment would have had, the task
// pane anchors it on the change, and the daemon records its id. Without
// this the change kept the review file and Word apart for good: every
// refresh named it and no Propose all run got past it.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  emptyRevision,
  replaceComments,
  addChange,
  setChangeComment,
} from "../daemon/revision-store.mjs";
import { createReviewService } from "../daemon/review-service.mjs";
import { applyTrackedEdit, formatRationaleComment } from "../daemon/review-tools.mjs";

const R = { what: "w", why: "y", addresses: "a" };
const TEXT = "[R1.1] w / y / a";

// A fake Word paragraph list. runs: the text of each paragraph as pieces
// [text, kind?] with kind "ins" | "del" (a pending tracked run; Word keeps a
// deleted run in the text and finds it in a search).
function fakeDoc(paras, { mode = "uniqueLocalId", failComment = () => false } = {}) {
  const state = { comments: [], inserts: 0 };
  const coll = (items) => ({ items, load() {} });
  const items = paras.map((pieces, k) => {
    const text = pieces.map((x) => x[0]).join("");
    const p = { k, text, pieces, load() {} };
    if (mode === "uniqueLocalId") p.uniqueLocalId = `U${k}`;
    // A hit: where it starts in the paragraph text.
    const hit = (start, q) => ({
      p,
      start,
      end: start + q.length,
      load() {},
      compareLocationWith(other) {
        const end = start + q.length;
        if (start === other.start && end === other.end) return { value: "Equal" };
        if (start >= other.start && end <= other.end) return { value: "Inside" };
        if (end <= other.start) return { value: "Before" };
        if (start >= other.end) return { value: "After" };
        return { value: "OverlapsBefore" };
      },
      insertComment(content) {
        state.inserts += 1;
        if (failComment(state.inserts)) throw new Error("GeneralException: insertComment");
        const c = {
          id: `C${state.comments.length + 1}`,
          content,
          paragraph: k,
          start,
          load() {},
          getRange: () => hit(start, q),
        };
        state.comments.push(c);
        return c;
      },
    });
    p.search = (q) => {
      const out = [];
      for (let i = text.indexOf(q); i !== -1; i = text.indexOf(q, i + 1)) out.push(hit(i, q));
      return coll(out);
    };
    p.getComments = () => coll(state.comments.filter((c) => c.paragraph === k));
    return p;
  });
  const document = { load() {}, body: { paragraphs: coll(items) } };
  globalThis.Word = { run: async (fn) => fn({ document, sync: async () => {} }) };
  return state;
}

const { addChangeComment } = await import("../taskpane/shared/tools-word.js");

// "held at 30 degrees": "degrees" replaced by "degrees Celsius" (narrowed).
const para = (...pieces) => pieces;
const REPLACE_PENDING = para(
  ["held at 30 "],
  ["degrees Celsius", "ins"],
  ["degrees", "del"],
  [" and"],
);
const change = (over = {}) => ({
  kind: "replace",
  find: "degrees",
  text: "degrees Celsius",
  paragraph_id: "U1",
  paragraph_index: 1,
  comment_text: TEXT,
  ...over,
});

test("task pane: the comment goes on the deleted words of a pending replace, not on the inserted ones", async () => {
  const s = fakeDoc([para(["Title"]), REPLACE_PENDING]);
  const r = await addChangeComment(change());
  assert.deepEqual([r.ok, r.word_comment_id, r.existed], [true, "C1", false]);
  assert.equal(s.comments.length, 1);
  // "held at 30 degrees Celsius" is 26 long: the deleted "degrees" starts there.
  assert.deepEqual([s.comments[0].content, s.comments[0].start], [TEXT, 26]);
});

test("task pane: a second press finds the comment and adds no second one", async () => {
  const s = fakeDoc([para(["Title"]), REPLACE_PENDING]);
  const a = await addChangeComment(change());
  const b = await addChangeComment(change());
  assert.deepEqual([b.ok, b.word_comment_id, b.existed], [true, a.word_comment_id, true]);
  assert.deepEqual([s.comments.length, s.inserts], [1, 1]);
});

test("task pane: accepted replace, pure insertion, deletion; the paragraph by id, or by index without ids", async () => {
  // Accepted replace: the old words are gone, the new ones are the anchor.
  let s = fakeDoc([para(["Title"]), para(["held at 30 degrees Celsius and"])]);
  let r = await addChangeComment(change());
  assert.deepEqual([r.ok, s.comments[0].start], [true, 11]);
  // Pure insertion: its own text.
  s = fakeDoc([para(["Title"]), para(["for four days"], [" (OD600)", "ins"], ["."])]);
  r = await addChangeComment(change({ kind: "insert", find: "four days", text: " (OD600)" }));
  assert.deepEqual([r.ok, s.comments[0].start], [true, 13]);
  // Pending deletion: the deleted words are still there.
  s = fakeDoc([para(["Title"]), para(["Cells were "], ["carefully ", "del"], ["held."])]);
  r = await addChangeComment(change({ kind: "delete", find: "carefully ", text: "" }));
  assert.deepEqual([r.ok, s.comments[0].start], [true, 11]);
  // The paragraph moved: found by its id. Without ids, by its index.
  s = fakeDoc([para(["Title"]), para(["New."]), REPLACE_PENDING]);
  r = await addChangeComment(change({ paragraph_id: "U2", paragraph_index: 1 }));
  assert.deepEqual([r.ok, s.comments[0].paragraph], [true, 2]);
  s = fakeDoc([para(["Title"]), REPLACE_PENDING], { mode: "index" });
  r = await addChangeComment(change({ paragraph_id: null }));
  assert.deepEqual([r.ok, s.comments[0].paragraph], [true, 1]);
});

test("task pane: no place for the comment is an answer, not a comment somewhere else", async () => {
  // The change's words are not in the paragraph (an accepted deletion, or the author rewrote it).
  let s = fakeDoc([para(["Title"]), para(["Cells were held."])]);
  let r = await addChangeComment(change({ kind: "delete", find: "carefully ", text: "" }));
  assert.deepEqual([r.ok, r.code, s.comments.length], [false, "CommentAnchorMissing", 0]);
  // The words are there twice and nothing says which: no comment.
  s = fakeDoc([para(["Title"]), para(["degrees and degrees"])]);
  r = await addChangeComment(change({ text: "kelvin" }));
  assert.deepEqual([r.ok, r.code, s.comments.length], [false, "CommentAnchorMissing", 0]);
  // The paragraph is gone.
  s = fakeDoc([para(["Title"])]);
  r = await addChangeComment(change());
  assert.deepEqual([r.ok, r.code], [false, "ParagraphNotFound"]);
  // Only a tool comment text is accepted.
  await assert.rejects(addChangeComment(change({ comment_text: "a note" })), /comment_text/);
});

test("task pane: an insert that throws is tried once more, and then is an error", async () => {
  let s = fakeDoc([para(["Title"]), REPLACE_PENDING], { failComment: (n) => n === 1 });
  const r = await addChangeComment(change());
  assert.deepEqual([r.ok, s.comments.length, s.inserts], [true, 1, 2]);
  s = fakeDoc([para(["Title"]), REPLACE_PENDING], { failComment: () => true });
  await assert.rejects(addChangeComment(change()), /GeneralException/);
  assert.equal(s.comments.length, 0);
});

// ---- store and service -----------------------------------------------------------

const row = (over = {}) => ({
  word_comment_id: null,
  comment_missing: "CommentInsertFailed",
  paragraph_index: 1,
  paragraph_id: "U1",
  kind: "replace",
  find: "degrees",
  text: "degrees Celsius",
  before: "b",
  after: "a",
  rationale: R,
  ...over,
});

test("store: setChangeComment gives a change without a comment its id, once, and no id of another change", () => {
  let rev = replaceComments(emptyRevision("/m/p.docx"), [{ ref: "R1.1", reviewer: 1, text: "c" }]);
  rev = addChange(rev, "R1.1", row()).revision;
  rev = addChange(rev, "R1.1", row({ word_comment_id: "7", comment_missing: null })).revision;
  const next = setChangeComment(rev, "R1.1", "R1.1-1", 12);
  const ch = next.comments[0].changes[0];
  assert.deepEqual([ch.word_comment_id, ch.comment_missing], ["12", null]);
  assert.equal(
    rev.comments[0].changes[0].word_comment_id,
    null,
    "the revision passed in is not changed",
  );
  assert.throws(() => setChangeComment(next, "R1.1", "R1.1-1", 13), /already has/);
  assert.throws(() => setChangeComment(rev, "R1.1", "R1.1-1", "7"), /another change/);
  assert.throws(() => setChangeComment(rev, "R1.1", "R1.1-9", 1), /No change/);
  assert.throws(() => setChangeComment(rev, "R1.1", "R1.1-1", null), /comment id/);
});

test("service: the comment text of a change is the one its edit would have carried; the id is recorded", async () => {
  const dir = await mkdtemp(join(tmpdir(), "change-comment-"));
  const doc = join(dir, "paper.docx");
  const lines = [];
  const svc = createReviewService({
    getContext: () => ({ activeDoc: doc, caps: ["target_hash", "author_insertion_guard"] }),
    callTool: null,
    pushTurn: () => {},
    log: (s) => lines.push(s),
    env: { CC_OFFICE_WSL_PATHS: "0" },
  });
  const ask = (type, payload = {}) =>
    new Promise((resolve) =>
      svc.handlers[type]({ type, request_id: "q", ...payload }, resolve, "k", "word"),
    );
  try {
    await ask("review_save_comments", { comments: [{ ref: "R1.1", reviewer: 1, text: "Unit." }] });
    // The edit goes in, its comment does not.
    const rationale = { ...R, requirement_ids: [] };
    let sent = null;
    const r = await applyTrackedEdit(
      {
        ...svc.trackedEditDeps("k"),
        call: async (name, args) => {
          sent = args.comment_text;
          return {
            ok: false,
            code: "CommentInsertFailed",
            applied: true,
            paragraph_index: 1,
            paragraph_id: "U1",
            before: "It was held at 30 degrees for 5 days.",
            after: "It was kept warm.",
            target_checked: true,
          };
        },
      },
      {
        comment_ref: "R1.1",
        find: "held at 30 degrees for 5 days",
        replace: "kept warm",
        rationale,
      },
      { ref: "R1.1" },
    );
    assert.equal(r.comment_missing, "CommentInsertFailed");
    assert.match(sent, /사라지는 사실/, "the comment of this edit names the lost facts");

    const t = await ask("review_comment_text", { ref: "R1.1", change_id: "R1.1-1" });
    assert.equal(t.ok, true, t.error);
    assert.equal(t.comment_text, sent, "the same text the task pane was given with the edit");
    assert.equal(t.comment_text, formatRationaleComment("R1.1", rationale, { lost: r.lost_facts }));
    assert.deepEqual(t.known_comment_ids, [], "no comment is on record yet");
    assert.deepEqual(
      [
        t.change.kind,
        t.change.find,
        t.change.text,
        t.change.paragraph_id,
        t.change.paragraph_index,
      ],
      ["replace", "held at 30 degrees for 5 days", "kept warm", "U1", 1],
    );
    // What the task pane needs to anchor a deletion whose old words Word
    // dropped where the edit would have (droppedDeletionAnchor).
    assert.deepEqual(
      [t.change.before, t.change.narrowed_from, t.change.comment_missing],
      ["It was held at 30 degrees for 5 days.", null, "CommentInsertFailed"],
    );

    const set = await ask("review_comment_set", {
      ref: "R1.1",
      change_id: "R1.1-1",
      word_comment_id: 55,
    });
    assert.equal(set.ok, true, set.error);
    const ch = set.revision.comments[0].changes[0];
    assert.deepEqual([ch.word_comment_id, ch.comment_missing], ["55", null]);
    assert.ok(lines.includes("[review] comment added later R1.1-1 comment=55"));
    // Once it has its comment there is nothing to add.
    const again = await ask("review_comment_text", { ref: "R1.1", change_id: "R1.1-1" });
    assert.equal(again.ok, false);
    assert.match(again.error, /already has/);
    assert.equal(
      (await ask("review_comment_set", { ref: "R1.1", change_id: "R1.1-1", word_comment_id: 56 }))
        .ok,
      false,
    );
    assert.equal(
      (await ask("review_comment_text", { ref: "R1.1", change_id: "R1.1-9" })).ok,
      false,
    );

    // A second change with the same rationale, again without its comment: the
    // id of the first comment goes out with the edit and with the comment
    // text, so the task pane takes it for neither.
    let known = null;
    const r2 = await applyTrackedEdit(
      {
        ...svc.trackedEditDeps("k"),
        call: async (name, args) => {
          known = args.known_comment_ids;
          return {
            ok: false,
            code: "CommentInsertFailed",
            applied: true,
            paragraph_index: 1,
            paragraph_id: "U1",
            before: "It was kept warm.",
            after: "It was kept warm throughout.",
            target_checked: true,
          };
        },
      },
      { comment_ref: "R1.1", find: "kept warm", insert_after: " throughout", rationale },
      { ref: "R1.1" },
    );
    assert.equal(r2.comment_missing, "CommentInsertFailed", JSON.stringify(r2));
    assert.deepEqual(known, ["55"]);
    const t2 = await ask("review_comment_text", { ref: "R1.1", change_id: "R1.1-2" });
    assert.deepEqual([t2.ok, t2.known_comment_ids], [true, ["55"]]);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

// Codex review of the recovery path (P2): the comment to reuse was looked for
// by its text alone. Two changes of one comment in one paragraph can carry
// the same rationale text; the second one then got the comment of the first,
// which the daemon refuses ("belongs to another change"), and could not be
// given a comment at all.
test("task pane: a comment with the same text on another change of the paragraph is not taken for this one", async () => {
  // "hot" -> "warm" (has its comment) and "cold" -> "cool" (has none), same rationale.
  const two = para(
    ["A "],
    ["warm", "ins"],
    ["hot", "del"],
    [" day and a "],
    ["cool", "ins"],
    ["cold", "del"],
    [" night."],
  );
  const first = change({ find: "hot", text: "warm" });
  const second = change({ find: "cold", text: "cool" });
  const s = fakeDoc([para(["Title"]), two]);
  const a = await addChangeComment(first);
  assert.deepEqual([a.word_comment_id, a.existed], ["C1", false]);
  // By its place: the comment of the first change does not lie on the second.
  const b = await addChangeComment(second);
  assert.deepEqual([b.ok, b.word_comment_id, b.existed], [true, "C2", false]);
  assert.deepEqual(
    s.comments.map((c) => c.start),
    [6, 24],
  );
  // Each finds its own on a second press.
  assert.deepEqual(
    [
      (await addChangeComment(first)).word_comment_id,
      (await addChangeComment(second)).word_comment_id,
    ],
    ["C1", "C2"],
  );
  assert.equal(s.comments.length, 2);
});

test("task pane: a comment the daemon already has on another change is never reused, wherever it lies", async () => {
  const s = fakeDoc([para(["Title"]), REPLACE_PENDING]);
  const a = await addChangeComment(change());
  // The same place, but the daemon says C1 is recorded on another change.
  const b = await addChangeComment(change({ known_comment_ids: [a.word_comment_id] }));
  assert.deepEqual([b.ok, b.word_comment_id, b.existed], [true, "C2", false]);
  assert.equal(s.comments.length, 2);
});
