// toolTrackedEdit (taskpane/shared/tools-word.js) on a fake Word object
// model with failure injection. Queued operations run at context.sync(), as
// in Office.js; a failing sync drops its queue. The point under test: once
// the sync that carries the body edit succeeded, any later failure (anchor
// search, insertComment, the comment sync, the track-mode restore) must
// still return applied:true, so the daemon records the change.
import { test } from "node:test";
import assert from "node:assert/strict";

const PARA = "Cells were held at 30 degrees for 2 h.";

function fakeWord({ fail = () => false, throwOnInsertComment = false } = {}) {
  const state = { applied: false, comments: [], mode: "Off", pending: [], syncs: 0 };
  const coll = (items) => ({ items, load() {} });
  const occurrences = (text, q) => {
    const out = [];
    for (let i = text.indexOf(q); i !== -1; i = text.indexOf(q, i + 1)) out.push(q);
    return out;
  };
  let para;
  const range = (text) => ({
    text,
    load() {},
    paragraphs: { getFirst: () => para },
    search: (q) => coll(occurrences(text, q).map(range)),
    insertText(s) {
      state.pending.push({ kind: "edit", run: () => (state.applied = true) });
      return range(s);
    },
    insertComment(content) {
      if (throwOnInsertComment) throw new Error("GeneralException: insertComment");
      const c = { id: "C1", load() {} };
      state.pending.push({ kind: "comment", run: () => state.comments.push(content) });
      return c;
    },
    compareLocationWith: () => ({ value: "Before" }),
  });
  para = { ...range(PARA), uniqueLocalId: "P1", style: "Normal" };
  const document = {
    load() {},
    get changeTrackingMode() {
      return state.mode;
    },
    set changeTrackingMode(v) {
      state.pending.push({ kind: "mode", run: () => (state.mode = v) });
    },
    body: { paragraphs: coll([para]), search: (q) => coll(occurrences(PARA, q).map(range)) },
  };
  const context = {
    document,
    async sync() {
      state.syncs += 1;
      const kinds = state.pending.map((p) => p.kind);
      const q = state.pending;
      state.pending = [];
      if (fail({ kinds, applied: state.applied })) throw new Error("GeneralException: sync");
      for (const p of q) p.run();
    },
  };
  globalThis.Word = {
    run: async (fn) => fn(context),
    InsertLocation: { replace: "Replace", after: "After" },
    ChangeTrackingMode: { trackAll: "TrackAll" },
  };
  return state;
}

const { toolTrackedEdit } = await import("../taskpane/shared/tools-word.js");
const REPLACE = { find: "30 degrees", replace: "30 °C", comment_text: "[R1.1] unit / R1.1 / q1" };
const INSERT = {
  find: "for 2 h.",
  insert_after: " Samples were run in triplicate.",
  comment_text: "[R1.1] n / R1.1 / q2",
};

const partial = (r) => [r.ok, r.applied, r.code, r.paragraph_index, r.paragraph_id, r.narrowed];

test("happy path: comment anchored, ok with its id", async () => {
  const s = fakeWord();
  const r = await toolTrackedEdit(REPLACE);
  assert.deepEqual([r.ok, r.word_comment_id, r.paragraph_index], [true, "C1", 0]);
  assert.deepEqual([s.applied, s.comments.length, s.mode], [true, 1, "Off"]);
});

test("insertComment throws after the edit applied: applied:true, CommentInsertFailed", async () => {
  const s = fakeWord({ throwOnInsertComment: true });
  const r = await toolTrackedEdit(REPLACE);
  assert.equal(s.applied, true);
  assert.deepEqual(partial(r), [false, true, "CommentInsertFailed", 0, "P1", false]);
  assert.equal(r.before, PARA);
  assert.equal(r.after, "Cells were held at 30 °C for 2 h.");
  assert.match(r.message, /applied/i);
  assert.equal(s.mode, "Off", "track mode restored");
});

test("the comment sync fails after the edit applied: applied:true", async () => {
  const s = fakeWord({ fail: ({ kinds }) => kinds.includes("comment") });
  const r = await toolTrackedEdit(REPLACE);
  assert.equal(s.applied, true);
  assert.deepEqual(partial(r), [false, true, "CommentInsertFailed", 0, "P1", false]);
});

test("the anchor search sync fails after the edit applied: applied:true", async () => {
  let failed = false;
  const s = fakeWord({
    fail: ({ kinds, applied }) => {
      if (failed || !applied || kinds.length) return false;
      failed = true;
      return true;
    },
  });
  const r = await toolTrackedEdit(REPLACE);
  assert.equal(s.applied, true);
  assert.deepEqual(partial(r), [false, true, "CommentInsertFailed", 0, "P1", false]);
});

test("pure insertion: insertComment throws, still applied:true", async () => {
  const s = fakeWord({ throwOnInsertComment: true });
  const r = await toolTrackedEdit(INSERT);
  assert.equal(s.applied, true);
  assert.deepEqual(partial(r), [false, true, "CommentInsertFailed", 0, "P1", false]);
  assert.equal(r.after, "Cells were held at 30 degrees for 2 h. Samples were run in triplicate.");
});

test("a failed track-mode restore does not undo the success result", async () => {
  const s = fakeWord({ fail: ({ kinds, applied }) => applied && kinds.join() === "mode" });
  const r = await toolTrackedEdit(REPLACE);
  assert.deepEqual([r.ok, r.word_comment_id, s.comments.length], [true, "C1", 1]);
});

test("the edit sync itself fails: nothing applied, the error propagates", async () => {
  const s = fakeWord({ fail: ({ kinds }) => kinds.includes("edit") });
  await assert.rejects(toolTrackedEdit(REPLACE), /sync/);
  assert.equal(s.applied, false);
});
