// toolTrackedEdit (taskpane/shared/tools-word.js) on a fake Word object
// model with failure injection. Queued operations run at context.sync(), as
// in Office.js; a failing sync drops its queue. The point under test: once
// the sync that carries the body edit succeeded, any later failure (anchor
// search, insertComment, the comment sync, the track-mode restore) must
// still return applied:true, so the daemon records the change.
import { test } from "node:test";
import assert from "node:assert/strict";

const PARA = "Cells were held at 30 degrees for 2 h.";

// fail(batch): the sync throws and its queue is dropped. failAfter(batch):
// the queue runs and the sync throws all the same (Word did it and then
// reported an error). error: what a failing sync throws.
function fakeWord({
  fail = () => false,
  failAfter = () => false,
  throwOnInsertComment = false,
  error = () => new Error("GeneralException: sync"),
  // Comments already in the paragraph, on other words than this edit.
  comments = [],
  // What the paragraph holds besides text: counts of fields, inline
  // pictures, footnotes. noPlainApi: this Word cannot list them.
  fields = 0,
  pictures = 0,
  footnotes = 0,
  noPlainApi = false,
} = {}) {
  const state = {
    applied: false,
    comments: comments.map((c) => ({ ...c, foreign: true })),
    mode: "Off",
    pending: [],
    syncs: 0,
  };
  const coll = (items) => ({ items, load() {} });
  const occurrences = (text, q) => {
    const out = [];
    for (let i = text.indexOf(q); i !== -1; i = text.indexOf(q, i + 1)) out.push(q);
    return out;
  };
  let para;
  let lastComment = 0;
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
      const c = { id: `C${++lastComment}`, load() {} };
      state.pending.push({
        kind: "comment",
        run: () => state.comments.push({ id: c.id, content }),
      });
      return c;
    },
    // A comment this edit added lies on its words; another one does not.
    getComments: () =>
      coll(
        state.comments.map((c) => ({
          ...c,
          getRange: () => ({
            compareLocationWith: () => ({ value: c.foreign ? "Before" : "Inside" }),
          }),
        })),
      ),
    compareLocationWith: () => ({ value: "Before" }),
    // The paragraph is compared before the edit (expected_before_hash).
    getRange() {
      return this;
    },
    getReviewedText: () => ({ value: text }),
  });
  para = { ...range(PARA), uniqueLocalId: "P1", style: "Normal" };
  const things = (n) => ({
    items: Array.from({ length: n }, () => ({})),
    load() {
      if (noPlainApi) throw new Error("ApiNotFound");
    },
  });
  Object.assign(para, {
    fields: things(fields),
    inlinePictures: things(pictures),
    footnotes: things(footnotes),
    endnotes: things(0),
    contentControls: things(0),
  });
  const document = {
    load() {},
    getSelection: () => ({ select() {} }),
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
      const batch = { kinds, applied: state.applied, sync: state.syncs };
      if (fail(batch)) throw error();
      for (const p of q) p.run();
      if (failAfter(batch)) throw error();
    },
  };
  globalThis.Word = {
    run: async (fn) => fn(context),
    InsertLocation: { replace: "Replace", after: "After" },
    ChangeTrackingMode: { trackAll: "TrackAll" },
    ChangeTrackingVersion: { current: "Current" },
  };
  return state;
}

const { toolTrackedEdit } = await import("../taskpane/shared/tools-word.js");
const { paragraphHash } = await import("../daemon/doc-diff.mjs");
// The paragraph is as the daemon read it: the comparison before the edit passes.
const expected_before_hash = paragraphHash(PARA);
const REPLACE = {
  find: "30 degrees",
  replace: "30 °C",
  comment_text: "[R1.1] unit / R1.1 / q1",
  expected_before_hash,
};
const INSERT = {
  find: "for 2 h.",
  insert_after: " Samples were run in triplicate.",
  comment_text: "[R1.1] n / R1.1 / q2",
  expected_before_hash,
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

test("the anchor search sync fails once after the edit applied: the comment is added on the second try", async () => {
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
  assert.deepEqual([r.ok, r.word_comment_id, s.comments.length], [true, "C1", 1]);
  assert.equal(r.comment_attempts, 2);
  assert.deepEqual(
    r.comment_errors.map((e) => [e.step, e.message]),
    [["anchor_search", "GeneralException: sync"]],
  );
});

// The s25 acceptance run: Word applied a tracked replace and the comment
// step threw; the change was recorded with no rationale comment
// (CommentInsertFailed) and nothing said what Word had answered.
test("the comment sync fails once: one more try adds the comment, and the result says what failed first", async () => {
  let n = 0;
  const officeError = () =>
    Object.assign(new Error("GeneralException"), {
      code: "GeneralException",
      debugInfo: { code: "GeneralException", errorLocation: "Range.insertComment" },
    });
  const s = fakeWord({
    fail: ({ kinds }) => kinds.includes("comment") && ++n === 1,
    error: officeError,
  });
  const r = await toolTrackedEdit(REPLACE);
  assert.deepEqual([r.ok, r.applied, r.code], [true, undefined, undefined]);
  assert.deepEqual([s.comments.length, s.comments[0].content], [1, REPLACE.comment_text]);
  assert.equal(r.word_comment_id, s.comments[0].id);
  assert.equal(r.comment_attempts, 2);
  assert.deepEqual(r.comment_errors, [
    {
      step: "insert_comment",
      code: "GeneralException",
      location: "Range.insertComment",
      message: "GeneralException",
    },
  ]);
  assert.equal(s.mode, "Off");
});

test("the comment sync keeps failing: applied without its comment, with both errors and no second edit", async () => {
  let edits = 0;
  const s = fakeWord({
    fail: ({ kinds }) => (kinds.includes("edit") && edits++, kinds.includes("comment")),
    error: () => Object.assign(new Error("GeneralException"), { code: "GeneralException" }),
  });
  const r = await toolTrackedEdit(REPLACE);
  assert.deepEqual(partial(r), [false, true, "CommentInsertFailed", 0, "P1", false]);
  assert.deepEqual([s.comments.length, edits], [0, 1]);
  assert.deepEqual(
    r.comment_errors.map((e) => [e.step, e.code, e.location]),
    [
      ["insert_comment", "GeneralException", null],
      ["insert_comment", "GeneralException", null],
    ],
  );
  assert.match(r.message, /applied.*GeneralException/i);
  // Nothing of the manuscript is in what goes to the log.
  assert.doesNotMatch(JSON.stringify(r.comment_errors), /degrees|Cells were held/);
});

test("the comment went in although its sync threw: the second try finds it and adds no second one", async () => {
  let n = 0;
  const s = fakeWord({ failAfter: ({ kinds }) => kinds.includes("comment") && ++n === 1 });
  const r = await toolTrackedEdit(REPLACE);
  assert.deepEqual([r.ok, s.comments.length], [true, 1]);
  assert.equal(r.word_comment_id, s.comments[0].id);
  assert.equal(r.comment_attempts, 2);
});

test("the happy path says nothing about tries", async () => {
  fakeWord();
  const r = await toolTrackedEdit(REPLACE);
  assert.deepEqual([r.comment_attempts, r.comment_errors], [undefined, undefined]);
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

// Codex review (P2): the second try took any comment of the paragraph with
// the same text for its own, also the rationale comment of an earlier change
// with the same wording. Two changes then shared one comment id.
test("the second try does not take the comment of another change for its own", async () => {
  let n = 0;
  const other = { id: "K9", content: REPLACE.comment_text };
  let s = fakeWord({
    comments: [other],
    fail: ({ kinds }) => kinds.includes("comment") && ++n === 1,
  });
  let r = await toolTrackedEdit(REPLACE);
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.notEqual(r.word_comment_id, "K9");
  assert.deepEqual(
    s.comments.map((c) => c.id),
    ["K9", r.word_comment_id],
  );
  // The comment went in although its sync threw, and the daemon knows it as
  // the comment of another change: not reused either.
  n = 0;
  s = fakeWord({ failAfter: ({ kinds }) => kinds.includes("comment") && ++n === 1 });
  r = await toolTrackedEdit({ ...REPLACE, known_comment_ids: ["C1"] });
  assert.deepEqual([r.ok, r.word_comment_id, s.comments.length], [true, "C2", 2]);
});

// ---- the parts of a rewritten paragraph (daemon/review-tools.mjs applyTrackedRewrite)

test("the task pane names the capability a rewrite needs", async () => {
  const { PANE_CAPS } = await import("../taskpane/shared/tools-word.js");
  assert.deepEqual(PANE_CAPS, ["target_hash", "tracked_rewrite", "field_guard"]);
});

test("no_comment: the edit is applied and no comment is added", async () => {
  const s = fakeWord();
  const { comment_text, ...bare } = REPLACE;
  const r = await toolTrackedEdit({ ...bare, no_comment: true });
  assert.deepEqual(
    [r.ok, r.word_comment_id, r.no_comment, r.target_checked],
    [true, null, true, true],
  );
  assert.deepEqual([s.applied, s.comments.length, s.mode], [true, 0, "Off"]);
  assert.equal(r.after, "Cells were held at 30 °C for 2 h.");
  // Without the flag a comment text is still required.
  await assert.rejects(() => toolTrackedEdit(bare), /comment_text must start with/);
});

test("plain_only: a paragraph with a field, a picture or a footnote is not touched", async () => {
  for (const what of [{ fields: 1 }, { pictures: 2 }, { footnotes: 1 }]) {
    const s = fakeWord(what);
    const r = await toolTrackedEdit({ ...REPLACE, plain_only: true });
    assert.deepEqual(
      [r.ok, r.code, s.applied, s.comments.length],
      [false, "ParagraphNotPlain", false, 0],
    );
    assert.match(r.message, /Nothing was changed/);
  }
  // Plain text: applied as usual. Without the flag nothing is looked at.
  let s = fakeWord();
  assert.equal((await toolTrackedEdit({ ...REPLACE, plain_only: true })).ok, true);
  assert.equal(s.applied, true);
  s = fakeWord({ fields: 1 });
  assert.equal((await toolTrackedEdit(REPLACE)).ok, true);
  // A Word that cannot list them: not rewritten as a whole.
  s = fakeWord({ noPlainApi: true });
  const r = await toolTrackedEdit({ ...REPLACE, plain_only: true });
  assert.deepEqual([r.ok, r.code, s.applied], [false, "ParagraphNotPlain", false]);
});
