// The rationale comment of an edit whose old words Word dropped. A replace or
// a deletion inside a sibling change's pending insertion leaves no tracked
// deletion: Word removes inserted text outright. The task pane looked for
// the old words to put the comment on, found nothing and gave up
// (CommentAnchorMissing), so the change stood in the document with no
// comment (replays of 2026-10-05: four changes, all of this shape). The
// comment now goes on what stands: the new words of a replace, the words
// around the place of a deletion. Invented text (bridge works).
import { test } from "node:test";
import assert from "node:assert/strict";
import { paragraphHash } from "../daemon/doc-diff.mjs";
import { anchorQueries, wordsAround } from "../taskpane/shared/review-view.js";

// A fake Word paragraph list whose text changes with an edit. pending:
// [start, end) of the paragraph text that is a pending tracked insertion.
// A replace inside it drops the old words (Word's behaviour); elsewhere the
// old words stay as tracked deleted text, after the new ones.
function fakeDoc(paras, { pending = null } = {}) {
  const state = { comments: [] };
  const coll = (items) => ({ items, load() {} });
  const hits = (text, q) => {
    const out = [];
    for (let i = text.indexOf(q); i !== -1; i = text.indexOf(q, i + q.length)) out.push(i);
    return out;
  };
  const range = (p, start, q) => ({
    owner: p,
    start,
    text: q,
    load() {},
    paragraphs: { getFirst: () => p },
    // A range searches the paragraph text where it stands now.
    search: (s) =>
      coll(
        hits(q == null ? p.text : p.text.slice(start, start + q.length), s).map((i) =>
          range(p, (q == null ? 0 : start) + i, s),
        ),
      ),
    insertText(s) {
      // The part of the old words inside the pending insertion is dropped;
      // the rest stays (tracked deleted).
      const out = (n) => !pending || start + n < pending[0] || start + n >= pending[1];
      const kept = [...q].filter((_, n) => out(n)).join("");
      p.text = p.text.slice(0, start) + s + kept + p.text.slice(start + q.length);
      return range(p, start, s);
    },
    insertComment(content) {
      const c = { id: `C${state.comments.length + 1}`, content, on: q, load() {} };
      state.comments.push(c);
      return c;
    },
    compareLocationWith: () => ({ value: "Before" }),
    getReviewedText: () => ({ value: p.text }),
  });
  const items = paras.map((text, k) => {
    const p = { k, text, style: "Normal", uniqueLocalId: `U${k}`, load() {} };
    p.getRange = () => range(p, 0, null);
    p.search = (q) => coll(hits(p.text, q).map((i) => range(p, i, q)));
    p.getComments = () => coll(state.comments);
    return p;
  });
  const document = {
    load() {},
    changeTrackingMode: "Off",
    getSelection: () => ({ select() {} }),
    body: {
      paragraphs: coll(items),
      search: (q) => coll(items.flatMap((p) => hits(p.text, q).map((i) => range(p, i, q)))),
    },
  };
  globalThis.Word = {
    run: async (fn) => fn({ document, sync: async () => {} }),
    InsertLocation: { replace: "Replace", after: "After" },
    ChangeTrackingMode: { trackAll: "TrackAll" },
    ChangeTrackingVersion: { current: "Current" },
  };
  return { state, items };
}

const { toolTrackedEdit } = await import("../taskpane/shared/tools-word.js");
const COMMENT = "[R3.1] w / y / a";
const HEAD = "Crews checked the north span daily. ";
const I1 = "New bearings with thick steel plates were fitted.";
const TEXT = HEAD + I1;
const INSERTION = [HEAD.length, TEXT.length];
const edit = (find, replace, extra = {}) => ({
  find,
  replace,
  comment_text: COMMENT,
  expected_before_hash: paragraphHash(TEXT),
  ...extra,
});

test("wordsAround: the words on both sides of the place of a deletion", () => {
  assert.equal(wordsAround(TEXT, "thick ", "thick ", 1), "with steel");
  assert.equal(wordsAround(TEXT, "thick ", "thick ", 2), "bearings with steel plates");
  // A narrowed find: the place is where the narrowed part was.
  assert.equal(wordsAround(TEXT, "with thick steel", "thick ", 1), "with steel");
  // At the end of the paragraph: the words before it only.
  assert.equal(wordsAround(TEXT, " were fitted.", " were fitted.", 1), "plates");
  assert.equal(wordsAround(TEXT, "absent", "absent", 1), "");
});

test("a deletion inside a pending insertion gets its comment on the words around it", async () => {
  const { state, items } = fakeDoc(["Title", TEXT], { pending: INSERTION });
  const r = await toolTrackedEdit(edit("thick ", ""));
  assert.equal(items[1].text, HEAD + I1.replace("thick ", ""), "Word dropped the words");
  assert.deepEqual([r.ok, r.code, r.word_comment_id], [true, undefined, "C1"]);
  assert.deepEqual(
    state.comments.map((c) => [c.content, c.on]),
    [[COMMENT, "with steel"]],
  );
});

test("a narrowed deletion inside a pending insertion: the same", async () => {
  const { state } = fakeDoc(["Title", TEXT], { pending: INSERTION });
  const r = await toolTrackedEdit(
    edit("with thick steel", "with steel", { narrow: { find: "thick ", replace: "" } }),
  );
  assert.deepEqual([r.ok, r.narrowed, r.word_comment_id], [true, true, "C1"]);
  assert.deepEqual(
    state.comments.map((c) => c.on),
    ["with steel"],
  );
});

test("a replace inside a pending insertion gets its comment on its new words", async () => {
  const { state } = fakeDoc(["Title", TEXT], { pending: INSERTION });
  const r = await toolTrackedEdit(edit("thick steel plates", "shims"));
  assert.deepEqual([r.ok, r.word_comment_id], [true, "C1"]);
  assert.deepEqual(
    state.comments.map((c) => c.on),
    ["shims"],
  );
});

test("controls: the old words still there take the comment, as before", async () => {
  // Not inside an insertion: the old words stay as tracked deleted text.
  let { state } = fakeDoc(["Title", TEXT]);
  let r = await toolTrackedEdit(edit("thick ", ""));
  assert.deepEqual([r.ok, state.comments.map((c) => c.on)], [true, ["thick "]]);
  ({ state } = fakeDoc(["Title", TEXT]));
  r = await toolTrackedEdit(edit("thick steel plates", "shims"));
  assert.deepEqual([r.ok, state.comments.map((c) => c.on)], [true, ["thick steel plates"]]);
});

test("controls: words around the place found more than once give no comment, as before", async () => {
  // Two words on each side are once in the paragraph: those take it.
  const twice = "Crews fitted with steel plates. New bearings with thick steel plates were fitted.";
  let { state } = fakeDoc(["Title", twice], { pending: [32, twice.length] });
  let r = await toolTrackedEdit({
    ...edit("thick ", ""),
    expected_before_hash: paragraphHash(twice),
  });
  assert.deepEqual([r.ok, state.comments.map((c) => c.on)], [true, ["bearings with steel plates"]]);
  // Up to three words on each side, joined, are there twice: the words on
  // one side that are there once take it (five words before the place,
  // from the end of the first sentence), never a place that is there twice.
  const first = "Crews fitted it with steel plates there. ";
  const same = `${first}Crews fitted it with thick steel plates there.`;
  ({ state } = fakeDoc(["Title", same], { pending: [first.length, same.length] }));
  r = await toolTrackedEdit({ ...edit("thick ", ""), expected_before_hash: paragraphHash(same) });
  assert.deepEqual(
    [r.ok, state.comments.map((c) => c.on)],
    [true, ["there. Crews fitted it with"]],
  );
});

// Replay of 2026-10-05, change R6.5-11: a deletion that took a sibling's
// pending insertion and the words before it. Word dropped the inserted
// words and kept the others as tracked deleted text, so the words on the two
// sides of the place were not next to each other and no joined text was
// found: CommentAnchorMissing. The words on one side are. Invented text.
const LEAD = "Spans were checked (Fig. 2A). ";
const OLD = "This repair ";
const NEW = "keeps the deck level";
const TAIL = "restores the load path. Loads were logged (Fig. 2A).";
const SHAPE = LEAD + OLD + NEW + TAIL;
const NEW_AT = [LEAD.length + OLD.length, LEAD.length + OLD.length + NEW.length];

test("anchorQueries: joined words first, then one side, then the first words of the paragraph", () => {
  const q = anchorQueries(SHAPE, OLD + NEW, OLD + NEW);
  assert.deepEqual(q.slice(0, 5), [
    "2A). restores",
    "(Fig. 2A). restores the",
    "checked (Fig. 2A). restores the load",
    "2A).",
    "restores",
  ]);
  assert.equal(q.at(-1), "Spans were checked (Fig. 2A). This repair keeps");
  assert.equal(new Set(q).size, q.length);
  assert.ok(q.every((t) => t && t.length <= 255));
  assert.deepEqual(anchorQueries(SHAPE, "absent", "absent"), []);
});

test("a deletion of kept words and a pending insertion gets its comment on one side of its place", async () => {
  const { state, items } = fakeDoc(["Title", SHAPE], { pending: NEW_AT });
  const r = await toolTrackedEdit({
    ...edit(OLD + NEW, ""),
    expected_before_hash: paragraphHash(SHAPE),
  });
  assert.equal(items[1].text, LEAD + OLD + TAIL, "Word dropped the inserted words, kept the rest");
  // "2A)." is there twice; the first word after the place is there once.
  assert.deepEqual([r.ok, r.code, r.word_comment_id], [true, undefined, "C1"]);
  assert.deepEqual(
    state.comments.map((c) => [c.content, c.on]),
    [[COMMENT, "restores"]],
  );
});

test("a deletion with no words of its paragraph there once gets its comment on the paragraph", async () => {
  // Two sentences the same once the dropped word is gone; six words on
  // either side of the place and the first eight words of the paragraph
  // stay inside one of them.
  const first =
    "Crews fitted it with new and very good steel plates near the old pier on the north side. ";
  const same = `${first}Crews fitted it with new and very good steel thick plates near the old pier on the north side.`;
  const { state } = fakeDoc(["Title", same], { pending: [first.length, same.length] });
  const r = await toolTrackedEdit({
    ...edit("thick ", ""),
    expected_before_hash: paragraphHash(same),
  });
  assert.deepEqual([r.ok, r.word_comment_id], [true, "C1"]);
  assert.deepEqual(
    state.comments.map((c) => c.on),
    [null],
    "on the paragraph range",
  );
});
