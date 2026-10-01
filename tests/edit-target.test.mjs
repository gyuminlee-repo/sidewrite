// Which paragraph a tracked edit lands in. Word decides it in toolTrackedEdit
// (taskpane/shared/tools-word.js): one hit of `find` in the document, or,
// with several, one hit inside the paragraph of paragraph_hint. The daemon
// judges the edit before Word is called (forbid list, lost facts, the form
// an unanswered edit would have) and must pick the same paragraph:
// editTarget (daemon/review-tools.mjs). Both get the same inputs here and
// their answers are compared (Codex review 13).
import { test } from "node:test";
import assert from "node:assert/strict";
import { editTarget } from "../daemon/review-tools.mjs";

// A fake Word document of plain paragraphs. mode "uniqueLocalId": every
// paragraph has one (U0, U1, ...); mode "index": none, so the task pane
// falls back to p<N> ids.
function fakeDoc(texts, mode) {
  const coll = (items) => ({ items, load() {} });
  const hits = (text, q) => {
    const out = [];
    for (let i = text.indexOf(q); i !== -1; i = text.indexOf(q, i + q.length)) out.push(i);
    return out;
  };
  const range = (p) => ({
    load() {},
    paragraphs: { getFirst: () => p },
    search: () => coll([]),
    insertText: () => range(p),
    insertComment: () => ({ id: "C1", load() {} }),
    compareLocationWith: () => ({ value: "Before" }),
  });
  const paras = texts.map((text, k) => {
    const p = {
      text,
      style: "Normal",
      ...(mode === "uniqueLocalId" ? { uniqueLocalId: `U${k}` } : {}),
      load() {},
      search: (q) => coll(hits(text, q).map(() => range(p))),
    };
    return p;
  });
  const document = {
    load() {},
    changeTrackingMode: "Off",
    body: {
      paragraphs: coll(paras),
      search: (q) => coll(paras.flatMap((p) => hits(p.text, q).map(() => range(p)))),
    },
  };
  globalThis.Word = {
    run: async (fn) => fn({ document, sync: async () => {} }),
    InsertLocation: { replace: "Replace", after: "After" },
    ChangeTrackingMode: { trackAll: "TrackAll" },
  };
}

const { toolTrackedEdit } = await import("../taskpane/shared/tools-word.js");

const TEXTS = [
  "This is a novel material.",
  "This is a conventional material.",
  "Plain text only.",
  "One material and another material.",
];
// [find, hint as a paragraph position or a literal string, expected]
const CASES = [
  ["conventional", null, 1],
  ["conventional", 0, 1], // one hit in the document: the hint is not read
  ["conventional", "nowhere", 1],
  ["material", null, "Ambiguous"],
  ["material", 0, 0],
  ["material", 1, 1],
  ["material", 2, "NotFound"],
  ["material", 3, "Ambiguous"],
  ["material", "nowhere", "HintNotFound"],
  ["material", "p99", "HintNotFound"],
  ["material", "p1", { uniqueLocalId: "HintNotFound", index: 1 }],
  ["material", "U1", { uniqueLocalId: 1, index: "HintNotFound" }],
  ["absent words", null, "NotFound"],
  ["absent words", 1, "NotFound"],
  ["Material", null, "NotFound"], // case matters
  ["This is a", 1, 1],
  ["This is a", null, "Ambiguous"],
];

for (const mode of ["uniqueLocalId", "index"]) {
  test(`editTarget picks the paragraph toolTrackedEdit edits (${mode} ids)`, async () => {
    const ids = TEXTS.map((_, k) => (mode === "uniqueLocalId" ? `U${k}` : null));
    for (const [find, h, want] of CASES) {
      const hint = typeof h === "number" ? (mode === "uniqueLocalId" ? `U${h}` : `p${h}`) : h;
      const expected = want !== null && typeof want === "object" ? want[mode] : want;
      fakeDoc(TEXTS, mode);
      const r = await toolTrackedEdit({
        find,
        replace: "x",
        ...(hint ? { paragraph_hint: hint } : {}),
        comment_text: "[R1.1] w / y / a",
      });
      const word = r.ok ? r.paragraph_index : r.code;
      const t = editTarget(TEXTS, ids, find, hint);
      const daemon = t.code ?? t.index;
      const name = `${find} / ${hint}`;
      assert.equal(word, expected, `Word: ${name}`);
      assert.equal(daemon, word, `daemon vs Word: ${name}`);
    }
    console.log(`edit target corpus (${mode}): ${CASES.length} cases, same answer on both sides`);
  });
}

test("editTarget without known ids reads a p<N> hint only; other hints are not found", () => {
  assert.deepEqual(editTarget(TEXTS, null, "material", "p1"), { index: 1, code: null });
  assert.deepEqual(editTarget(TEXTS, null, "material", "U1"), {
    index: null,
    code: "HintNotFound",
  });
  assert.deepEqual(editTarget([], null, "material", null), { index: null, code: "NotFound" });
  assert.deepEqual(editTarget(TEXTS, null, "  ", null), { index: null, code: "NotFound" });
});
