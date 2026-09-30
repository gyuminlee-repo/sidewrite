import { test } from "node:test";
import assert from "node:assert/strict";
import {
  normalizeParaText,
  locateParagraph,
  classifyChange,
  classifyRevision,
  summarizeStates,
} from "../daemon/revision-classify.mjs";
import * as classify from "../daemon/revision-classify.mjs";

// Measured in Word on 2026-09-30 (session f2748a6f, office_spike_* outputs:
// comment_on_tracked x3, revision_state before and after the user acted).
// Paragraph texts are verbatim, including U+0005 comment anchors. The user
// accepted p9, rejected p14 and accepted-then-edited p22.
const A = "\u0005";
const HEAD9 =
  "Cultures were grown in a defined mineral medium in sealed flasks with shaking. For every run ";
const TAIL9 =
  " and the headspace was refreshed once per day. Optical density was read at 600 nm in a plate reader.";
const OLD9 = "the incubation temperature was held at 30 degrees";
const NEW9 = "the incubation temperature was maintained at 30 °C";
const TAIL14 =
  " by sonication followed by centrifugation, and the supernatant was kept on ice until use. Protein concentration was estimated with a dye-binding assay against a standard curve.";
const TAIL22 =
  " measurable activity that declined over two hours on ice. Heat treatment removed the activity completely, which suggests that the signal depends on a protein component rather than on a small molecule.";

const change9 = {
  id: "R1.1-1",
  kind: "replace",
  word_comment_id: "1259839512",
  paragraph_index: 9,
  paragraph_id: null,
  before: HEAD9 + OLD9 + TAIL9,
  after: HEAD9 + NEW9 + TAIL9,
};
const change14 = {
  id: "R1.2-1",
  kind: "replace",
  word_comment_id: "1499887155",
  paragraph_index: 14,
  paragraph_id: null,
  before: "A crude extract was prepared" + TAIL14,
  after: "A soluble extract was prepared" + TAIL14,
};
const change22 = {
  id: "R1.3-1",
  kind: "replace",
  word_comment_id: "1165562882",
  paragraph_index: 22,
  paragraph_id: null,
  before: "The crude extract showed" + TAIL22,
  after: "The soluble extract showed" + TAIL22,
};

const pendingSnapshot = {
  paragraphs: [
    { index: 9, id: null, text: HEAD9 + NEW9 + A + OLD9 + A + TAIL9 + A, w_ins: 2, w_del: 1 },
    {
      index: 14,
      id: null,
      text: "A soluble extract was prepared" + A + "A crude extract was prepared" + A + TAIL14 + A,
      w_ins: 2,
      w_del: 1,
    },
    {
      index: 22,
      id: null,
      text: "The soluble extract showed" + A + "The crude extract showed" + A + TAIL22 + A,
      w_ins: 2,
      w_del: 1,
    },
  ],
  comments: [
    { id: "1259839512", paragraph_index: 9 },
    { id: "1499887155", paragraph_index: 14 },
    { id: "1165562882", paragraph_index: 22 },
  ],
};

const settledSnapshot = {
  paragraphs: [
    { index: 9, id: null, text: HEAD9 + NEW9 + A + A + TAIL9 + A, w_ins: 0, w_del: 0 },
    { index: 14, id: null, text: "A crude extract was prepared" + A + TAIL14, w_ins: 0, w_del: 0 },
    {
      index: 22,
      id: null,
      text: "The clarified extract showed" + A + A + TAIL22 + A,
      w_ins: 0,
      w_del: 0,
    },
  ],
  // The deleted-range comments ("b" in the spike) all survived.
  comments: [
    { id: "1259839512", paragraph_index: 9 },
    { id: "1499887155", paragraph_index: 14 },
    { id: "1165562882", paragraph_index: 22 },
  ],
};

test("normalizeParaText drops U+0005 and collapses whitespace", () => {
  assert.equal(normalizeParaText(`a${A}${A}  b\t c ${A}`), "a b c");
  assert.equal(normalizeParaText(null), "");
});

test("measured pending state: all three replaces are pending", () => {
  for (const ch of [change9, change14, change22]) {
    assert.equal(classifyChange(ch, pendingSnapshot).state, "pending", ch.id);
  }
});

test("measured settled state: accepted, rejected, user_edited", () => {
  assert.equal(classifyChange(change9, settledSnapshot).state, "accepted");
  assert.equal(classifyChange(change14, settledSnapshot).state, "rejected");
  assert.equal(classifyChange(change22, settledSnapshot).state, "user_edited");
});

test("located through the rationale comment when the index moved", () => {
  // One paragraph inserted above p9: every index shifts by +1.
  const shifted = {
    paragraphs: settledSnapshot.paragraphs.map((p) => ({ ...p, index: p.index + 1 })),
    comments: settledSnapshot.comments.map((c) => ({
      ...c,
      paragraph_index: c.paragraph_index + 1,
    })),
  };
  const r = classifyChange(change9, shifted);
  assert.deepEqual([r.state, r.paragraph_index, r.method], ["accepted", 10, "comment"]);
});

test("located by unique text when the comment is gone (tool comments cleaned up)", () => {
  const shifted = {
    paragraphs: settledSnapshot.paragraphs.map((p) => ({ ...p, index: p.index - 1 })),
    comments: [],
  };
  const r = classifyChange(change14, shifted);
  assert.deepEqual([r.state, r.paragraph_index, r.method], ["rejected", 13, "text"]);
});

test("paragraph_id wins over a stale index", () => {
  const snap = {
    paragraphs: [
      { index: 0, id: "AAA", text: change9.before, w_ins: 0, w_del: 0 },
      { index: 9, id: "BBB", text: change9.after, w_ins: 0, w_del: 0 },
    ],
    comments: [],
  };
  const r = classifyChange({ ...change9, paragraph_id: "AAA", word_comment_id: null }, snap);
  assert.deepEqual([r.state, r.paragraph_index, r.method], ["rejected", 0, "paragraph_id"]);
});

test("identical paragraphs make the text match ambiguous: unlocated, never a guess", () => {
  const snap = {
    paragraphs: [
      { index: 3, id: null, text: change9.after, w_ins: 0, w_del: 0 },
      { index: 30, id: null, text: change9.after, w_ins: 0, w_del: 0 },
    ],
    comments: [],
  };
  assert.equal(locateParagraph({ ...change9, word_comment_id: null }, snap).method, "ambiguous");
  assert.equal(classifyChange({ ...change9, word_comment_id: null }, snap).state, "unlocated");
});

test("near-duplicate paragraphs (cell-free vs crude) are not confused", () => {
  const snap = {
    paragraphs: [
      {
        index: 10,
        id: null,
        text: "A cell-free extract was prepared" + TAIL14,
        w_ins: 0,
        w_del: 0,
      },
      { index: 14, id: null, text: "A crude extract was prepared" + TAIL14, w_ins: 0, w_del: 0 },
    ],
    comments: [],
  };
  const r = classifyChange({ ...change14, word_comment_id: null }, snap);
  assert.deepEqual([r.state, r.paragraph_index], ["rejected", 14]);
});

// Pure insertion and pure deletion: synthetic, built from the measured
// behavior (pending text shows inserted and deleted text; an inserted-range
// comment disappears on reject).
const insertChange = {
  id: "R1.3-1",
  kind: "insert",
  word_comment_id: "777",
  paragraph_index: 18,
  paragraph_id: null,
  before: "Growth reached a plateau after four days.",
  after: "Growth reached a plateau after four days (optical density at 600 nm).",
};

test("pure insertion: pending while marked, accepted, rejected when comment and text are gone", () => {
  const pending = {
    paragraphs: [{ index: 18, id: null, text: insertChange.after + A, w_ins: 1, w_del: 0 }],
    comments: [{ id: "777", paragraph_index: 18 }],
  };
  assert.equal(classifyChange(insertChange, pending).state, "pending");
  const accepted = {
    paragraphs: [{ index: 18, id: null, text: insertChange.after, w_ins: 0, w_del: 0 }],
    comments: [{ id: "777", paragraph_index: 18 }],
  };
  assert.equal(classifyChange(insertChange, accepted).state, "accepted");
  const rejected = {
    paragraphs: [{ index: 18, id: null, text: insertChange.before, w_ins: 0, w_del: 0 }],
    comments: [],
  };
  assert.equal(classifyChange(insertChange, rejected).state, "rejected");
  const rejectedThenEdited = {
    paragraphs: [
      { index: 18, id: null, text: "Growth plateaued after four days.", w_ins: 0, w_del: 0 },
    ],
    comments: [],
  };
  // Not locatable and the text is neither form: a missing comment alone is no
  // proof of a reject (Codex review 7), so the change is not found.
  assert.equal(classifyChange(insertChange, rejectedThenEdited).state, "unlocated");
});

test("accepted insertion stays accepted after tool comments are cleaned up", () => {
  const snap = {
    paragraphs: [{ index: 18, id: null, text: insertChange.after, w_ins: 0, w_del: 0 }],
    comments: [],
  };
  assert.equal(classifyChange(insertChange, snap).state, "accepted");
});

test("pure deletion: pending text still reads as before, so marks mean pending", () => {
  const del = {
    id: "R1.2-1",
    kind: "delete",
    word_comment_id: "55",
    paragraph_index: 4,
    paragraph_id: null,
    before: "Keep this. Drop this sentence.",
    after: "Keep this.",
  };
  const pending = {
    paragraphs: [
      { index: 4, id: null, text: "Keep this. Drop this sentence." + A, w_ins: 0, w_del: 1 },
    ],
    comments: [{ id: "55", paragraph_index: 4 }],
  };
  assert.equal(classifyChange(del, pending).state, "pending");
  const accepted = {
    paragraphs: [{ index: 4, id: null, text: "Keep this. " + A, w_ins: 0, w_del: 0 }],
    comments: [{ id: "55", paragraph_index: 4 }],
  };
  assert.equal(classifyChange(del, accepted).state, "accepted");
});

test("classifyRevision and summarizeStates cover every change", () => {
  const rev = {
    comments: [
      { ref: "R1.1", changes: [change9] },
      { ref: "R1.2", changes: [change14] },
      { ref: "R1.3", changes: [change22] },
    ],
  };
  const rows = classifyRevision(rev, settledSnapshot);
  assert.deepEqual(
    rows.map((r) => [r.ref, r.change_id, r.state]),
    [
      ["R1.1", "R1.1-1", "accepted"],
      ["R1.2", "R1.2-1", "rejected"],
      ["R1.3", "R1.3-1", "user_edited"],
    ],
  );
  assert.deepEqual(summarizeStates(rows), {
    pending: 0,
    accepted: 1,
    rejected: 1,
    user_edited: 1,
    unlocated: 0,
  });
});

// ---- per-change context (several tool changes in one paragraph) ------------
//
// A change recorded with `context` (a few words either side of the range the
// task pane changed) is judged on its own range, not on the whole paragraph.
const withContext = (ch, find, text, extra = {}) => ({
  ...ch,
  find,
  text,
  context: classify.changeContext(ch.before, find, { kind: ch.kind, ...extra }),
});

test("changeContext: words around the applied range, narrowed range, insertion point", () => {
  const before = "the incubation temperature was held at 30 degrees for a day";
  assert.deepEqual(classify.changeContext(before, "held at 30", { kind: "replace" }), {
    left: "incubation temperature was ",
    right: " degrees for a",
  });
  // Narrowed (L3): the task pane changed only "held" inside "was held at".
  assert.deepEqual(
    classify.changeContext(before, "was held at", { kind: "replace", narrowFind: "held" }),
    { left: "incubation temperature was ", right: " at 30 degrees" },
  );
  // Insertion: the new text goes right after `find`.
  assert.deepEqual(classify.changeContext(before, "30 degrees", { kind: "insert" }), {
    left: "at 30 degrees",
    right: " for a day",
  });
  assert.deepEqual(classify.changeContext("A crude extract was prepared by", "A crude", {}), {
    left: "",
    right: " extract was prepared",
  });
  assert.equal(classify.changeContext(before, "absent words", {}), null);
  assert.equal(classify.changeContext("a b a b", "a b", {}), null, "ambiguous find");
});

test("measured Word fixtures give the same states through the context rule", () => {
  const c9 = withContext(change9, OLD9, NEW9);
  const c14 = withContext(
    change14,
    "A crude extract was prepared",
    "A soluble extract was prepared",
  );
  const c22 = withContext(change22, "The crude extract showed", "The soluble extract showed");
  for (const ch of [c9, c14, c22]) {
    assert.ok(ch.context, ch.id);
    assert.equal(classifyChange(ch, pendingSnapshot).state, "pending", ch.id);
  }
  assert.equal(classifyChange(c9, settledSnapshot).state, "accepted");
  assert.equal(classifyChange(c14, settledSnapshot).state, "rejected");
  assert.equal(classifyChange(c22, settledSnapshot).state, "user_edited");
});

test("insertion and deletion fixtures give the same states through the context rule", () => {
  const ins = withContext(insertChange, "four days", " (optical density at 600 nm)");
  const para = (text, w_ins = 0, w_del = 0) => ({ index: 18, id: null, text, w_ins, w_del });
  const snap = (p, comments = [{ id: "777", paragraph_index: 18 }]) => ({
    paragraphs: [p],
    comments,
  });
  assert.equal(classifyChange(ins, snap(para(insertChange.after + A, 1))).state, "pending");
  assert.equal(classifyChange(ins, snap(para(insertChange.after))).state, "accepted");
  assert.equal(classifyChange(ins, snap(para(insertChange.before), [])).state, "rejected");
  assert.equal(
    classifyChange(ins, snap(para("Growth plateaued after four days."), [])).state,
    "rejected",
  );
  const del = withContext(
    {
      id: "R1.2-1",
      kind: "delete",
      word_comment_id: "55",
      paragraph_index: 18,
      paragraph_id: null,
      before: "Keep this. Drop this sentence.",
      after: "Keep this.",
    },
    " Drop this sentence.",
    "",
  );
  assert.equal(
    classifyChange(
      del,
      snap(para("Keep this. Drop this sentence." + A, 0, 1), [{ id: "55", paragraph_index: 18 }]),
    ).state,
    "pending",
  );
  assert.equal(
    classifyChange(del, snap(para("Keep this." + A), [{ id: "55", paragraph_index: 18 }])).state,
    "accepted",
  );
  assert.equal(
    classifyChange(
      del,
      snap(para("Keep this. Drop this sentence."), [{ id: "55", paragraph_index: 18 }]),
    ).state,
    "rejected",
  );
});

// Two tool changes three words apart in one paragraph. B was recorded while A
// was pending, so the paragraph text B saw holds A's inserted and deleted
// text back to back (the measured pending form).
const P0 =
  "Cells were grown at 30 degrees in a mineral medium with shaking and were harvested after two days.";
const A_OLD = "30 degrees";
const A_NEW = "30 °C";
const B_OLD = "mineral medium";
const B_NEW = "defined medium";
const pendingA = P0.replace(A_OLD, A_NEW + A_OLD);
const twoA = () =>
  withContext(
    {
      id: "R1.1-1",
      kind: "replace",
      word_comment_id: "a",
      paragraph_index: 5,
      paragraph_id: null,
      before: P0,
      after: P0.replace(A_OLD, A_NEW),
    },
    A_OLD,
    A_NEW,
  );
const twoB = () =>
  withContext(
    {
      id: "R1.2-1",
      kind: "replace",
      word_comment_id: "b",
      paragraph_index: 5,
      paragraph_id: null,
      before: pendingA,
      after: pendingA.replace(B_OLD, B_NEW),
    },
    B_OLD,
    B_NEW,
  );
const onePara = (text, marks = 0) => ({
  paragraphs: [{ index: 5, id: null, text, w_ins: marks, w_del: marks }],
  comments: [
    { id: "a", paragraph_index: 5 },
    { id: "b", paragraph_index: 5 },
  ],
});
const states = (snap) => [classifyChange(twoA(), snap).state, classifyChange(twoB(), snap).state];

test("two changes in one paragraph: both accepted", () => {
  const t = P0.replace(A_OLD, A_NEW).replace(B_OLD, B_NEW);
  assert.deepEqual(states(onePara(t)), ["accepted", "accepted"]);
});

test("two changes in one paragraph: A accepted while B is still pending", () => {
  const t = P0.replace(A_OLD, A_NEW).replace(B_OLD, B_NEW + A + B_OLD + A);
  assert.deepEqual(states(onePara(t, 1)), ["accepted", "pending"]);
});

test("two changes in one paragraph: A rejected, B accepted", () => {
  const t = P0.replace(B_OLD, B_NEW);
  assert.deepEqual(states(onePara(t)), ["rejected", "accepted"]);
});

test("two changes in one paragraph: both accepted, then the author rewrote B only", () => {
  const t = P0.replace(A_OLD, A_NEW).replace(B_OLD, "synthetic broth");
  assert.deepEqual(states(onePara(t)), ["accepted", "user_edited"]);
});

test("a change recorded without context keeps the whole-paragraph rule", () => {
  const t = P0.replace(A_OLD, A_NEW).replace(B_OLD, B_NEW);
  const oldA = { ...twoA(), context: null };
  assert.equal(classifyChange(oldA, onePara(t)).state, "user_edited");
});

test("a change with no rationale comment id is located by index or text and classified", () => {
  const bare = { ...change9, word_comment_id: null };
  assert.deepEqual(
    [classifyChange(bare, pendingSnapshot).state, classifyChange(bare, settledSnapshot).state],
    ["pending", "accepted"],
  );
  // No comment and no valid index: a pending replace reads as neither
  // sentence, so it is unlocated rather than guessed.
  const lost = { ...bare, paragraph_index: 3 };
  assert.deepEqual(
    [classifyChange(lost, pendingSnapshot).state, classifyChange(lost, settledSnapshot).state],
    ["unlocated", "accepted"],
  );
});

// Codex review 3: with the context words gone, a bare substring match took
// "7" inside the author's "70" as the accepted suggestion.
const P_DAYS = "Samples were kept for 5 days before the assay.";
const days = () =>
  withContext(
    {
      id: "R1.4-1",
      kind: "replace",
      word_comment_id: "40",
      paragraph_index: 0,
      paragraph_id: null,
      before: P_DAYS,
      after: P_DAYS.replace("5", "7"),
    },
    "5",
    "7",
  );
const para = (text) => ({
  paragraphs: [{ index: 0, id: null, text, w_ins: 0, w_del: 0 }],
  comments: [
    { id: "40", paragraph_index: 0 },
    { id: "41", paragraph_index: 0 },
  ],
});

test("a one-character change inside a longer number the author typed is user_edited", () => {
  assert.equal(
    classifyChange(days(), para("Samples were kept for 70 days before the assay.")).state,
    "user_edited",
  );
  assert.equal(
    classifyChange(days(), para("Samples stayed 70 days in the cold.")).state,
    "user_edited",
  );
});

test("the plain accepted one-character change stays accepted", () => {
  assert.equal(classifyChange(days(), para(P_DAYS.replace("5", "7"))).state, "accepted");
  assert.equal(classifyChange(days(), para(P_DAYS)).state, "rejected");
  // Context words rewritten, the new value once on word boundaries.
  assert.equal(
    classifyChange(days(), para("Samples stayed 7 days in the cold.")).state,
    "accepted",
  );
});

test("a one-character change found several times without context is not called accepted", () => {
  const t = "Samples stayed 7 days in the cold at 7 sites.";
  assert.notEqual(classifyChange(days(), para(t)).state, "accepted");
});

test("Korean counters keep the context rule: 5개 -> 7개 accepted, 70개 edited", () => {
  const before = "시료 5개를 준비했다.";
  const ch = withContext(
    {
      id: "R1.5-1",
      kind: "replace",
      word_comment_id: "41",
      paragraph_index: 0,
      paragraph_id: null,
      before,
      after: before.replace("5", "7"),
    },
    "5",
    "7",
  );
  assert.equal(classifyChange(ch, para("시료 7개를 준비했다.")).state, "accepted");
  assert.equal(classifyChange(ch, para("시료 70개를 준비했다.")).state, "user_edited");
});

// Codex review 7: a missing rationale comment is evidence of a rejected
// insertion only when the text reads the old form; the comment cleanup
// removes it from accepted and edited insertions too.
const grow = (extra = {}) =>
  withContext(
    {
      id: "R1.6-1",
      kind: "insert",
      word_comment_id: "60",
      paragraph_index: 0,
      paragraph_id: "P60",
      before: "Cells grow.",
      after: "Cells rapidly grow.",
      ...extra,
    },
    "Cells",
    " rapidly",
    { kind: "insert" },
  );
const growPara = (text, comments = []) => ({
  paragraphs: [{ index: 0, id: "P60", text, w_ins: 0, w_del: 0 }],
  comments,
});

test("insertion accepted then reworded, comment cleaned up: user_edited, not rejected", () => {
  const withComment = growPara("Cells quickly grow.", [{ id: "60", paragraph_index: 0 }]);
  assert.equal(classifyChange(grow(), withComment).state, "user_edited");
  assert.equal(classifyChange(grow(), growPara("Cells quickly grow.")).state, "user_edited");
  // Recorded before context existed: the whole-paragraph rule says the same.
  const old = { ...grow(), context: null };
  assert.equal(classifyChange(old, growPara("Cells quickly grow.")).state, "user_edited");
});

test("insertion accepted and comment cleaned up: accepted; rejected when the text is the old one", () => {
  assert.equal(classifyChange(grow(), growPara("Cells rapidly grow.")).state, "accepted");
  assert.equal(classifyChange(grow(), growPara("Cells grow.")).state, "rejected");
  const old = { ...grow(), context: null };
  assert.equal(classifyChange(old, growPara("Cells rapidly grow.")).state, "accepted");
  assert.equal(classifyChange(old, growPara("Cells grow.")).state, "rejected");
});

test("unlocated insertion: a gone comment means rejected only when every match is the old text", () => {
  const noId = { ...grow(), paragraph_id: null, context: null, paragraph_index: 5 };
  const twice = (a, b) => ({
    paragraphs: [
      { index: 0, id: null, text: a, w_ins: 0, w_del: 0 },
      { index: 1, id: null, text: b, w_ins: 0, w_del: 0 },
    ],
    comments: [],
  });
  assert.equal(classifyChange(noId, twice("Cells grow.", "Cells grow.")).state, "rejected");
  assert.equal(
    classifyChange(noId, twice("Cells grow.", "Cells rapidly grow.")).state,
    "unlocated",
  );
  // The comment was removed by the cleanup, not by Word: no evidence.
  const cleaned = { ...noId, comment_cleaned_at: "2026-09-30T00:00:00.000Z" };
  assert.equal(classifyChange(cleaned, twice("Cells grow.", "Cells grow.")).state, "unlocated");
});
