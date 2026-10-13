// "Check other places": a second turn of a comment that is given the places
// its changes did not touch (taskpane/shared/follow-places.js) and told to
// make them agree. All texts here are made up (a bridge inspection report).
import { plainPane } from "./helpers/plain-pane.mjs";
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { followNote, followPlan } from "../daemon/paragraph-check.mjs";
import { followState, MAX_CHARS } from "../taskpane/shared/follow-places.js";
import { createReviewService } from "../daemon/review-service.mjs";
import { applyTrackedEdit, applyTrackedRewrite, askAuthor } from "../daemon/review-tools.mjs";
import {
  RevisionSchema,
  loadRevision,
  saveRevision,
  replaceComments,
  setFollowThrough,
  answerQuestion,
} from "../daemon/revision-store.mjs";

const PARAGRAPHS = [
  "Deck survey",
  "Rust was eliminated on all 14 spans after recoating in 2031 (Fig. 2).",
  "The coating held through the winter.",
  "Figure 2. Rust coverage per span.",
  "Cable tension",
  "Tension stayed within limits.",
];
const COMMENTS = [
  { ref: "R1.1", reviewer: 1, text: "Saying rust was eliminated is too strong." },
  { ref: "R1.2", reviewer: 1, text: "Name the paint supplier." },
];

const T0 = "2026-01-01T00:00:00.000Z"; // the status refresh
const REV = {
  snapshot: { taken_at: T0, paragraphs: PARAGRAPHS, ids: ["k0", "k1", "k2", "k3", "k4", "k5"] },
  comments: [
    {
      ...COMMENTS[0],
      changes: [
        {
          id: "R1.1-1",
          kind: "replace",
          find: "was eliminated",
          text: "was reduced",
          paragraph_index: 1,
          paragraph_id: "k1",
          rationale: { what: "w", why: "y", addresses: "a", requirement_ids: ["q1"] },
          narrowed_from: null,
          state: "pending",
          created_at: "2026-01-02T00:00:00.000Z", // after the refresh: put into the text
        },
        {
          id: "R1.1-2",
          kind: "delete",
          find: " through the winter",
          text: "",
          paragraph_index: 2,
          paragraph_id: "k2",
          rationale: { what: "w", why: "y", addresses: "a", requirement_ids: [] },
          narrowed_from: null,
          state: "rejected",
          created_at: "2025-12-31T00:00:00.000Z", // before the refresh, which found it rejected
        },
      ],
    },
    { ...COMMENTS[1], changes: [] },
  ],
};

test("followPlan: the changes of the comment, their paragraphs as they read now, the other places", () => {
  const plan = followPlan(REV, "R1.1");
  assert.deepEqual(plan.changes, [
    {
      id: "R1.1-1",
      kind: "replace",
      before: "was eliminated",
      after: "was reduced",
      requirement_ids: ["q1"],
    },
  ]);
  assert.deepEqual(plan.changed, [
    {
      index: 1,
      id: "k1",
      text: "Rust was reduced on all 14 spans after recoating in 2031 (Fig. 2).",
    },
  ]);
  assert.deepEqual(
    plan.candidates.map((p) => [p.id, p.kind, p.why, p.text]),
    [
      ["k0", "heading", "heading above a changed paragraph", "Deck survey"],
      ["k3", "caption", "legend of Fig. 2", "Figure 2. Rust coverage per span."],
      // Not k2: the author rejected a change of this comment there.
    ],
  );
  assert.deepEqual(followPlan(REV, "R1.2").candidates, []);
});

test("followNote: the instruction a follow-through turn carries", () => {
  const note = followNote(REV, "R1.1", "F1");
  assert.equal(
    note,
    [
      "[Other places of R1.1]",
      "This turn does one thing. The changes made for R1.1 so far:",
      '- R1.1-1 (q1): "was eliminated" → "was reduced"',
      'The 2 places below are candidates to check. Each is given as [paragraph id] (what it is; why it is listed) and its text as it stands now. Most listed places need no change: a place is listed for where it stands, which is no sign of a problem. For each place decide whether it still states what these changes withdrew, weakened or made conditional, or still uses the name of a quantity or a method that they changed. Where it does, change it so that it agrees with them: office_tracked_edit or office_tracked_rewrite, comment_ref "R1.1", with the requirement_ids of the change it follows. Where agreeing takes data, a number, an analysis or a citation the manuscript does not have, ask with office_ask_author and name the place. Where it already agrees, leave it.',
      "Only the listed places may be changed in this turn: an edit of any other paragraph is refused. The general rules above about changing a claim everywhere it is made, finishing every request and going to every place of a term do not apply in this turn.",
      "<<<F1",
      "[k0] (heading; heading above a changed paragraph) Deck survey",
      "[k3] (caption; legend of Fig. 2) Figure 2. Rust coverage per span.",
      ">>>F1",
    ].join("\n"),
  );
  // A deletion and an insertion, and a long text cut short.
  const rev = structuredClone(REV);
  rev.comments[0].changes = [
    { ...rev.comments[0].changes[0], kind: "delete", find: "Rust ", text: "" },
    {
      ...rev.comments[0].changes[0],
      id: "R1.1-3",
      kind: "insert",
      find: "recoating",
      text: ` ${"x".repeat(300)}`,
      rationale: null,
    },
  ];
  const lines = followNote(rev, "R1.1", "F1").split("\n");
  assert.equal(lines[2], '- R1.1-1 (q1): took out "Rust "');
  assert.equal(lines[3], `- R1.1-3: after "recoating" put in " ${"x".repeat(196)}..."`);
  // Nothing to check: no note.
  assert.throws(
    () => followNote(REV, "R1.2", "F1"),
    /R1\.2 has no change in the document, so there is no other place to check/,
  );
  // No change in the document, and the author answered a question that
  // names its place: the note is about that place alone.
  const answered = structuredClone(REV);
  answered.comments[1].questions = [
    {
      id: "a1",
      question: "Who supplied the paint?",
      why: "w",
      answer: "Harbour Coatings.",
      paragraph_id: "k2",
    },
    { id: "a2", question: "Which batch?", why: "w", answer: "", paragraph_id: "k5" },
  ];
  assert.deepEqual(followPlan(answered, "R1.2").candidates, [
    {
      index: 2,
      id: "k2",
      kind: "body",
      why: "the author's answer to a1 goes here",
      answer_ids: ["a1"],
      text: PARAGRAPHS[2],
    },
  ]);
  assert.equal(
    followNote(answered, "R1.2", "F1"),
    [
      "[Other places of R1.2]",
      "This turn does one thing. No change made for R1.2 is in the document. The place below is given as [paragraph id] (what it is; why it is listed) and its text as it stands now.",
      'The author answered question a1 of this comment (the answer stands under "Author answers" above), and the answer goes into [k2]. Write it in there as far as the comment asks for it: office_tracked_edit or office_tracked_rewrite, comment_ref "R1.2". Where the place already says it, leave it.',
      "Only the listed places may be changed in this turn: an edit of any other paragraph is refused. The general rules above about changing a claim everywhere it is made, finishing every request and going to every place of a term do not apply in this turn.",
      "<<<F1",
      "[k2] (body; the author's answer to a1 goes here) The coating held through the winter.",
      ">>>F1",
    ].join("\n"),
  );
  // With changes: the line of the answer stands after the changes.
  const both = structuredClone(REV);
  both.comments[0].questions = [answered.comments[1].questions[0]];
  both.comments[0].questions[0].paragraph_id = "k5";
  const withBoth = followNote(both, "R1.1", "F1").split("\n");
  assert.match(
    withBoth[3],
    /^The author answered question a1 of this comment .* goes into \[k5\]\./,
  );
  assert.match(withBoth[4], /^The 3 places below are candidates to check\./);
  assert.equal(
    withBoth[7],
    "[k5] (body; the author's answer to a1 goes here) Tension stayed within limits.",
  );
  const alone = structuredClone(REV);
  alone.snapshot = { taken_at: T0, paragraphs: [PARAGRAPHS[1]], ids: ["k1"] };
  assert.throws(
    () => followNote(alone, "R1.1", "F1"),
    /No other place was found for the changes of R1\.1/,
  );
});

test("the field of the review file: old files load, a comment-list save keeps it", () => {
  let rev = RevisionSchema.parse({
    version: 1,
    manuscript: "m",
    updated_at: "t",
    comments: [{ ref: "R1.1", reviewer: 1, text: "c" }],
  });
  // No default: an old file saves as it was loaded.
  assert.ok(!("follow_through" in rev.comments[0]));
  const record = {
    at: "2026-01-03T00:00:00.000Z",
    candidate_ids: ["k0", "k3"],
    change_ids: ["R1.1-2"],
    questions: 1,
    seen_change_ids: ["R1.1-1", "R1.1-2"],
  };
  rev = setFollowThrough(rev, "R1.1", record);
  assert.deepEqual(rev.comments[0].follow_through, record);
  assert.equal(setFollowThrough(rev, "R9.9", record), null);
  assert.deepEqual(RevisionSchema.parse(JSON.parse(JSON.stringify(rev))), rev);
  // The answers a check saw: optional, no default (a record from before it
  // has no such field and saves as it was loaded).
  assert.ok(!("seen_answer_ids" in rev.comments[0].follow_through));
  const seen = setFollowThrough(rev, "R1.1", { ...record, seen_answer_ids: ["a1"] });
  assert.deepEqual(RevisionSchema.parse(JSON.parse(JSON.stringify(seen))), seen);
  assert.deepEqual(seen.comments[0].follow_through.seen_answer_ids, ["a1"]);
  const saved = replaceComments(rev, [{ ref: "R1.1", reviewer: 1, text: "c edited" }]);
  assert.deepEqual(saved.comments[0].follow_through, record);
});

const R = { what: "w", why: "y", addresses: "a" };

async function setup(paragraphs = PARAGRAPHS) {
  const dir = await mkdtemp(join(tmpdir(), "follow-svc-"));
  const doc = join(dir, "paper.docx");
  const events = [];
  const lines = [];
  const pushed = [];
  // Word, as the task pane answers for it: the paragraph the edit went into.
  const word = { paragraph_index: 1 };
  const wordResult = () => ({
    ok: true,
    word_comment_id: 5,
    paragraph_index: word.paragraph_index,
    before: word.before ?? "b",
    after: word.after ?? "a",
    target_checked: true,
  });
  const svc = createReviewService({
    getContext: () => ({ activeDoc: doc, caps: ["target_hash"] }),
    callTool: async () => wordResult(),
    pushTurn: (key, host, text, turn) => pushed.push({ key, host, text, turn }),
    oneshot: null,
    sendEvent: (key, obj) => events.push(obj),
    log: (s) => lines.push(s),
    env: { CC_OFFICE_WSL_PATHS: "0" },
  });
  const ask = (type, payload = {}) =>
    new Promise((resolve) => {
      svc.handlers[type]({ type, request_id: "q1", ...plainPane(payload) }, resolve, "k", "word");
    });
  await ask("review_save_comments", { comments: COMMENTS });
  await ask("review_refresh", {
    snapshot: {
      paragraphs: paragraphs.map((text, index) => ({
        index,
        id: `p${index}`,
        text,
        reviewed_text: text,
        w_ins: 0,
        w_del: 0,
      })),
      comments: [],
    },
  });
  const deps = { ...svc.trackedEditDeps("k"), call: async () => wordResult() };
  const edit = (find, replace, turn) =>
    applyTrackedEdit(deps, { comment_ref: "R1.1", find, replace, rationale: R }, turn);
  const turnDone = async (turn, outcome) => {
    const n = () => events.filter((e) => e.event === "turn_done").length;
    const before = n();
    svc.turnEnded("k", turn, outcome);
    for (let i = 0; i < 100 && n() === before; i++) await new Promise((r) => setTimeout(r, 5));
    assert.equal(n(), before + 1);
  };
  return {
    dir,
    svc,
    ask,
    lines,
    pushed,
    word,
    deps,
    edit,
    turnDone,
    load: () => loadRevision(join(dir, "paper.revision.json")),
    cleanup: () => rm(dir, { recursive: true, force: true }),
  };
}

// The first turn of the comment, and what a follow-through turn does to the legend.
const FIRST = ["was eliminated on all 14 spans", "was reduced on each of the 14 spans"];
const CAPTION = ["Rust coverage per span.", "Rust coverage on the 14 spans."];
const rewriteOf = (t, paragraph_id, text, turn) =>
  applyTrackedRewrite(
    { ...t.deps, paneCan: () => true },
    { comment_ref: "R1.1", paragraph_id, text, what: "w", why: "y", addresses: "a" },
    turn,
  );

test("review_propose with follow_through: the turn is given the other places of the comment", async () => {
  const t = await setup();
  try {
    // No change yet: no turn.
    const none = await t.ask("review_propose", {
      ref: "R1.1",
      paragraphs: [],
      follow_through: true,
    });
    assert.equal(none.ok, false);
    assert.match(none.error, /R1\.1 has no change in the document/);
    assert.equal(t.pushed.length, 0);

    const first = await t.edit(...FIRST, { ref: "R1.1" });
    assert.equal(first.ok, true, first.message);

    // Not with a rewrite in one turn.
    const both = await t.ask("review_propose", {
      ref: "R1.1",
      paragraphs: [],
      follow_through: true,
      rewrite_paragraph: 1,
    });
    assert.equal(both.ok, false);
    assert.match(both.error, /either writes a paragraph anew or checks the other places/);
    assert.equal(t.pushed.length, 0);

    const r = await t.ask("review_propose", { ref: "R1.1", paragraphs: [], follow_through: true });
    assert.equal(r.ok, true, r.error);
    const { text, turn } = t.pushed.at(-1);
    assert.match(
      text,
      /\n\n\[Other places of R1\.1\]\nThis turn does one thing\. The changes made for R1\.1 so far:\n- R1\.1-1: "was eliminated on all 14 spans" → "was reduced on each of the 14 spans"\nThe 3 places below/,
    );
    assert.ok(text.includes("[p3] (caption; legend of Fig. 2) Figure 2. Rust coverage per span."));
    assert.equal(turn.ref, "R1.1");
    assert.deepEqual(turn.follow.candidate_ids, ["p0", "p3", "p2"]);
    // The sources of numbers: what the changes of this comment put in and took out.
    assert.deepEqual(turn.follow.sources, [FIRST[1], FIRST[0]]);
    // A turn without the flag carries nothing of it.
    await t.ask("review_propose", { ref: "R1.1", paragraphs: [] });
    assert.deepEqual(t.pushed.at(-1).turn, { ref: "R1.1" });
    assert.ok(!t.pushed.at(-1).text.includes("[Other places of"));
  } finally {
    await t.cleanup();
  }
});

test("review_propose with follow_through: an answer stays seen after another turn and rejection", async () => {
  const t = await setup();
  try {
    const asked = await askAuthor(
      t.deps,
      {
        comment_ref: "R1.1",
        question: "On how many spans was rust left?",
        why: "The report gives no count.",
        paragraph_id: "p2",
      },
      { ref: "R1.1" },
    );
    assert.equal(asked.ok, true, asked.message);
    // Not answered yet: no turn.
    const open = await t.ask("review_propose", {
      ref: "R1.1",
      paragraphs: [],
      follow_through: true,
    });
    assert.match(open.error, /R1\.1 has no change in the document/);
    const a = await t.ask("review_question_answer", {
      ref: "R1.1",
      id: "a1",
      answer: "On 3 spans.",
    });
    assert.equal(a.ok, true, a.error);
    assert.equal(followState(a.revision.comments[0], a.revision.snapshot).offer, 1);

    const r = await t.ask("review_propose", { ref: "R1.1", paragraphs: [], follow_through: true });
    assert.equal(r.ok, true, r.error);
    const { text, turn } = t.pushed.at(-1);
    assert.ok(
      text.includes("- [R1.1 a1] (for [p2]) On how many spans was rust left? Answer: On 3 spans."),
    );
    assert.ok(
      text.includes(
        "\n\n[Other places of R1.1]\nThis turn does one thing. No change made for R1.1 is in the document.",
      ),
    );
    assert.ok(text.includes("[p2] (body; the author's answer to a1 goes here) The coating held"));
    assert.deepEqual(turn.follow.candidate_ids, ["p2"]);
    assert.deepEqual(turn.follow.sources, []);
    // Only that place may be changed, and the number of the answer may be written there.
    const other = await t.edit("Tension stayed", "Tension held", turn);
    assert.equal(other.code, "NotAFollowPlace");
    t.word.paragraph_index = 2;
    t.word.before = PARAGRAPHS[2];
    t.word.after = "On 3 spans the coating held through the winter.";
    const put = await t.edit("The coating held", "On 3 spans the coating held", turn);
    assert.equal(put.ok, true, put.message);
    await t.turnDone(turn, { ok: true, text: "Written in." });
    const c = (await t.load()).comments[0];
    assert.deepEqual(
      [
        c.follow_through.candidate_ids,
        c.follow_through.change_ids,
        c.follow_through.seen_answer_ids,
      ],
      [["p2"], ["R1.1-1"], ["a1"]],
    );
    const again = await t.ask("review_propose", {
      ref: "R1.1",
      paragraphs: [],
      follow_through: true,
    });
    assert.equal(again.ok, true, again.error);
    const next = t.pushed.at(-1).turn;
    assert.deepEqual(next.follow.answers, []);
    await t.turnDone(next, { ok: true, text: "No change needed." });

    // Word rejects the answer-driven change and refreshes the original text.
    const refreshed = await t.ask("review_refresh", {
      snapshot: {
        paragraphs: PARAGRAPHS.map((text, index) => ({
          index,
          id: `p${index}`,
          text,
          reviewed_text: text,
          w_ins: 0,
          w_del: 0,
        })),
        comments: [],
      },
    });
    assert.equal(refreshed.ok, true, refreshed.error);
    const rejected = await t.load();
    assert.equal(rejected.comments[0].changes[0].state, "rejected");
    assert.equal(followState(rejected.comments[0], rejected.snapshot).offer, null);
    assert.deepEqual(rejected.comments[0].follow_through.seen_answer_ids, ["a1"]);
  } finally {
    await t.cleanup();
  }
});

async function answerAt(t, answer, paragraph_id = "p2") {
  const asked = await askAuthor(
    t.deps,
    { comment_ref: "R1.1", question: `Detail ${answer}?`, why: "Missing detail.", paragraph_id },
    { ref: "R1.1" },
  );
  assert.equal(asked.ok, true, asked.message);
  const rev = await t.load();
  const id = rev.comments[0].questions.at(-1).id;
  const saved = await t.ask("review_question_answer", { ref: "R1.1", id, answer });
  assert.equal(saved.ok, true, saved.error);
  return id;
}

async function startFollow(t) {
  const reply = await t.ask("review_propose", {
    ref: "R1.1",
    paragraphs: [],
    follow_through: true,
  });
  assert.equal(reply.ok, true, reply.error);
  return t.pushed.at(-1).turn;
}

test("H1: two answers at one paragraph stay fresh, appear once and both become seen", async () => {
  const t = await setup();
  try {
    await answerAt(t, "One winter.");
    await t.turnDone(await startFollow(t), { ok: true, text: "No change needed." });
    await answerAt(t, "The coating held.");
    const rev = await t.load();
    assert.equal(followState(rev.comments[0], rev.snapshot).offer, 1);
    const plan = followPlan(rev, "R1.1");
    assert.deepEqual(
      plan.candidates.map((p) => p.id),
      ["p2"],
    );
    assert.deepEqual(plan.candidates[0].answer_ids, ["a1", "a2"]);
    const turn = await startFollow(t);
    assert.deepEqual(turn.follow.candidate_ids, ["p2"]);
    for (const id of ["a1", "a2"])
      assert.ok(
        t.pushed.at(-1).text.includes(`The author answered question ${id} of this comment`),
      );
    await t.turnDone(turn, { ok: true, text: "No change needed." });
    const done = await t.load();
    assert.deepEqual(done.comments[0].follow_through.seen_answer_ids, ["a1", "a2"]);
    assert.equal(followState(done.comments[0], done.snapshot).offer, null);
  } finally {
    await t.cleanup();
  }
});

for (const changed of [false, true]) {
  test(`H2: changed answer reoffers after ${changed ? "a live edit" : "no change"}; identical save does not`, async () => {
    const t = await setup();
    try {
      await answerAt(t, "One winter.");
      const first = await startFollow(t);
      if (changed) {
        t.word.paragraph_index = 2;
        t.word.before = PARAGRAPHS[2];
        t.word.after = "The coating held through one winter.";
        const edit = await t.edit("through the winter", "through one winter", first);
        assert.equal(edit.ok, true, edit.message);
      }
      await t.turnDone(first, { ok: true, text: "Checked." });
      const save = (answer) => t.ask("review_question_answer", { ref: "R1.1", id: "a1", answer });
      const same = await save("  One winter.  ");
      assert.equal(same.ok, true, same.error);
      assert.deepEqual(same.revision.comments[0].follow_through.seen_answer_ids, ["a1"]);
      assert.equal(followState(same.revision.comments[0], same.revision.snapshot).offer, null);
      const updated = await save("Two winters.");
      assert.equal(updated.ok, true, updated.error);
      assert.deepEqual(updated.revision.comments[0].follow_through.seen_answer_ids, []);
      assert.ok(followState(updated.revision.comments[0], updated.revision.snapshot).offer > 0);
      const next = await startFollow(t);
      assert.ok(next.follow.candidate_ids.includes("p2"));
      assert.ok(t.pushed.at(-1).text.includes("Answer: Two winters."));
      await t.turnDone(next, { ok: true, text: "Checked." });
      const done = await t.load();
      assert.deepEqual(done.comments[0].follow_through.seen_answer_ids, ["a1"]);
      assert.equal(followState(done.comments[0], done.snapshot).offer, null);
    } finally {
      await t.cleanup();
    }
  });
}

for (const [ok, live] of [
  [false, false],
  [true, false],
  [false, true],
  [true, true],
]) {
  test(`H2: ${ok ? "successful" : "failed"} turn preserves freshness of late answer updates${live ? " with a live edit" : ""}`, async () => {
    const t = await setup();
    try {
      await answerAt(t, "One winter.");
      await t.turnDone(await startFollow(t), { ok: true, text: "Checked." });
      await answerAt(t, "Coating held.");
      const turn = await startFollow(t);
      if (live) {
        t.word.paragraph_index = 2;
        t.word.before = PARAGRAPHS[2];
        t.word.after = "The coating endured through the winter.";
        const edit = await t.edit("The coating held", "The coating endured", turn);
        assert.equal(edit.ok, true, edit.message);
      }
      const saved = await t.ask("review_question_answer", {
        ref: "R1.1",
        id: "a1",
        answer: "Two winters.",
      });
      assert.equal(saved.ok, true, saved.error);
      await answerAt(t, "Inspection completed.");
      const before = await t.load();
      await t.turnDone(
        turn,
        ok ? { ok: true, text: "Checked old answers." } : { ok: false, error: "Interrupted." },
      );
      const after = await t.load();
      const c = after.comments[0];
      if (!ok) assert.deepEqual(c.follow_through, before.comments[0].follow_through);
      assert.deepEqual(c.follow_through.seen_answer_ids, ok ? ["a2"] : []);
      assert.ok(followState(c, after.snapshot).offer > 0);
      assert.ok(followPlan(after, "R1.1").candidates.some((p) => p.id === "p2"));
      await t.turnDone(await startFollow(t), { ok: true, text: "Checked current answers." });
      const done = await t.load();
      assert.deepEqual([...done.comments[0].follow_through.seen_answer_ids].sort(), [
        "a1",
        "a2",
        "a3",
      ]);
      assert.equal(followState(done.comments[0], done.snapshot).offer, null);
    } finally {
      await t.cleanup();
    }
  });
}

test("H2: identical trimmed save during a turn still becomes seen", async () => {
  const t = await setup();
  try {
    await answerAt(t, "One winter.");
    const turn = await startFollow(t);
    const saved = await t.ask("review_question_answer", {
      ref: "R1.1",
      id: "a1",
      answer: "  One winter.  ",
    });
    assert.equal(saved.ok, true, saved.error);
    await t.turnDone(turn, { ok: true, text: "Checked." });
    const done = await t.load();
    assert.deepEqual(done.comments[0].follow_through.seen_answer_ids, ["a1"]);
    assert.equal(followState(done.comments[0], done.snapshot).offer, null);
  } finally {
    await t.cleanup();
  }
});

test("H2: answer save invalidates only the changed question and preserves input history", () => {
  const rev = {
    comments: [
      {
        ref: "R1.1",
        questions: [
          { id: "a1", answer: "One winter." },
          { id: "a2", answer: "Coating held." },
        ],
        follow_through: {
          seen_answer_ids: ["a1", "a2"],
          candidate_ids: ["p2"],
          change_ids: [],
          questions: 0,
        },
      },
    ],
  };
  const same = answerQuestion(rev, "R1.1", "a1", " One winter. ");
  assert.deepEqual(same.comments[0].follow_through, rev.comments[0].follow_through);
  for (const text of ["Two winters.", ""]) {
    const edited = answerQuestion(rev, "R1.1", "a1", text);
    assert.deepEqual(edited.comments[0].follow_through, {
      ...rev.comments[0].follow_through,
      seen_answer_ids: ["a2"],
    });
    assert.deepEqual(rev.comments[0].follow_through.seen_answer_ids, ["a1", "a2"]);
    assert.equal(rev.comments[0].questions[0].answer, "One winter.");
  }
});

test("H2: a new answer reopens an already changed paragraph", async () => {
  const t = await setup();
  try {
    const edit = await t.edit(...FIRST, { ref: "R1.1" });
    assert.equal(edit.ok, true, edit.message);
    await t.turnDone(await startFollow(t), { ok: true, text: "Checked." });
    await answerAt(t, "The coating reduced rust.", "p1");
    const rev = await t.load();
    assert.ok(followState(rev.comments[0], rev.snapshot).offer > 0);
    const turn = await startFollow(t);
    assert.ok(turn.follow.candidate_ids.includes("p1"));
    await t.turnDone(turn, { ok: true, text: "Checked answer." });
    const done = await t.load();
    assert.deepEqual(done.comments[0].follow_through.seen_answer_ids, ["a1"]);
    assert.equal(followState(done.comments[0], done.snapshot).offer, null);
  } finally {
    await t.cleanup();
  }
});

test("number guard: a number of a change of this comment may go to its legend in a follow-through turn only", async () => {
  const t = await setup();
  try {
    assert.equal((await t.edit(...FIRST, { ref: "R1.1" })).ok, true);
    t.word.paragraph_index = 3;
    // An ordinary turn: 14 is not in the legend, the comment or an answer.
    const refused = await t.edit(...CAPTION, { ref: "R1.1" });
    assert.equal(refused.ok, false);
    assert.equal(refused.code, "NumberNeedsAuthor");
    assert.deepEqual(refused.numbers, ["14"]);

    await t.ask("review_propose", { ref: "R1.1", paragraphs: [], follow_through: true });
    const { turn } = t.pushed.at(-1);
    const ok = await t.edit(...CAPTION, turn);
    assert.equal(ok.ok, true, ok.message);
    // A number that only stands elsewhere in the changed paragraph (another
    // comment may have put it there) is refused as ever.
    t.word.paragraph_index = 0;
    const other = await t.edit("Deck survey", "Deck survey of 2031", turn);
    assert.equal(other.code, "NumberNeedsAuthor");
    // The same for a paragraph written anew (office_tracked_rewrite).
    t.word.paragraph_index = 2;
    const rewrite = (with_) =>
      rewriteOf(t, "p2", "The coating held on the 14 spans through the winter.", with_);
    assert.equal((await rewrite({ ref: "R1.1" })).code, "NumberNeedsAuthor");
    const anew = await rewrite(turn);
    assert.equal(anew.ok, true, anew.message);
  } finally {
    await t.cleanup();
  }
});

test("a follow-through turn changes the listed places only", async () => {
  const t = await setup();
  try {
    assert.equal((await t.edit(...FIRST, { ref: "R1.1" })).ok, true);
    await t.ask("review_propose", { ref: "R1.1", paragraphs: [], follow_through: true });
    const { turn } = t.pushed.at(-1);
    // A paragraph the comment changed, one that is no place, and words that
    // are in no paragraph: refused before Word is called.
    for (const [find, replace, at, says] of [
      ["after recoating", "after repainting", "1", /This paragraph is not one of them\./],
      ["within limits", "inside limits", "5", /This paragraph is not one of them\./],
      ["words the report does not have", "others", "?", /The words to change are in no paragraph/],
    ]) {
      const r = await t.edit(find, replace, turn);
      assert.equal(r.ok, false);
      assert.equal(r.code, "NotAFollowPlace", r.message);
      assert.match(r.message, says);
      assert.equal(t.lines.at(-1), `[review] not a follow place R1.1 paragraph=${at}`);
    }
    // Words that stand in several paragraphs ("Rust": the changed paragraph
    // and the legend; "s": twice in the legend). Each refusal says what to
    // give next, so the same call is not made again.
    const hinted = (find, replace, paragraph_hint) =>
      applyTrackedEdit(
        t.deps,
        { comment_ref: "R1.1", find, replace, paragraph_hint, rationale: R },
        turn,
      );
    for (const [find, hint, says] of [
      [
        "Rust",
        undefined,
        /occur in more than one place: give paragraph_hint with the id of a listed place/,
      ],
      ["Rust", "p99", /paragraph_hint p99 names no paragraph/],
      ["s", "p3", /occur more than once in p3: give more of the words around them/],
      ["Rust", "p1", /This paragraph is not one of them\./],
    ]) {
      const r = await hinted(find, "Corrosion", hint);
      assert.equal(r.code, "NotAFollowPlace", r.message);
      assert.match(r.message, says);
    }
    t.word.paragraph_index = 3;
    const byHint = await hinted("Rust", "Corrosion", "p3");
    assert.equal(byHint.ok, true, byHint.message);
    const anew = await rewriteOf(t, "p5", "Tension stayed inside limits.", turn);
    assert.equal(anew.code, "NotAFollowPlace", anew.message);
    assert.equal(t.lines.at(-1), "[review] not a follow place R1.1 paragraph=5");
    assert.equal((await t.load()).comments[0].changes.length, 2);
    // A listed place: both tools go through.
    t.word.paragraph_index = 2;
    const edit = await t.edit("held through", "lasted through", turn);
    assert.equal(edit.ok, true, edit.message);
    t.word.paragraph_index = 0;
    const head = await rewriteOf(t, "p0", "Deck inspection", turn);
    assert.equal(head.ok, true, head.message);
    // An ordinary turn of the comment is held to no list.
    t.word.paragraph_index = 5;
    const plain = await t.edit("within limits", "inside limits", { ref: "R1.1" });
    assert.equal(plain.ok, true, plain.message);
  } finally {
    await t.cleanup();
  }
});

test("a listed place written anew in more than three parts goes in whole", async () => {
  const LONG =
    "The coating held in spring. The drains ran clear in summer. The joints stayed dry in autumn. The bearings moved freely in winter.";
  const t = await setup(PARAGRAPHS.map((text, i) => (i === 2 ? LONG : text)));
  try {
    assert.equal((await t.edit(...FIRST, { ref: "R1.1" })).ok, true);
    await t.ask("review_propose", { ref: "R1.1", paragraphs: [], follow_through: true });
    const { turn } = t.pushed.at(-1);
    t.word.paragraph_index = 2;
    // The parts are not judged one by one against the list: the paragraph was.
    const r = await rewriteOf(
      t,
      "p2",
      "The coating lasted in spring. The drains ran free in summer. The joints stayed tight in autumn. The bearings moved well in winter.",
      turn,
    );
    assert.equal(r.ok, true, r.message);
    assert.ok(r.parts > 3 && r.change_ids.length === r.parts, JSON.stringify(r));
    await t.turnDone(turn, { ok: true, text: "The coating report now agrees." });
    const c = (await t.load()).comments[0];
    assert.deepEqual(c.follow_through.change_ids, r.change_ids);
    assert.deepEqual(c.follow_through.seen_change_ids, ["R1.1-1", ...r.change_ids]);
  } finally {
    await t.cleanup();
  }
});

test("an edit of a follow-through turn that a forbid rule blocked can be allowed afterwards", async () => {
  const t = await setup();
  try {
    await t.ask("ledger_add", {
      item: {
        kind: "forbid",
        text: "do not say novel",
        check: { type: "regex", pattern: String.raw`\bnovel\b`, flags: "i" },
      },
    });
    assert.equal((await t.edit(...FIRST, { ref: "R1.1" })).ok, true);
    await t.ask("review_propose", { ref: "R1.1", paragraphs: [], follow_through: true });
    const { turn } = t.pushed.at(-1);
    t.word.paragraph_index = 3;
    const blocked = await t.edit(CAPTION[0], "Novel rust coverage on the 14 spans.", turn);
    assert.equal(blocked.ok, false);
    assert.ok(blocked.blocked_id, blocked.message);
    // The blocked edit keeps what its turn was held to and allowed.
    const on = (await t.load()).comments[0].blocked.find((b) => b.id === blocked.blocked_id);
    assert.deepEqual(on.follow, {
      candidate_ids: ["p0", "p3", "p2"],
      sources: [FIRST[1], FIRST[0]],
    });
    await t.turnDone(turn, { ok: true });
    // "Allow exception": the number and the place are judged as in the turn.
    const r = await t.ask("review_override", { ref: "R1.1", blocked_id: blocked.blocked_id });
    assert.equal(r.ok, true, r.error);
    assert.equal(r.applied, true, JSON.stringify(r.result));
    // No text of the manuscript went to the log with it.
    assert.ok(!t.lines.some((l) => l.includes("14 spans")));

    // An item whose paragraph is not among the places on it stays refused.
    const again = await t.edit("The coating held", "The novel coating held", turn);
    assert.ok(again.blocked_id, again.message);
    const file = join(t.dir, "paper.revision.json");
    const rev = await t.load();
    rev.comments[0].blocked.find((b) => b.id === again.blocked_id).follow.candidate_ids = ["p0"];
    await saveRevision(file, rev);
    const no = await t.ask("review_override", { ref: "R1.1", blocked_id: again.blocked_id });
    assert.equal(no.applied, false);
    assert.equal(no.result.code, "NotAFollowPlace");
    assert.ok((await t.load()).comments[0].blocked.some((b) => b.id === again.blocked_id));
  } finally {
    await t.cleanup();
  }
});

for (const ownWork of [false, true]) {
  test(`follow-through owns only its tool receipts during a concurrent override (${ownWork ? "with" : "without"} own work)`, async () => {
    const t = await setup();
    try {
      assert.equal((await t.edit(...FIRST, { ref: "R1.1" })).ok, true);
      await t.ask("ledger_add", {
        item: {
          kind: "forbid",
          text: "Do not say novel",
          check: { type: "regex", pattern: "novel", flags: "i" },
        },
      });
      t.word.paragraph_index = 5;
      const blocked = await t.edit("Tension stayed", "Novel tension stayed", { ref: "R1.1" });
      assert.ok(blocked.blocked_id, blocked.message);
      const turn = await startFollow(t);
      // An exception first proposed by this follow-through turn has its
      // candidate/source constraints, but still belongs to the author's
      // separate approval request rather than the turn's tool receipts.
      const followBlocked = await t.edit("The coating held", "The novel coating held", turn);
      assert.ok(followBlocked.blocked_id, followBlocked.message);
      let owned = [];
      if (ownWork) {
        t.word.paragraph_index = 3;
        const changed = await t.edit(...CAPTION, turn);
        assert.equal(changed.ok, true, changed.message);
        owned = [changed.change_id];
        const asked = await askAuthor(
          t.deps,
          {
            comment_ref: "R1.1",
            question: "Was the deck heading checked?",
            why: "The library record is silent.",
          },
          turn,
        );
        assert.equal(asked.ok, true, asked.message);
      }
      // The author's exception belongs to a separate request even though
      // it saves a new change under the same comment while this turn runs.
      t.word.paragraph_index = 5;
      const override = await t.ask("review_override", {
        ref: "R1.1",
        blocked_id: blocked.blocked_id,
      });
      assert.equal(override.applied, true, JSON.stringify(override));
      const other = override.result.change_id;
      t.word.paragraph_index = 2;
      const followOverride = await t.ask("review_override", {
        ref: "R1.1",
        blocked_id: followBlocked.blocked_id,
      });
      assert.equal(followOverride.applied, true, JSON.stringify(followOverride));
      const askedElsewhere = await askAuthor(
        t.deps,
        {
          comment_ref: "R1.1",
          question: "Which library keeps the plans?",
          why: "The report does not say.",
        },
        { ref: "R1.1" },
      );
      assert.equal(askedElsewhere.ok, true, askedElsewhere.message);
      await t.turnDone(turn, { ok: true, text: "Checked the listed places." });
      const rev = await t.load();
      const c = rev.comments[0];
      assert.deepEqual(c.follow_through.change_ids, owned);
      assert.equal(c.follow_through.questions, ownWork ? 1 : 0);
      assert.deepEqual(c.follow_through.seen_change_ids, ["R1.1-1", ...owned]);
      assert.ok(!c.follow_through.seen_change_ids.includes(other));
      assert.ok(!c.follow_through.seen_change_ids.includes(followOverride.result.change_id));
      assert.ok(followState(c, rev.snapshot).offer > 0, "the unrelated edit still needs checking");
      assert.ok(
        t.lines.includes(
          `[review] follow R1.1 candidates=3 changes=${owned.length} questions=${ownWork ? 1 : 0}`,
        ),
      );
    } finally {
      await t.cleanup();
    }
  });
}

test("the end of a follow-through turn: one log line, and the comment says the check ran", async () => {
  const t = await setup();
  try {
    assert.equal((await t.edit(...FIRST, { ref: "R1.1" })).ok, true);
    // What the first turn said of the term is on file.
    await t.turnDone({ ref: "R1.1" }, { ok: true, text: "Done.\n\nSame term still at: p2" });
    assert.deepEqual((await t.load()).comments[0].same_term_at, ["p2"]);
    await t.ask("review_propose", { ref: "R1.1", paragraphs: [], follow_through: true });
    const { turn } = t.pushed.at(-1);

    // A turn that failed: logged, not marked, so it can be run again.
    await t.turnDone(turn, { ok: false, error: "Claude usage limit reached." });
    assert.ok(t.lines.includes("[review] follow R1.1 candidates=3 changes=0 questions=0"));
    assert.equal((await t.load()).comments[0].follow_through, undefined);

    t.word.paragraph_index = 3;
    assert.equal((await t.edit(...CAPTION, turn)).ok, true);
    const asked = await askAuthor(
      t.deps,
      {
        comment_ref: "R1.1",
        question: "Was the heading 'Deck survey' meant to cover the approach spans?",
        why: "The report does not say.",
      },
      turn,
    );
    assert.equal(asked.ok, true, asked.message);
    await t.turnDone(turn, { ok: true, text: "The legend now agrees." });
    assert.ok(t.lines.includes("[review] follow R1.1 candidates=3 changes=1 questions=1"));
    // Its answer has no line on the term: the list of the first turn stays.
    assert.deepEqual((await t.load()).comments[0].same_term_at, ["p2"]);
    const c = (await t.load()).comments[0];
    assert.deepEqual(
      [
        c.follow_through.candidate_ids,
        c.follow_through.change_ids,
        c.follow_through.questions,
        c.follow_through.seen_change_ids,
      ],
      [["p0", "p3", "p2"], ["R1.1-2"], 1, ["R1.1-1", "R1.1-2"]],
    );
    // No answer named a place: the record has no list of answers.
    assert.ok(!("seen_answer_ids" in c.follow_through));
    assert.match(c.follow_through.at, /^\d{4}-\d\d-\d\dT/);
    // An ordinary turn leaves the record as it is and logs no such line.
    const n = t.lines.filter((l) => l.startsWith("[review] follow ")).length;
    await t.turnDone({ ref: "R1.1" }, { ok: true });
    assert.equal(t.lines.filter((l) => l.startsWith("[review] follow ")).length, n);
    assert.deepEqual((await t.load()).comments[0].follow_through, c.follow_through);
  } finally {
    await t.cleanup();
  }
});

for (const [name, paragraphs, firstSize] of [
  ["paragraph limit", Array.from({ length: 21 }, (_, i) => `Span ${i} held firm.`), 20],
  ["character limit", ["a".repeat(25000) + ".", "b".repeat(25000) + "."], 1],
]) {
  test(`answer batches: no-change turns reach the remaining answer after the ${name}`, async () => {
    const t = await setup(paragraphs);
    try {
      for (let i = 0; i < paragraphs.length; i++) await answerAt(t, `Detail ${i}.`, `p${i}`);
      const first = await startFollow(t);
      assert.equal(first.follow.candidate_ids.length, firstSize);
      await t.turnDone(first, { ok: true, text: "No change needed." });
      const before = await t.load();
      assert.equal(followState(before.comments[0], before.snapshot).offer, 1);
      const second = await startFollow(t);
      assert.deepEqual(second.follow.candidate_ids, [`p${firstSize}`]);
      assert.deepEqual(
        second.follow.answers.map((a) => a.id),
        [`a${firstSize + 1}`],
      );
      await t.turnDone(second, { ok: true, text: "No change needed." });
      const done = await t.load();
      assert.equal(done.comments[0].follow_through.seen_answer_ids.length, paragraphs.length);
      assert.equal(followState(done.comments[0], done.snapshot).offer, null);
      assert.deepEqual(followPlan(done, "R1.1").candidates, []);
      assert.deepEqual(done.comments[0].changes, []);
      assert.deepEqual(done.comments[1], before.comments[1], "other comments keep their state");
      const repeat = await t.ask("review_propose", {
        ref: "R1.1",
        paragraphs: [],
        follow_through: true,
      });
      assert.equal(repeat.ok, false);
      assert.equal(t.pushed.length, 2, "finished answer-only checks do not enqueue another turn");
    } finally {
      await t.cleanup();
    }
  });
}

test("answer batches: changed and newly saved answers during the next batch remain pending", async () => {
  const paragraphs = Array.from({ length: 21 }, (_, i) => `Span ${i} held firm.`);
  const t = await setup(paragraphs);
  try {
    for (let i = 0; i < paragraphs.length; i++) await answerAt(t, `Detail ${i}.`, `p${i}`);
    await t.turnDone(await startFollow(t), { ok: true, text: "No change needed." });
    const second = await startFollow(t);
    assert.deepEqual(second.follow.candidate_ids, ["p20"]);
    const saved = await t.ask("review_question_answer", {
      ref: "R1.1",
      id: "a1",
      answer: "Revised detail.",
    });
    assert.equal(saved.ok, true, saved.error);
    await answerAt(t, "Another detail.", "p20");
    await t.turnDone(second, { ok: true, text: "Checked earlier answers." });
    const pending = await t.load();
    assert.deepEqual(
      pending.comments[0].follow_through.seen_answer_ids,
      Array.from({ length: 20 }, (_, i) => `a${i + 2}`),
    );
    assert.equal(followState(pending.comments[0], pending.snapshot).offer, 2);
    const last = await startFollow(t);
    assert.deepEqual(last.follow.candidate_ids, ["p0", "p20"]);
    assert.deepEqual(
      last.follow.answers.map((a) => a.id),
      ["a1", "a21", "a22"],
    );
    assert.ok(t.pushed.at(-1).text.includes("Answer: Revised detail."));
    await t.turnDone(last, { ok: true, text: "No change needed." });
    const done = await t.load();
    assert.equal(done.comments[0].follow_through.seen_answer_ids.length, 22);
    assert.equal(followState(done.comments[0], done.snapshot).offer, null);
    assert.deepEqual(done.comments[0].changes, []);
  } finally {
    await t.cleanup();
  }
});

test("answer batches: an oversized answer reaches the turn beside live-change context", async () => {
  const long = "The coating held. ".repeat(Math.ceil(MAX_CHARS / 18) + 1).trimEnd();
  assert.ok(long.length > MAX_CHARS);
  const t = await setup(PARAGRAPHS.map((text, i) => (i === 2 ? long : text)));
  try {
    assert.equal((await t.edit(...FIRST, { ref: "R1.1" })).ok, true);
    const id = await answerAt(t, "The coating was inspected.", "p2");
    const before = await t.load();
    const plan = followPlan(before, "R1.1");
    assert.deepEqual(
      plan.candidates.map((p) => p.id),
      ["p2"],
    );
    assert.equal(plan.candidates[0].text, long, "the oversized paragraph is not truncated");
    const turn = await startFollow(t);
    assert.deepEqual(turn.follow.candidate_ids, ["p2"]);
    assert.deepEqual(
      turn.follow.answers.map((a) => a.id),
      [id],
    );
    await t.turnDone(turn, { ok: true, text: "No change needed." });
    const done = await t.load();
    assert.deepEqual(done.comments[0].follow_through.seen_answer_ids, [id]);
    assert.equal(followState(done.comments[0], done.snapshot).offer, null);
    assert.deepEqual(done.comments[0].changes, before.comments[0].changes);
    assert.deepEqual(done.comments[1], before.comments[1]);
  } finally {
    await t.cleanup();
  }
});
