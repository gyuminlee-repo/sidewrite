// What only the author can know is asked, not guessed: office_ask_author
// keeps a question on the comment, the author answers it on the card, and
// the answer goes into the next turn like an author instruction (and is a
// source of numbers for the number guard). All texts here are made up.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  emptyRevision,
  replaceComments,
  addQuestion,
  answerQuestion,
  loadRevision,
} from "../daemon/revision-store.mjs";
import {
  askAuthor,
  askAuthorShape,
  applyTrackedEdit,
  REVIEW_NOTE_TOOLS,
  reviewTurnToolAllowed,
} from "../daemon/review-tools.mjs";
import { buildReviewPrompt, REVIEW_RULES } from "../daemon/review-prompt.mjs";
import { createReviewService } from "../daemon/review-service.mjs";

const COMMENTS = [
  {
    ref: "R1.1",
    reviewer: 1,
    text: "The run time in the Methods and in the Results do not agree.",
    requirements: ["make the run time consistent"],
  },
  { ref: "R1.2", reviewer: 1, text: "State the unit." },
];
const NOW = new Date("2026-01-02T03:04:05Z");

test("store: questions get ids a1, a2; an answer is kept with its time; a comment-list save keeps them", () => {
  let rev = replaceComments(emptyRevision(), COMMENTS);
  let r = addQuestion(
    rev,
    "R1.1",
    { requirement_id: "q1", question: "Which run time is right?", why: "Two values." },
    NOW,
  );
  assert.deepEqual(r.question, {
    id: "a1",
    requirement_id: "q1",
    question: "Which run time is right?",
    why: "Two values.",
    asked_at: NOW.toISOString(),
    answer: "",
    answered_at: null,
  });
  r = addQuestion(r.revision, "R1.1", { question: "How was growth followed?", why: "Not said." });
  assert.equal(r.question.id, "a2");
  assert.equal(r.question.requirement_id, null);
  // The same open question again (a later turn asks it once more): no second row.
  const again = addQuestion(r.revision, "R1.1", {
    question: "How was growth followed?",
    why: "Still not said.",
  });
  assert.equal(again.question.id, "a2");
  assert.equal(again.revision.comments[0].questions.length, 2);

  rev = answerQuestion(r.revision, "R1.1", "a1", "  3,020 hours.  ", NOW);
  const q = rev.comments[0].questions[0];
  assert.deepEqual([q.answer, q.answered_at], ["3,020 hours.", NOW.toISOString()]);
  // An emptied answer is no answer.
  const cleared = answerQuestion(rev, "R1.1", "a1", " ", NOW).comments[0].questions[0];
  assert.deepEqual([cleared.answer, cleared.answered_at], ["", null]);
  assert.throws(() => answerQuestion(rev, "R1.1", "a9", "x"), /No question a9 under R1\.1/);
  assert.throws(() => addQuestion(rev, "R9.9", { question: "x", why: "y" }), /No reviewer comment/);

  const saved = replaceComments(rev, [{ ref: "R1.1", reviewer: 1, text: "Edited text." }]);
  assert.equal(saved.comments[0].questions.length, 2);
  assert.equal(saved.comments[0].questions[0].answer, "3,020 hours.");
  // The next id goes on after a save.
  assert.equal(addQuestion(saved, "R1.1", { question: "Another?", why: "y" }).question.id, "a3");
});

function toolDeps(requirements = [{ id: "q1", text: "make the run time consistent" }]) {
  const d = {
    asked: [],
    lines: [],
    log: (s) => d.lines.push(s),
    getReviewContext: async () => ({ requirements }),
    onQuestion: async ({ ref, question }) => {
      d.asked.push([ref, question]);
      return { id: "a1", ...question };
    },
  };
  return d;
}

test("office_ask_author: only in the turn of its comment, with a known request id; the log holds no text", async () => {
  const ask = {
    comment_ref: "R1.1",
    requirement_id: "q1",
    question: "Which run time is right: 3,360 h or 3,020 h?",
    why: "The Methods and the Results give different values.",
  };
  let d = toolDeps();
  assert.match((await askAuthor(d, ask, null)).message, /only in a review turn/);
  assert.match(
    (await askAuthor(d, ask, { ref: "R1.2" })).message,
    /comment_ref R1\.1 does not match the comment of this turn R1\.2/,
  );
  const bad = await askAuthor(d, { ...ask, requirement_id: "q7" }, { ref: "R1.1" });
  assert.equal(bad.ok, false);
  assert.match(bad.message, /Unknown requirement id q7 for R1\.1\. Valid ids: q1\./);
  assert.deepEqual(d.asked, []);

  const r = await askAuthor(d, ask, { ref: "R1.1" });
  assert.deepEqual([r.ok, r.question_id], [true, "a1"]);
  assert.match(r.message, /Leave the text it is about as it is in this turn/);
  assert.deepEqual(d.asked, [
    ["R1.1", { requirement_id: "q1", question: ask.question, why: ask.why }],
  ]);
  assert.deepEqual(d.lines, ["[review] question R1.1 a1 requirement=q1"]);
  // Without a request id.
  d = toolDeps([]);
  const { requirement_id, ...plain } = ask;
  assert.equal((await askAuthor(d, plain, { ref: "R1.1" })).ok, true);
  assert.equal(d.asked[0][1].requirement_id, null);
  assert.deepEqual(d.lines, ["[review] question R1.1 a1 requirement=-"]);
  assert.deepEqual(Object.keys(askAuthorShape), [
    "comment_ref",
    "requirement_id",
    "question",
    "why",
  ]);
});

test("the question tool is a review-turn tool that writes no document", () => {
  assert.ok(REVIEW_NOTE_TOOLS.has("office_ask_author"));
  assert.equal(reviewTurnToolAllowed("office_ask_author"), true);
});

test("prompt: author answers go in like instructions; an open question is not asked again; the rule says when to ask", () => {
  const comment = {
    ref: "R1.1",
    reviewer: 1,
    text: "The run time does not agree.",
    questions: [
      {
        id: "a1",
        requirement_id: "q1",
        question: "Which run time?",
        why: "w",
        answer: "3,020 hours.",
      },
      {
        id: "a2",
        requirement_id: null,
        question: "How was growth followed?",
        why: "w",
        answer: "",
      },
    ],
  };
  const other = {
    ref: "R2.4",
    reviewer: 2,
    text: "n?",
    questions: [
      { id: "a1", requirement_id: null, question: "How many flasks?", why: "w", answer: "12" },
    ],
  };
  const { text } = buildReviewPrompt({ comment, others: [comment, other], fence: "F1" });
  assert.ok(
    text.includes(
      [
        "Author answers (what the author said when asked; they hold like author instructions, and a number in them may be written):",
        "- [R1.1 a1] Which run time? Answer: 3,020 hours.",
        "- [R2.4 a1] How many flasks? Answer: 12",
        "",
        "Asked for this comment and not answered yet (do not ask again; leave the text they are about as it is):",
        "- [a2] How was growth followed?",
      ].join("\n"),
    ),
    text,
  );
  // The block stands before the rules, after the instructions.
  assert.ok(text.indexOf("Author instructions:") < text.indexOf("Author answers ("));
  assert.ok(text.indexOf("Author answers (") < text.indexOf("\nRules:"));
  // No question on record: no block.
  const bare = buildReviewPrompt({ comment: { ...comment, questions: [] }, fence: "F1" }).text;
  assert.equal(bare.includes("Author answers ("), false);
  assert.equal(bare.includes("not answered yet"), false);
  // The rule.
  const rule = REVIEW_RULES.find((r) => r.includes("office_ask_author"));
  assert.match(rule, /^What only the author can know is asked, not guessed:/);
  assert.match(rule, /a number that differs between two places/);
  assert.match(rule, /how something was measured/);
  assert.match(rule, /a statistic that would have to be computed again/);
  assert.match(
    rule,
    /An edit that puts a number into the manuscript that nothing gives is refused/,
  );
  assert.ok(bare.includes(rule));
});

async function service() {
  const dir = await mkdtemp(join(tmpdir(), "review-q-"));
  const doc = join(dir, "paper.docx");
  const events = [];
  const lines = [];
  const svc = createReviewService({
    getContext: () => ({ activeDoc: doc, caps: ["target_hash"] }),
    callTool: async () => ({
      ok: true,
      word_comment_id: 5,
      paragraph_index: 1,
      before: "The culture ran for 3,360 hours.",
      after: "The culture ran for 3,020 hours.",
      target_checked: true,
    }),
    pushTurn: () => {},
    sendEvent: (key, obj) => events.push(obj),
    log: (s) => lines.push(s),
    env: { CC_OFFICE_WSL_PATHS: "0" },
  });
  const ask = (type, payload = {}) =>
    new Promise((resolve) =>
      svc.handlers[type]({ type, request_id: "q", ...payload }, resolve, "k", "word"),
    );
  return {
    dir,
    doc,
    svc,
    ask,
    events,
    lines,
    cleanup: () => rm(dir, { recursive: true, force: true }),
  };
}

test("service: a question is saved and the pane told; the answer lifts the number guard and is in the next prompt", async () => {
  const t = await service();
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    const para = (text, index) => ({
      index,
      id: `p${index}`,
      text,
      reviewed_text: text,
      w_ins: 0,
      w_del: 0,
    });
    await t.ask("review_refresh", {
      snapshot: {
        paragraphs: ["Methods", "The culture ran for 3,360 hours."].map(para),
        comments: [],
      },
    });
    const turn = { ref: "R1.1" };
    const deps = { ...t.svc.trackedEditDeps("k"), call: async () => ({ ok: true }) };
    const edit = {
      comment_ref: "R1.1",
      find: "3,360 hours",
      replace: "3,020 hours",
      rationale: { what: "w", why: "y", addresses: "a" },
    };
    // The model computed another value: refused.
    assert.equal((await applyTrackedEdit(deps, edit, turn)).code, "NumberNeedsAuthor");
    // It asks instead.
    const asked = await askAuthor(
      deps,
      {
        comment_ref: "R1.1",
        requirement_id: "q1",
        question: "Which run time is right?",
        why: "Two values.",
      },
      turn,
    );
    assert.deepEqual([asked.ok, asked.question_id], [true, "a1"]);
    assert.deepEqual(t.events.at(-1), {
      type: "review_event",
      event: "comment_noted",
      ref: "R1.1",
    });
    let rev = await loadRevision(join(t.dir, "paper.revision.json"));
    assert.equal(rev.comments[0].questions[0].question, "Which run time is right?");
    assert.equal(
      t.lines.some((l) => l.includes("Which run time")),
      false,
      "no question text in the log",
    );

    // The author answers on the card.
    const bad = await t.ask("review_question_answer", { ref: "R1.1", id: "a7", answer: "x" });
    assert.equal(bad.ok, false);
    const r = await t.ask("review_question_answer", {
      ref: "R1.1",
      id: "a1",
      answer: "3,020 hours; the other value was a typing error.",
    });
    assert.equal(r.ok, true, r.error);
    assert.equal(
      r.revision.comments[0].questions[0].answer,
      "3,020 hours; the other value was a typing error.",
    );
    const ctx = await t.svc.trackedEditDeps("k").getReviewContext("R1.1");
    assert.deepEqual(ctx.author_texts, ["3,020 hours; the other value was a typing error."]);
    // The next turn has the answer, and its number may be written.
    const p = await t.ask("review_preview", { ref: "R1.1", paragraphs: [] });
    assert.ok(
      p.prompt.includes(
        "- [R1.1 a1] Which run time is right? Answer: 3,020 hours; the other value was a typing error.",
      ),
    );
    const applied = {
      ok: true,
      word_comment_id: 5,
      paragraph_index: 1,
      before: "The culture ran for 3,360 hours.",
      after: "The culture ran for 3,020 hours.",
      target_checked: true,
    };
    const ok = await applyTrackedEdit(
      { ...t.svc.trackedEditDeps("k"), call: async () => applied },
      edit,
      turn,
    );
    assert.equal(ok.ok, true, ok.message);
  } finally {
    await t.cleanup();
  }
});
