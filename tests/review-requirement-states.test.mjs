// Every request of a comment has one of four states when a turn is over:
// changed (a change names it), needs_author (a question with its id waits
// for an answer), outside_this_file (the model said it belongs to another
// file), or open. Open is the default: a turn that handles one request of
// three and says nothing of the others leaves two open on the card.
// All texts here are made up.
import { test } from "node:test";
import assert from "node:assert/strict";
import {
  emptyRevision,
  replaceComments,
  setRequirementOutside,
} from "../daemon/revision-store.mjs";
import {
  requirementOutside,
  requirementOutsideShape,
  REVIEW_NOTE_TOOLS,
} from "../daemon/review-tools.mjs";
import { buildReviewPrompt, REVIEW_RULES } from "../daemon/review-prompt.mjs";
import * as view from "../taskpane/shared/review-view.js";

const NOW = new Date("2026-01-02T03:04:05Z");
const REQS = [
  { id: "q1", text: "state the unit" },
  { id: "q2", text: "say which run time is right" },
  { id: "q3", text: "add the raw data as a supplementary table" },
  { id: "q4", text: "soften the claim in the abstract" },
];
const change = (ids, state = "pending") => ({
  id: "R1.1-1",
  state,
  rationale: { what: "w", why: "y", addresses: "a", requirement_ids: ids },
});
const question = (requirement_id, answer = "") => ({
  id: "a1",
  requirement_id,
  question: "Which?",
  why: "w",
  answer,
});

test("requirementStates: changed, needs_author, outside_this_file, and open for a request nobody spoke of", () => {
  const comment = {
    ref: "R1.1",
    requirements: REQS,
    changes: [change(["q1"])],
    questions: [question("q2")],
    requirements_outside: [{ id: "q3", why: "It is a file of its own.", at: "t" }],
  };
  assert.deepEqual(
    view.requirementStates(comment).map((r) => [r.id, r.state]),
    [
      ["q1", "changed"],
      ["q2", "needs_author"],
      ["q3", "outside_this_file"],
      ["q4", "open"],
    ],
  );
  assert.equal(view.requirementStates(comment)[2].why, "It is a file of its own.");
  // The judged case: three requests, one change, silence on the rest.
  const silent = { requirements: REQS.slice(0, 3), changes: [change(["q1"])] };
  assert.deepEqual(
    view.requirementStates(silent).map((r) => r.state),
    ["changed", "open", "open"],
  );
  // A rejected change answers nothing; an answered question waits for nobody.
  const later = {
    requirements: REQS.slice(0, 2),
    changes: [change(["q1"], "rejected")],
    questions: [question("q2", "3,020 hours.")],
  };
  assert.deepEqual(
    view.requirementStates(later).map((r) => r.state),
    ["open", "open"],
  );
  // A change wins over a note about the same request.
  const both = {
    requirements: REQS.slice(0, 1),
    changes: [change(["q1"])],
    questions: [question("q1")],
    requirements_outside: [{ id: "q1", why: "w", at: "t" }],
  };
  assert.equal(view.requirementStates(both)[0].state, "changed");
  assert.deepEqual(view.requirementStates({}), []);
});

test("the card line, the Done question and the waiting list name each state", () => {
  const comment = {
    ref: "R1.1",
    requirements: REQS,
    changes: [{ ...change(["q1"]), state: "accepted" }],
    blocked: [],
    questions: [question("q2")],
    requirements_outside: [{ id: "q3", why: "A file of its own.", at: "t" }],
  };
  // What each request is, in words, for the list "What the reviewer asks".
  assert.deepEqual(view.requirementStates(comment).map(view.requestStateText), [
    "Changed",
    "Needs your answer",
    "In another file: A file of its own.",
    "Not yet",
  ]);
  // Only an open request counts as unaddressed.
  assert.deepEqual(
    view.unaddressedRequirements(comment).map((r) => r.id),
    ["q4"],
  );
  assert.deepEqual(view.doneBlockers(comment, []), [
    '"soften the claim in the abstract" has no change',
    '"say which run time is right" waits for the author\'s answer',
  ]);
  const waiting = view.pendingDecisions({ comments: [comment], ledger: [] }).requirements;
  assert.deepEqual(
    waiting.map((w) => w.text),
    [
      "R1.1: Claude waits for your answer about \u201csay which run time is right\u201d.",
      "R1.1: nothing changes yet for \u201csoften the claim in the abstract\u201d.",
    ],
  );
  // The line after a turn.
  assert.equal(
    view.openRequestsText(comment),
    "1 request of R1.1 has no change yet: \u201csoften the claim in the abstract\u201d.",
  );
  assert.equal(
    view.openRequestsText({
      ref: "R1.2",
      requirements: REQS.slice(0, 3),
      changes: [change(["q1"])],
    }),
    "2 requests of R1.2 have no change yet: \u201csay which run time is right\u201d; \u201cadd the raw data as a supplementary table\u201d.",
  );
  assert.equal(
    view.openRequestsText({
      ref: "R1.3",
      requirements: REQS.slice(0, 1),
      changes: [change(["q1"])],
    }),
    "",
  );
});

test("store: a request noted as outside this file; one note per request; a comment-list save keeps it", () => {
  let rev = replaceComments(emptyRevision(), [
    { ref: "R1.1", reviewer: 1, text: "t", requirements: ["state the unit", "add the raw data"] },
  ]);
  rev = setRequirementOutside(rev, "R1.1", { id: "q2", why: "Supplementary file." }, NOW);
  assert.deepEqual(rev.comments[0].requirements_outside, [
    { id: "q2", why: "Supplementary file.", at: NOW.toISOString() },
  ]);
  rev = setRequirementOutside(rev, "R1.1", { id: "q2", why: "Table S3." }, NOW);
  assert.deepEqual(
    rev.comments[0].requirements_outside.map((n) => n.why),
    ["Table S3."],
  );
  assert.throws(
    () => setRequirementOutside(rev, "R1.1", { id: "q9", why: "x" }),
    /No requirement q9 under R1\.1/,
  );
  const saved = replaceComments(rev, [{ ref: "R1.1", reviewer: 1, text: "Edited." }]);
  assert.equal(saved.comments[0].requirements_outside.length, 1);
});

test("office_requirement_outside: a listed request of the turn's comment; the log holds no text", async () => {
  const d = {
    noted: [],
    lines: [],
    log: (s) => d.lines.push(s),
    getReviewContext: async () => ({ requirements: REQS }),
    onRequirementOutside: async ({ ref, note }) => d.noted.push([ref, note]),
  };
  const args = { comment_ref: "R1.1", requirement_id: "q3", why: "It is a supplementary file." };
  assert.match((await requirementOutside(d, args, null)).message, /only in a review turn/);
  assert.match(
    (await requirementOutside(d, { ...args, requirement_id: "q9" }, { ref: "R1.1" })).message,
    /Unknown requirement id q9 for R1\.1\. Valid ids: q1, q2, q3, q4\./,
  );
  const r = await requirementOutside(d, args, { ref: "R1.1" });
  assert.equal(r.ok, true);
  assert.match(r.message, /q3 is shown to the author as outside this file/);
  assert.deepEqual(d.noted, [["R1.1", { id: "q3", why: "It is a supplementary file." }]]);
  assert.deepEqual(d.lines, ["[review] requirement outside R1.1 q3"]);
  assert.ok(REVIEW_NOTE_TOOLS.has("office_requirement_outside"));
  assert.deepEqual(Object.keys(requirementOutsideShape), ["comment_ref", "requirement_id", "why"]);
});

test("prompt: each listed request must end as a change, a question or an outside note; without a list the turn makes one", () => {
  const comment = { ref: "R1.1", reviewer: 1, text: "Several things.", requirements: REQS };
  const { text } = buildReviewPrompt({ comment, fence: "F1" });
  assert.ok(
    text.includes(
      "Every request above ends this turn in one of three ways: a change that names its id, office_ask_author with its id (it waits for the author), or office_requirement_outside with its id (what it asks for is in another file). A request with none of these is shown to the author as open.",
    ),
  );
  const bare = buildReviewPrompt({ comment: { ...comment, requirements: [] }, fence: "F1" }).text;
  assert.ok(
    bare.includes(
      "No list of requests was extracted for this comment: list what the comment asks for yourself first, then handle each one.",
    ),
  );
  assert.equal(bare.includes("Every request above ends this turn"), false);
  const rule = REVIEW_RULES.find((r) => r.includes("office_requirement_outside"));
  assert.match(rule, /^Do not pass over a request in silence\./);
  assert.ok(text.includes(rule));
});
