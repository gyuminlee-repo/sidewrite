// The words of the list of comments (review-view.js): the state of a comment,
// the header, the line the paste box shrinks to, the first line of a comment.
import { test } from "node:test";
import assert from "node:assert/strict";
import {
  clockText,
  commentState,
  firstLine,
  reviewLine,
  summaryText,
} from "../taskpane/shared/review-view.js";

const ch = (state) => ({ id: "x", state });
const base = { status: "in_progress", changes: [], blocked: [], questions: [], requirements: [] };

test("commentState: each state in words, from what the review file holds", () => {
  const now = Date.parse("2026-10-02T00:01:05Z");
  const working = { startedAt: "2026-10-02T00:00:00Z" };
  const cases = [
    [{ ...base }, {}, "not_started", "Not started", false],
    [{ ...base, status: "done" }, {}, "done", "Done", false],
    [{ ...base }, { working }, "working", "Working 1:05", false],
    // The work shown wins over done: a turn is running.
    [{ ...base, status: "done" }, { working }, "working", "Working 1:05", false],
    [{ ...base, changes: [ch("pending")] }, {}, "to_review", "1 change to review in Word", true],
    [
      { ...base, changes: [ch("pending"), ch("unlocated"), ch("accepted")] },
      {},
      "to_review",
      "2 changes to review in Word",
      true,
    ],
    [
      { ...base, changes: [ch("accepted"), ch("rejected")] },
      {},
      "decided",
      "All changes decided",
      false,
    ],
    [
      { ...base, questions: [{ id: "a1", answer: "" }] },
      {},
      "needs_answer",
      "Needs your answer",
      true,
    ],
    // An answered question asks nothing more.
    [
      { ...base, questions: [{ id: "a1", answer: "yes" }] },
      {},
      "not_started",
      "Not started",
      false,
    ],
    [
      { ...base, blocked: [{ id: "b", state: "blocked", blocked_by: ["L1"] }] },
      {},
      "needs_answer",
      "Needs your decision",
      true,
    ],
    // An edit that is with Word right now waits for nobody.
    [
      { ...base, blocked: [{ id: "b", state: "applying", blocked_by: [] }] },
      {},
      "not_started",
      "Not started",
      false,
    ],
    [{ ...base }, { noChange: true }, "no_change", "No change proposed", true],
    [
      {
        ...base,
        requirements: [{ id: "q1", text: "t" }],
        requirements_outside: [{ id: "q1", why: "w" }],
      },
      {},
      "no_change",
      "Belongs to another file",
      false,
    ],
  ];
  for (const [c, opts, kind, label, attention] of cases) {
    assert.deepEqual(
      commentState(c, { now, ...opts }),
      { kind, label, attention },
      `${kind}: ${label}`,
    );
  }
});

test("summaryText: how many comments, how many need the author", () => {
  const s = (kind, attention = false) => ({ kind, attention });
  assert.equal(summaryText([], []), "");
  assert.equal(summaryText([{}], [s("not_started")]), "1 comment. None needs you right now.");
  assert.equal(
    summaryText([{}, {}, {}], [s("to_review", true), s("done"), s("needs_answer", true)]),
    "3 comments, 1 done. 2 need you.",
  );
  assert.equal(
    summaryText([{}, {}], [s("to_review", true), s("done")]),
    "2 comments, 1 done. 1 needs you.",
  );
});

test("reviewLine: the line the paste box shrinks to", () => {
  const refs = (...r) => r.map((ref) => ({ ref }));
  assert.equal(
    reviewLine(refs("R1.1", "R1.2", "R1.3", "R1.4", "R1.5")),
    "Review: 5 comments from 1 reviewer",
  );
  assert.equal(reviewLine(refs("R1.1", "R2.1")), "Review: 2 comments from 2 reviewers");
  assert.equal(
    reviewLine(refs("E.1", "R1.1")),
    "Review: 2 comments from 1 reviewer and the editor",
  );
  assert.equal(reviewLine(refs("R1.1")), "Review: 1 comment from 1 reviewer");
});

test("firstLine: the first line, cut on a word end", () => {
  assert.equal(firstLine("One line.\nSecond."), "One line.");
  assert.equal(firstLine("\n\n  First real line \nx"), "First real line");
  assert.equal(firstLine(""), "");
  const long = "word ".repeat(40).trim();
  const cut = firstLine(long, 30);
  assert.ok(cut.endsWith("...") && cut.length <= 33, cut);
  assert.ok(!/ \.\.\.$/.test(cut));
});

test("clockText: m:ss from a start time", () => {
  assert.equal(clockText(0, 65_000), "1:05");
  assert.equal(clockText("2026-10-02T00:00:00Z", Date.parse("2026-10-02T00:10:09Z")), "10:09");
  assert.equal(clockText(5000, 1000), "0:00", "never negative");
});
