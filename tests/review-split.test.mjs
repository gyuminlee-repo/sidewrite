import { test } from "node:test";
import assert from "node:assert/strict";
import {
  splitReviewRegex,
  splitReview,
  splitPrompt,
  requirementsFrom,
  MAX_REQUIREMENTS,
} from "../daemon/review-split.mjs";

const TWO_REVIEWERS = `Dear Editor,

Reviewer #1: Comments to the Author
Comment 1: The temperature unit is missing.
Please use degrees Celsius.
Comment 2: The extract name is misleading.
  1. It is not cell-free.
  2. Call it a clarified lysate.
Reviewer 2 raised a similar point, which I support.

Reviewer #2:
1. Section 3 should say how growth was measured.
2) Q10 values were not reported.
3.5 mM is an unusual concentration here.
`;

// "Dear Editor," before the first reviewer head is kept as E.1.
test("reviewer and comment heads give R{reviewer}.{n} refs", () => {
  const c = splitReviewRegex(TWO_REVIEWERS);
  assert.deepEqual(
    c.map((x) => x.ref),
    ["E.1", "R1.1", "R1.2", "R2.1", "R2.2"],
  );
  assert.equal(c[0].text, "Dear Editor,");
  assert.equal(c[1].text, "The temperature unit is missing.\nPlease use degrees Celsius.");
});

test("numbered list inside a Comment-head block stays body text", () => {
  const c = splitReviewRegex(TWO_REVIEWERS);
  assert.match(c[2].text, /1\. It is not cell-free\.\n {2}2\. Call it a clarified lysate\./);
});

test("trap: 'Reviewer 2 raised ...' in a body is not a reviewer head", () => {
  const c = splitReviewRegex(TWO_REVIEWERS);
  assert.match(c[2].text, /Reviewer 2 raised a similar point/);
});

test("trap: 'Q10 values' and '3.5 mM' do not start comments", () => {
  const c = splitReviewRegex(TWO_REVIEWERS);
  assert.equal(c[4].ref, "R2.2");
  assert.match(
    c[4].text,
    /^Q10 values were not reported\.\n3\.5 mM is an unusual concentration here\.$/,
  );
});

test("Q heads with punctuation and Point heads are recognized", () => {
  const c = splitReviewRegex("Q1. First?\nQ2: Second?\nPoint 3 Third.");
  assert.deepEqual(
    c.map((x) => [x.ref, x.text]),
    [
      ["R1.1", "First?"],
      ["R1.2", "Second?"],
      ["R1.3", "Third."],
    ],
  );
});

test("no reviewer head means reviewer 1; preamble before the first comment is dropped", () => {
  const c = splitReviewRegex("Thanks for the paper.\n1. Add units.\n2. Fix Fig. 2.");
  assert.deepEqual(
    c.map((x) => x.ref),
    ["R1.1", "R1.2"],
  );
});

test("custom patterns from a profile replace the defaults", () => {
  const c = splitReviewRegex("Referee 3\nRemark 1 a\nRemark 2 b", {
    reviewerHead: String.raw`^Referee\s+(\d+)$`,
    commentHead: String.raw`^Remark\s+(\d+)\b`,
  });
  assert.deepEqual(
    c.map((x) => x.ref),
    ["R3.1", "R3.2"],
  );
});

test("splitReview uses regex when it finds more than one comment (no LLM call)", async () => {
  let called = 0;
  const r = await splitReview(TWO_REVIEWERS, {
    oneshot: async () => (called++, { ok: true, json: {} }),
  });
  assert.equal(r.method, "regex");
  assert.equal(called, 0);
});

const PROSE =
  "The paper is clear. However the temperature unit is missing. The extract name is also misleading.";

test("splitReview falls back to one LLM call and keeps only verbatim items", async () => {
  let req;
  const r = await splitReview(PROSE, {
    oneshot: async (q) => {
      req = q;
      return {
        ok: true,
        json: {
          comments: [
            { reviewer: 1, number: 1, text: "However the temperature unit is missing." },
            { reviewer: 1, number: 2, text: "The extract   name is also misleading." },
            { reviewer: 1, number: 3, text: "The statistics are wrong." },
          ],
        },
      };
    },
  });
  assert.equal(r.method, "llm");
  assert.deepEqual(
    r.comments.map((c) => c.ref),
    ["R1.1", "R1.2"],
  );
  assert.match(r.warnings[0], /dropped LLM item 3/);
  assert.equal(req.prompt, splitPrompt(PROSE));
});

test("splitReview: LLM failure returns the regex result with a warning", async () => {
  const r = await splitReview(PROSE, { oneshot: async () => ({ ok: false, error: "timed out" }) });
  assert.equal(r.method, "llm");
  assert.deepEqual(r.comments, []);
  assert.match(r.warnings[0], /timed out/);
});

test("splitReview flags duplicate refs", async () => {
  const r = await splitReview("Comment 1: a\nComment 1: b");
  assert.match(r.warnings[0], /duplicate ref R1\.1/);
});

test("L5 regex split leaves requirements empty for the user", () => {
  for (const c of splitReviewRegex(TWO_REVIEWERS)) assert.deepEqual(c.requirements, []);
});

test("L5 LLM split returns the requirements of each kept comment with q-ids", async () => {
  const r = await splitReview(PROSE, {
    oneshot: async () => ({
      ok: true,
      json: {
        comments: [
          {
            reviewer: 1,
            number: 1,
            text: "However the temperature unit is missing.",
            requirements: ["state the unit", "  state   the unit ", 7, "use degrees Celsius"],
          },
          { reviewer: 1, number: 2, text: "The extract name is also misleading." },
        ],
      },
    }),
  });
  assert.deepEqual(r.comments[0].requirements, [
    { id: "q1", text: "state the unit" },
    { id: "q2", text: "use degrees Celsius" },
  ]);
  assert.deepEqual(r.comments[1].requirements, []);
  assert.match(splitPrompt(PROSE), /"requirements":\[/);
});

test("L5 requirementsFrom caps count and length", () => {
  const many = Array.from({ length: 15 }, (_, i) => `ask ${i} ${"x".repeat(400)}`);
  const r = requirementsFrom(many);
  assert.equal(r.length, MAX_REQUIREMENTS);
  assert.equal(r[0].text.length, 300);
  assert.deepEqual(requirementsFrom("not a list"), []);
});

// A reviewer block with no numbered heads is kept whole as R{n}.1 with a
// warning, instead of being dropped because another reviewer numbered theirs.
const MIXED = `Reviewer 1:
Please report the growth temperature for every culture.

Reviewer 2:
1. Add error bars to Fig. 2.
2. State the number of replicates.
`;

test("a reviewer block without numbered comments is kept whole, with a warning", async () => {
  let called = 0;
  const r = await splitReview(MIXED, {
    oneshot: async () => (called++, { ok: true, json: {} }),
  });
  assert.equal(r.method, "regex");
  assert.equal(called, 0);
  assert.deepEqual(
    r.comments.map((c) => [c.ref, c.text]),
    [
      ["R1.1", "Please report the growth temperature for every culture."],
      ["R2.1", "Add error bars to Fig. 2."],
      ["R2.2", "State the number of replicates."],
    ],
  );
  assert.equal(r.warnings.length, 1);
  assert.match(r.warnings[0], /Reviewer 1.*not split.*R1\.1/);
  assert.deepEqual(splitReviewRegex(MIXED).length, 3);
});

test("an unsplit block does not count toward skipping the LLM split", async () => {
  const one = "Reviewer 1:\nPlease report the temperature.\n\nReviewer 2:\n1. Add error bars.\n";
  let called = 0;
  const r = await splitReview(one, {
    oneshot: async () => (called++, { ok: false, error: "timed out" }),
  });
  assert.equal(called, 1);
  assert.deepEqual(
    r.comments.map((c) => c.ref),
    ["R1.1", "R2.1"],
  );
  assert.ok(r.warnings.some((w) => /R1\.1/.test(w)));
  assert.ok(r.warnings.some((w) => /timed out/.test(w)));
});

test("an LLM split that leaves out a kept-whole reviewer says so", async () => {
  const one = "Reviewer 1:\nPlease report the temperature.\n\nReviewer 2:\n1. Add error bars.\n";
  const r = await splitReview(one, {
    oneshot: async () => ({
      ok: true,
      json: { comments: [{ reviewer: 2, number: 1, text: "Add error bars." }] },
    }),
  });
  assert.equal(r.method, "llm");
  assert.deepEqual(
    r.comments.map((c) => [c.ref, c.text]),
    [
      ["R1.1", "Please report the temperature."],
      ["R2.1", "Add error bars."],
    ],
  );
  assert.ok(r.warnings.some((w) => /Reviewer 1.*no comment.*R1\.1/.test(w)));
});

// Text before the first reviewer head (the editor letter) is kept as E.1 and
// a reviewer preamble before the first numbered comment as R{n}.0, each with
// a warning, instead of being dropped.
const WITH_LETTER = `Dear Dr. Lee,
Please shorten the abstract to 200 words.

Reviewer 1:
The study is useful but the methods need detail.
1. Report the growth temperature.
2. Add error bars to Fig. 2.

Reviewer 2:
1. State the number of replicates.
`;

test("editor letter and reviewer preamble are kept as E.1 and R1.0 with warnings", async () => {
  let called = 0;
  const r = await splitReview(WITH_LETTER, {
    oneshot: async () => (called++, { ok: true, json: {} }),
  });
  assert.equal(r.method, "regex");
  assert.equal(called, 0);
  assert.deepEqual(
    r.comments.map((c) => [c.ref, c.reviewer, c.text]),
    [
      ["E.1", 0, "Dear Dr. Lee,\nPlease shorten the abstract to 200 words."],
      ["R1.0", 1, "The study is useful but the methods need detail."],
      ["R1.1", 1, "Report the growth temperature."],
      ["R1.2", 1, "Add error bars to Fig. 2."],
      ["R2.1", 2, "State the number of replicates."],
    ],
  );
  assert.equal(r.warnings.length, 2);
  assert.match(r.warnings[0], /E\.1.*delete it/i);
  assert.match(r.warnings[1], /Reviewer 1.*R1\.0.*delete it/i);
});

test("whitespace-only text before a reviewer or a first comment makes no comment", async () => {
  const r = await splitReview("  \n\nReviewer 1:\n   \n1. a\n2. b\n");
  assert.deepEqual(
    r.comments.map((c) => c.ref),
    ["R1.1", "R1.2"],
  );
  assert.deepEqual(r.warnings, []);
});

test("kept letter and preamble do not count toward skipping the LLM split", async () => {
  const one = "Dear Editor,\nThanks.\n\nReviewer 1:\nOverall fine.\n1. Add units.\n";
  let called = 0;
  const r = await splitReview(one, {
    oneshot: async () => (
      called++,
      { ok: true, json: { comments: [{ reviewer: 1, number: 1, text: "Add units." }] } }
    ),
  });
  assert.equal(called, 1);
  assert.equal(r.method, "llm");
  assert.deepEqual(
    r.comments.map((c) => c.ref),
    ["E.1", "R1.0", "R1.1"],
  );
  assert.equal(r.warnings.filter((w) => /delete it/i.test(w)).length, 2);
});
