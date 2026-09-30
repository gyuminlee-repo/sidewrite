import { test } from "node:test";
import assert from "node:assert/strict";
import {
  splitReviewRegex,
  splitReview,
  splitPrompt,
  SPLIT_SYSTEM,
  uncoveredText,
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

test("text the regex split drops (no reviewer head) is named on the regex paths too", async () => {
  const one = "Please also cite the method.\n1. Add units.\n2. Add error bars.\n";
  const r = await splitReview(one);
  assert.equal(r.method, "regex");
  assert.ok(r.warnings.some((w) => /1 sentence\(s\) .*in no comment.*cite the method/.test(w)));
  const failed = await splitReview("Please also cite the method.\n1. Add units.\n", {
    oneshot: async () => ({ ok: false, error: "timed out" }),
  });
  assert.ok(failed.warnings.some((w) => /in no comment.*cite the method/.test(w)));
  const clean = await splitReview(TWO_REVIEWERS);
  assert.ok(!clean.warnings.some((w) => /in no comment/.test(w)), clean.warnings.join(" | "));
});

// Text on the reviewer head line after the colon is the reviewer's own
// words: it opens the block (R{n}.0 before numbered comments, R{n}.1 in an
// unnumbered block) and counts in the check for text in no comment.
test("text after a reviewer head on the same line is kept, one-line comments too", async () => {
  const review =
    "Reviewer 1: Please report the sample size.\nComment 1: Add error bars.\nComment 2: State the temperature.\n";
  const r = await splitReview(review);
  assert.equal(r.method, "regex");
  assert.deepEqual(
    r.comments.map((c) => [c.ref, c.text]),
    [
      ["R1.0", "Please report the sample size."],
      ["R1.1", "Add error bars."],
      ["R1.2", "State the temperature."],
    ],
  );
  assert.ok(r.warnings.some((w) => /R1\.0/.test(w)));
  assert.equal(r.warnings.filter((w) => /in no comment/.test(w)).length, 0);
});

test("the no-comment check reads the reviewer head line after its colon", () => {
  const review = "Reviewer 1: Please report the sample size.\nComment 1: Add error bars.\n";
  const lost = uncoveredText(review, [{ text: "Add error bars." }]);
  assert.deepEqual(lost, ["Please report the sample size."]);
  // A head with nothing after it adds nothing to check.
  assert.deepEqual(
    uncoveredText("Reviewer 2:\nComment 1: Add error bars.\n", [{ text: "Add error bars." }]),
    [],
  );
});

test("a journal label on the reviewer head line makes no comment and no warning", async () => {
  const review =
    "Reviewer 1: Comments to the Author\nComment 1: Add error bars.\nComment 2: State n.\n";
  const r = await splitReview(review);
  assert.deepEqual(
    r.comments.map((c) => c.ref),
    ["R1.1", "R1.2"],
  );
  assert.deepEqual(r.warnings, []);
});

// Label path (document structure recovery): the model gets the non-empty
// lines of the review with numbers and answers one type per line; the code
// builds the comments from the source lines. labels(...) answers with the
// given types in line order; ["reviewer_head", n] carries the reviewer number.
const labels =
  (...types) =>
  async () => ({
    ok: true,
    json: {
      lines: types.map((t, i) =>
        Array.isArray(t) ? { n: i + 1, type: t[0], reviewer: t[1] } : { n: i + 1, type: t },
      ),
    },
  });
const RH = (n) => ["reviewer_head", n];
const refsTexts = (r) => r.comments.map((c) => [c.ref, c.text]);

test("the prompt numbers the non-empty lines only and asks for labels, not text", async () => {
  let req;
  const r = await splitReview("Please add units.\n\n  Also state n.\n", {
    oneshot: async (q) => {
      req = q;
      return labels("comment_start", "comment_start")();
    },
  });
  assert.equal(req.prompt, splitPrompt(["Please add units.", "  Also state n."]));
  assert.match(req.prompt, /^1: Please add units\.$/m);
  assert.match(req.prompt, /^2: {3}Also state n\.$/m);
  assert.equal(req.system, SPLIT_SYSTEM);
  assert.equal(r.method, "llm");
  assert.deepEqual(refsTexts(r), [
    ["R1.1", "Please add units."],
    ["R1.2", "Also state n."],
  ]);
  for (const c of r.comments) assert.deepEqual(c.requirements, []);
  assert.deepEqual(r.labels, {
    editor: 0,
    reviewer_head: 0,
    comment_start: 2,
    continue: 0,
    drop: 0,
  });
});

test("a blank line joins the block of the line before it", async () => {
  const r = await splitReview("Please add units.\n\nin every table.\nAlso state n.\n", {
    oneshot: labels("comment_start", "continue", "comment_start"),
  });
  assert.deepEqual(refsTexts(r), [
    ["R1.1", "Please add units.\n\nin every table."],
    ["R1.2", "Also state n."],
  ]);
});

test("empty review text asks no model", async () => {
  let called = 0;
  const r = await splitReview("  \n\n", { oneshot: async () => (called++, labels()()) });
  assert.equal(called, 0);
  assert.deepEqual(r.comments, []);
});

// Codex review 2: a reviewer who did not number the points.
test("labels split unnumbered reviewer text; a block with no comment start is R{n}.1", async () => {
  const review =
    "Reviewer 1:\nPlease report the temperature.\nAlso give the sample size.\n\nReviewer 2:\nAdd error bars.\nto every panel.\n";
  const r = await splitReview(review, {
    oneshot: labels(RH(1), "comment_start", "comment_start", RH(2), "continue", "continue"),
  });
  assert.equal(r.method, "llm");
  assert.deepEqual(refsTexts(r), [
    ["R1.1", "Please report the temperature."],
    ["R1.2", "Also give the sample size."],
    ["R2.1", "Add error bars.\nto every panel."],
  ]);
  assert.ok(r.warnings.some((w) => /Reviewer 2.*not split.*R2\.1/.test(w)));
  assert.ok(!r.warnings.some((w) => /in no comment/.test(w)));
});

// Codex review 5: the request on the reviewer head line.
test("text after a reviewer head on its line opens the block (R{n}.0 or R{n}.1)", async () => {
  const review = "Reviewer 1: Please report the sample size.\nAlso add error bars.\n";
  const pre = await splitReview(review, { oneshot: labels(RH(1), "comment_start") });
  assert.deepEqual(refsTexts(pre), [
    ["R1.0", "Please report the sample size."],
    ["R1.1", "Also add error bars."],
  ]);
  assert.ok(pre.warnings.some((w) => /Reviewer 1.*R1\.0.*delete it/i.test(w)));
  const one = await splitReview(review, { oneshot: labels(RH(1), "continue") });
  assert.deepEqual(refsTexts(one), [
    ["R1.1", "Please report the sample size.\nAlso add error bars."],
  ]);
});

test("a journal label on the reviewer head line adds no text", async () => {
  const review = "Reviewer 1: Comments to the Author\nPlease add units.\nAlso state n.\n";
  const r = await splitReview(review, { oneshot: labels(RH(1), "comment_start", "comment_start") });
  assert.deepEqual(refsTexts(r), [
    ["R1.1", "Please add units."],
    ["R1.2", "Also state n."],
  ]);
  assert.deepEqual(r.warnings, []);
});

test("a reviewer head the regex does not know keeps its whole line", async () => {
  const r = await splitReview("Second referee - please add units.\nAlso state n.\n", {
    oneshot: labels(RH(2), "comment_start"),
  });
  assert.deepEqual(refsTexts(r), [
    ["R2.0", "Second referee - please add units."],
    ["R2.1", "Also state n."],
  ]);
});

// Codex review 3 and 4: the editor letter and a general assessment.
test("editor lines are E.1, a later editor run E.2, the assessment R1.0", async () => {
  const review = `Dear Dr. Lee,
Please shorten the abstract to 200 words.

Reviewer 1:
The study is useful but the methods need detail.
1. Report the growth temperature.
Comment 3: Add error bars to Fig. 2.
Editor note: also fix the title.
`;
  const r = await splitReview(review, {
    oneshot: labels(
      "editor",
      "editor",
      RH(1),
      "continue",
      "comment_start",
      "comment_start",
      "editor",
    ),
  });
  assert.equal(r.method, "llm");
  assert.deepEqual(
    r.comments.map((c) => [c.ref, c.reviewer, c.text]),
    [
      ["E.1", 0, "Dear Dr. Lee,\nPlease shorten the abstract to 200 words."],
      ["E.2", 0, "Editor note: also fix the title."],
      ["R1.0", 1, "The study is useful but the methods need detail."],
      ["R1.1", 1, "Report the growth temperature."],
      ["R1.3", 1, "Add error bars to Fig. 2."],
    ],
  );
  assert.ok(r.warnings.some((w) => /E\.1.*delete it/i.test(w)));
  assert.ok(r.warnings.some((w) => /R1\.0.*delete it/i.test(w)));
  // E.2 comes after a reviewer: its warning does not call it text before one.
  const e2 = r.warnings.filter((w) => /E\.2/.test(w));
  assert.equal(e2.length, 1, r.warnings.join(" | "));
  assert.match(e2[0], /^Editor text after a reviewer .*E\.2.*delete it/i);
});

test("a first line labelled continue starts a comment", async () => {
  const r = await splitReview("Please add units.\nAlso state n.\n", {
    oneshot: labels("continue", "comment_start"),
  });
  assert.deepEqual(
    r.comments.map((c) => c.ref),
    ["R1.1", "R1.2"],
  );
});

test("dropped lines are named in a warning and left out of the comments", async () => {
  const r = await splitReview("Page 2 of 5\nPlease add units.\nAlso state n.\n----\n", {
    oneshot: labels("drop", "comment_start", "comment_start", "drop"),
  });
  assert.deepEqual(refsTexts(r), [
    ["R1.1", "Please add units."],
    ["R1.2", "Also state n."],
  ]);
  assert.ok(
    r.warnings.some((w) => /2 line\(s\) were dropped.*"Page 2 of 5".*"----"/.test(w)),
    r.warnings.join(" | "),
  );
  assert.ok(!r.warnings.some((w) => /more than half/.test(w)));
  assert.equal(r.labels.drop, 2);
});

test("dropping more than half of the lines warns", async () => {
  const r = await splitReview("Page 1\nPage 2\nPlease add units.\n", {
    oneshot: labels("drop", "drop", "comment_start"),
  });
  assert.ok(
    r.warnings.some((w) => /more than half/.test(w)),
    r.warnings.join(" | "),
  );
});

// Codex review 3: an empty or malformed reply keeps the regex split.
const ONE = "Reviewer 1:\nPlease add error bars to Fig. 2.\nAlso state n.\n";
const line = (n, type, extra = {}) => ({ n, type, ...extra });
const BAD = {
  "empty lines": { lines: [] },
  "old comments format": { comments: [{ reviewer: 1, number: 1, text: "Also state n." }] },
  "no object": { foo: 1 },
  "missing line": { lines: [line(1, "reviewer_head", { reviewer: 1 }), line(2, "comment_start")] },
  "duplicate line": {
    lines: [
      line(1, "reviewer_head", { reviewer: 1 }),
      line(2, "comment_start"),
      line(2, "continue"),
    ],
  },
  "line out of range": {
    lines: [
      line(1, "reviewer_head", { reviewer: 1 }),
      line(2, "comment_start"),
      line(4, "continue"),
    ],
  },
  "unknown type": {
    lines: [
      line(1, "reviewer_head", { reviewer: 1 }),
      line(2, "comment_start"),
      line(3, "request"),
    ],
  },
  "reviewer head without a number": {
    lines: [line(1, "reviewer_head"), line(2, "comment_start"), line(3, "continue")],
  },
  "reviewer number not an integer": {
    lines: [
      line(1, "reviewer_head", { reviewer: "1" }),
      line(2, "comment_start"),
      line(3, "continue"),
    ],
  },
  "text returned": {
    lines: [line(1, "reviewer_head", { reviewer: 1 }), line(2, "comment_start"), "continue"],
  },
};
for (const [name, json] of Object.entries(BAD)) {
  test(`LLM reply with ${name} keeps the regex split and warns`, async () => {
    const r = await splitReview(ONE, { oneshot: async () => ({ ok: true, json }) });
    assert.equal(r.method, "regex");
    assert.deepEqual(refsTexts(r), [["R1.1", "Please add error bars to Fig. 2.\nAlso state n."]]);
    assert.ok(
      r.warnings.some((w) =>
        /model split returned no usable line labels; kept the regex split/.test(w),
      ),
      r.warnings.join(" | "),
    );
    assert.equal(r.labels, undefined);
  });
}

// Lossless by construction: whatever labels the model gives (drop aside),
// every non-empty line is in exactly one comment. Lines match no head
// pattern, so no marker is stripped; tokens are unique so a lost or doubled
// line cannot hide behind another.
function mulberry32(seed) {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
const TYPES = ["editor", "reviewer_head", "comment_start", "continue"];

test("property: random labels without drop keep every non-empty line exactly once", async () => {
  const rand = mulberry32(20260930);
  for (let run = 0; run < 50; run++) {
    const count = 1 + Math.floor(rand() * 25);
    const src = [];
    const types = [];
    for (let i = 0; i < count; i++) {
      if (rand() < 0.2) src.push(rand() < 0.5 ? "" : "   ");
      src.push(`${rand() < 0.3 ? "  " : ""}w${run}x${i} asks for item ${i}.`);
      const t = TYPES[Math.floor(rand() * TYPES.length)];
      types.push(t === "reviewer_head" ? RH(1 + Math.floor(rand() * 3)) : t);
    }
    const r = await splitReview(src.join("\n"), { oneshot: labels(...types) });
    assert.equal(r.method, "llm", `run ${run}`);
    const want = src
      .map((l) => l.trim())
      .filter(Boolean)
      .sort();
    const got = r.comments
      .flatMap((c) => c.text.split("\n"))
      .map((l) => l.trim())
      .filter(Boolean)
      .sort();
    assert.deepEqual(got, want, `run ${run}`);
  }
});
