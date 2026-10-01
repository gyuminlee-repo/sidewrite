// A claim changed in one place and left in three others: after a turn the
// card names the content words the changes of a comment took out that the
// manuscript still holds elsewhere. A report, nothing is refused.
// All texts here are made up.
import { test } from "node:test";
import assert from "node:assert/strict";
import { removedTerms, stillPresent, stillPresentLines } from "../taskpane/shared/residue.js";
import { buildReviewPrompt, REVIEW_RULES } from "../daemon/review-prompt.mjs";

const ch = (id, find, text, extra = {}) => ({
  id,
  kind: text === "" ? "delete" : "replace",
  state: "pending",
  find,
  text,
  narrowed_from: null,
  ...extra,
});

test("removedTerms: content words of the replaced text that the new text no longer has", () => {
  assert.deepEqual(
    removedTerms(
      ch(
        "c1",
        "is essential for the growth of strain X",
        "was associated with the growth of strain X",
      ),
    ),
    ["essential"],
  );
  // A word that stays in another form is not removed; short and common words are not terms.
  assert.deepEqual(
    removedTerms(ch("c1", "Deletion inhibits growth", "Deletion inhibited growth")),
    [],
  );
  assert.deepEqual(removedTerms(ch("c1", "which was also then very high", "high")), []);
  // A deletion takes out every content word of its text; numbers are not terms.
  assert.deepEqual(
    removedTerms(ch("c1", "This demonstrates a causal mechanism in 12 strains.", "")),
    ["demonstrates", "causal", "mechanism", "strains"],
  );
  // A narrowed change: by the find and text the model asked for.
  assert.deepEqual(
    removedTerms(
      ch("c1", "essential", "associated", {
        narrowed_from: { find: "is essential for", text: "is associated with" },
      }),
    ),
    ["essential"],
  );
  // An insertion takes nothing out.
  assert.deepEqual(
    removedTerms({ id: "c1", kind: "insert", find: "essential", text: " and more" }),
    [],
  );
});

const PAPER = [
  "Gene A is essential for growth on methanol.", // 0 abstract
  "Introduction",
  "Earlier work suggested a role of gene A.",
  "Results",
  "Deletion of gene A was associated with slower growth.", // 4: the changed paragraph, as it reads now
  "Figure 2. Gene A is essential for growth.", // 5 legend
  "These data confirm the Essential role of gene A.", // 6
  "Growth was followed for 48 h.",
];

test("stillPresent: the removed term with the other paragraphs that hold it, rarest first", () => {
  const comment = {
    ref: "R1.1",
    changes: [
      ch("R1.1-1", "is essential for slower growth", "was associated with slower growth"),
      ch("R1.1-2", "These data prove a role", "These data suggest a role", { state: "rejected" }),
    ],
  };
  assert.deepEqual(stillPresent(comment, PAPER), [
    { term: "essential", change_ids: ["R1.1-1"], places: [0, 5, 6] },
  ]);
  // Nothing left elsewhere: nothing to say.
  assert.deepEqual(stillPresent(comment, [PAPER[4], PAPER[7]]), []);
  assert.deepEqual(stillPresent({ ref: "R1.1", changes: [] }, PAPER), []);
  assert.deepEqual(stillPresent(comment, undefined), []);
  // Two changes that took the same term out are named together.
  const two = {
    ref: "R1.1",
    changes: [
      ch("R1.1-1", "is essential for", "was associated with"),
      ch("R1.1-3", "an essential gene", "a gene"),
    ],
  };
  assert.deepEqual(stillPresent(two, PAPER)[0].change_ids, ["R1.1-1", "R1.1-3"]);
});

test("stillPresent: a word that is in many paragraphs is no claim term; the list is short", () => {
  const many = Array.from({ length: 20 }, (_, i) => `Growth of culture ${i} was significant.`);
  const comment = {
    ref: "R1.1",
    changes: [ch("R1.1-1", "a significant and essential effect", "an effect")],
  };
  assert.deepEqual(stillPresent(comment, [...many, "An essential point."]), [
    { term: "essential", change_ids: ["R1.1-1"], places: [20] },
  ]);
  // At most `maxTerms` terms, the ones in the fewest places first.
  const paper = ["alpha1x bravo2x", "bravo2x", "charlie delta", "charlie"];
  const c = { ref: "R1.1", changes: [ch("R1.1-1", "alphaa bravoo charlie deltaa", "")] };
  const text = ["alphaa", "bravoo bravoo", "charlie", "bravoo charlie deltaa", "charlie", "bravoo"];
  assert.deepEqual(
    stillPresent(c, text, { maxTerms: 2 }).map((r) => [r.term, r.places.length]),
    [
      ["alphaa", 1],
      ["deltaa", 1],
    ],
  );
  assert.equal(stillPresent(c, paper).length, 1); // only "charlie" is there as a word
});

test("stillPresentLines: the card lines, with paragraph ids", () => {
  const comment = {
    ref: "R1.1",
    changes: [ch("R1.1-1", "is essential for slower growth", "was associated with slower growth")],
  };
  assert.deepEqual(stillPresentLines(comment, { paragraphs: PAPER, ids: null }), [
    '"essential" (taken out by R1.1-1) is still present in 3 other places: p0, p5, p6.',
  ]);
  const ids = PAPER.map((_, i) => `id${i}`);
  assert.deepEqual(stillPresentLines(comment, { paragraphs: [PAPER[4], PAPER[5]], ids }), [
    '"essential" (taken out by R1.1-1) is still present in 1 other place: id1.',
  ]);
  assert.deepEqual(stillPresentLines(comment, null), []);
});

test("prompt rules: the same claim everywhere, withdrawal instead of a caveat, and the check before the turn ends", () => {
  const comment = { ref: "R1.1", reviewer: 1, text: "This claim is not supported." };
  const { text } = buildReviewPrompt({ comment, fence: "F1" });
  const rule = (start) => {
    const r = REVIEW_RULES.find((x) => x.startsWith(start));
    assert.ok(r, start);
    assert.ok(text.includes(r));
    return r;
  };
  const everywhere = rule(
    "A claim you change for this comment is changed everywhere the manuscript makes it",
  );
  assert.match(
    everywhere,
    /abstract, introduction, results, discussion, conclusion, headings, figure and table legends/,
  );
  assert.match(everywhere, /That is part of what the comment asks for/);
  const withdraw = rule("When the reviewer says a claim does not hold");
  assert.match(withdraw, /take the claim out or rewrite it as what was observed/);
  assert.match(withdraw, /Do not keep the claim and put a caveat next to it/);
  assert.match(withdraw, /no sentence and no paragraph may state a claim and its denial together/);
  assert.match(withdraw, /only when the reviewer asks for that/);
  const check = rule("Before you end the turn, check:");
  assert.match(check, /every request has a change, a question or a note/);
  assert.match(check, /no changed paragraph states a claim and its denial together/);
  assert.match(check, /End your answer with the line "Same term still at: "/);
  // Appended: the pack of rules 2 to 4 keeps its place.
  assert.ok(REVIEW_RULES.indexOf(everywhere) > 6);
});
