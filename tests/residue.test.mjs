// A claim changed in one place and left in three others: after a turn the
// card names the content words the changes of a comment took out that the
// manuscript still holds elsewhere. A report, nothing is refused.
// All texts here are made up.
import { test } from "node:test";
import assert from "node:assert/strict";
import {
  removedTerms,
  claimTerms,
  stillPresent,
  stillPresentLines,
} from "../taskpane/shared/residue.js";
import { buildReviewPrompt, REVIEW_RULES } from "../daemon/review-prompt.mjs";

const ch = (id, find, text, extra = {}) => ({
  id,
  kind: text === "" ? "delete" : "replace",
  state: "pending",
  find,
  text,
  narrowed_from: null,
  // The reason these changes give names the words they took out.
  rationale: { what: "w", why: `Took out: ${find}`, addresses: "a" },
  ...extra,
});

// What the made-up reviewer wrote: it names the claims of these tests.
const SAID =
  "It is not shown that the gene is essential or that the effect is significant; nothing proves it. The same for alphaa, bravoo, charlie, deltaa and echooo.";

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
    text: SAID,
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
  assert.deepEqual(stillPresent({ ref: "R1.1", text: SAID, changes: [] }, PAPER), []);
  assert.deepEqual(stillPresent(comment, undefined), []);
  // Two changes that took the same term out are named together.
  const two = {
    ref: "R1.1",
    text: SAID,
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
    text: SAID,
    changes: [ch("R1.1-1", "a significant and essential effect", "an effect")],
  };
  assert.deepEqual(stillPresent(comment, [...many, "An essential point."]), [
    { term: "essential", change_ids: ["R1.1-1"], places: [20] },
  ]);
  // In a short paper too: a word in more than a fifth of the paragraphs (and
  // in more than 3) is the paper's vocabulary. Seen in a real turn on a
  // made-up paper: "growth", taken out with one sentence of a legend.
  const short = [
    "Growth of strain X",
    "Deletion lowered growth.",
    "The mutant grew slowly.",
    "Growth was followed daily.",
    "Slower growth is a target.",
    "Figure 2. Growth of both strains. The mutant grew at half the rate.",
    "Gene A is essential.",
  ];
  const legend = {
    ref: "R1.1",
    text: SAID,
    changes: [ch("R1.1-1", "Gene A is essential for growth.", "The mutant grew at half the rate.")],
  };
  assert.deepEqual(
    stillPresent(legend, short).map((r) => r.term),
    ["essential"],
  );
  // At most `maxTerms` terms, the ones in the fewest places first.
  const paper = ["alpha1x bravo2x", "bravo2x", "charlie delta", "charlie"];
  const c = {
    ref: "R1.1",
    text: SAID,
    changes: [ch("R1.1-1", "alphaa bravoo charlie deltaa", "")],
  };
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
    text: SAID,
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

// A real revision of 20 comments gave 88 such lines, 8 for most comments:
// "suggests", "primary", "likely", "provides" and the like, the words any
// rewording drops. The term of a claim is one the reviewer named and the
// model names again as its reason, so a removed word is reported only when
// the comment and the why (or addresses) of its change both hold it, in any
// form, and at most three for a comment. (What a change says it did quotes
// the replaced text and names every word of it: measured on that revision,
// counting it too left 121 terms, the reason alone 58, reason and comment 25.)
test("a claim term: a removed word that the comment and the reason of its change both name", () => {
  const reworded = ch(
    "R1.1-1",
    "This finding suggests that gene A is essential and provides a primary target",
    "Gene A was associated with growth and may be a target",
    {
      rationale: {
        what: 'Replaced "This finding suggests that gene A is essential and provides a primary target".',
        why: "The mutant still grows, so essentiality is not shown.",
        addresses: "the claim",
      },
    },
  );
  const said = "The mutant grows, so gene A is not essential. This finding is overstated.";
  assert.deepEqual(removedTerms(reworded), [
    "finding",
    "suggests",
    "essential",
    "provides",
    "primary",
  ]);
  assert.deepEqual(claimTerms(reworded, said), ["essential"]);
  // Named by the reason and not by the reviewer, or the other way round: not a claim term.
  assert.deepEqual(claimTerms(reworded, "Please shorten this paragraph."), []);
  assert.deepEqual(
    claimTerms({ ...reworded, rationale: { what: "w", why: "Clarity.", addresses: "a" } }, said),
    [],
  );
  assert.deepEqual(claimTerms({ ...reworded, rationale: undefined }, said), []);
  const paper = [
    "This suggests a primary role.",
    "Gene A is essential.",
    "It provides a finding.",
    "An essential and primary gene.",
  ];
  assert.deepEqual(stillPresent({ ref: "R1.1", text: said, changes: [reworded] }, paper), [
    { term: "essential", change_ids: ["R1.1-1"], places: [1, 3] },
  ]);
  // At most three terms for a comment unless asked otherwise.
  const many = ch("R1.1-2", "alphaa bravoo charlie deltaa echooo", "");
  const text = ["alphaa", "bravoo", "charlie", "deltaa", "echooo", "x", "y", "z", "w", "v"];
  assert.equal(stillPresent({ ref: "R1.1", text: SAID, changes: [many] }, text).length, 3);
});

// What the model itself says is left ("Same term still at: p4, p7", the
// last line the rules ask for) comes first on the card.
test("the paragraphs the turn itself named come first in the card lines", () => {
  const comment = { ref: "R1.1", text: SAID, changes: [], same_term_at: ["p4", "p7"] };
  assert.deepEqual(stillPresentLines(comment, { paragraphs: ["a"], ids: null }), [
    "The turn says the same term is still at: p4, p7.",
  ]);
  assert.deepEqual(
    stillPresentLines({ ...comment, same_term_at: [] }, { paragraphs: ["a"], ids: null }),
    [],
  );
});
