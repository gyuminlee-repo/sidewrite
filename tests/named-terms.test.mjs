// The words a reviewer names are looked for in the whole manuscript: the
// turn is told every place (headings, legends and methods too), and the card
// shows where they still stand after it. All texts here are made up.
import { test } from "node:test";
import assert from "node:assert/strict";
import {
  quotedTerms,
  namedTerms,
  paragraphKinds,
  termPlaces,
  termBlock,
  namedTermLines,
} from "../taskpane/shared/named-terms.js";
import { buildReviewPrompt, REVIEW_RULES } from "../daemon/review-prompt.mjs";
import { extractRequirements, requirementPrompt } from "../daemon/review-split.mjs";
import { emptyRevision, replaceComments } from "../daemon/revision-store.mjs";

test("quotedTerms: what the comment puts in quotation marks or calls a term", () => {
  assert.deepEqual(
    quotedTerms(
      'The use of "convergence" is not justified, and “upregulated/downregulated” should be replaced. ' +
        "The term trade-off is never defined; the words robust and novel are overused. It's the authors' choice.",
    ),
    ["convergence", "upregulated", "downregulated", "trade-off", "robust", "novel"],
  );
  // Single quotes, not apostrophes; a long quotation is a sentence, not a term.
  assert.deepEqual(quotedTerms("Replace 'adaptive' here. The strain's rate isn't shown."), [
    "adaptive",
  ]);
  assert.deepEqual(
    quotedTerms(
      'The authors write "we show beyond any doubt that the gene is required for growth".',
    ),
    [],
  );
  // "the term used in line 3": no term is named.
  assert.deepEqual(quotedTerms("The term used in line 3 is unclear, and the term is vague."), []);
  assert.deepEqual(quotedTerms('Twice: "novel" and "Novel".'), ["novel"]);
  assert.deepEqual(quotedTerms(""), []);
});

test("namedTerms: the comment's own first, then the stored ones its text still holds", () => {
  const comment = {
    text: 'The claim of "convergence" is too strong; metabolic burden is asserted without data.',
    named_terms: ["metabolic burden", "Convergence", "not in the text", ""],
  };
  assert.deepEqual(namedTerms(comment), ["convergence", "metabolic burden"]);
  assert.deepEqual(namedTerms({ text: "Nothing named." }), []);
});

const PAPER = [
  "Convergence of strain ZK-7 lines on glycerol", // 0 heading (title)
  "Abstract", // 1
  "The lines showed convergence, and two genes were upregulated.", // 2 body
  "Materials and Methods", // 3
  "Genes were called upregulated when the ratio was above 2.", // 4 methods
  "Results", // 5
  "3.1 Convergent mutations in glpK", // 6 heading
  "The growth rate rose in all lines.", // 7 body
  "Figure 2. Convergence of the growth rate. Upregulation of glpK is shown in red.", // 8 caption
  "Table 1 Strains used.", // 9 caption
];

test("paragraphKinds: caption, heading, methods, body", () => {
  assert.deepEqual(paragraphKinds(PAPER), [
    "heading",
    "heading",
    "body",
    "heading",
    "methods",
    "heading",
    "heading",
    "body",
    "caption",
    "caption",
  ]);
});

test("termPlaces: every paragraph that holds a term, in any form of the word, with what the paragraph is", () => {
  const comment = {
    text: 'The term "convergence" is not justified; "upregulated" needs a threshold; "epistasis" too.',
  };
  const rows = termPlaces(comment, PAPER, null);
  assert.deepEqual(
    rows.map((r) => [r.term, r.places.map((p) => `${p.id}:${p.kind}`)]),
    [
      ["convergence", ["p0:heading", "p2:body", "p6:heading", "p8:caption"]],
      ["upregulated", ["p2:body", "p4:methods", "p8:caption"]],
      ["epistasis", []],
    ],
  );
  assert.deepEqual(termBlock(comment, PAPER, null), [
    "Terms this comment names, and every place they occur (paragraph id, and what the paragraph is):",
    '- "convergence": p0 (heading), p2 (body), p6 (heading), p8 (caption)',
    '- "upregulated": p2 (body), p4 (methods), p8 (caption)',
    '- "epistasis": not in the manuscript as written',
    "Go to every place: change it there, or say in your end report why it stays. Headings, figure and table legends and the methods count like the body.",
  ]);
  assert.deepEqual(termBlock({ text: "Nothing named." }, PAPER, null), []);
  // Task pane ids are used when the document has them.
  const ids = PAPER.map((_, i) => `id${i}`);
  assert.equal(termPlaces(comment, PAPER, ids)[1].places[1].id, "id4");
  // After the turn: the body was changed, the legend and the methods were not.
  const after = [...PAPER];
  after[0] = "Parallel changes in strain ZK-7 lines on glycerol";
  after[2] = "The lines changed in parallel, and two genes had higher transcript levels.";
  after[6] = "3.1 Recurrent mutations in glpK";
  assert.deepEqual(namedTermLines(comment, { paragraphs: after, ids: null }), [
    'The reviewer names "convergence": still present in 1 place: p8 (caption).',
    'The reviewer names "upregulated": still present in 2 places: p4 (methods), p8 (caption).',
  ]);
  assert.deepEqual(namedTermLines(comment, null), []);
});

test("prompt: the block of named terms with their places, and the rule that every place is handled", () => {
  const comment = {
    ref: "R1.1",
    reviewer: 1,
    text: 'The term "convergence" is not justified.',
  };
  const { text } = buildReviewPrompt({
    comment,
    fence: "F1",
    manuscript: { paragraphs: PAPER, ids: null },
  });
  assert.ok(
    text.includes(
      [
        "Terms this comment names, and every place they occur (paragraph id, and what the paragraph is):",
        '- "convergence": p0 (heading), p2 (body), p6 (heading), p8 (caption)',
      ].join("\n"),
    ),
    text,
  );
  // After the comment, before the rules.
  assert.ok(text.indexOf(">>>F1") < text.indexOf("Terms this comment names"));
  assert.ok(text.indexOf("Terms this comment names") < text.indexOf("\nRules:"));
  const none = buildReviewPrompt({
    comment: { ...comment, text: "Shorten it." },
    fence: "F1",
  }).text;
  assert.equal(none.includes("Terms this comment names"), false);
  const rule = REVIEW_RULES.find((r) => r.startsWith("The reviewer's own words come first"));
  assert.ok(rule);
  assert.match(rule, /headings, figure and table legends and the methods/);
  assert.ok(text.includes(rule));
});

test("requirement labelling also returns the terms each comment names; only those its text holds are kept", async () => {
  const comments = [
    {
      ref: "R1.1",
      reviewer: 1,
      text: "Convergence is claimed without support. Please tone it down.",
      requirements: [],
    },
    { ref: "R1.2", reviewer: 1, text: "Report the unit.", requirements: [] },
  ];
  assert.match(
    requirementPrompt([{ ref: "R1.1", sentences: ["Report the unit."] }]),
    /"terms":\[\{"ref":"R1\.2","terms":\["<word or short phrase>"\]\}\]/,
  );
  const reply =
    (terms) =>
    async ({ prompt }) => ({
      ok: true,
      json: {
        sentences: [...prompt.matchAll(/^\[(R\d+\.\d+) s(\d+)\] /gm)].map((m) => ({
          ref: m[1],
          s: Number(m[2]),
          type: "request",
        })),
        ...(terms === undefined ? {} : { terms }),
      },
    });
  const r = await extractRequirements(comments, {
    oneshot: reply([
      { ref: "R1.1", terms: ["convergence", "epistasis", 7, "x".repeat(80)] },
      { ref: "R9.9", terms: ["unit"] },
    ]),
  });
  assert.deepEqual(
    r.comments.map((c) => c.named_terms ?? []),
    [["convergence"], []],
  );
  // No terms in the reply, or not a list: the labels still count.
  for (const bad of [undefined, "none", [{ ref: "R1.1", terms: "convergence" }]]) {
    const x = await extractRequirements(comments, { oneshot: reply(bad) });
    assert.deepEqual(
      x.comments.map((c) => [c.requirements.length, c.named_terms ?? []]),
      [
        [2, []],
        [1, []],
      ],
    );
    assert.deepEqual(x.warnings, []);
  }
});

test("store: named terms come in with the comment list and stay when a later save does not send them", () => {
  let rev = replaceComments(emptyRevision(), [
    { ref: "R1.1", reviewer: 1, text: "Convergence is claimed.", named_terms: ["convergence"] },
  ]);
  assert.deepEqual(rev.comments[0].named_terms, ["convergence"]);
  rev = replaceComments(rev, [
    { ref: "R1.1", reviewer: 1, text: "Convergence is claimed, twice." },
  ]);
  assert.deepEqual(rev.comments[0].named_terms, ["convergence"]);
  rev = replaceComments(rev, [{ ref: "R1.1", reviewer: 1, text: "x", named_terms: [] }]);
  assert.deepEqual(rev.comments[0].named_terms, []);
});
