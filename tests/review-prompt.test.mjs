import { test } from "node:test";
import assert from "node:assert/strict";
import { buildReviewPrompt, formatPromptLog, REVIEW_RULES } from "../daemon/review-prompt.mjs";

const comment = {
  ref: "R1.3",
  reviewer: 1,
  text: "In section 3, please say how growth was followed,\nand describe the substrate as novel.",
};
const ledger = [
  { id: "L1", kind: "instruction", text: "in vitro 는 하이픈 없이 쓴다", scope: "manuscript" },
  { id: "L2", kind: "forbid", text: "novel 을 쓰지 않는다", scope: "manuscript" },
  { id: "L3", kind: "instruction", text: "다중비교 문장을 지우지 않는다", scope: "manuscript" },
  { id: "L4", kind: "instruction", text: "응답서는 두괄식", scope: "response" },
];

test("the comment goes in verbatim with its ref", () => {
  const { text } = buildReviewPrompt({ comment, ledger });
  assert.ok(text.includes(`<<<\n${comment.text}\n>>>`));
  assert.ok(text.startsWith("[Review turn R1.3]"));
  assert.ok(text.includes('Call office_tracked_edit with comment_ref "R1.3".'));
});

test("every manuscript instruction and forbid item is in the prompt; response-scope items are not", () => {
  const b = buildReviewPrompt({ comment, ledger });
  for (const i of ledger.slice(0, 3)) assert.ok(b.text.includes(`- [${i.id}] ${i.text}`), i.id);
  assert.ok(!b.text.includes("L4"));
  assert.deepEqual([b.instructionIds, b.forbidIds, b.omittedForbidIds], [["L1", "L3"], ["L2"], []]);
});

test("rules: G7 (no revision history), no unrequested edits, facts only, language", () => {
  const { text } = buildReviewPrompt({ comment, ledger, lang: "English" });
  assert.match(text, /Do not write the revision history into the manuscript/);
  assert.match(text, /Do not delete or add anything the reviewer did not request/);
  assert.match(text, /facts in the manuscript or the review/);
  assert.match(text, /Write the rationale in English/);
  assert.equal(text.match(/^\d+\. /gm).length, REVIEW_RULES.length);
});

test("selected paragraphs are listed with id and index", () => {
  const { text } = buildReviewPrompt({
    comment,
    ledger,
    paragraphs: [{ id: "4A1B", index: 18, text: "Growth reached a plateau after four days." }],
  });
  assert.match(text, /- id=4A1B index=18: Growth reached a plateau after four days\./);
  assert.ok(!text.includes("No paragraphs were selected"));
});

test("includeForbid=false omits the forbid list and reports what was omitted", () => {
  const b = buildReviewPrompt({ comment, ledger, includeForbid: false });
  assert.ok(!b.text.includes("[L2]"));
  assert.deepEqual([b.forbidIds, b.omittedForbidIds], [[], ["L2"]]);
});

test("formatPromptLog is the C3 evidence line", () => {
  const b = buildReviewPrompt({ comment, ledger });
  assert.equal(
    formatPromptLog("R1.3", b, { paragraphs: 1 }),
    "[review] prompt R1.3 instructions=L1,L3 forbid=L2 omitted_forbid=- paragraphs=1",
  );
  assert.equal(
    formatPromptLog("R1.3", buildReviewPrompt({ comment, ledger: [] }), { preview: true }),
    "[review] prompt R1.3 instructions=- forbid=- omitted_forbid=- paragraphs=0 (preview)",
  );
});

test("L5 requirements are listed with their ids; none listed means no block", () => {
  const comment = {
    ref: "R1.3",
    reviewer: 1,
    text: "Say how growth was followed and call it novel.",
    requirements: [
      { id: "q1", text: "say how growth was followed" },
      { id: "q2", text: "describe the substrate as novel" },
    ],
  };
  const { text } = buildReviewPrompt({ comment });
  assert.match(text, /rationale\.requirement_ids/);
  assert.match(
    text,
    /- \[q1\] say how growth was followed\n- \[q2\] describe the substrate as novel/,
  );
  const plain = buildReviewPrompt({ comment: { ...comment, requirements: [] } }).text;
  assert.doesNotMatch(plain, /requirement_ids/);
});

test("an editor comment E.1 is introduced as the editor, not reviewer 0", () => {
  const { text } = buildReviewPrompt({
    comment: { ref: "E.1", reviewer: 0, text: "Shorten the abstract." },
  });
  assert.ok(text.includes("Editor, comment E.1 (verbatim):"));
  assert.ok(!text.includes("Reviewer 0"));
});

// S2-30: the 2026-10-01 acceptance run left requirement_ids empty on a change
// that answered part of R1.3. Partial counts; empty only when unrelated.
test("L5 prompt: cite every request a change handles, even in part; empty only when unrelated", () => {
  const comment = {
    ref: "R1.3",
    reviewer: 1,
    text: "Say how growth was followed.",
    requirements: [{ id: "q1", text: "say how growth was followed" }],
  };
  const { text } = buildReviewPrompt({ comment });
  assert.match(text, /- \[q1\] say how growth was followed\n/);
  assert.match(text, /rationale\.requirement_ids/);
  assert.match(text, /part of a request[^\n]*still put its id/i);
  assert.match(text, /empty only when the change answers none of them/i);
  assert.doesNotMatch(text, /not sure/);
});
