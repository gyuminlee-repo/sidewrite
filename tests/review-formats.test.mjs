// Format corpus of the regex split: tests/fixtures/review_formats/.
// <family>.json holds short cases, kind "positive" (the format must split
// this way) or "trap" (a line that looks like a head must stay body text).
// letters.json gives the expected split of the whole letters in letters/.
// Per case: refs in order, texts of some refs, the section label per ref,
// and, where given, whether the split is strong enough to skip the model.
// Every case also runs the two lossless checks: uncoveredText names nothing,
// and the end of every input line is in a comment, a section label or a
// head line.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { splitReview, splitReviewRegex, uncoveredText } from "../daemon/review-split.mjs";

const DIR = new URL("./fixtures/review_formats/", import.meta.url);
const json = (name) => JSON.parse(readFileSync(new URL(name, DIR), "utf8"));

const FAMILIES = [
  "reviewer_heads",
  "editor_heads",
  "comment_heads",
  "sections",
  "traps",
  "boilerplate",
  "combined_heads",
  "sublists",
  "questionnaire",
  "multi_block",
];

// Head lines the split may consume without putting them in a comment.
const HEAD_LINE =
  /^(?:Reviewer|Referee|Comments? (?:from|to|and)|Report of|(?:Associate |Academic |Handling |Guest )?Editor|심사|리뷰어|편집)/i;

const squash = (s) => s.replace(/\s+/g, " ").trim();

function lostLines(text, comments) {
  const hay = comments.map((c) => `${squash(c.text)} | ${c.section ?? ""}`).join(" | ");
  return text
    .split(/\r\n?|\n/)
    .map((l) => squash(l.replace(/^(?:\s*>)+/, "")))
    .filter((l) => l && !HEAD_LINE.test(l) && !/^\(?\d+[.)]$/.test(l) && !/^[*=]{5,}$/.test(l))
    .filter((l) => !hay.includes(l.replace(/[:.]$/, "").slice(-8).trim()));
}

async function check(text, want) {
  const comments = splitReviewRegex(text);
  assert.deepEqual(
    comments.map((c) => c.ref),
    want.refs,
  );
  for (const [ref, t] of Object.entries(want.texts ?? {})) {
    assert.equal(comments.find((c) => c.ref === ref)?.text, t, ref);
  }
  assert.deepEqual(
    Object.fromEntries(comments.filter((c) => c.section).map((c) => [c.ref, c.section])),
    want.sections ?? {},
  );
  for (const c of comments) {
    assert.match(c.ref, /^(?:R\d+|E)\.\d+$/);
    assert.equal(c.reviewer, c.ref.startsWith("E") ? 0 : Number(c.ref.slice(1).split(".")[0]));
    assert.deepEqual(c.requirements, []);
  }
  assert.deepEqual(uncoveredText(text, comments), want.uncovered ?? []);
  assert.deepEqual(lostLines(text, comments), []);
  if (want.strong !== undefined) {
    let called = 0;
    const r = await splitReview(text, {
      oneshot: async () => (called++, { ok: false, error: "off" }),
    });
    assert.equal(called, want.strong ? 0 : 1, "model call");
    assert.deepEqual(
      r.comments.map((c) => c.ref),
      want.refs,
    );
  }
}

for (const family of FAMILIES) {
  const { cases } = json(`${family}.json`);
  test(`${family}: at least 5 positive and 5 trap cases`, () => {
    assert.ok(cases.filter((c) => c.kind === "positive").length >= 5);
    assert.ok(cases.filter((c) => c.kind === "trap").length >= 5);
  });
  for (const c of cases) test(`${family} ${c.kind}: ${c.name}`, () => check(c.text, c));
}

for (const letter of json("letters.json").letters) {
  test(`letter ${letter.file}`, () =>
    check(readFileSync(new URL(`letters/${letter.file}`, DIR), "utf8"), letter));
}

// Regex split and label split of the same letter: whatever the labels, the
// label split keeps every line the regex split keeps.
test("label split of each letter loses no line the regex split has", async () => {
  for (const letter of json("letters.json").letters) {
    const text = readFileSync(new URL(`letters/${letter.file}`, DIR), "utf8");
    const count = text.split("\n").filter((l) => l.trim()).length;
    const r = await splitReview(`Reviewer 1:\n${text.replace(/^(?=\S)/gm, "x ")}`, {
      oneshot: async () => ({
        ok: true,
        json: {
          lines: Array.from({ length: count + 1 }, (_, i) => ({
            n: i + 1,
            type: i ? (i % 3 ? "continue" : "comment_start") : "reviewer_head",
            ...(i ? {} : { reviewer: 1 }),
          })),
        },
      }),
    });
    assert.equal(r.method, "llm", letter.file);
    const got = r.comments.flatMap((c) => c.text.split("\n")).filter((l) => l.trim()).length;
    assert.equal(got, count, letter.file);
  }
});
