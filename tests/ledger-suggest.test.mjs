import { test } from "node:test";
import assert from "node:assert/strict";
import {
  reviewSentences,
  suggestLedgerRegex,
  suggestLedger,
  suggestPrompt,
  SUGGEST_SYSTEM,
} from "../daemon/ledger-suggest.mjs";

const REVIEW = [
  "Reviewer 1",
  'Please do not use the word "novel" for this approach.',
  "Avoid overstating the clinical relevance.",
  'Keep the sentence "No correction for multiple comparisons was applied" in Methods.',
  "The figures are clear.",
].join("\n");

test("regex: quoted forbid becomes an absent regex, keep becomes present, cue without quote is an LLM rule", () => {
  const c = suggestLedgerRegex(REVIEW);
  const novel = c.find((x) => x.text.includes("novel"));
  assert.equal(novel.kind, "forbid");
  assert.deepEqual(novel.check, { type: "regex", pattern: "novel", flags: "i", expect: "absent" });
  assert.equal(novel.sentence_index, 1);
  const keep = c.find((x) => x.text.startsWith("Keep"));
  assert.equal(keep.check.expect, "present");
  assert.equal(keep.check.pattern, "No correction for multiple comparisons was applied");
  const over = c.find((x) => x.text.startsWith("Avoid"));
  assert.deepEqual([over.kind, over.check.type], ["instruction", "llm"]);
  assert.equal(
    c.some((x) => x.text.includes("figures")),
    false,
  );
  for (const x of c) assert.equal(x.scope, "manuscript");
});

test("sentences: one per line, a heading stays apart, a wrapped line is joined", () => {
  assert.deepEqual(reviewSentences(REVIEW).length, 5);
  assert.deepEqual(
    reviewSentences(
      'Reviewer 2\nPlease do not call the\n"new" method novel. It is not.\n2. Keep it.',
    ),
    ["Reviewer 2", 'Please do not call the "new" method novel.', "It is not.", "Keep it."],
  );
});

test("regex: a pattern is the quoted phrase as literal text", () => {
  const [c] = suggestLedgerRegex('Never write "p < 0.05 (two-sided)".');
  assert.equal(c.check.pattern, String.raw`p < 0\.05 \(two-sided\)`);
  assert.ok(new RegExp(c.check.pattern, c.check.flags).test("P < 0.05 (two-sided)"));
  // Apostrophes are not quotes.
  const [d] = suggestLedgerRegex("Don't overstate the authors' claims.");
  assert.deepEqual([d.kind, d.check], ["instruction", { type: "llm" }]);
  const [e] = suggestLedgerRegex("Do not delete 'in vitro' anywhere.");
  assert.deepEqual(
    [e.kind, e.check.pattern, e.check.expect],
    ["instruction", "in vitro", "present"],
  );
});

test("model labels sentences by index; text is copied by code, never by the model", async () => {
  const asked = [];
  const oneshot = async (req) => {
    assert.equal(req.system, SUGGEST_SYSTEM);
    asked.push(req.prompt);
    return {
      ok: true,
      json: {
        labels: [
          { i: 4, type: "instruction" },
          { i: 99, type: "forbid" },
          { i: 1, type: "bogus" },
        ],
      },
    };
  };
  const r = await suggestLedger(REVIEW, { oneshot });
  assert.equal(r.method, "regex+llm");
  assert.ok(
    r.candidates.some((x) => x.text === "The figures are clear." && x.check.type === "llm"),
  );
  assert.equal(r.candidates.filter((x) => x.text.includes("novel")).length, 1);
  assert.ok(r.warnings.some((w) => /99/.test(w)));
  assert.deepEqual(
    r.candidates.map((x) => x.sentence_index),
    [1, 2, 3, 4],
  );
  // Only sentences the regex pass left open are sent, under their index.
  assert.equal(asked.length, 1);
  assert.equal(asked[0], suggestPrompt(reviewSentences(REVIEW), [0, 4]));
  assert.match(asked[0], /\n4: The figures are clear\.$/);
  assert.equal(r.asked, true);
});

test("a failed model call keeps the regex candidates", async () => {
  const r = await suggestLedger(REVIEW, { oneshot: async () => ({ ok: false, error: "x" }) });
  assert.equal(r.method, "regex");
  assert.equal(r.candidates.length, 3);
  assert.ok(r.warnings.some((w) => /x/.test(w)));
});

test("no model, or nothing left to ask: regex only and no call", async () => {
  const a = await suggestLedger(REVIEW, {});
  assert.deepEqual([a.method, a.candidates.length, a.asked], ["regex", 3, false]);
  let calls = 0;
  const b = await suggestLedger('Do not write "novel".', {
    oneshot: async () => (calls++, { ok: true, json: { labels: [] } }),
  });
  assert.deepEqual([b.method, b.candidates.length, calls], ["regex", 1, 0]);
});
