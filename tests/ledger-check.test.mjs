import { test } from "node:test";
import assert from "node:assert/strict";
import {
  checkRegexItem,
  checkLlmItem,
  checkLedger,
  batchParagraphs,
  forbidHits,
  ledgerPrompt,
  checkLedgerIncremental,
} from "../daemon/ledger-check.mjs";

const IN_VITRO = {
  id: "L1",
  kind: "instruction",
  text: "in vitro, in vivo, in situ 는 하이픈 없이 쓴다(이탤릭 규칙의 텍스트 대리 검사)",
  check: {
    type: "regex",
    pattern: String.raw`\bin-(?:vitro|vivo|situ)\b`,
    flags: "i",
    expect: "absent",
  },
};
const NOVEL = {
  id: "L2",
  kind: "forbid",
  text: "novel 이라는 표현을 쓰지 않는다",
  check: { type: "regex", pattern: String.raw`\bnovel(?:ty)?\b`, flags: "i", expect: "absent" },
};
const KEEP_STATS = {
  id: "L3",
  kind: "instruction",
  text: "다중비교 보정 문장을 지우지 않는다",
  check: {
    type: "regex",
    pattern: "No correction for multiple comparisons was applied",
    flags: "",
    expect: "present",
  },
};

const PARAS = [
  { index: 3, text: "Earlier reports measured this activity in-vitro only.", reviewed_text: null },
  {
    index: 5,
    text: "Enzyme activity was measured in vitro spectrophotometrically.",
    reviewed_text: null,
  },
  {
    index: 7,
    text: "This is a Novel use. The same decline was seen In-Vitro.",
    reviewed_text: null,
  },
  {
    index: 9,
    text: "Two conditions. No correction for multiple comparisons was applied.",
    reviewed_text: "Two conditions. No correction for multiple comparisons was applied.",
  },
];

test("regex absent-rule finds every hyphenated form and skips the correct one", () => {
  const r = checkRegexItem(IN_VITRO, PARAS);
  assert.equal(r.status, "violations");
  assert.deepEqual(
    r.violations.map((v) => [v.paragraph_index, v.match]),
    [
      [3, "in-vitro"],
      [7, "In-Vitro"],
    ],
  );
});

test("regex checks read reviewed_text: a pending deletion already counts as gone", () => {
  const paras = [{ index: 1, text: "A novel host.\u0005", reviewed_text: "A host." }];
  assert.equal(checkRegexItem(NOVEL, paras).status, "ok");
});

test("present-rule: ok while the sentence exists, violation at the pending deletion", () => {
  assert.equal(checkRegexItem(KEEP_STATS, PARAS).status, "ok");
  const pendingDelete = [
    { index: 9, text: PARAS[3].text, reviewed_text: "Two conditions." },
    { index: 10, text: "Other.", reviewed_text: "Other." },
  ];
  const r = checkRegexItem(KEEP_STATS, pendingDelete);
  assert.equal(r.status, "violations");
  assert.deepEqual(
    r.violations.map((v) => [v.paragraph_index, v.missing]),
    [[9, true]],
  );
  const accepted = [{ index: 9, text: "Two conditions.", reviewed_text: "Two conditions." }];
  assert.deepEqual(checkRegexItem(KEEP_STATS, accepted).violations[0].paragraph_index, null);
});

test("invalid regex is an error, not a crash", () => {
  const r = checkRegexItem({ ...NOVEL, check: { type: "regex", pattern: "(", flags: "" } }, PARAS);
  assert.equal(r.status, "error");
});

test("forbidHits: only forbid items with regex absent-checks", () => {
  const ledger = [
    IN_VITRO,
    NOVEL,
    KEEP_STATS,
    { id: "L4", kind: "forbid", text: "x", check: { type: "llm" } },
  ];
  assert.deepEqual(forbidHits(ledger, "a novel substrate"), ["L2"]);
  assert.deepEqual(
    forbidHits(ledger, "measured in-vitro"),
    [],
    "instruction items are not pre-blocked",
  );
  assert.deepEqual(forbidHits(ledger, "the novelty of it"), ["L2"]);
  assert.deepEqual(forbidHits(ledger, "a nov el idea"), []);
  assert.deepEqual(forbidHits(ledger, ""), []);
});

test("forbidHits pre-blocks manuscript-scope rules only, like the review prompt", async () => {
  const { buildReviewPrompt } = await import("../daemon/review-prompt.mjs");
  const rule = (id, scope, pattern) => ({
    id,
    kind: "forbid",
    text: pattern,
    scope,
    check: { type: "regex", pattern, flags: "i", expect: "absent" },
  });
  const ledger = [
    rule("L1", "response", "we thank the reviewer"),
    rule("L2", "supplement", "see main text"),
    rule("L3", "manuscript", "novel"),
    { ...rule("L4", undefined, "groundbreaking"), scope: undefined }, // old item: manuscript
  ];
  const text = "We thank the reviewer; see main text for this novel, groundbreaking result.";
  assert.deepEqual(forbidHits(ledger, text), ["L3", "L4"]);
  const prompt = buildReviewPrompt({ comment: { ref: "R1.1", reviewer: 1, text: "x" }, ledger });
  assert.deepEqual(prompt.forbidIds, forbidHits(ledger, text));
});

test("batchParagraphs respects the size limit and keeps order", () => {
  const ps = [0, 1, 2, 3].map((i) => ({ index: i, text: "x".repeat(40) }));
  const b = batchParagraphs(ps, 100);
  assert.deepEqual(
    b.map((x) => x.map((p) => p.index)),
    [
      [0, 1],
      [2, 3],
    ],
  );
  assert.equal(batchParagraphs([{ index: 0, text: "y".repeat(500) }], 100).length, 1);
});

const LLM_ITEM = {
  id: "L5",
  kind: "instruction",
  text: "종명은 전체 이름을 한 번만 쓴다",
  check: { type: "llm" },
};

test("LLM check: one call per batch, fabricated quotes and foreign indexes dropped", async () => {
  const calls = [];
  const r = await checkLlmItem(LLM_ITEM, PARAS, {
    maxChars: 120,
    oneshot: async (req) => {
      calls.push(req.prompt);
      return {
        ok: true,
        json: {
          violations: [
            { paragraph_index: 3, quote: "in-vitro only", reason: "r" },
            { paragraph_index: 3, quote: "words that are not there", reason: "r" },
            { paragraph_index: 99, quote: "x", reason: "r" },
          ],
        },
      };
    },
  });
  assert.equal(calls.length, batchParagraphs(PARAS, 120).length);
  assert.equal(r.status, "violations");
  assert.equal(r.violations.length, 1);
  assert.equal(calls[0], ledgerPrompt(LLM_ITEM, batchParagraphs(PARAS, 120)[0]));
});

test("LLM check: batch limit and one-shot failure are errors", async () => {
  let r = await checkLlmItem(LLM_ITEM, PARAS, {
    maxChars: 10,
    maxBatches: 2,
    oneshot: async () => ({ ok: true, json: {} }),
  });
  assert.match(r.error, /needs 4 LLM batches/);
  r = await checkLlmItem(LLM_ITEM, PARAS, { oneshot: async () => ({ ok: false, error: "limit" }) });
  assert.deepEqual([r.status, r.error], ["error", "limit"]);
  r = await checkLlmItem(LLM_ITEM, PARAS, {});
  assert.equal(r.status, "error");
});

test("checkLedger runs every item in order", async () => {
  const r = await checkLedger([IN_VITRO, NOVEL, KEEP_STATS], PARAS);
  assert.deepEqual(
    r.map((x) => [x.id, x.status, x.violations.length]),
    [
      ["L1", "violations", 2],
      ["L2", "violations", 1],
      ["L3", "ok", 0],
    ],
  );
});

// L10. The incremental check must give the same violations as a full check
// of the same document; it only reads fewer paragraphs.
const asStored = (item, o) => ({
  ...item,
  status: o.status,
  violations: o.violations,
  checked_hashes: o.checked_hashes,
});
const pairs = (vs) => vs.map((v) => [v.paragraph_index, v.match]);

async function sameAsFull(items, paragraphs, deps = {}) {
  const inc = await checkLedgerIncremental(items, paragraphs, deps);
  const full = await checkLedgerIncremental(items, paragraphs, { ...deps, full: true });
  inc.forEach((o, i) => {
    assert.equal(o.status, full[i].status, o.id);
    assert.deepEqual(pairs(o.violations), pairs(full[i].violations), o.id);
  });
  return inc;
}

test("L10 first check is full; the next one reads only changed paragraphs", async () => {
  const [first] = await checkLedgerIncremental([IN_VITRO], PARAS);
  assert.deepEqual(
    [first.mode, first.checked, first.total, first.violations.length],
    ["full", 4, 4, 2],
  );
  const item = asStored(IN_VITRO, first);
  const edited = PARAS.map((p) =>
    p.index === 5 ? { ...p, text: "Enzyme activity was measured in-vivo." } : p,
  );
  const [next] = await sameAsFull([item], edited);
  assert.deepEqual([next.mode, next.checked, next.violations.length], ["incremental", 1, 3]);
  assert.deepEqual(pairs(next.violations), [
    [3, "in-vitro"],
    [5, "in-vivo"],
    [7, "In-Vitro"],
  ]);
});

test("L10 a fixed paragraph drops its violation; an unchanged copy keeps it", async () => {
  const twin = { index: 11, text: PARAS[0].text, reviewed_text: null };
  const [first] = await checkLedgerIncremental([IN_VITRO], [...PARAS, twin]);
  const item = asStored(IN_VITRO, first);
  const fixed = PARAS.map((p) => (p.index === 7 ? { ...p, text: "This is a Novel use." } : p));
  const [next] = await sameAsFull([item], [...fixed, twin]);
  assert.deepEqual(pairs(next.violations), [
    [3, "in-vitro"],
    [11, "in-vitro"],
  ]);
  assert.equal(next.checked, 1);
});

test("L10 present-rules, never-checked and errored items run in full", async () => {
  const keep = asStored(KEEP_STATS, (await checkLedgerIncremental([KEEP_STATS], PARAS))[0]);
  const [o] = await checkLedgerIncremental([keep], PARAS);
  assert.deepEqual([o.mode, o.checked], ["full", 4]);
  const bad = { ...IN_VITRO, check: { type: "regex", pattern: "(", flags: "" } };
  const [e] = await checkLedgerIncremental([bad], PARAS);
  assert.deepEqual([e.status, e.checked_hashes], ["error", null]);
  const [again] = await checkLedgerIncremental([asStored(bad, e)], PARAS);
  assert.equal(again.mode, "full");
});

test("L10 LLM items: one-shot calls only for changed paragraphs", async () => {
  const prompts = [];
  const oneshot = async (req) => {
    prompts.push(req.prompt);
    return { ok: true, json: { violations: [] } };
  };
  const [first] = await checkLedgerIncremental([LLM_ITEM], PARAS, { oneshot, maxChars: 1 });
  assert.equal(prompts.length, 4);
  const edited = PARAS.map((p) => (p.index === 3 ? { ...p, text: "Changed." } : p));
  prompts.length = 0;
  const [next] = await checkLedgerIncremental([asStored(LLM_ITEM, first)], edited, {
    oneshot,
    maxChars: 1,
  });
  assert.equal(prompts.length, 1);
  assert.match(prompts[0], /\[3\] Changed\./);
  assert.equal(next.mode, "incremental");
});

// A response or supplement rule is not a manuscript rule: the manuscript
// check marks it skipped (not ok, no violations counted) in both modes.
test("checkLedgerIncremental skips non-manuscript items in incremental and full runs", async () => {
  const RESPONSE_NOVEL = { ...NOVEL, id: "L4", scope: "response" };
  const SUPP_NOVEL = { ...NOVEL, id: "L5", scope: "supplement" };
  const items = [IN_VITRO, RESPONSE_NOVEL, SUPP_NOVEL];
  for (const full of [false, true]) {
    const [inVitro, response, supp] = await checkLedgerIncremental(items, PARAS, { full });
    assert.equal(inVitro.status, "violations");
    assert.equal(inVitro.violations.length, 2);
    for (const [o, id] of [
      [response, "L4"],
      [supp, "L5"],
    ]) {
      assert.equal(o.id, id);
      assert.equal(o.status, "skipped");
      assert.equal(o.skipped, "scope");
      assert.deepEqual(o.violations, []);
      assert.equal(o.checked, 0);
      assert.equal(o.checked_hashes, null);
    }
  }
  // An item stored after an earlier check is still skipped incrementally.
  const stored = { ...RESPONSE_NOVEL, status: "ok", violations: [], checked_hashes: ["x"] };
  const [again] = await checkLedgerIncremental([stored], PARAS);
  assert.equal(again.status, "skipped");
});
