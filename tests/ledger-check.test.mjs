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
  inScope,
  inManuscriptScope,
} from "../daemon/ledger-check.mjs";
import { ledgerReply, ledgerJson, readLedgerPrompt } from "./fixtures/ledger-reply.mjs";

test("inScope keeps the manuscript default", () => {
  assert.equal(inScope({}, "manuscript"), true);
  assert.equal(inScope({ scope: "response" }, "response"), true);
  assert.equal(inScope({ scope: "response" }, "manuscript"), false);
  assert.equal(inScope({}, "response"), false);
  assert.equal(inScope(null, "manuscript"), true);
  for (const item of [{}, { scope: "manuscript" }, { scope: "response" }, { scope: "supplement" }])
    assert.equal(inManuscriptScope(item), inScope(item, "manuscript"));
});

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

test("LLM check: one call per batch, a fabricated quote is dropped", async () => {
  const calls = [];
  const r = await checkLlmItem(LLM_ITEM, PARAS, {
    maxChars: 120,
    oneshot: async (req) => {
      calls.push(req.prompt);
      return ledgerReply(req, PARAS, {
        violations: [
          { paragraph_index: 3, quote: "in-vitro only", reason: "r" },
          { paragraph_index: 3, quote: "words that are not there", reason: "r" },
        ],
      });
    },
  });
  const batches = batchParagraphs(PARAS, 120);
  assert.equal(calls.length, batches.length);
  assert.equal(r.status, "violations");
  assert.equal(r.violations.length, 1);
  assert.deepEqual(r.unsure, []);
  // The request: the paragraphs under numbers of the batch, and its id.
  const { batch_id, paragraphs } = readLedgerPrompt(calls[0]);
  assert.match(batch_id, /^b\d+$/);
  assert.deepEqual(
    paragraphs.map((x) => x.id),
    batches[0].map((_, k) => k),
  );
  assert.equal(calls[0], ledgerPrompt(LLM_ITEM, batches[0], batch_id));
  // Every batch has an id of its own.
  assert.equal(new Set(calls.map((c) => readLedgerPrompt(c).batch_id)).size, calls.length);
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
    return ledgerReply(req, []);
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
  assert.match(prompts[0], /\[0\] Changed\./);
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

// An empty match must advance by one code point in unicode mode. Advancing by
// one UTF-16 unit lands inside a surrogate pair and exec() matches the same
// place again forever.
test("regex empty matches advance by code point under the u and v flags", () => {
  const para = (text) => [{ index: 0, text }];
  for (const flags of ["u", "v"]) {
    const item = { id: "L9", check: { type: "regex", pattern: "(?=.)", flags, expect: "absent" } };
    const astral = checkRegexItem(item, para("\u{1D400}x"));
    assert.deepEqual(
      astral.violations.map((v) => v.start),
      [0, 2],
    );
    const ascii = checkRegexItem(item, para("ab"));
    assert.deepEqual(
      ascii.violations.map((v) => v.start),
      [0, 1],
    );
  }
  // Without a unicode flag the pattern sees UTF-16 units, one match each.
  const plain = {
    id: "L9",
    check: { type: "regex", pattern: "(?=[^])", flags: "", expect: "absent" },
  };
  assert.equal(checkRegexItem(plain, para("\u{1D400}")).violations.length, 2);
  assert.deepEqual(
    forbidHits([{ ...plain, kind: "forbid", check: { ...plain.check, flags: "u" } }], "\u{1D400}"),
    ["L9"],
  );
});

// H4. The reply holds one verdict for every paragraph of the batch, under
// the batch id of the request. A model that skips paragraphs and answers
// "no violations" used to read like a clean pass, and the skipped paragraphs
// were cached as checked. Anything but exactly one verdict per paragraph
// asked about is a failed check: an error, nothing cached.
const P3 = [
  { index: 0, text: "It grew on a novel substrate.", reviewed_text: null },
  { index: 4, text: "The medium was warm.", reviewed_text: null },
  { index: 6, text: "Cells were counted twice.", reviewed_text: null },
];
// A fake that answers the first request with `bad(json of a good reply)` and,
// asked again, with `second` (default: the same bad reply).
function replies(first, second = first) {
  const prompts = [];
  const oneshot = async (req) => {
    prompts.push(req.prompt);
    const make = prompts.length === 1 ? first : second;
    return asResult(make(ledgerJson(req, P3), req));
  };
  return { oneshot, prompts };
}
// What the one-shot runner returns: the JSON of the reply, or its failure.
const asResult = (out) => (out && out.ok === false ? out : { ok: true, json: out });
const V = { paragraph_id: 0, quote: "a novel substrate", reason: "r" };
const BAD_REPLIES = [
  // [name, the reply made from a good one, what the error names]
  [
    "a paragraph has no verdict",
    (j) => ({ ...j, verdicts: j.verdicts.slice(0, 2) }),
    /no verdict for paragraph\(s\) 2/,
  ],
  ["no verdicts at all (the old form)", () => ({ violations: [] }), /batch_id|verdicts/],
  [
    "an empty verdict list",
    (j) => ({ ...j, verdicts: [] }),
    /no verdict for paragraph\(s\) 0, 1, 2/,
  ],
  [
    "a paragraph has two verdicts",
    (j) => ({ ...j, verdicts: [...j.verdicts, [1, "ok"]] }),
    /paragraph 1 has more than one verdict/,
  ],
  [
    "a verdict for a paragraph outside the batch",
    (j) => ({ ...j, verdicts: [...j.verdicts, [3, "ok"]] }),
    /paragraph 3.*not in the batch/,
  ],
  [
    "a negative paragraph number",
    (j) => ({ ...j, verdicts: [[-1, "ok"], ...j.verdicts.slice(1)] }),
    /paragraph -1.*not in the batch/,
  ],
  [
    "a paragraph number that is a fraction",
    (j) => ({ ...j, verdicts: [[0.5, "ok"], ...j.verdicts.slice(1)] }),
    /verdicts\[0\]/,
  ],
  ["another batch id", (j) => ({ ...j, batch_id: "b0" }), /batch_id/],
  ["no batch id", (j) => ({ ...j, batch_id: undefined }), /batch_id/],
  [
    "a verdict that is none of the three",
    (j) => ({ ...j, verdicts: [[0, "fine"], ...j.verdicts.slice(1)] }),
    /verdicts\[0\]/,
  ],
  [
    "a verdict that is not a pair",
    (j) => ({ ...j, verdicts: [0, ...j.verdicts.slice(1)] }),
    /verdicts\[0\]/,
  ],
  [
    "verdict violation with no entry in violations",
    (j) => ({ ...j, verdicts: [[0, "violation"], ...j.verdicts.slice(1)] }),
    /paragraph 0.*"violation".*no entry in violations/,
  ],
  [
    "an entry in violations for a paragraph judged ok",
    (j) => ({ ...j, violations: [V] }),
    /violations\[0\].*paragraph 0.*"ok"/,
  ],
  [
    "an entry in unsure for a paragraph judged ok",
    (j) => ({ ...j, unsure: [{ paragraph_id: 1, reason: "?" }] }),
    /unsure\[0\].*paragraph 1.*"ok"/,
  ],
  [
    "a violation without a quote",
    (j) => ({
      ...j,
      verdicts: [[0, "violation"], ...j.verdicts.slice(1)],
      violations: [{ paragraph_id: 0, reason: "r" }],
    }),
    /violations\[0\]\.quote/,
  ],
  [
    "a violation without a reason",
    (j) => ({
      ...j,
      verdicts: [[0, "violation"], ...j.verdicts.slice(1)],
      violations: [{ paragraph_id: 0, quote: "novel" }],
    }),
    /violations\[0\]\.reason/,
  ],
  [
    "a violation without a paragraph",
    (j) => ({ ...j, violations: [{ quote: "novel", reason: "r" }] }),
    /violations\[0\]\.paragraph_id/,
  ],
  ["violations is not a list", (j) => ({ ...j, violations: "none" }), /violations/],
  ["unsure is not a list", (j) => ({ ...j, unsure: "x" }), /unsure/],
  ["an unsure entry that is a number", (j) => ({ ...j, unsure: [4] }), /unsure\[0\]/],
  ["the reply is a list", () => [], /not a JSON object/],
  ["the reply is null", () => null, /not a JSON object/],
  ["the reply is not JSON", () => ({ ok: false, error: "reply was not JSON" }), /not JSON/],
];

test("LLM check: every paragraph judged ok is a pass, and all of them are cached", async () => {
  const { oneshot, prompts } = replies((j) => j);
  const r = await checkLlmItem(LLM_ITEM, P3, { oneshot });
  assert.deepEqual([r.status, r.violations, r.unsure], ["ok", [], []]);
  assert.equal(prompts.length, 1, "no second request for a good reply");
  const [o] = await checkLedgerIncremental([LLM_ITEM], P3, { oneshot });
  assert.deepEqual([o.status, o.checked_hashes.length], ["ok", 3]);
});

test("LLM check: a reply that is not one verdict per paragraph of the batch fails the check, asked twice", async () => {
  for (const [name, bad, why] of BAD_REPLIES) {
    const a = replies(bad);
    const verdicts = [];
    const r = await checkLlmItem(LLM_ITEM, P3, {
      oneshot: a.oneshot,
      onDecision: (d) => verdicts.push(d.verdict),
    });
    assert.equal(r.status, "error", name);
    assert.match(r.error, why, name);
    assert.deepEqual([r.violations, r.unsure ?? []], [[], []], name);
    assert.deepEqual(verdicts, ["error", "error", "error"], name);
    // Asked once more, with what was wrong, about the same batch.
    assert.equal(a.prompts.length, 2, name);
    assert.ok(a.prompts[1].startsWith(a.prompts[0]), name);
    assert.match(a.prompts[1].slice(a.prompts[0].length), why, name);
    assert.equal(readLedgerPrompt(a.prompts[1]).batch_id, readLedgerPrompt(a.prompts[0]).batch_id);
    // Nothing is cached: no paragraph of a failed batch counts as checked.
    const [o] = await checkLedgerIncremental([LLM_ITEM], P3, { oneshot: replies(bad).oneshot });
    assert.deepEqual([o.status, o.checked_hashes], ["error", null], name);
  }
  console.log(`ledger reply form: ${BAD_REPLIES.length} unusable replies, each an error`);
});

test("LLM check: a good reply to the second request is used; the first one is not", async () => {
  for (const [name, bad] of BAD_REPLIES) {
    const a = replies(bad, (j) => ({
      ...j,
      verdicts: [[0, "violation"], ...j.verdicts.slice(1)],
      violations: [V],
    }));
    const r = await checkLlmItem(LLM_ITEM, P3, { oneshot: a.oneshot });
    assert.equal(a.prompts.length, 2, name);
    assert.deepEqual(
      [r.status, r.violations.map((v) => [v.paragraph_index, v.match]), r.unsure],
      ["violations", [[0, "a novel substrate"]], []],
      name,
    );
  }
  // A failed model call (limit, refusal, timeout, a server whose answer was
  // no JSON at all) is not a form error: no second request.
  for (const error of [
    "Claude usage limit reached.",
    "the reply was not JSON from the server",
    "agy output was not JSON",
  ]) {
    let n = 0;
    const r = await checkLlmItem(LLM_ITEM, P3, {
      oneshot: async () => (n++, { ok: false, error }),
    });
    assert.deepEqual([r.status, r.error, n], ["error", error, 1]);
  }
});

test("LLM check: paragraph numbers given as numeric strings are read; lists may be left out when empty", async () => {
  const a = replies((j) => ({
    batch_id: j.batch_id,
    verdicts: [
      ["0", "violation"],
      ["1", "ok"],
      ["2", "unsure"],
    ],
    violations: [{ ...V, paragraph_id: "0" }],
  }));
  const r = await checkLlmItem(LLM_ITEM, P3, { oneshot: a.oneshot });
  assert.equal(a.prompts.length, 1);
  assert.deepEqual(
    [r.status, r.violations.map((v) => v.paragraph_index), r.unsure],
    ["violations", [0], [{ paragraph_index: 6, reason: "" }]],
  );
});

// Abstain: a paragraph the model cannot judge comes back as "unsure", a
// normal result. No confidence number is asked for or read.
test("LLM check: unsure paragraphs are a result, and incremental checks read them again", async () => {
  const prompts = [];
  const UNSURE = { unsure: [{ paragraph_index: 5, reason: "cannot tell" }] };
  const oneshot = async (req) => {
    prompts.push(req);
    return ledgerReply(req, PARAS, UNSURE);
  };
  const r = await checkLlmItem(LLM_ITEM, PARAS, { oneshot });
  assert.equal(r.status, "unsure");
  assert.deepEqual(r.unsure, [{ paragraph_index: 5, reason: "cannot tell" }]);
  assert.match(prompts[0].system, /unsure/);
  assert.match(prompts[0].prompt, /"unsure"/);
  assert.doesNotMatch(prompts[0].prompt, /confidence|probability/i);

  const [first] = await checkLedgerIncremental([LLM_ITEM], PARAS, { oneshot, maxChars: 1 });
  assert.equal(first.status, "unsure");
  assert.equal(first.unsure.length, 1);
  assert.equal(first.unsure[0].paragraph_index, 5);
  // Only the three paragraphs with verdict ok are stored.
  assert.equal(first.checked_hashes.length, PARAS.length - 1);
  assert.ok(!first.checked_hashes.includes(first.unsure[0].h));
  prompts.length = 0;
  const quiet = async (req) => {
    prompts.push(req);
    return ledgerReply(req, PARAS);
  };
  const [next] = await checkLedgerIncremental(
    [{ ...asStored(LLM_ITEM, first), unsure: first.unsure }],
    PARAS,
    { oneshot: quiet, maxChars: 1 },
  );
  assert.equal(next.mode, "incremental");
  assert.equal(prompts.length, 1);
  assert.match(prompts[0].prompt, /\[0\] Enzyme activity was measured in vitro/);
  assert.deepEqual([next.status, next.unsure, next.checked_hashes.length], ["ok", [], 4]);
});

test("LLM check: violations outrank unsure in the status of the item", async () => {
  const both = await checkLlmItem(LLM_ITEM, PARAS, {
    oneshot: async (req) =>
      ledgerReply(req, PARAS, {
        violations: [{ paragraph_index: 3, quote: "in-vitro only", reason: "r" }],
        unsure: [{ paragraph_index: 5, reason: "?" }],
      }),
  });
  assert.equal(both.status, "violations");
  assert.deepEqual(
    both.unsure.map((u) => u.paragraph_index),
    [5],
  );
});

// Codex review 9: a paragraph read again keeps only what the new check
// says: an old violation does not survive it.
test("LLM check: an unsure paragraph is not stored; a recheck drops its old violation", async () => {
  const unsure5 = async (req) =>
    ledgerReply(req, PARAS, { unsure: [{ paragraph_index: 5, reason: "cannot tell" }] });
  const decisions = [];
  const r = await checkLlmItem(LLM_ITEM, PARAS, {
    oneshot: unsure5,
    onDecision: (d) => decisions.push(d),
  });
  assert.deepEqual([r.status, r.violations, r.unsure.length], ["unsure", [], 1]);
  assert.equal(decisions.find((d) => d.paragraph_index === 5).verdict, "unsure");

  const [first] = await checkLedgerIncremental([LLM_ITEM], PARAS, { oneshot: unsure5 });
  assert.deepEqual([first.status, first.violations], ["unsure", []]);
  assert.equal(first.checked_hashes.length, PARAS.length - 1, "the unsure paragraph is not stored");

  // Codex reproduction: the stored item holds a violation for the paragraph
  // it also left unchecked (as the old code stored it). The recheck finds
  // nothing there, so the old violation must go and the hash is stored.
  const stored = {
    ...asStored(LLM_ITEM, first),
    status: "violations",
    violations: [
      { paragraph_index: 5, start: null, match: "in vitro", reason: "r", h: first.unsure[0].h },
    ],
    unsure: first.unsure,
  };
  const prompts = [];
  const quiet = async (req) => {
    prompts.push(req);
    return ledgerReply(req, PARAS);
  };
  const [next] = await checkLedgerIncremental([stored], PARAS, { oneshot: quiet });
  assert.equal(next.mode, "incremental");
  assert.match(prompts[0].prompt, /\[0\] Enzyme activity was measured in vitro/);
  assert.deepEqual([next.status, next.violations, next.unsure], ["ok", [], []]);
  assert.equal(next.checked_hashes.length, PARAS.length);

  // A paragraph skipped by the cache keeps its violation.
  const keepItem = {
    ...stored,
    checked_hashes: next.checked_hashes,
    violations: [{ ...stored.violations[0] }],
  };
  const [kept] = await checkLedgerIncremental([keepItem], PARAS, { oneshot: quiet });
  assert.deepEqual([kept.checked, kept.violations.map((v) => v.paragraph_index)], [0, [5]]);
});

// Codex review 12: a violation whose quote is not in the paragraph was
// dropped, the paragraph counted as ok and its hash was cached, so the next
// incremental check skipped it ("An in-vitro assay." / quote "in-vitro
// assays": first check ok, next check checked 0).
test("LLM check: a violation with a quote that is not in the paragraph leaves it unsure, and it is read again", async () => {
  const paras = [
    { index: 1, text: "An in-vitro assay.", reviewed_text: null },
    { index: 2, text: "Nothing to see here.", reviewed_text: null },
  ];
  const prompts = [];
  const oneshot = async (req) => {
    prompts.push(req.prompt);
    return ledgerReply(req, paras, {
      violations: [{ paragraph_index: 1, quote: "in-vitro assays", reason: "hyphen" }],
    });
  };
  const verdicts = [];
  const r = await checkLlmItem(LLM_ITEM, paras, {
    oneshot,
    onDecision: (d) => verdicts.push([d.paragraph_index, d.verdict]),
  });
  assert.equal(r.status, "unsure");
  assert.deepEqual(r.violations, []); // the quote is no evidence
  assert.deepEqual(r.unsure, [
    {
      paragraph_index: 1,
      reason: "model reported a violation but its quote was not found (hyphen)",
    },
  ]);
  assert.deepEqual(verdicts, [
    [1, "unsure"],
    [2, "ok"],
  ]);

  const [first] = await checkLedgerIncremental([LLM_ITEM], paras, { oneshot });
  assert.deepEqual([first.status, first.checked_hashes.length], ["unsure", 1]);
  prompts.length = 0;
  const [next] = await checkLedgerIncremental(
    [{ ...asStored(LLM_ITEM, first), unsure: first.unsure }],
    paras,
    { oneshot },
  );
  assert.deepEqual([next.mode, next.checked, next.status], ["incremental", 1, "unsure"]);
  assert.match(prompts[0], /\[0\] An in-vitro assay\./);
  assert.doesNotMatch(prompts[0], /Nothing to see here/);
});

test("LLM check: a quote that differs only in case or spacing is located; an empty quote is unsure", async () => {
  const reply =
    (violations, paras = PARAS) =>
    async (req) =>
      ledgerReply(req, paras, { violations });
  const r = await checkLlmItem(LLM_ITEM, PARAS, {
    oneshot: reply([
      { paragraph_index: 7, quote: "in-vitro", reason: "r" }, // text has "In-Vitro"
      { paragraph_index: 3, quote: "activity   in-vitro\nonly", reason: "r" },
      { paragraph_index: 5, quote: "", reason: "r" },
    ]),
  });
  assert.deepEqual(
    r.violations.map((v) => [v.paragraph_index, v.match]),
    [
      [7, "In-Vitro"], // the words as the paragraph has them
      [3, "activity in-vitro only"],
    ],
  );
  assert.deepEqual(
    r.unsure.map((u) => u.paragraph_index),
    [5],
  );
  // A quote of regex characters is literal text, not a pattern.
  const abc = [{ index: 1, text: "abc", reviewed_text: null }];
  const lit = await checkLlmItem(LLM_ITEM, abc, {
    oneshot: reply([{ paragraph_index: 1, quote: ".*", reason: "r" }], abc),
  });
  assert.deepEqual([lit.status, lit.violations.length], ["unsure", 0]);
  // A paragraph with one located quote stays a violation; its bad quote is dropped.
  const mixed = await checkLlmItem(LLM_ITEM, PARAS, {
    oneshot: reply([
      { paragraph_index: 3, quote: "in-vitro only", reason: "r" },
      { paragraph_index: 3, quote: "words that are not there", reason: "r" },
    ]),
  });
  assert.deepEqual(
    [mixed.status, mixed.violations.length, mixed.unsure.length],
    ["violations", 1, 0],
  );
});

// Codex review 12: the forbid pre-check read only the replacement string, so
// a forbidden phrase formed across the edit boundary passed (forbid "novel
// substrate"; "This is a novel material."; material -> substrate).
test("forbidHitsAfterEdit: a forbidden phrase the edit forms with the words next to it", async () => {
  const { forbidHitsAfterEdit } = await import("../daemon/ledger-check.mjs");
  const rule = (pattern, flags = "", extra = {}) => ({
    id: "L1",
    kind: "forbid",
    text: "t",
    check: { type: "regex", pattern, flags, expect: "absent" },
    ...extra,
  });
  const PHRASE = [rule("novel substrate")];
  const WORD = [rule(String.raw`\bnovel\b`, "i")];
  const hit = (ledger, texts, find, text, kind = "replace") =>
    forbidHitsAfterEdit(ledger, texts, { find, text, kind });

  // Created across the boundary: blocked.
  const created = [
    // the Codex reproduction
    [PHRASE, ["This is a novel material."], "material", "substrate", "replace"],
    // pure insertion after an anchor
    [PHRASE, ["This is a novel."], "novel", " substrate", "insert"],
    // partial replace: only the tail of the phrase is new
    [PHRASE, ["It used a novel carbon source."], "carbon source", "substrate mix", "replace"],
    // a deletion that joins two words into the phrase
    [PHRASE, ["It used a novel bright substrate."], "bright ", "", "delete"],
    // a deletion that turns a longer word into the forbidden one
    [WORD, ["The novelty is clear."], "ty", "", "delete"],
    // regex with flags, whitespace class across the boundary
    [
      [rule(String.raw`NOVEL\s+substrate`, "i")],
      ["A Novel material."],
      "material",
      "Substrate",
      "replace",
    ],
    // `find` occurs in two paragraphs (paragraph_hint picks one): either may be the one
    [PHRASE, ["A plain material.", "A novel material."], "material", "substrate", "replace"],
  ];
  for (const [ledger, texts, find, text, kind] of created) {
    assert.deepEqual(hit(ledger, texts, find, text, kind), ["L1"], `${texts[0]} / ${find}`);
  }

  // Not created by this edit: passes.
  const kept = [
    // a violation elsewhere in the paragraph, unrelated edit
    [PHRASE, ["A novel substrate grows. The medium was warm."], "warm", "hot", "replace"],
    // a violation right next to the edit that the edit does not touch
    [WORD, ["A novel material."], "material", "substrate", "replace"],
    // the same rule already matched at this place before the edit
    [[rule(String.raw`novel \w+`)], ["A novel material."], "material", "substrate", "replace"],
    // the edit removes the violation
    [PHRASE, ["A novel substrate."], "novel substrate", "new substrate", "replace"],
    [PHRASE, ["A novel substrate."], "substrate", "material", "replace"],
    [PHRASE, ["A novel substrate."], " substrate", "", "delete"],
    // `find` is not in the known paragraphs: nothing to say here
    [PHRASE, ["Another paragraph."], "material", "substrate", "replace"],
    [PHRASE, [], "material", "substrate", "replace"],
  ];
  for (const [ledger, texts, find, text, kind] of kept) {
    assert.deepEqual(hit(ledger, texts, find, text, kind), [], `${texts[0]} / ${find}`);
  }

  // Same rule filter as forbidHits: forbid, regex, expect absent, manuscript scope.
  const P = ["This is a novel material."];
  const others = [
    rule("novel substrate", "", { kind: "instruction" }),
    rule("novel substrate", "", { scope: "response" }),
    { id: "L1", kind: "forbid", text: "t", check: { type: "llm" } },
    { id: "L1", kind: "forbid", text: "t", check: { type: "regex", pattern: "(", flags: "" } },
    // A must-stay ("present") rule is not pre-checked: deleting its phrase
    // directly is not blocked either, the ledger check reports it afterwards.
    rule("This is a novel material", "", {
      kind: "instruction",
      check: { type: "regex", pattern: "This is a novel material", flags: "", expect: "present" },
    }),
  ];
  for (const item of others) assert.deepEqual(hit([item], P, "material", "substrate"), []);
  console.log(`forbidHitsAfterEdit corpus: ${created.length} created, ${kept.length} not created`);
});

// Codex review 13: an entry whose paragraph number is not in the batch was
// dropped and the whole batch stored as ok, hashes cached. With one verdict
// per paragraph such a reply is no result at all (H4).
test("LLM check: a reply about a paragraph that was not asked about fails the batch, and nothing of it is cached", async () => {
  const stray = (j) => ({
    ...j,
    verdicts: [...j.verdicts, [7, "violation"]],
    violations: [{ paragraph_id: 7, quote: "a novel substrate", reason: "r" }],
  });
  const verdicts = [];
  const r = await checkLlmItem(LLM_ITEM, P3, {
    oneshot: replies(stray).oneshot,
    onDecision: (d) => verdicts.push([d.paragraph_index, d.verdict]),
  });
  assert.deepEqual([r.status, r.violations], ["error", []]);
  assert.match(r.error, /paragraph 7.*not in the batch/);
  assert.deepEqual(verdicts, [
    [0, "error"],
    [4, "error"],
    [6, "error"],
  ]);
  const [first] = await checkLedgerIncremental([LLM_ITEM], P3, { oneshot: replies(stray).oneshot });
  assert.deepEqual([first.status, first.checked_hashes], ["error", null]);
  // The next check reads all three again, in full.
  const [next] = await checkLedgerIncremental([asStored(LLM_ITEM, first)], P3, {
    oneshot: async (req) => ledgerReply(req, P3),
  });
  assert.deepEqual([next.mode, next.checked, next.status], ["full", 3, "ok"]);
});

// Class audit (Codex review 13): a reply that cannot be fully validated never
// gives a cached pass. Each row is one way a ledger check can end.
test("LLM check audit: no malformed or failed reply leaves checked_hashes with an unjudged paragraph", async () => {
  const paras = P3.slice(0, 2);
  const good = (marks) => async (req) => ledgerReply(req, paras, marks);
  const cases = [
    // [name, oneshot, expected status, expected cached hash count]
    [
      "model call failed (limit, refusal, timeout)",
      async () => ({ ok: false, error: "x" }),
      "error",
      null,
    ],
    // (In a batch of two the first row leaves no paragraph out.)
    ...BAD_REPLIES.slice(1).map(([name, bad]) => [
      name,
      async (req) => asResult(bad(ledgerJson(req, paras), req)),
      "error",
      null,
    ]),
    [
      "quote not in its paragraph",
      good({ violations: [{ paragraph_index: 4, quote: "cold", reason: "r" }] }),
      "unsure",
      1,
    ],
    [
      "empty quote",
      good({ violations: [{ paragraph_index: 4, quote: "", reason: "r" }] }),
      "unsure",
      1,
    ],
    ["verdict unsure", good({ unsure: [{ paragraph_index: 4, reason: "?" }] }), "unsure", 1],
    [
      "a located violation",
      good({ violations: [{ paragraph_index: 4, quote: "warm", reason: "r" }] }),
      "violations",
      2,
    ],
    ["a clean pass", good(), "ok", 2],
  ];
  for (const [name, oneshot, status, cached] of cases) {
    const [o] = await checkLedgerIncremental([LLM_ITEM], paras, { oneshot });
    assert.equal(o.status, status, name);
    assert.equal(o.checked_hashes === null ? null : o.checked_hashes.length, cached, name);
  }
  // A later batch that fails does not keep the earlier batch as checked.
  let n = 0;
  const [half] = await checkLedgerIncremental([LLM_ITEM], paras, {
    maxChars: 1,
    oneshot: async (req) => (n++ ? { ok: false, error: "x" } : ledgerReply(req, paras)),
  });
  assert.deepEqual([half.status, half.checked_hashes], ["error", null]);
  // Too many batches, or no runner: an error, nothing cached.
  const [many] = await checkLedgerIncremental([LLM_ITEM], paras, {
    maxChars: 1,
    maxBatches: 1,
    oneshot: good(),
  });
  assert.deepEqual([many.status, many.checked_hashes], ["error", null]);
  const [none] = await checkLedgerIncremental([LLM_ITEM], paras, {});
  assert.deepEqual([none.status, none.checked_hashes], ["error", null]);
  // A regex item whose pattern does not compile.
  const badRegex = { ...IN_VITRO, check: { type: "regex", pattern: "(", flags: "" } };
  const [re] = await checkLedgerIncremental([badRegex], paras);
  assert.deepEqual([re.status, re.checked_hashes], ["error", null]);
  console.log(`ledger reply audit: ${cases.length} ways a check can end`);
});
