// The stage 4 flow of the acceptance runner (Tools tab): its evidence
// functions, and that what it asks of the driver and the judges exists. The
// flow itself needs Word and does not run here.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import * as J from "../scripts/e2e/judge.mjs";

const read = (p) => readFileSync(new URL(p, import.meta.url), "utf8");

test("stage 4: its feature list, with gaps, and the manuscript name", () => {
  assert.equal(J.defaultFeatures("s4"), "S4-30..S4-33,S4-37,S4-38");
  assert.deepEqual(J.featureIds(J.defaultFeatures("s4")), [
    "S4-30",
    "S4-31",
    "S4-32",
    "S4-33",
    "S4-37",
    "S4-38",
  ]);
  assert.equal(J.stageOf("S4-37,S4-38"), "s4");
  assert.equal(J.docStem("s4"), "paper");
  assert.throws(() => J.featureIds("S4-30,S3-30"), /one stage/);
  assert.throws(() => J.featureIds("S4-30,,S4-31"), /S<stage>-<n>/);
  // A plain range reads as before.
  assert.deepEqual(J.featureIds("S3-30..S3-32"), ["S3-30", "S3-31", "S3-32"]);
});

const PARAS = [{ index: 0, text: "One.", w_ins: 0, w_del: 0, tracked: [] }];
const step = (head, answer) => ({ message_error: false, selection: [head, answer] });

test("selection: answers, an unchanged document and no tool call", () => {
  const good = {
    translate: step("Translation into German (sonnet) for: ...", "Eins."),
    ask: step("Answer (sonnet) for: ...", "One."),
    before: PARAS,
    after: structuredClone(PARAS),
    log: "[tools] selection_translate chars=4 lang=German model=sonnet\n",
  };
  assert.deepEqual(J.selectionEvidence(good), {
    pass: true,
    tool_calls: 0,
    document_unchanged: true,
    answer_chars: [5, 4],
  });
  const changed = [{ ...PARAS[0], text: "One. Two." }];
  assert.equal(J.selectionEvidence({ ...good, after: changed }).pass, false);
  assert.equal(J.selectionEvidence({ ...good, before: [], after: [] }).pass, false);
  const called = J.selectionEvidence({
    ...good,
    log: "[diag] model called tool: office_tracked_edit\n",
  });
  assert.deepEqual([called.pass, called.tool_calls], [false, 1]);
  assert.equal(
    J.selectionEvidence({ ...good, log: "[agent] turn complete (success)\n" }).pass,
    false,
  );
  assert.equal(J.selectionEvidence({ ...good, ask: step("Answer", " ") }).pass, false);
  assert.equal(
    J.selectionEvidence({ ...good, ask: { message_error: true, selection: [] } }).pass,
    false,
  );
});

const GUIDELINE =
  "We welcome work.  The abstract must not exceed 250 words. Provide 4 to 6 keywords.";
const RULE = {
  id: "L1",
  text: "The abstract must not exceed 250 words.",
  check: { type: "count", target: "section_words", section: "^abstract", max: 250 },
  source: { kind: "guideline", sentence: "The abstract must not exceed 250 words." },
};

test("guideline: every candidate cites the guideline and the kept one is the one rule added", () => {
  const good = {
    extract: {
      rules: [
        { kind: "Count rule", sentence: "The abstract must not exceed 250 words." },
        { kind: "Count rule", sentence: "Provide 4 to 6 keywords." },
      ],
    },
    add: { message: "1 rule added to the ledger." },
    guideline: GUIDELINE,
    ledgerBefore: [],
    ledgerAfter: [RULE],
    pick: "abstract must not exceed",
  };
  const ev = J.guidelineEvidence(good);
  assert.deepEqual([ev.pass, ev.candidates, ev.all_cited, ev.added.length], [true, 2, true, 1]);
  assert.equal(ev.added[0].id, "L1");
  // A candidate whose sentence is not in the guideline.
  const forged = { rules: [...good.extract.rules, { kind: "x", sentence: "Use blue ink." }] };
  assert.equal(J.guidelineEvidence({ ...good, extract: forged }).pass, false);
  // Nothing added, two added, a rule without its sentence, another rule than the kept one.
  assert.equal(J.guidelineEvidence({ ...good, ledgerAfter: [] }).pass, false);
  assert.equal(
    J.guidelineEvidence({ ...good, ledgerAfter: [RULE, { ...RULE, id: "L2" }] }).pass,
    false,
  );
  assert.equal(
    J.guidelineEvidence({ ...good, ledgerAfter: [{ ...RULE, source: null }] }).pass,
    false,
  );
  assert.equal(J.guidelineEvidence({ ...good, pick: "keywords" }).pass, false);
  assert.equal(J.guidelineEvidence({ ...good, ledgerBefore: [RULE] }).pass, false);
  assert.equal(J.guidelineEvidence({ ...good, extract: { rules: [] } }).pass, false);
});

const CHECK = {
  checks: [
    "Format rules on: statistics symbols in italics, empty tables.",
    'Paragraph 6: "p" should be italic in "significant (p < 0.05) in"',
    'Paragraph 7: "Examplea fictiva" should be italic in "Cells of Examplea fictiva taken"',
    'Paragraph 8: the digits of "CO2" should be subscript in "share of CO2 in it"',
    "Paragraph 9: first-line indent is 0 pt, the others have 36 pt",
    "Paragraph 15: empty table (2 x 2)",
    "L1: 270 found, at most 250 allowed. Rule: The abstract must not exceed 250 words.",
  ],
};

test("count: the row of the adopted rule says the planted abstract is over its limit", () => {
  assert.deepEqual(J.countEvidence({ check: CHECK, rule: RULE }), {
    pass: true,
    rule_id: "L1",
    found: 270,
    max: 250,
    row: CHECK.checks[6],
  });
  const ok = { checks: ["L1 ok (120): The abstract must not exceed 250 words."] };
  assert.equal(J.countEvidence({ check: ok, rule: RULE }).pass, false);
  assert.equal(J.countEvidence({ check: CHECK, rule: undefined }).pass, false);
  assert.equal(
    J.countEvidence({ check: CHECK, rule: { ...RULE, check: { type: "llm" } } }).pass,
    false,
  );
  const other = { ...RULE, check: { ...RULE.check, max: 300 } };
  assert.equal(J.countEvidence({ check: CHECK, rule: other }).pass, false);
});

test("format: exactly the planted findings, and the document as it was", () => {
  const good = { check: CHECK, before: PARAS, after: structuredClone(PARAS) };
  const ev = J.formatEvidence(good);
  assert.deepEqual(
    [ev.pass, ev.findings, ev.kinds, ev.document_unchanged],
    [true, 5, { italic: 2, subscript: 1, indent: 1, empty_table: 1 }, true],
  );
  const missing = { checks: CHECK.checks.filter((t) => !/empty table/.test(t)) };
  assert.equal(J.formatEvidence({ ...good, check: missing }).pass, false);
  const extra = { checks: [...CHECK.checks, 'Paragraph 3: "p" should be italic in "x"'] };
  assert.equal(J.formatEvidence({ ...good, check: extra }).pass, false);
  assert.equal(J.formatEvidence({ ...good, after: [{ ...PARAS[0], w_ins: 1 }] }).pass, false);
});

test("marked copy: both modes, each equal to the script's file", () => {
  const modes = [
    { mode: "track_changes", name: "paper_marked.docx", script_exit: 0, equal: true },
    { mode: "color", name: "paper_marked_20260101-000000.docx", script_exit: 0, equal: true },
  ];
  assert.equal(J.markedEvidence(modes).pass, true);
  assert.equal(J.markedEvidence(modes.slice(0, 1)).pass, false);
  assert.equal(J.markedEvidence([modes[0], { ...modes[1], equal: false }]).pass, false);
  assert.equal(J.markedEvidence([modes[0], { ...modes[1], script_exit: 2 }]).pass, false);
  assert.equal(J.markedEvidence([modes[0], { ...modes[1], name: null }]).pass, false);
  assert.equal(J.markedEvidence(undefined).pass, false);
});

test("anonymized copy: the tab and the runner's own verifier both count 0, the text is the same", () => {
  const rows = [
    "Written: paper_anonymized.docx.",
    "Verifier: total 0 left; text unchanged: yes; tracked changes unchanged: yes; structure intact: yes.",
  ];
  const verifier = { total: 0, body_equal: true, markup_equal: true, structure_ok: true, x: 1 };
  const ev = J.anonymizeEvidence({ step: { anonymized: rows }, verifier, exists: true });
  assert.deepEqual(ev, {
    pass: true,
    verifier: { total: 0, body_equal: true, markup_equal: true, structure_ok: true },
    rows,
  });
  const no = (more) =>
    J.anonymizeEvidence({ step: { anonymized: rows }, verifier, exists: true, ...more }).pass;
  assert.equal(no({ exists: false }), false);
  assert.equal(no({ verifier: { ...verifier, total: 1 } }), false);
  assert.equal(no({ verifier: { ...verifier, body_equal: false } }), false);
  assert.equal(no({ verifier: null }), false);
  assert.equal(no({ step: { anonymized: ["The copy did not pass its check"] } }), false);
});

test("the judges of the Tools tab steps", () => {
  const j = (name, r) => J.judge(name, r).pass;
  assert.equal(
    j("open-tools", { doc: "Saved file: paper.docx. Profile: s4-acceptance. Copies..." }),
    true,
  );
  assert.equal(
    j("open-tools", { doc: "Save the document to a local folder to use the tools." }),
    false,
  );
  assert.equal(
    j("tools-translate", step("Translation into German (sonnet) for: x", "Eins.")),
    true,
  );
  assert.equal(j("tools-translate", step("Answer (sonnet) for: x", "One.")), false);
  assert.equal(j("tools-ask", step("Answer (sonnet) for: x", "One.")), true);
  assert.equal(j("tools-ask", { message_error: true, message: "failed", selection: [] }), false);
  assert.equal(
    j("tools-extract", {
      message: "3 rule candidates. Untick the ones ...",
      rules: [{ kind: "x" }],
    }),
    true,
  );
  assert.equal(
    j("tools-extract", { message: "No rule was found in the guideline text.", rules: [] }),
    false,
  );
  assert.equal(j("tools-add", { kept: 1, message: "1 rule added to the ledger." }), true);
  assert.equal(j("tools-add", { kept: 2, message: "2 rules added to the ledger." }), false);
  assert.equal(j("tools-check", { message: "6 findings. Nothing was changed.", checks: [] }), true);
  assert.equal(j("tools-check", { message: "The checks found nothing.", checks: [] }), false);
  assert.equal(
    j("tools-marked", { marked: ["Written: paper_marked.docx (tracked changes)."] }),
    true,
  );
  assert.equal(
    j("tools-marked", { marked: ["No marked copy was written.", "paragraph 12 ..."] }),
    false,
  );
  assert.equal(j("tools-anonymize", { anonymized: ["Written: paper_anonymized.docx."] }), true);
  assert.equal(j("tools-anonymize", { ok: false, error: "pane not reachable" }), false);
});

test("the stage 4 flow uses only pane steps, judges and judge exports that exist, and writes every marker of its features", () => {
  const runner = read("../scripts/e2e/run-acceptance.mjs");
  const driver = read("../scripts/e2e/pane-driver.cjs");
  const html = read("../taskpane/word/index.html");
  const flow = runner.slice(
    runner.indexOf("async function mainS4()"),
    runner.indexOf("// Close the Word this run started"),
  );
  assert.ok(flow.length > 1000);
  assert.match(runner, /s4: mainS4/);
  const names = (re, text = flow) => [...new Set([...text.matchAll(re)].map((m) => m[1]))];
  const steps = names(/\bpane\(\s*"([a-z-]+)"/g);
  assert.ok(steps.length >= 9, steps.join(","));
  for (const s of steps)
    assert.match(driver, new RegExp(`\\n  (?:async )?"?${s}"?(?:\\(page|: )`), `pane step ${s}`);
  for (const [, s, rest] of flow.matchAll(/\bnote\(\s*"([a-z-]+)",([^;]*);/g))
    if (!/judge: "none"/.test(rest)) assert.equal(typeof J.JUDGES[s], "function", `judge ${s}`);
  // Every name the runner imports from judge.mjs is exported there.
  const imported = /import \{([^}]+)\} from "\.\/judge\.mjs";/.exec(runner)[1];
  for (const name of imported
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean))
    assert.ok(name in J, `judge.mjs exports ${name}`);
  // One marker per feature of the stage's default range, named as its verify command reads it.
  assert.deepEqual(names(/\bmarker\(\s*"(s4-[a-z]+)"/g).sort(), [
    "s4-anonymize",
    "s4-count",
    "s4-format",
    "s4-marked",
    "s4-profile",
    "s4-selection",
  ]);
  // The ids the Tools steps of the driver touch are in the Word page.
  const ids = names(/["#](tools-[a-z-]+)/g, driver).filter(
    (id) => !steps.includes(id) && id !== "tools-view",
  );
  assert.ok(ids.length >= 12, ids.join(","));
  for (const id of ids) assert.ok(html.includes(`id="${id}"`), `index.html has #${id}`);
});
