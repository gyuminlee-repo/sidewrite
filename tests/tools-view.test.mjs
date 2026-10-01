// Wording of the Tools tab (taskpane/shared/tools-view.js).
import { test } from "node:test";
import assert from "node:assert/strict";
import {
  docLine,
  selectionHead,
  candidateKind,
  candidateCite,
  rulesAdded,
  fileTooLarge,
  markedLines,
  checkRows,
  formatFindingText,
  anonymizeLines,
} from "../taskpane/shared/tools-view.js";

test("the document line names the saved file, or says to save", () => {
  assert.match(
    docLine({ doc_name: "paper.docx", profile_name: "generic" }),
    /^Saved file: paper\.docx\. Profile: generic\./,
  );
  assert.match(docLine(null), /^Save the document/);
});

test("the head of a selection answer quotes the start of the selection", () => {
  assert.equal(
    selectionHead("translate", "One\n two.", { lang: "Korean", model: "sonnet" }),
    'Translation into Korean (sonnet) for: "One two."',
  );
  const long = selectionHead("ask", "x".repeat(200), { model: "haiku" });
  assert.equal(long, `Answer (haiku) for: "${"x".repeat(80)}..."`);
});

test("a rule candidate says how it is checked and cites its sentence", () => {
  const source = { kind: "guideline", sentence: "Provide 4 to 6 keywords." };
  assert.equal(
    candidateKind({ check: { type: "count", target: "keywords", min: 4, max: 6 } }),
    "Count rule: at least 4 and at most 6 keywords",
  );
  assert.equal(
    candidateKind({
      check: { type: "count", target: "section_words", section: "^abstract", max: 250 },
    }),
    "Count rule: at most 250 words in the section",
  );
  assert.equal(
    candidateKind({
      check: { type: "count", target: "section_order", order: ["abstract", "methods"] },
    }),
    "Count rule: section order abstract, methods",
  );
  assert.equal(candidateKind({ check: { type: "llm" } }), "Rule the model checks");
  assert.equal(
    candidateKind({ check: { type: "llm" }, profile: { italics: { statistics: true } } }),
    'Profile setting: {"italics":{"statistics":true}}',
  );
  assert.equal(candidateCite({ source }), 'Guideline: "Provide 4 to 6 keywords."');
});

test("what was added is counted in words", () => {
  assert.equal(
    rulesAdded({ added: 1, skipped: 0, profile_keys: [] }),
    "1 rule added to the ledger.",
  );
  assert.equal(
    rulesAdded({ added: 2, skipped: 1, profile_keys: ["italics"] }),
    "2 rules added to the ledger; 1 already there; profile setting saved: italics.",
  );
});

test("the marked copy: what was written, or the refusal as the tool said it", () => {
  assert.deepEqual(
    markedLines({
      ok: true,
      name: "paper_marked.docx",
      mode: "track_changes",
      report: { paragraphs: { n_eq: 18, n_mod: 6, n_ins: 1, n_del: 0 }, warnings: ["w1"] },
      notice: "Read from disk.",
    }),
    [
      { text: "Written: paper_marked.docx (tracked changes)." },
      { text: "Paragraphs: 6 changed, 1 added, 0 removed, 18 unchanged." },
      { text: "w1", bad: true },
      { text: "Read from disk." },
    ],
  );
  const refusal = 'paragraph 12 ("The cells were"): a text direction mark (w:dir) was added';
  assert.deepEqual(markedLines({ ok: false, error: refusal }), [
    { text: "No marked copy was written.", bad: true },
    { text: refusal, bad: true },
  ]);
  assert.equal(fileTooLarge(1024), false);
  assert.equal(fileTooLarge(80 * 1024 * 1024), true);
});

test("a format finding names its paragraph and what should be", () => {
  assert.equal(
    formatFindingText({
      paragraph_index: 9,
      rule: "stat",
      find: "p",
      context: " significant (p < 0.05). ",
    }),
    'Paragraph 9: "p" should be italic in "significant (p < 0.05)."',
  );
  assert.equal(
    formatFindingText({
      paragraph_index: 4,
      rule: "subscript",
      find: "CO2",
      context: "of CO2 was",
    }),
    'Paragraph 4: the digits of "CO2" should be subscript in "of CO2 was"',
  );
  assert.equal(
    formatFindingText({
      paragraph_index: 8,
      rule: "indent",
      found: 0,
      want: { first_line_indent: 36 },
      context: "The cultures",
    }),
    "Paragraph 8: first-line indent is 0 pt, the others have 36 pt",
  );
  assert.equal(
    formatFindingText({ paragraph_index: 10, rule: "empty_table", found: { rows: 2, cols: 2 } }),
    "Paragraph 10: empty table (2 x 2)",
  );
});

test("the rows of a check: rules on, findings with their paragraph, count results", () => {
  const rows = checkRows({
    format: {
      rules: ["statistics symbols in italics"],
      error: null,
      findings: [{ paragraph_index: 9, rule: "stat", find: "p", context: "(p < 0.05)" }],
    },
    counts: [
      {
        id: "L1",
        text: "At most 10 words.",
        status: "violations",
        violations: [{ paragraph_index: 0, count: 12, min: null, max: 10 }],
      },
      { id: "L2", text: "3 to 6 keywords.", status: "ok", count: 3, violations: [] },
      {
        id: "L3",
        text: "Order.",
        status: "violations",
        violations: [{ paragraph_index: 5, count: null, min: null, max: null }],
      },
      { id: "L4", text: "Tables.", status: "error", error: "no section", violations: [] },
    ],
  });
  assert.deepEqual(rows, [
    { text: "Format rules on: statistics symbols in italics, empty tables." },
    { text: 'Paragraph 9: "p" should be italic in "(p < 0.05)"', index: 9, bad: true },
    { text: "L1: 12 found, at most 10 allowed. Rule: At most 10 words.", index: 0, bad: true },
    { text: "L2 ok (3): 3 to 6 keywords." },
    { text: "L3: out of order. Rule: Order.", index: 5, bad: true },
    { text: "L4 could not be checked: no section. Rule: Tables.", bad: true },
  ]);
  assert.deepEqual(
    checkRows({ format: { rules: [], error: "not read", findings: [] }, counts: [] }),
    [
      { text: "No format rule is on in the profile; only empty tables are reported." },
      { text: "not read", bad: true },
      { text: "No count rule is in the ledger." },
    ],
  );
  assert.deepEqual(
    checkRows({ format: { rules: [], error: null, findings: [] }, counts: [] }).map((r) => r.text),
    [
      "No format rule is on in the profile; only empty tables are reported.",
      "Format: nothing found.",
      "No count rule is in the ledger.",
    ],
  );
});

test("the anonymized copy: the verifier's totals, what is kept, what is left", () => {
  const report = {
    total: 0,
    body_equal: true,
    markup_equal: true,
    structure_ok: true,
    categories: {
      names: { status: "cleared", count: 0, what: "creator" },
      title: { status: "kept by option", count: 1, what: "document title in core.xml" },
    },
    name_mentions: {},
  };
  assert.deepEqual(
    anonymizeLines({
      ok: true,
      name: "paper_anonymized.docx",
      report,
      notice: "Text is not changed.",
    }),
    [
      { text: "Written: paper_anonymized.docx." },
      {
        text: "Verifier: total 0 left; text unchanged: yes; tracked changes unchanged: yes; structure intact: yes.",
        bad: false,
      },
      { text: "Kept: title (1): document title in core.xml", bad: false },
      { text: "Text is not changed." },
    ],
  );
  const failed = anonymizeLines({
    ok: false,
    name: "paper_anonymized.INCOMPLETE.docx",
    error: "The copy is incomplete: still present: names (2).",
    report: {
      ...report,
      total: 2,
      body_equal: false,
      categories: { names: { status: "not cleared", count: 2, what: "creator" } },
      name_mentions: { "A. Writer": 3 },
    },
  });
  assert.deepEqual(failed, [
    {
      text: "The copy did not pass its check and was set aside as paper_anonymized.INCOMPLETE.docx. Do not submit it.",
      bad: true,
    },
    { text: "The copy is incomplete: still present: names (2).", bad: true },
    {
      text: "Verifier: total 2 left; text unchanged: NO; tracked changes unchanged: yes; structure intact: yes.",
      bad: true,
    },
    { text: "Still present: names (2): creator", bad: true },
    { text: 'The text still has the name "A. Writer" (3 times).', bad: true },
  ]);
});
