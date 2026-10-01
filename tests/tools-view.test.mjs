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
