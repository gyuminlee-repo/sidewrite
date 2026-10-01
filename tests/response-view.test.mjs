import { test } from "node:test";
import assert from "node:assert/strict";
import {
  bytesToBase64,
  pdfTooLarge,
  PDF_FRAME_LIMIT,
  STALE_LABELS,
  staleText,
  translationMissing,
  shownCitations,
  citationText,
  checkLine,
  checksSummary,
  canExport,
  proposalRows,
  overridesFromForm,
  openRequirements,
  undecidedChanges,
  undecidedText,
  nextStep,
  exportMessage,
} from "../taskpane/shared/response-view.js";
import { locationLine } from "../daemon/response-lines.mjs";
import { withDefaults } from "../daemon/journal-profile.mjs";

const P = withDefaults({}).response_letter;

const L = {
  location_line: "[Revised in {section}, {lines}]",
  removed_line: "[Removed from {section}, {lines}]",
  line_one: "line {a}",
  line_range: "lines {a}-{b}",
};

test("base64 of bytes in chunks equals Buffer's", () => {
  const b = Uint8Array.from({ length: 20000 }, (_, i) => i % 251);
  assert.equal(bytesToBase64(b), Buffer.from(b).toString("base64"));
  assert.equal(bytesToBase64(new Uint8Array(0)), "");
  // A length that is no multiple of the chunk size, and one byte.
  const odd = Uint8Array.from({ length: 8193 }, (_, i) => 255 - (i % 256));
  assert.equal(bytesToBase64(odd), Buffer.from(odd).toString("base64"));
  assert.equal(bytesToBase64(Uint8Array.of(0)), "AA==");
});

test("a PDF that cannot fit one WebSocket frame is refused before sending", () => {
  assert.equal(PDF_FRAME_LIMIT, 100 * 1024 * 1024);
  assert.equal(pdfTooLarge(10), false);
  assert.equal(pdfTooLarge(100 * 1024 * 1024), true);
  assert.deepEqual(
    [pdfTooLarge(PDF_FRAME_LIMIT - 4096), pdfTooLarge(PDF_FRAME_LIMIT - 4095)],
    [false, true],
  );
});

test("staleText", () => {
  assert.equal(staleText([]), "");
  assert.equal(staleText(undefined), "");
  assert.equal(
    staleText(["quote_changed", "lines_old"]),
    `${STALE_LABELS.quote_changed}; ${STALE_LABELS.lines_old}`,
  );
  assert.equal(STALE_LABELS.lines_old, "Line numbers are older than the manuscript");
  // A reason this pane does not know is shown as it is, not dropped.
  assert.equal(staleText(["other"]), "other");
});

test("shownCitations: what a card shows of its citations, and nothing else of them", () => {
  const cit = {
    change_ids: ["R1.1-1"],
    kind: "revised",
    section: "2. Methods",
    paragraph_index: 4,
    quote: "Cells were kept at 30 °C.",
    lines: { start: 9, end: 9 },
    lines_pdf: "sha",
    problem: null,
  };
  assert.deepEqual(shownCitations({ response: { citations: [cit] } }), [
    {
      kind: "revised",
      quote: "Cells were kept at 30 °C.",
      paragraph_index: 4,
      change_ids: ["R1.1-1"],
    },
  ]);
  for (const none of [{ response: null }, { response: { citations: [] } }, {}, undefined])
    assert.deepEqual(shownCitations(none), []);
});

test("translationMissing: only with the second language on, an answer and no translation of it", () => {
  const on = { gloss_enabled: true, gloss_lang: "German" };
  const c = (response) => ({ ref: "R1.1", response });
  const line = "No German translation of this answer; the letter prints the answer without one.";
  assert.equal(translationMissing(c({ text: "We agree.", gloss: "" }), on), line);
  assert.equal(translationMissing(c({ text: "We agree." }), on), line);
  assert.equal(translationMissing(c({ text: "We agree.", gloss: " \n" }), on), line);
  assert.equal(translationMissing(c({ text: "We agree.", gloss: "Einverstanden." }), on), "");
  assert.equal(translationMissing(c({ text: "  ", gloss: "" }), on), "");
  assert.equal(translationMissing(c(null), on), "");
  // Off, no language named, or no profile read yet.
  for (const off of [{ gloss_enabled: false, gloss_lang: "German" }, { gloss_enabled: true }, null])
    assert.equal(translationMissing(c({ text: "We agree.", gloss: "" }), off), "");
});

test("citationText", () => {
  assert.deepEqual(
    citationText(
      { kind: "revised", section: "3. Results", quote: "Q.", lines: { start: 4, end: 6 } },
      L,
    ),
    { line: "[Revised in 3. Results, lines 4-6]", quote: "Q." },
  );
  assert.deepEqual(
    citationText({ kind: "revised", section: "3. Results", quote: "Q.", lines: null }, L),
    { line: "[Revised in 3. Results]", quote: "Q." },
  );
  assert.deepEqual(
    citationText({ kind: "removed", section: "1. Introduction", quote: "", lines: null }, L),
    { line: "[Removed from 1. Introduction]", quote: null },
  );
  assert.deepEqual(
    citationText({ kind: "revised", section: null, quote: "", problem: "unlocated" }, L),
    { line: "Not located: unlocated", quote: null },
  );
});

test("citationText fills a frame as the daemon does (response-lines.mjs), on four inputs", () => {
  const cases = [
    { kind: "revised", section: "3. Results", quote: "Q.", lines: { start: 4, end: 6 } },
    { kind: "revised", section: "Cost of $1 {lines}", quote: "Q.", lines: { start: 7, end: 7 } },
    { kind: "removed", section: "1. Introduction", quote: "", lines: null },
    { kind: "removed", section: null, quote: "", lines: { start: 12, end: 13 } },
  ];
  for (const c of cases) {
    // With the profile's own keys and without them (the pane has defaults).
    assert.equal(citationText(c, P).line, locationLine(c, P), JSON.stringify(c));
    assert.equal(citationText(c, L).line, locationLine(c, P), JSON.stringify(c));
  }
  assert.equal(citationText(cases[3], L).line, "[Removed]");
  assert.equal(
    citationText({ kind: "revised", section: null, quote: "Q.", lines: { start: 4, end: 6 } }, L)
      .line,
    "[Revised, lines 4-6]",
  );
});

test("checkLine", () => {
  assert.equal(
    checkLine({
      id: "quote_exact",
      level: "fail",
      ref: "R1.1",
      detail: "quote is not in paragraph 3",
    }),
    "FAIL R1.1 quote_exact: quote is not in paragraph 3",
  );
  assert.equal(
    checkLine({ id: "pending_changes", level: "warn", ref: null, detail: "2 paragraph(s)" }),
    "WARN pending_changes: 2 paragraph(s)",
  );
});

test("checksSummary and canExport", () => {
  assert.equal(checksSummary({ fail: 0, warn: 0 }), "All checks passed");
  assert.equal(checksSummary({ fail: 0, warn: 1 }), "All checks passed, 1 warning");
  assert.equal(checksSummary({ fail: 2, warn: 0 }), "2 failed");
  assert.equal(checksSummary({ fail: 2, warn: 3 }), "2 failed, 3 warnings");
  assert.deepEqual(canExport(null), { ok: false, label: "Run checks first" });
  assert.deepEqual(canExport({ fail: 0, warn: 2 }), { ok: true, label: "Export letter (.docx)" });
  assert.deepEqual(canExport({ fail: 1, warn: 0 }), { ok: false, label: "Export anyway" });
});

test("proposalRows: one row per changed key with its evidence", () => {
  const proposal = {
    overrides: {
      colors: { response_body: "C00000", response_revised: "0070C0" },
      response_letter: { location_line: "[Changed in {section}, {lines}]" },
    },
    evidence: {
      colors: {
        C00000: { paragraphs: 5, role: "response_body" },
        "0070C0": { paragraphs: 9, role: "response_revised", theme: ["accent1"] },
      },
    },
  };
  const profile = {
    colors: { response_body: null, response_revised: "0070C0" },
    response_letter: { location_line: "[Revised in {section}, {lines}]" },
  };
  assert.deepEqual(proposalRows(proposal, profile), [
    {
      key: "colors.response_body",
      current: "(none)",
      proposed: "C00000",
      evidence: "5 paragraphs",
    },
    {
      key: "response_letter.location_line",
      current: "[Revised in {section}, {lines}]",
      proposed: "[Changed in {section}, {lines}]",
      evidence: "",
    },
  ]);
  assert.deepEqual(proposalRows({ overrides: {}, evidence: {} }, profile), []);
  // No profile loaded: every proposed key is a change. One paragraph is singular.
  assert.deepEqual(
    proposalRows(
      {
        overrides: { colors: { response_kr: "7F7F7F" } },
        evidence: { colors: { "7F7F7F": { paragraphs: 1 } } },
      },
      null,
    ),
    [{ key: "colors.response_kr", current: "(none)", proposed: "7F7F7F", evidence: "1 paragraph" }],
  );
});

test("overridesFromForm: empty text removes the key", () => {
  assert.deepEqual(overridesFromForm({ lang: " German ", rationaleLang: "", gloss: true }), {
    rationale_lang: null,
    response_letter: { lang: "German", gloss_enabled: true },
  });
  assert.deepEqual(
    overridesFromForm({ lang: "", rationaleLang: "Korean", gloss: false, glossLang: " Korean " }),
    {
      rationale_lang: "Korean",
      response_letter: { lang: null, gloss_enabled: false, gloss_lang: "Korean" },
    },
  );
  assert.equal(overridesFromForm({ glossLang: "" }).response_letter.gloss_lang, null);
});

const COMMENT = {
  declined: null,
  requirements: [
    { id: "q1", text: "a" },
    { id: "q2", text: "b" },
    { id: "q3", text: "c" },
    { id: "q4", text: "d" },
  ],
  changes: [
    { state: "accepted", rationale: { requirement_ids: ["q1"] } },
    { state: "pending", rationale: { requirement_ids: ["q2"] } },
    { state: "rejected", rationale: { requirement_ids: ["q3"] } },
  ],
  response: { declined_requirements: [{ id: "q4", reason: "r" }] },
};

test("openRequirements", () => {
  assert.deepEqual(openRequirements(COMMENT), [
    { id: "q2", text: "b", pending: true },
    { id: "q3", text: "c", pending: false },
  ]);
  assert.deepEqual(openRequirements({ ...COMMENT, declined: { reason: "x" } }), []);
  assert.deepEqual(openRequirements({ requirements: [{ id: "q1", text: "a" }] }), [
    { id: "q1", text: "a", pending: false },
  ]);
});

test("undecidedChanges and undecidedText: what still waits for a decision in Word", () => {
  assert.deepEqual(undecidedChanges(COMMENT), { pending: 1, unlocated: 0 });
  assert.equal(
    undecidedText(COMMENT),
    "Decision left in Word: 1 pending. Only accepted changes are cited.",
  );
  const both = {
    changes: [{ state: "pending" }, { state: "unlocated" }, { state: "unlocated" }],
  };
  assert.deepEqual(undecidedChanges(both), { pending: 1, unlocated: 2 });
  assert.equal(
    undecidedText(both),
    "Decision left in Word: 1 pending, 2 not found. Only accepted changes are cited.",
  );
  // Nothing left: accepted, edited by the author and rejected are decisions.
  for (const state of ["accepted", "user_edited", "rejected"])
    assert.equal(undecidedText({ changes: [{ state }] }), "", state);
  assert.equal(undecidedText({}), "");
});

test("nextStep: the one thing to do next", () => {
  const answered = { ref: "R1.1", response: { text: "Done.", citations: [], stale: [] } };
  const cited = (lines, stale = []) => ({
    ref: "R1.2",
    response: { text: "Done.", stale, citations: [{ kind: "revised", quote: "Q.", lines }] },
  });
  const step = (comments, summary = null, exported = null) =>
    nextStep({ comments, response_export: exported }, summary);
  assert.deepEqual(step([]), {
    id: "split",
    text: "Next: split the review into comments on the Review tab.",
  });
  assert.equal(nextStep(null, null).id, "split");
  assert.deepEqual(step([answered, { ref: "R1.2" }, { ref: "R1.3", response: { text: " " } }]), {
    id: "draft",
    text: "Next: write or draft the 2 missing responses (or decline the comment).",
  });
  assert.equal(
    step([{ ref: "R1.2" }]).text,
    "Next: write or draft the 1 missing response (or decline the comment).",
  );
  // A declined comment needs no response.
  assert.equal(step([{ ref: "R1.2", declined: { reason: "x" } }]).id, "check");
  assert.deepEqual(step([cited(null)]), {
    id: "locate",
    text: "Next: update line numbers (Word renders the manuscript to PDF).",
  });
  assert.equal(step([cited({ start: 1, end: 2 }, ["lines_old"])]).id, "locate");
  assert.deepEqual(step([answered, cited({ start: 1, end: 2 })]), {
    id: "check",
    text: "Next: run checks.",
  });
  assert.deepEqual(step([answered], { fail: 2, warn: 0 }), {
    id: "fix",
    text: "Next: fix the 2 failed checks below, then run checks again.",
  });
  assert.equal(
    step([answered], { fail: 1, warn: 0 }).text,
    "Next: fix the 1 failed check below, then run checks again.",
  );
  assert.deepEqual(step([answered], { fail: 0, warn: 1 }), {
    id: "export",
    text: "Next: export the letter.",
  });
  // The manuscript changed under an answer: read it again before anything else.
  assert.deepEqual(step([answered, cited({ start: 1, end: 2 }, ["quote_changed", "lines_old"])]), {
    id: "reread",
    text: "Next: read the 1 response marked out of date, then save or draft it again.",
  });
  assert.equal(
    step([cited(null, ["change_state"]), cited(null, ["quote_changed"])]).text,
    "Next: read the 2 responses marked out of date, then save or draft them again.",
  );
  // Missing answers come before a stale summary of an earlier check.
  assert.equal(step([{ ref: "R1.2" }], { fail: 0, warn: 0 }).id, "draft");
});

test("exportMessage: the path, and that an existing file was kept", () => {
  assert.equal(
    exportMessage({
      exported: true,
      path: "/m/paper.response.docx",
      wanted: "/m/paper.response.docx",
    }),
    "Exported: /m/paper.response.docx",
  );
  assert.equal(
    exportMessage({
      exported: true,
      path: "/m/paper.response-2.docx",
      wanted: "/m/paper.response.docx",
    }),
    "Exported under a new name: /m/paper.response-2.docx (/m/paper.response.docx was already there or open in Word and was kept).",
  );
  assert.equal(
    exportMessage({
      exported: true,
      paths: ["/m/paper.response.R1.txt", "/m/paper.response.R2-2.txt"],
      wanted: ["/m/paper.response.R1.txt", "/m/paper.response.R2.txt"],
    }),
    "Exported: /m/paper.response.R1.txt, /m/paper.response.R2-2.txt (/m/paper.response.R2.txt was already there and was kept).",
  );
  assert.equal(
    exportMessage({ exported: false, summary: { fail: 2, warn: 0 } }),
    "Not exported: 2 failed. Fix them, or press Export anyway.",
  );
});

test("undecidedText: an edit Word did not answer is told on the card as well", () => {
  const c = {
    changes: [{ state: "pending" }],
    blocked: [
      { id: "B1", state: "outcome_unknown" },
      { id: "B2", state: "blocked" },
    ],
  };
  assert.equal(
    undecidedText(c),
    "Decision left in Word: 1 pending. Only accepted changes are cited. 1 edit got no answer from Word: settle it on the Review tab.",
  );
  assert.equal(
    undecidedText({
      changes: [],
      blocked: [{ state: "outcome_unknown" }, { state: "outcome_unknown" }],
    }),
    "2 edits got no answer from Word: settle them on the Review tab.",
  );
  assert.equal(undecidedText({ changes: [], blocked: [{ state: "blocked" }] }), "");
});
