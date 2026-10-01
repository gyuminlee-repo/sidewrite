// The fake manuscript of the stage 4 acceptance run, read the way the pane
// reads it (taskpane/shared/format-runs.js) and checked by the daemon's rules
// with the profile of the run: exactly the planted findings come out.
import { test } from "node:test";
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { formatParagraphs } from "../taskpane/shared/format-runs.js";
import { findFormatIssues } from "../daemon/format-check.mjs";
import { checkCountItem } from "../daemon/count-check.mjs";
import { withDefaults } from "../daemon/journal-profile.mjs";
import { S4_PLANTED, formatEvidence, countEvidence } from "../scripts/e2e/judge.mjs";
import { checkRows } from "../taskpane/shared/tools-view.js";

const here = (p) => fileURLToPath(new URL(p, import.meta.url));
const SCRIPT = here("../scripts/e2e/make_aux_fixture.py");
const python = (args) =>
  JSON.parse(execFileSync("python3", [SCRIPT, ...args], { encoding: "utf8", maxBuffer: 64 << 20 }));

const rows = python(["--dump", here("./fixtures/aux_manuscript.docx")]);
const profile = withDefaults(python(["--profile"]));
const format = formatParagraphs(
  rows.map((p) => ({
    ooxml: `<pkg:part pkg:name="/word/document.xml">${p.xml}</pkg:part>`,
    first_line: p.first_line,
    left: 0,
    is_list: false,
    table_level: p.table_level,
  })),
  [
    {
      level: 1,
      rows: 2,
      values: [
        ["", ""],
        ["", ""],
      ],
    },
  ],
);
const paragraphs = rows.map((p, index) => ({ index, text: p.text }));
const headings = rows.flatMap((p, index) => (p.heading ? [{ index, level: 1, text: p.text }] : []));

test("the runs read from the fixture are its text", () => {
  assert.equal(format.tables_read, true);
  for (const [i, p] of format.paragraphs.entries())
    assert.equal(p.runs.map((r) => r.text).join(""), rows[i].text);
});

test("the format rules of the run's profile find exactly what the fixture plants", () => {
  const findings = findFormatIssues(format.paragraphs, profile, { headings });
  assert.deepEqual(findings.map((f) => f.rule).sort(), [
    "empty_table",
    "indent",
    "species",
    "stat",
    "subscript",
  ]);
  // As the acceptance run judges the rows of the tab.
  const abstract = {
    id: "L1",
    text: "The abstract must not exceed 250 words.",
    check: {
      type: "count",
      target: "section_words",
      section: "^(?:[\\d.]+\\s+)?(?:abstract|summary)\\b",
      max: 250,
    },
  };
  const counts = [
    { text: abstract.text, ...checkCountItem(abstract, paragraphs, { headings, profile }) },
  ];
  assert.equal(counts[0].status, "violations");
  const tab = {
    checks: checkRows({ format: { rules: ["x"], error: null, findings }, counts }).map(
      (r) => r.text,
    ),
  };
  const before = paragraphs;
  const ev = formatEvidence({ check: tab, before, after: structuredClone(before) });
  assert.deepEqual([ev.pass, ev.findings, ev.kinds], [true, 5, S4_PLANTED]);
  const count = countEvidence({ check: tab, rule: abstract });
  assert.deepEqual([count.pass, count.found, count.max], [true, 270, 250]);
});
