// The checks of the Tools tab (daemon/aux-service.mjs format_check): format
// rules of the profile and count rules of the ledger, report only.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createAuxService, AUX_JOBS, formatRulesOn } from "../daemon/aux-service.mjs";
import {
  emptyRevision,
  saveRevision,
  revisionPathFor,
  addLedgerItem,
  setProfile,
} from "../daemon/revision-store.mjs";
import { formatParagraphs } from "../taskpane/shared/format-runs.js";

const ENV = { CC_OFFICE_WSL_PATHS: "0", CC_OFFICE_CONFIG_DIR: "/nonexistent-config" };
const words = (n) => Array.from({ length: n }, (_, i) => `w${i}`).join(" ");
const BODY =
  "The cultures were grown in shaken flasks at thirty degrees for two days and then sampled every six hours until the end of the run.";

// A manuscript with planted violations: an abstract of 12 words (limit 10), a
// "p" that is not italic, one body paragraph without the indent of the
// others, and an empty table.
const TEXTS = [
  "Abstract",
  words(12),
  "Keywords: methane, growth, flask",
  "Methods",
  BODY,
  `${BODY} Again.`,
  `${BODY} Thrice.`,
  `${BODY} And a fourth time.`,
  `${BODY} The fifth is flush.`,
  "The difference was significant (p < 0.05).",
  "",
];
const HEADINGS = [
  { index: 0, level: 1, text: "Abstract" },
  { index: 3, level: 1, text: "Methods" },
];
const snapshot = () => ({
  paragraphs: TEXTS.map((text, index) => ({ index, id: `P${index}`, text, w_ins: 0, w_del: 0 })),
  comments: [],
});
const ooxml = (text) =>
  `<pkg:part pkg:name="/word/document.xml"><w:p><w:r><w:t>${text.replace(/</g, "&lt;")}</w:t></w:r></w:p></pkg:part>`;
const format = () =>
  formatParagraphs(
    TEXTS.map((text, i) => ({
      ooxml: ooxml(text),
      first_line: i >= 4 && i <= 7 ? 36 : 0,
      left: 0,
      is_list: false,
      table_level: i === 10 ? 1 : 0,
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

async function setup({ overrides, rules = [] } = {}) {
  const dir = await mkdtemp(join(tmpdir(), "fcheck-"));
  const doc = join(dir, "paper.docx");
  await writeFile(doc, "docx");
  let rev = emptyRevision(doc);
  for (const item of rules) rev = addLedgerItem(rev, item).revision;
  if (overrides) rev = setProfile(rev, { overrides });
  await saveRevision(revisionPathFor(doc), rev);
  const service = createAuxService({
    getContext: () => ({ activeDoc: doc }),
    log: () => {},
    env: ENV,
  });
  async function check(msg) {
    let reply;
    await service.handlers.format_check(
      { request_id: "q", ...msg },
      (r) => (reply = r),
      "k",
      "word",
    );
    return reply;
  }
  const saved = () => readFile(revisionPathFor(doc), "utf8");
  return { dir, check, saved, done: () => rm(dir, { recursive: true }) };
}

const ABSTRACT_LIMIT = {
  kind: "instruction",
  text: "The abstract must not exceed 10 words.",
  scope: "manuscript",
  check: {
    type: "count",
    target: "section_words",
    section: "^(?:[\\d.]+\\s+)?(?:abstract|summary)\\b",
    max: 10,
  },
  source: { kind: "guideline", sentence: "The abstract must not exceed 10 words." },
};
const KEYWORDS = {
  kind: "instruction",
  text: "Provide 3 to 6 keywords.",
  scope: "manuscript",
  check: { type: "count", target: "keywords", min: 3, max: 6 },
};
const NOT_A_COUNT = { kind: "forbid", text: "novel", scope: "manuscript", check: { type: "llm" } };

test("format_check answers at once (it is not a job) and stores nothing", async () => {
  assert.equal(AUX_JOBS.has("format_check"), false);
  const t = await setup({ rules: [ABSTRACT_LIMIT] });
  try {
    const before = await t.saved();
    const r = await t.check({ snapshot: snapshot(), headings: HEADINGS, format: format() });
    assert.equal(r.type, "format_check_result");
    assert.equal(r.ok, true);
    assert.equal(await t.saved(), before);
  } finally {
    await t.done();
  }
});

test("the planted violations are found: italics, indent, empty table, word limit", async () => {
  const t = await setup({
    overrides: { italics: { statistics: true }, first_line_indent: "consistent" },
    rules: [ABSTRACT_LIMIT, KEYWORDS, NOT_A_COUNT],
  });
  try {
    const r = await t.check({ snapshot: snapshot(), headings: HEADINGS, format: format() });
    assert.deepEqual(r.format.rules, ["statistics symbols in italics", "first-line indent"]);
    assert.equal(r.format.error, null);
    assert.deepEqual(
      r.format.findings.map((f) => [f.paragraph_index, f.rule, f.find ?? null]),
      [
        [8, "indent", null],
        [9, "stat", "p"],
        [10, "empty_table", null],
      ],
    );
    // Count rules only; the model-checked rule is not in the report.
    assert.deepEqual(
      r.counts.map((c) => [c.id, c.status, c.text]),
      [
        ["L1", "violations", ABSTRACT_LIMIT.text],
        ["L2", "ok", KEYWORDS.text],
      ],
    );
    assert.deepEqual(r.counts[0].violations, [
      { paragraph_index: 0, start: null, match: null, count: 12, min: null, max: 10 },
    ]);
    assert.equal(r.counts[1].count, 3);
  } finally {
    await t.done();
  }
});

test("with the generic profile no style rule runs: only the empty table is reported", async () => {
  const t = await setup();
  try {
    const r = await t.check({ snapshot: snapshot(), headings: HEADINGS, format: format() });
    assert.deepEqual(r.format.rules, []);
    assert.deepEqual(
      r.format.findings.map((f) => f.rule),
      ["empty_table"],
    );
    assert.deepEqual(r.counts, []);
  } finally {
    await t.done();
  }
});

test("formatting that does not line up with the text is not checked and said so; the counts still run", async () => {
  const t = await setup({ overrides: { italics: { statistics: true } }, rules: [ABSTRACT_LIMIT] });
  try {
    const short = format();
    short.paragraphs.pop();
    let r = await t.check({ snapshot: snapshot(), headings: HEADINGS, format: short });
    assert.match(r.format.error, /read from 10 paragraphs and the text from 11/);
    assert.deepEqual(r.format.findings, []);
    assert.equal(r.counts[0].status, "violations");
    r = await t.check({ snapshot: snapshot(), headings: HEADINGS });
    assert.match(r.format.error, /read from 0 paragraphs/);
    const unmatched = { ...format(), tables_read: false };
    r = await t.check({ snapshot: snapshot(), headings: HEADINGS, format: unmatched });
    assert.match(r.format.error, /tables could not be matched/);
    assert.ok(r.format.findings.some((f) => f.rule === "stat"));
    const none = await t.check({ headings: HEADINGS, format: format() });
    assert.deepEqual([none.ok, none.error], [false, "The request holds no document snapshot."]);
  } finally {
    await t.done();
  }
});

test("a heading the snapshot does not confirm is not used: the count rule says it found no section", async () => {
  const t = await setup({ rules: [ABSTRACT_LIMIT] });
  try {
    const r = await t.check({
      snapshot: snapshot(),
      headings: [{ index: 0, level: 1, text: "Something else" }],
      format: format(),
    });
    assert.equal(r.counts[0].status, "error");
    assert.match(r.counts[0].error, /no section matches/);
  } finally {
    await t.done();
  }
});

test("the names of the format rules a profile turns on", () => {
  assert.deepEqual(formatRulesOn({}), []);
  assert.deepEqual(
    formatRulesOn({
      italics: { statistics: true, genera: ["Methylococcus"], genes: [] },
      subscripts: ["CO2"],
      first_line_indent: "indented",
    }),
    [
      "statistics symbols in italics",
      "species names in italics",
      "formula digits in subscript",
      "first-line indent",
    ],
  );
});
