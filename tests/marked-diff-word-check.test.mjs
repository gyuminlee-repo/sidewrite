// The pure parts of the Word check of the marked-diff tool
// (scripts/e2e/marked-diff-word-lib.mjs): the selection of pairs, the
// comparison of two structures as Word reports them, the verdict on a pair.

import { test } from "node:test";
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import {
  KINDS,
  PAIRS,
  compareStructure,
  exitCode,
  jobsFor,
  judgePair,
  renderRows,
  show,
  uncovered,
} from "../scripts/e2e/marked-diff-word-lib.mjs";

const HELPER = fileURLToPath(new URL("../scripts/e2e/marked_diff_pairs.py", import.meta.url));

const shape = (paragraphs, tables = [], bookmarks = []) => ({ paragraphs, tables, bookmarks });
const job = (s, extra = {}) => ({
  ok: true,
  dirty_on_open: false,
  revisions_left: 0,
  raw: { paragraphs: s.paragraphs.length, tables: s.tables.length },
  shape: s,
  ...extra,
});
const CLEAN = shape(
  ["One.\r", "Two new.\r"],
  [{ rows: 1, columns: 2, cells: ["a\r\u0007", "b\r\u0007"] }],
  [{ name: "mark1", text: "Two" }],
);
const BASE = shape(["One.\r", "Two.\r", "\r"], CLEAN.tables);
const good = () => ({
  tool: { track: 0, color: 0 },
  accept: job(CLEAN, { raw: { paragraphs: 4, tables: 1 } }),
  reject: job({ ...BASE, bookmarks: CLEAN.bookmarks }),
  clean: job(CLEAN),
  base: job(BASE),
  color: job(shape([]), { raw: { paragraphs: 4, tables: 1 } }),
});

test("the selection is fixed, 30 to 45 pairs, and covers every kind of change", () => {
  assert.ok(PAIRS.length >= 30 && PAIRS.length <= 45, String(PAIRS.length));
  for (const k of ["cell-merge", "picture", "field"]) assert.ok(KINDS.includes(k), k);
  assert.equal(new Set(PAIRS.map((p) => p.name)).size, PAIRS.length);
  assert.deepEqual(uncovered(), []);
  for (const p of PAIRS) for (const k of p.covers) assert.ok(KINDS.includes(k), `${p.name}: ${k}`);
  assert.deepEqual(uncovered(PAIRS.filter((p) => !p.covers.includes("tab"))), ["tab"]);
});

test("every selected pair is in the corpus and is one the tool is expected to write", () => {
  const corpus = JSON.parse(execFileSync("python3", [HELPER, "--list"], { encoding: "utf8" }));
  const refused = new Map(corpus.map((c) => [c.name, c.refused]));
  for (const p of PAIRS) {
    assert.ok(refused.has(p.name), `not in the corpus: ${p.name}`);
    assert.equal(refused.get(p.name), false, `expected to be refused: ${p.name}`);
  }
});

test("two equal structures do not differ; bookmarks count only when asked", () => {
  assert.equal(compareStructure(CLEAN, structuredClone(CLEAN)), null);
  const other = { ...CLEAN, bookmarks: [] };
  assert.match(compareStructure(other, CLEAN), /^bookmark mark1: expected "Two", missing$/);
  assert.equal(compareStructure(other, CLEAN, { bookmarks: false }), null);
  assert.match(
    compareStructure({ ...CLEAN, bookmarks: [{ name: "mark1", text: "Two new" }] }, CLEAN),
    /^bookmark mark1: expected "Two", got "Two new"$/,
  );
  assert.match(
    compareStructure(
      { ...CLEAN, bookmarks: [...CLEAN.bookmarks, { name: "x", text: "y" }] },
      CLEAN,
    ),
    /^bookmark x: not expected/,
  );
});

test("the first difference names the paragraph and shows 80 characters of both texts", () => {
  const long = "x".repeat(200);
  const d = compareStructure(
    shape(["One.\r", `${long}\r`]),
    shape(["One.\r", `A page break\u000c${long}\r`]),
  );
  assert.equal(
    d,
    `paragraph 2: expected ${JSON.stringify(`A page break\u000c${long}`.slice(0, 80))}, got ${JSON.stringify(long.slice(0, 80))}`,
  );
  assert.ok(d.includes("\\f"), "a control character is shown, not printed");
  assert.equal(show("a\u000bb"), '"a\\u000bb"');
  // A break of another kind is a difference: U+000C is a page break, U+000E a column break.
  assert.match(compareStructure(shape(["a\u000cb\r"]), shape(["a\u000eb\r"])), /^paragraph 1: /);
});

test("a paragraph too many or too few is named with its place and its text", () => {
  assert.equal(
    compareStructure(shape(["One.\r", "\r", "Two.\r"]), shape(["One.\r", "Two.\r"])),
    'paragraph 2: expected "Two.\\r", got "\\r"',
  );
  assert.equal(
    compareStructure(shape(["One.\r", "Two.\r", "\r"]), shape(["One.\r", "Two.\r"])),
    'paragraph count: expected 2, got 3; paragraph 3 is extra: "\\r"',
  );
  assert.equal(
    compareStructure(shape(["One.\r"]), shape(["One.\r", "Two.\r"])),
    'paragraph count: expected 2, got 1; paragraph 2 is missing: "Two.\\r"',
  );
});

test("tables are compared by their number, rows, columns and cell texts", () => {
  const t = (rows, columns, cells) => shape(["p\r"], [{ rows, columns, cells }]);
  assert.equal(compareStructure(shape(["p\r"]), t(1, 1, ["a"])), "table count: expected 1, got 0");
  assert.equal(
    compareStructure(t(2, 1, ["a", "b"]), t(1, 1, ["a"])),
    "table 1 rows: expected 1, got 2",
  );
  assert.equal(
    compareStructure(t(1, 2, ["a", "b"]), t(1, 1, ["a"])),
    "table 1 columns: expected 1, got 2",
  );
  assert.equal(
    compareStructure(t(1, 1, ["x"]), t(1, 1, ["a"])),
    'table 1 cell 1: expected "a", got "x"',
  );
});

test("tables are compared by where their cells stand, so a merge that is left behind shows", () => {
  const t = (uniform, at) =>
    shape(["p\r"], [{ rows: 2, columns: 2, uniform, cells: at.map(() => "a\r\u0007"), at }]);
  const four = t(true, ["1,1", "1,2", "2,1", "2,2"]);
  assert.equal(compareStructure(four, four), null);
  // A cell merged down the first column: Word reports three cells.
  assert.equal(
    compareStructure(t(false, ["1,1", "1,2", "2,2"]), four),
    "table 1 uniform (no merged cells): expected true, got false",
  );
  // Merged across in one, merged down in the other: neither is uniform.
  assert.equal(
    compareStructure(t(false, ["1,1", "2,1", "2,2"]), t(false, ["1,1", "1,2", "2,2"])),
    'table 1 cell position 2: expected "1,2", got "2,1"',
  );
  // A report without these fields (an older script) compares as before.
  const old = shape(["p\r"], [{ rows: 2, columns: 2, cells: ["a", "b", "c", "d"] }]);
  assert.equal(compareStructure(old, old), null);
});

test("paragraphs are compared by their pictures and fields too", () => {
  const p = (inline, fields) => ({ ...shape(["One.\r", "Two.\r"]), inline, fields });
  assert.equal(compareStructure(p([0, 1], [0, 0]), p([0, 1], [0, 0])), null);
  assert.equal(
    compareStructure(p([0, 1], [0, 0]), p([0, 0], [0, 0])),
    "paragraph 2 pictures: expected 0, got 1",
  );
  assert.equal(
    compareStructure(p([0, 0], [0, 0]), p([0, 0], [1, 0])),
    "paragraph 1 fields: expected 1, got 0",
  );
  // The texts are told first: a count that differs because a paragraph is missing is not a second finding.
  assert.match(
    compareStructure({ ...shape(["One.\r"]), inline: [0] }, p([0, 1], [0, 0])),
    /^paragraph count: /,
  );
});

test("a pair fails when rejecting leaves a picture or a merged cell", () => {
  const a = good();
  a.reject = job({ ...BASE, inline: [0, 1, 0] });
  a.base = job({ ...BASE, inline: [0, 0, 0] });
  assert.equal(
    judgePair(a).first,
    "reject-all is not the baseline: paragraph 2 pictures: expected 0, got 1",
  );
  const b = good();
  b.reject = job({ ...BASE, tables: [{ ...CLEAN.tables[0], uniform: false }] });
  b.base = job({ ...BASE, tables: [{ ...CLEAN.tables[0], uniform: true }] });
  assert.match(judgePair(b).first, /^reject-all is not the baseline: table 1 uniform/);
});

test("a pair passes when Word opens everything and both sides match", () => {
  assert.deepEqual(judgePair(good()), { pass: true, failures: [], first: "" });
});

test("a file Word does not open fails the pair, and the error is the first line", () => {
  const r = good();
  r.accept = { ok: false, error: "open: The file is corrupt and cannot be opened." };
  const v = judgePair(r);
  assert.equal(v.pass, false);
  assert.match(v.first, /^Word does not open the track-changes output: open: The file is corrupt/);
  const c = good();
  c.color = { ok: false, error: "open: unreadable content" };
  assert.match(judgePair(c).first, /^Word does not open the colour output/);
  const missing = good();
  delete missing.reject;
  assert.match(judgePair(missing).first, /the job did not run/);
});

test("an input Word cannot read is named as a fault of the test pair", () => {
  const r = good();
  r.clean = { ok: false, error: "open: nope" };
  const v = judgePair(r);
  assert.equal(v.pass, false);
  assert.match(v.first, /the clean input \(a fault of the test pair\)/);
  assert.ok(!v.failures.some((f) => f.startsWith("accept-all")));
});

test("a wrong accept or reject side, leftover revisions and a repair on opening all fail", () => {
  const a = good();
  a.accept = job(shape(["One.\r", "Two new.\r", "\r"], CLEAN.tables, CLEAN.bookmarks), {
    raw: { paragraphs: 4, tables: 1 },
  });
  assert.match(
    judgePair(a).first,
    /^accept-all is not the clean document: paragraph count: expected 2, got 3/,
  );
  const b = good();
  b.reject = job(shape(["One.\r", "Two.\r"], CLEAN.tables));
  assert.match(
    judgePair(b).first,
    /^reject-all is not the baseline: paragraph count: expected 3, got 2/,
  );
  const c = good();
  c.accept.revisions_left = 2;
  assert.match(judgePair(c).first, /^2 revision\(s\) left after AcceptAll$/);
  const d = good();
  d.accept.dirty_on_open = true;
  assert.match(judgePair(d).first, /changed the track-changes output while opening it/);
  // When Word marks the clean input as changed on opening too, that is how Word reads these files.
  d.clean.dirty_on_open = true;
  assert.equal(judgePair(d).pass, true);
});

test("the colour output must have the paragraphs and tables of the track-changes output", () => {
  const r = good();
  r.color.raw = { paragraphs: 5, tables: 1 };
  assert.equal(
    judgePair(r).first,
    "the colour output has 5 paragraphs, 1 tables, the track-changes output 4 paragraphs, 1 tables",
  );
});

test("a file the tool did not write fails the pair and gets no Word job", () => {
  const r = good();
  r.tool.track = 2;
  delete r.accept;
  delete r.reject;
  assert.match(judgePair(r).first, /did not write the track-changes file \(exit 2\)/);
  const jobs = jobsFor(7, {
    track: null,
    color: "C:\\r\\c.docx",
    clean: "C:\\r\\k.docx",
    base: "C:\\r\\b.docx",
  });
  assert.deepEqual(
    jobs.map((j) => [j.id, j.action]),
    [
      ["7:clean", "none"],
      ["7:base", "none"],
      ["7:color", "none"],
    ],
  );
  const all = jobsFor(1, { track: "t", color: "c", clean: "k", base: "b" });
  assert.deepEqual(
    all.map((j) => `${j.id} ${j.action} ${j.path}`),
    ["1:accept accept t", "1:reject reject t", "1:clean none k", "1:base none b", "1:color none c"],
  );
});

test("the run exits 0 only when every pair passed and Word was closed properly", () => {
  const pass = [{ step: "01 a", result: "PASS", got: "" }];
  const fail = [...pass, { step: "02 b", result: "FAIL", got: "paragraph 2: expected ..." }];
  assert.equal(exitCode({ rows: pass }), 0);
  assert.equal(exitCode({ rows: fail }), 1);
  assert.equal(exitCode({ rows: [] }), 1);
  assert.equal(exitCode({ rows: pass, wordError: "timed out" }), 1);
  assert.equal(exitCode({ rows: pass, cleanupFail: true }), 1);
  assert.equal(renderRows(fail).split("\n").length, 3);
  assert.match(renderRows(fail), /^02 b\s+FAIL\s+paragraph 2: expected \.\.\.$/m);
});
