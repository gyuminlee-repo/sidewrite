import { test } from "node:test";
import assert from "node:assert/strict";
import { sectionsOf, wordCount, checkCountItem } from "../daemon/count-check.mjs";

const P = (index, text, more = {}) => ({ index, text, ...more });
const doc = [
  P(0, "A study of growth"),
  P(1, "Abstract"),
  P(2, "one two three four five six"),
  P(3, "Keywords: growth; carbon; enzyme"),
  P(4, "1. Introduction"),
  P(5, "Intro text."),
  P(6, "2. Methods"),
  P(7, "Method text."),
  P(8, "2.1 Strains"),
  P(9, "Three strains were used."),
  P(10, "3. Results"),
  P(11, "Fig. 1. A curve."),
  P(12, "Fig. 2. A gel."),
  P(13, "Table 1. Strains."),
  P(14, "Figure 2 shows the gel again and Table 4 is only named in passing."),
];
const H = (index, level) => ({ index, level, text: doc[index].text });
const headings = [H(0, 1), H(1, 1), H(4, 1), H(6, 1), H(8, 2), H(10, 1)];
const item = (check) => ({ id: "L1", text: "t", check: { type: "count", ...check } });
const run = (check, paragraphs = doc, deps = { headings }) =>
  checkCountItem(item(check), paragraphs, deps);

test("words are whitespace tokens", () => {
  assert.equal(wordCount(" a  b\tc "), 3);
  assert.equal(wordCount(""), 0);
  assert.equal(wordCount("   "), 0);
  assert.equal(wordCount("well-known 3.5 mL"), 3);
});

test("sections follow the headings; a section runs to the next heading of its level or above", () => {
  const secs = sectionsOf(doc, { headings });
  assert.deepEqual(
    secs.map((s) => [s.title, s.index, s.level]),
    [
      ["A study of growth", 0, 1],
      ["Abstract", 1, 1],
      ["1. Introduction", 4, 1],
      ["2. Methods", 6, 1],
      ["2.1 Strains", 8, 2],
      ["3. Results", 10, 1],
    ],
  );
  // Methods holds its own text and the text of its subsection, not the subheading.
  assert.deepEqual(
    secs[3].body.map((p) => p.index),
    [7, 9],
  );
  assert.deepEqual(
    secs[4].body.map((p) => p.index),
    [9],
  );
});

test("without headings a profile regex finds the sections; without either there are none", () => {
  const flat = sectionsOf(doc, { headRe: "^(?:Abstract|\\d+\\.\\s)" });
  assert.deepEqual(
    flat.map((s) => s.title),
    ["Abstract", "1. Introduction", "2. Methods", "3. Results"],
  );
  assert.deepEqual(sectionsOf(doc), []);
  // Headings win over the regex when both are given.
  assert.equal(sectionsOf(doc, { headings, headRe: "^Abstract" }).length, 6);
});

test("section_words over the limit is one violation at the heading, with the count", () => {
  const r = run({ target: "section_words", section: "^abstract$", max: 5 });
  assert.equal(r.id, "L1");
  assert.equal(r.status, "violations");
  assert.deepEqual(r.violations, [
    { paragraph_index: 1, start: null, match: null, count: 6, min: null, max: 5 },
  ]);
});

test("section_words within the limit is ok and reports the count; the Keywords line is not section text", () => {
  const r = run({ target: "section_words", section: "^abstract$", max: 250 });
  assert.deepEqual([r.status, r.violations, r.count], ["ok", [], 6]);
  assert.equal(run({ target: "section_words", section: "methods", min: 6, max: 6 }).status, "ok");
  assert.equal(run({ target: "section_words", section: "methods", min: 7 }).violations[0].count, 6);
});

test("section_words reads the text as if pending tracked changes were accepted", () => {
  const edited = doc.map((p) =>
    p.index === 2 ? { ...p, reviewed_text: "one two three" } : { ...p },
  );
  assert.equal(run({ target: "section_words", section: "^abstract$", max: 250 }, edited).count, 3);
});

test("a missing section, a bad section pattern and a check without a bound are errors, not passes", () => {
  const missing = run({ target: "section_words", section: "^conclusion", max: 9 });
  assert.deepEqual([missing.status, missing.violations], ["error", []]);
  assert.match(missing.error, /no section matches/);
  assert.equal(run({ target: "section_words", section: "(", max: 9 }).status, "error");
  assert.equal(run({ target: "section_words", section: "abstract" }).status, "error");
  assert.equal(run({ target: "section_words", max: 9 }).status, "error");
  assert.equal(run({ target: "keywords" }).status, "error");
  assert.equal(run({ target: "paragraphs", max: 3 }).status, "error");
  // No headings and no regex: the document has no sections to count in.
  assert.equal(
    run({ target: "section_words", section: "abstract", max: 9 }, doc, {}).status,
    "error",
  );
});

test("keywords are counted on the Keywords line", () => {
  const r = run({ target: "keywords", min: 4, max: 6 });
  assert.deepEqual(r.violations, [
    { paragraph_index: 3, start: null, match: null, count: 3, min: 4, max: 6 },
  ]);
  assert.equal(run({ target: "keywords", min: 3, max: 6 }).status, "ok");
  const comma = [P(0, "Key words. growth, carbon, enzyme, methane,")];
  assert.equal(run({ target: "keywords", max: 3 }, comma).violations[0].count, 4);
  const none = [P(0, "The keywords of this study are few.")];
  assert.equal(run({ target: "keywords", max: 3 }, none).status, "error");
});

test("figures and tables are counted by distinct caption numbers, not by mentions", () => {
  const figs = run({ target: "figures", max: 1 });
  assert.deepEqual(figs.violations, [
    { paragraph_index: null, start: null, match: null, count: 2, min: null, max: 1 },
  ]);
  assert.equal(run({ target: "figures", max: 2 }).status, "ok");
  assert.deepEqual(
    [run({ target: "tables", max: 1 }).status, run({ target: "tables", max: 1 }).count],
    ["ok", 1],
  );
  assert.equal(run({ target: "tables", min: 2 }).violations[0].count, 1);
});

test("a profile caption pattern replaces the default one", () => {
  const ps = [P(0, "Scheme 1 - A route."), P(1, "Scheme 2 - Another route."), P(2, "Fig. 1. X.")];
  const profile = { legend_heads: { main: "^Scheme\\s+(\\d+)" } };
  assert.equal(run({ target: "figures", max: 5 }, ps, { profile }).count, 2);
  assert.equal(run({ target: "figures", max: 5 }, ps, {}).count, 1);
});

test("section_order: the named sections must come in the given order", () => {
  assert.equal(run({ target: "section_order", order: ["introduction", "methods"] }).status, "ok");
  const r = run({ target: "section_order", order: ["methods", "introduction", "results"] });
  assert.equal(r.status, "violations");
  assert.deepEqual(r.violations, [
    { paragraph_index: 4, start: null, match: null, count: null, min: null, max: null },
  ]);
  const gone = run({ target: "section_order", order: ["introduction", "discussion"] });
  assert.equal(gone.status, "error");
  assert.match(gone.error, /discussion/);
  assert.equal(run({ target: "section_order", order: [] }).status, "error");
  assert.equal(run({ target: "section_order" }).status, "error");
});

test("the result has the shape checkRegexItem gives: id, status, violations with paragraph_index, start, match", () => {
  for (const r of [
    run({ target: "keywords", max: 1 }),
    run({ target: "keywords", max: 9 }),
    run({ target: "nothing" }),
  ]) {
    assert.equal(r.id, "L1");
    assert.ok(["ok", "violations", "error"].includes(r.status));
    assert.ok(Array.isArray(r.violations));
    for (const v of r.violations) {
      assert.ok("paragraph_index" in v && "start" in v && "match" in v);
    }
  }
});
