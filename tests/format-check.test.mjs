import { test } from "node:test";
import assert from "node:assert/strict";
import { findFormatIssues, STAT, bodyParagraphs } from "../daemon/format-check.mjs";

const run = (text, f = {}) => ({ text, italic: false, sub: false, sup: false, ...f });
const para = (index, ...runs) => ({ index, runs: runs.map(([text, f]) => run(text, f)) });
const profile = {
  italics: {
    genera: ["Escherichia", "Bacillus"],
    epithets: ["coli", "subtilis"],
    genes: ["lacZ", "recA", "mmoX"],
    statistics: true,
  },
  subscripts: ["CO2", "OD600", "H2O"],
};
const STYLE_RULES = ["stat", "species", "gene", "subscript", "indent"];
const kinds = (ps, pr = profile, extra = {}) =>
  findFormatIssues(ps, pr, extra).map((f) => `${f.rule}:${f.find ?? f.paragraph_index}`);
const finds = (rule, text, pr = profile) =>
  findFormatIssues([para(0, [text])], pr)
    .filter((f) => f.rule === rule)
    .map((f) => f.find);

// ---------------------------------------------------------------- corpus
// Written before the rules. Every text is one plain (unformatted) run.
// positive: [text, the words the rule must report, in order].
// negative: texts the rule must not report anything in (substring traps).
const CORPUS = {
  stat: {
    positive: [
      ["significant (P < 0.05)", ["P"]],
      ["p = 0.012 after correction", ["p"]],
      ["adjusted q values below 0.1", ["q"]],
      ["t(2) = 8.96", ["t"]],
      ["n = 3 per group", ["n"]],
      ["P-values were adjusted", ["P"]],
      ["P≤0.01 and Q-adjusted", ["P", "Q"]],
      ["(n ≥ 5; P < 0.001)", ["n", "P"]],
      ["the Q value was low", ["Q"]],
      ["P > 0.05 in both", ["P"]],
    ],
    negative: [
      "Smith P.E. used the method",
      "Park P. E. and Lee Q. 2020",
      "strain Pm02 grew well",
      "n replicates were pooled",
      "plasmid pQ = none",
      "the buffer had pH = 7.0",
      "the test(2) call returned",
      "Top = 5 and step = 2",
      "P450 enzymes were assayed",
      "ATP > ADP in all samples",
      "OD600 = 0.5 at harvest",
      "the mean = 4 and the span = 2",
      "a t-test and a t test were used",
      "ΔP < 0.2 bar across the filter",
      "values of q and p are given below",
    ],
  },
  species: {
    positive: [
      ["Escherichia coli grew fast", ["Escherichia coli"]],
      ["E. coli grew fast", ["E. coli"]],
      ["Bacillus subtilis and B. subtilis", ["Bacillus subtilis", "B. subtilis"]],
      ["Escherichia sp. isolates", ["Escherichia"]],
      ["a soil genus (Bacillus)", ["Bacillus"]],
      ["Escherichia coli in a table", ["Escherichia coli"]],
      ["Escherichia-like cells", ["Escherichia"]],
      ["cells of E.coli", ["E.coli"]],
    ],
    negative: [
      "colitis was induced in mice",
      "coliform counts were low",
      "the family Escherichiaceae",
      "bacillus-shaped cells",
      "E. colitis is not a species",
      "the epithet coli alone",
      "Subtilisin is a protease",
      "Bacilli were counted",
      "Staphylococcus aureus is not in the profile",
      "PseudoBacillus is one word",
    ],
  },
  gene: {
    positive: [
      ["lacZ was deleted", ["lacZ"]],
      ["the recA gene", ["recA"]],
      ["the ΔlacZ strain", ["lacZ"]],
      ["a lacZ-GFP fusion", ["lacZ"]],
      ["one operon (mmoX)", ["mmoX"]],
      ["mmoX, recA and lacZ", ["mmoX", "recA", "lacZ"]],
    ],
    negative: [
      "LacZ protein was purified",
      "the lacZY operon",
      "the placZ vector",
      "RecA-dependent repair",
      "the recA2 allele",
      "lacz in lower case",
      "the mmoXYZ cluster",
    ],
  },
  subscript: {
    positive: [
      ["CO2 was fed", ["CO2"]],
      ["an OD600 of 1.0", ["OD600"]],
      ["H2O and CO2", ["H2O", "CO2"]],
      ["CO2-fixing cells", ["CO2"]],
      ["one gas (CO2)", ["CO2"]],
      ["grown on CO2.", ["CO2"]],
    ],
    negative: [
      "the CO2H group",
      "supercritical sCO2",
      "CO20 is a typo",
      "H2O2 is another compound",
      "an OD6000 reading",
      "co2 in lower case",
      "COD600 is one word",
    ],
  },
};

for (const [rule, { positive, negative }] of Object.entries(CORPUS)) {
  test(`${rule} corpus: at least five positive and five negative cases`, () => {
    assert.ok(positive.length >= 5 && negative.length >= 5);
  });
  for (const [text, want] of positive) {
    test(`${rule} reports ${JSON.stringify(want)} in ${JSON.stringify(text)}`, () => {
      assert.deepEqual(finds(rule, text), want);
    });
  }
  for (const text of negative) {
    test(`${rule} reports nothing in ${JSON.stringify(text)}`, () => {
      assert.deepEqual(finds(rule, text), []);
    });
  }
}

test("STAT is the statistics-symbol pattern and is exported", () => {
  assert.ok(STAT instanceof RegExp);
  assert.deepEqual("P < 0.05, n = 3".match(STAT), ["P", "n"]);
});

// ---------------------------------------------------------------- formatting already right
test("a symbol or a name that already has the format passes", () => {
  const ok = [
    para(0, ["P", { italic: true }], [" < 0.05 in "], ["Escherichia coli", { italic: true }]),
    para(1, ["Escherichia", { italic: true }], [" "], ["coli", { italic: true }], [" grew"]),
    para(2, ["E. coli", { italic: true }], [" carries "], ["lacZ", { italic: true }]),
    para(3, ["CO"], ["2", { sub: true }], [" and OD"], ["600", { sub: true }]),
    para(4, ["H"], ["2", { sub: true }], ["O was added"]),
    para(5, ["(", {}], ["n", { italic: true }], [" = 3)"]),
  ];
  assert.deepEqual(kinds(ok), []);
});

test("a half-formatted name or a digit in superscript is still a finding", () => {
  assert.deepEqual(
    kinds([
      para(0, ["Escherichia", { italic: true }], [" coli grew"]),
      para(1, ["CO"], ["2", { sup: true }]),
      para(2, ["la", { italic: true }], ["cZ"]),
    ]),
    ["species:Escherichia coli", "subscript:CO2", "gene:lacZ"],
  );
});

test("findings come in document order and say where and what to change", () => {
  const ps = [para(1, ["Escherichia coli lacZ grew on CO2 to OD600 1.0 (P < 0.05, P = 0.2)"])];
  assert.deepEqual(kinds(ps), [
    "species:Escherichia coli",
    "gene:lacZ",
    "subscript:CO2",
    "subscript:OD600",
    "stat:P",
    "stat:P",
  ]);
  const f = findFormatIssues(ps, profile);
  const text = ps[0].runs[0].text;
  assert.deepEqual(
    [f[0].paragraph_index, f[0].start, f[0].want, f[0].spans],
    [1, 0, { italic: true }, [{ start: 0, length: 16 }]],
  );
  // Only the digits of a formula go down.
  assert.deepEqual(
    [f[2].want, f[2].spans],
    [{ sub: true }, [{ start: text.indexOf("CO2") + 2, length: 1 }]],
  );
  assert.deepEqual(f[3].spans, [{ start: text.indexOf("OD600") + 2, length: 3 }]);
  // Two P in one paragraph: each finding carries words that occur once and
  // the place of its letter inside them.
  for (const x of [f[4], f[5]]) {
    assert.equal(text.split(x.locate.text).length, 2);
    assert.equal(x.locate.text[x.locate.offset], "P");
    assert.equal(text.indexOf(x.locate.text) + x.locate.offset, x.start);
    assert.ok(x.context.includes("P"));
  }
  assert.notEqual(f[4].start, f[5].start);
});

test("no style rule runs unless the profile turns it on: an empty profile gives no style finding", () => {
  const LONG_TEXT =
    "This sentence is long enough to count as body text for the indent rule and it ends with a period.";
  const indented = (index, first_line) => ({
    index,
    runs: [run(LONG_TEXT)],
    indent: { first_line, left: 0, source: "direct" },
  });
  // Everything a style rule could report: symbols, names, a formula, one odd indent.
  const ps = [
    para(0, ["Escherichia coli lacZ on CO2, P < 0.05, p = 0.01, n = 3, t(2) = 8.9, q values"]),
    indented(1, 36),
    indented(2, 36),
    indented(3, 36),
    indented(4, 36),
    indented(5, 0),
  ];
  for (const empty of [
    undefined,
    {},
    { italics: {} },
    { italics: { genera: [], epithets: [], genes: [] }, subscripts: [] },
    { italics: { statistics: false }, first_line_indent: null },
    { italics: { statistics: "yes" } },
    { italics: { statistics: 1 } },
    { italics: null, subscripts: null },
  ]) {
    assert.deepEqual(findFormatIssues(ps, empty), [], JSON.stringify(empty));
  }
  // The same document with every rule turned on reports every kind.
  const all = { ...profile, first_line_indent: "consistent" };
  assert.deepEqual(
    [...new Set(findFormatIssues(ps, all).map((f) => f.rule))].sort(),
    [...STYLE_RULES].sort(),
  );
});

test("each style rule is turned on by its own profile key and by nothing else", () => {
  const ps = [para(0, ["Escherichia coli lacZ on CO2, P < 0.05"])];
  assert.deepEqual(kinds(ps, { italics: { statistics: true } }), ["stat:P"]);
  assert.deepEqual(kinds(ps, { italics: { statistics: false, genes: ["lacZ"] } }), ["gene:lacZ"]);
  assert.deepEqual(kinds(ps, { italics: { genes: ["lacZ"] } }), ["gene:lacZ"]);
  assert.deepEqual(kinds(ps, { italics: { genes: ["lacZ"], statistics: true } }), [
    "gene:lacZ",
    "stat:P",
  ]);
  assert.deepEqual(kinds(ps, { italics: { genera: ["Escherichia"], epithets: ["coli"] } }), [
    "species:Escherichia coli",
  ]);
  assert.deepEqual(kinds(ps, { subscripts: ["CO2"] }), ["subscript:CO2"]);
});

test("names with regex characters in the profile are read as plain text", () => {
  const pr = { italics: { genes: ["bla(TEM)", "a.b"] }, subscripts: ["NH4+"] };
  assert.deepEqual(kinds([para(0, ["bla(TEM) and axb and a.b with NH4+ ions"])], pr), [
    "gene:bla(TEM)",
    "gene:a.b",
    "subscript:NH4+",
  ]);
});

// ---------------------------------------------------------------- references
test("the references section is skipped for text rules, up to its end heading", () => {
  const ps = [
    para(3, ["P < 0.05"]),
    para(4, ["References"]),
    para(5, ["Park P. E. 2020. P < 0.05 in Escherichia coli."]),
    para(6, ["Supplementary notes"]),
    para(7, ["n = 3"]),
  ];
  // No end heading: the list runs to the end of the document.
  assert.deepEqual(kinds(ps, { ...profile, ref_section_head: "references" }), ["stat:P"]);
  assert.deepEqual(kinds(ps), ["stat:P"]);
  assert.deepEqual(kinds(ps, { ...profile, ref_section_end_heads: ["^supplementary"] }), [
    "stat:P",
    "stat:n",
  ]);
  // A sentence that only mentions the word does not open the section.
  const prose = [
    para(0, ["The references below were checked by hand and n = 3 of them were wrong."]),
  ];
  const pr = { italics: { statistics: true }, ref_section_head: "references" };
  assert.deepEqual(kinds([...prose, para(1, ["q = 0.01"])], pr), ["stat:n", "stat:q"]);
});

// ---------------------------------------------------------------- empty table
const table = (index, t) => ({ index, runs: [], table: t });
const EMPTY_TABLES = [
  { rows: 2, cols: 2, empty: true },
  { rows: 1, cols: 1, empty: true },
  { rows: 2, cols: 2, cells: ["", "", "", ""] },
  { rows: 1, cols: 3, cells: [" ", " ", "\t"] },
  { rows: 3, cols: 1, cells: ["\u0007", "\r", ""] },
];
const FILLED_TABLES = [
  { rows: 2, cols: 2, empty: false },
  { rows: 2, cols: 2, cells: ["Strain", "", "", ""] },
  { rows: 1, cols: 2, cells: ["0", ""] },
  { rows: 1, cols: 2, cells: ["-", "-"] },
  { rows: 2, cols: 2 },
  undefined,
];

test("empty table: a table whose cells hold no text is one finding at its paragraph", () => {
  assert.ok(EMPTY_TABLES.length >= 5 && FILLED_TABLES.length >= 5);
  EMPTY_TABLES.forEach((t, i) => {
    const f = findFormatIssues([table(i + 3, t)], profile);
    assert.deepEqual(
      f.map((x) => [x.rule, x.paragraph_index, x.found]),
      [["empty_table", i + 3, { rows: t.rows, cols: t.cols }]],
    );
  });
  FILLED_TABLES.forEach((t, i) => {
    assert.deepEqual(findFormatIssues([table(i, t)], profile), [], JSON.stringify(t));
  });
});

test("an empty table is reported even inside the references section", () => {
  const ps = [
    para(2, ["References"]),
    table(3, { rows: 2, cols: 2, empty: true }),
    para(5, ["Park P. E. 2020. P < 0.05"]),
  ];
  assert.deepEqual(
    findFormatIssues(ps, { ...profile, ref_section_head: "references" }).map((x) => x.rule),
    ["empty_table"],
  );
});

// ---------------------------------------------------------------- first-line indent
const LONG =
  "This sentence is long enough to count as body text for the indent rule and it ends with a period.";
const body = (index, first_line, more = {}) => ({
  index,
  runs: [run(LONG)],
  indent: { first_line, left: 0, source: "direct" },
  ...more,
});
const withText = (index, first_line, text) => ({ ...body(index, first_line), runs: [run(text)] });
const headings = [{ index: 0, level: 1, text: "1. Introduction" }];
const doc = [
  para(0, ["1. Introduction"]),
  body(1, 0),
  body(2, 36),
  body(3, 0),
  body(4, 36),
  body(5, 36),
];

test("body paragraphs: no headings, captions, list items, table cells, short lines, unfinished lines or references", () => {
  const ps = [
    ...doc,
    body(6, 0, { is_list: true }),
    body(7, 0, { in_table: true }),
    withText(8, 0, `Fig. 1. ${LONG}`),
    withText(9, 0, `Table 2. ${LONG}`),
    withText(10, 0, "A short line."),
    withText(11, 0, LONG.slice(0, -1)),
    { ...body(12, 0), table: { rows: 1, cols: 1, empty: false } },
    { index: 13, runs: [run(LONG)] },
    para(14, ["References"]),
    body(15, 0),
  ];
  assert.deepEqual(
    bodyParagraphs(ps, {}, { headings }).map((p) => p.index),
    [1, 2, 3, 4, 5],
  );
  // A profile caption pattern is read too.
  const scheme = [...doc, withText(6, 0, `Scheme 1. ${LONG}`)];
  assert.equal(bodyParagraphs(scheme, {}, { headings }).length, 6);
  const pr = { legend_heads: { scheme: "^Scheme\\s+\\d+" } };
  assert.equal(bodyParagraphs(scheme, pr, { headings }).length, 5);
});

test("expect indented: an unindented body paragraph is a finding; the first one after a heading is skipped when asked", () => {
  const pr = { first_line_indent: { expect: "indented", skip_after_heading: true } };
  assert.deepEqual(kinds(doc, pr, { headings }), ["indent:3"]);
  assert.deepEqual(kinds(doc, { first_line_indent: "indented" }, { headings }), [
    "indent:1",
    "indent:3",
  ]);
  const f = findFormatIssues(doc, pr, { headings })[0];
  // The wanted value is the document's own most common indent.
  assert.deepEqual([f.rule, f.want.first_line_indent, f.found], ["indent", 36, 0]);
});

test("expect indented with min_pt: a smaller indent is a finding and the wanted value is min_pt", () => {
  const f = findFormatIssues(
    [para(0, ["H"]), body(1, 10), body(2, 36)],
    { first_line_indent: { expect: "indented", min_pt: 18 } },
    { headings },
  );
  assert.deepEqual(
    f.map((x) => [x.paragraph_index, x.want.first_line_indent, x.found]),
    [[1, 18, 10]],
  );
});

test("expect none: indented body paragraphs are findings and the wanted value is 0", () => {
  assert.deepEqual(kinds(doc, { first_line_indent: "none" }, { headings }), [
    "indent:2",
    "indent:4",
    "indent:5",
  ]);
  const f = findFormatIssues(doc, { first_line_indent: "none" }, { headings });
  assert.deepEqual(
    f.map((x) => x.want.first_line_indent),
    [0, 0, 0],
  );
  assert.deepEqual(
    kinds(
      [para(0, ["H"]), body(1, { chars: 2 }), body(2, 0)],
      { first_line_indent: "none" },
      { headings },
    ),
    ["indent:1"],
  );
});

test("expect consistent: only outliers of a clear majority (80%) are findings, and a mixed document gives none", () => {
  const same = { first_line_indent: "consistent" };
  const mostly = [para(0, ["H"]), body(1, 36), body(2, 36), body(3, 36), body(4, 36), body(5, 0)];
  assert.deepEqual(kinds(mostly, same, { headings }), ["indent:5"]);
  assert.equal(findFormatIssues(mostly, same, { headings })[0].want.first_line_indent, 36);
  assert.deepEqual(kinds(mostly, { first_line_indent: { expect: "consistent" } }, { headings }), [
    "indent:5",
  ]);
  const flush = [para(0, ["H"]), body(1, 0), body(2, 0), body(3, 0), body(4, 0), body(5, 24)];
  assert.deepEqual(kinds(flush, same, { headings }), ["indent:5"]);
  assert.equal(findFormatIssues(flush, same, { headings })[0].want.first_line_indent, 0);
  assert.deepEqual(kinds(doc, same, { headings }), []);
  // All alike: nothing to report.
  assert.deepEqual(kinds([para(0, ["H"]), body(1, 36), body(2, 36)], same, { headings }), []);
});

test("without an indent setting in the profile the indent rule does not run, whatever the document looks like", () => {
  const mostly = [para(0, ["H"]), body(1, 36), body(2, 36), body(3, 36), body(4, 36), body(5, 0)];
  for (const pr of [
    {},
    { first_line_indent: null },
    { first_line_indent: "majority" },
    { first_line_indent: { min_pt: 18 } },
    { first_line_indent: { expect: null, skip_after_heading: true } },
  ]) {
    assert.deepEqual(kinds(mostly, pr, { headings }), [], JSON.stringify(pr));
  }
});

test("a hanging indent is not an indent; an unknown indent is not judged; a chars-based indent counts as indented", () => {
  assert.deepEqual(
    kinds(
      [
        para(0, ["H"]),
        body(1, -18),
        body(2, null),
        body(3, { chars: 2 }),
        body(4, 36),
        body(5, 36),
        body(6, 36),
        body(7, 36),
      ],
      { first_line_indent: "indented" },
      { headings },
    ),
    ["indent:1"],
  );
});

test("when no indent value was ever seen, the finding asks for an indent without a value", () => {
  const f = findFormatIssues(
    [para(0, ["H"]), body(1, 0), body(2, 0)],
    { first_line_indent: "indented" },
    { headings },
  );
  assert.deepEqual(
    f.map((x) => x.want.first_line_indent),
    [null, null],
  );
});

test("style-inherited indents are judged like direct ones unless the caller says they cannot be read", () => {
  const ps = [
    para(0, ["H"]),
    body(1, 36),
    { ...body(2, 0), indent: { first_line: 0, left: 0, source: "style" } },
  ];
  const pr = { first_line_indent: "indented" };
  assert.deepEqual(kinds(ps, pr, { headings }), ["indent:2"]);
  assert.deepEqual(kinds(ps, pr, { headings, styleIndentsKnown: false }), []);
});

test("indent findings never appear for text that is not body text", () => {
  const ps = [
    para(0, ["1. Introduction"]),
    body(1, 36),
    body(2, 36),
    body(3, 36),
    body(4, 36),
    body(5, 0, { is_list: true }),
    body(6, 0, { in_table: true }),
    withText(7, 0, "Fig. 1. A curve."),
    withText(8, 0, "A short line."),
    para(9, ["References"]),
    body(10, 0),
  ];
  assert.deepEqual(kinds(ps, { first_line_indent: "indented" }, { headings }), []);
  assert.deepEqual(kinds(ps, {}, { headings }), []);
});
