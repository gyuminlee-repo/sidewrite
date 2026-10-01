// L1-L3 known-answer corpora. Each list has at least five inputs that must
// fire and five that must not, with substring traps ("3 mM" inside "30 mM",
// "v2" inside a word, "3" inside "3.5").
import { test } from "node:test";
import assert from "node:assert/strict";
import {
  extractFacts,
  factPresent,
  lostFacts,
  unsourcedTerms,
  narrowEdit,
  startsSentence,
} from "../daemon/edit-facts.mjs";

const kinds = (text, opts) => extractFacts(text, opts).map((f) => `${f.kind}:${f.text}`);

test("extractFacts: numbers with units, versions, parentheses, names", () => {
  assert.deepEqual(kinds("centrifuged at 4000 × g for 10 min (Eppendorf 5810R)"), [
    "number:4000 × g",
    "number:10 min",
    "paren:Eppendorf 5810R",
  ]);
  assert.deepEqual(kinds("processed with MaxQuant 2.4.1 and Boltz-2 v2.1"), [
    "name:MaxQuant",
    "version:2.4.1",
    "name:Boltz-2",
    "version:v2.1",
  ]);
  assert.deepEqual(kinds("stored at -20 °C in 3.5 mM buffer"), ["number:-20 °C", "number:3.5 mM"]);
  // One fact per parenthesis: the names inside are not repeated.
  assert.deepEqual(kinds("filters (Merck, Germany)"), ["paren:Merck, Germany"]);
});

test("extractFacts traps: v2 or digits inside a word, sentence starts, 3 mice", () => {
  assert.deepEqual(kinds("the dev2.3 branch and the pv2.1x plasmid"), []);
  assert.deepEqual(kinds("a Q10 value near CO2"), ["name:Q10", "name:CO2"]);
  assert.deepEqual(kinds("The values were clear. Cultures grew.", { sentenceStart: true }), []);
  assert.deepEqual(kinds("Boltz-2 predicted it", { sentenceStart: true }), ["name:Boltz-2"]);
  assert.deepEqual(kinds("3 mice"), ["number:3"]);
});

test("factPresent: whole tokens only, unit spacing ignored", () => {
  assert.equal(factPresent("30 °C", "held at 30°C"), true);
  assert.equal(factPresent("3 mM", "in 30 mM buffer"), false);
  assert.equal(factPresent("3 mM", "in 13 mM buffer"), false);
  assert.equal(factPresent("3", "for 3.5 h"), false);
  assert.equal(factPresent("Amicon", "amicon filters"), true);
});

// L1 corpus: [deleted, inserted, opts, expected lost facts]
const LOST_POSITIVE = [
  [
    "centrifuged at 4000 × g for 10 min (Eppendorf 5810R)",
    "centrifuged",
    {},
    ["4000 × g", "10 min", "Eppendorf 5810R"],
  ],
  [
    "Boltz-2 (v2.1.0) predicted the complex",
    "the complex was predicted",
    { sentenceStart: true },
    ["Boltz-2", "v2.1.0"],
  ],
  [
    "concentrated with Amicon Ultra filters (Merck)",
    "concentrated by ultrafiltration",
    {},
    ["Amicon", "Ultra", "Merck"],
  ],
  ["processed with MaxQuant 2.4.1", "processed with MaxQuant", {}, ["2.4.1"]],
  ["in 3 mM buffer", "in 30 mM buffer", {}, ["3 mM"]],
  ["held at 30 °C", "held at 130 °C", {}, ["30 °C"]],
  ["for 3 h", "for 3.5 h", {}, ["3 h"]],
];
const LOST_NEGATIVE = [
  ["held at 30 °C", "held at 30°C", {}],
  ["held at 30 degrees", "held at 30 °C", {}], // unit written another way
  ["30 degrees", "30 degrees Celsius", {}],
  ["The values were clear", "Values were clear", { sentenceStart: true }],
  ["a Q10 value near CO2", "the Q10 value near CO2", {}],
  ["Amicon Ultra filters (Merck) at 4 °C", "Amicon Ultra centrifugal filters (Merck) at 4 °C", {}],
  ["the dev2.3 branch", "the branch", {}],
  ["was significantly increased", "increased markedly", {}],
];

test("L1 lostFacts: positive corpus", () => {
  for (const [del, ins, opts, want] of LOST_POSITIVE) {
    assert.deepEqual(lostFacts(del, ins, opts), want, del);
  }
});

test("L1 lostFacts: negative corpus", () => {
  for (const [del, ins, opts] of LOST_NEGATIVE)
    assert.deepEqual(lostFacts(del, ins, opts), [], del);
  console.log(
    `lostFacts corpus: ${LOST_POSITIVE.length} positive, ${LOST_NEGATIVE.length} negative`,
  );
});

const DOC = [
  "Cultures were grown in sealed flasks with shaking at 30 °C.",
  "A cell-free extract was prepared by sonication; enzyme assays followed.",
  "Growth studies on methanol are cited here.",
];
// L2 corpus: [inserted, extra sources, expected unsourced terms]
const TERMS_POSITIVE = [
  ["with 0.1 % antifoam", [], ["0.1 %", "antifoam"]],
  ["using a Branson sonicator", [], ["Branson", "sonicator"]],
  ["in 12 sealed flasks", [], ["12"]],
  ["a clarified lysate", [], ["clarified", "lysate"]],
  ["grown on ethanol", [], ["ethanol"]], // trap: "methanol" contains it
];
const TERMS_NEGATIVE = [
  ["a clarified lysate", ["Please call it a clarified lysate."]],
  ["the flask and one assay study", []], // plural and -ies forms
  ["read the Optical Density", ["optical density was read"]],
  ["was also then used", []], // stopwords
  ["at 30°C", []], // unit spacing
  ["a pH of", []], // short words
];

test("L2 unsourcedTerms: positive corpus", () => {
  for (const [ins, extra, want] of TERMS_POSITIVE) {
    assert.deepEqual(unsourcedTerms(ins, [...DOC, ...extra]), want, ins);
  }
});

test("L2 unsourcedTerms: negative corpus", () => {
  for (const [ins, extra] of TERMS_NEGATIVE) {
    assert.deepEqual(unsourcedTerms(ins, [...DOC, ...extra]), [], ins);
  }
  console.log(
    `unsourcedTerms corpus: ${TERMS_POSITIVE.length} positive, ${TERMS_NEGATIVE.length} negative`,
  );
});

test("L3 narrowEdit: only the words that change", () => {
  assert.deepEqual(
    narrowEdit("A cell-free extract was prepared", "A clarified lysate was prepared"),
    {
      find: "cell-free extract",
      replace: "clarified lysate",
    },
  );
  assert.deepEqual(narrowEdit("the red big dog", "the big dog"), { find: "red ", replace: "" });
  assert.deepEqual(narrowEdit("30 min at room temperature", "300 min at room temperature"), {
    find: "30",
    replace: "300",
  });
});

test("L3 narrowEdit: a pure insertion keeps the word before it", () => {
  assert.deepEqual(narrowEdit("30 degrees", "30 degrees Celsius"), {
    find: "degrees",
    replace: "degrees Celsius",
  });
  assert.deepEqual(narrowEdit("grown in flasks", "grown in sealed flasks"), {
    find: "in ",
    replace: "in sealed ",
  });
});

test("L3 narrowEdit: null when the full range must stay", () => {
  assert.equal(narrowEdit("a b a b", "a c a b"), null); // cut "b" twice inside find
  assert.equal(narrowEdit("a  b", "a b"), null); // white space only
  assert.equal(narrowEdit("same", "same"), null);
  assert.equal(narrowEdit("flasks", "sealed flasks"), null); // nothing to cut
  assert.equal(narrowEdit("incubated", "incubation"), null); // never inside a word
});

test("startsSentence finds the phrase in the snapshot and reads what is before it", () => {
  const texts = ["Title", "Cultures were grown. Amicon filters were used."];
  assert.equal(startsSentence("Amicon filters", texts), true);
  assert.equal(startsSentence("Cultures were", texts), true);
  assert.equal(startsSentence("were grown", texts), false);
  assert.equal(startsSentence("not in the text", texts), false);
});
