// A guideline sentence can turn a format rule of the profile on. The only
// one that needs no list of names is "statistics symbols in italics"
// (profile italics.statistics). As with count rules, the model only labels
// the sentence; the code checks that the sentence says it.

import { test } from "node:test";
import assert from "node:assert/strict";
import {
  candidatesFromLabels,
  guidelinePrompt,
  FORMAT_SETTINGS,
  FORMAT_WORDS,
  GUIDELINE_TYPES,
  GUIDELINE_WARNING,
} from "../daemon/guideline-extract.mjs";
import { findFormatIssues } from "../daemon/format-check.mjs";

const label = (rule) => ({ s: 1, type: "format", rule });
const STAT = { setting: "italics.statistics" };
const one = (sentence, rule = STAT) => {
  const r = candidatesFromLabels([sentence], { sentences: [label(rule)] });
  assert.equal(r.candidates.length, 1);
  assert.equal(r.candidates[0].text, sentence);
  assert.deepEqual(r.candidates[0].source, { kind: "guideline", sentence });
  assert.deepEqual(r.candidates[0].check, { type: "llm" });
  return r;
};

const SAYS_IT = [
  "Statistical symbols such as P and n are set in italics.",
  "Italicize statistical symbols.",
  "P values should be italicised.",
  "Use italics for statistics symbols (P, t, n).",
  "Symbols for statistics (P, n) must be in italic type.",
  "The p-value symbol is italicized throughout.",
];
const DOES_NOT = [
  "Species names are set in italics.",
  "Gene symbols are italicized.",
  "Statistical symbols are not set in italics.",
  "Do not italicize statistical symbols.",
  "Statistical symbols are set in roman type, never in italics.",
  "Report statistics in a table.",
  "The Italian statistics office publishes the data.",
  "Statistical symbols are set upright; gene names are in italics.",
  "Non-italic statistical symbols are preferred.",
];

test("format is a label type, and the prompt offers the settings the code knows", () => {
  assert.ok(GUIDELINE_TYPES.includes("format"));
  assert.deepEqual(Object.keys(FORMAT_SETTINGS), ["italics.statistics"]);
  const prompt = guidelinePrompt(["A."]);
  assert.match(prompt, /"format"/);
  assert.match(prompt, /"italics\.statistics"/);
  assert.ok(SAYS_IT.length >= 5 && DOES_NOT.length >= 5);
});

for (const sentence of SAYS_IT) {
  test(`a profile setting is proposed for ${JSON.stringify(sentence)}`, () => {
    const r = one(sentence);
    assert.deepEqual(r.warnings, []);
    assert.deepEqual(r.candidates[0].profile, { italics: { statistics: true } });
  });
}

for (const sentence of DOES_NOT) {
  test(`no profile setting from ${JSON.stringify(sentence)}`, () => {
    const r = one(sentence);
    assert.equal("profile" in r.candidates[0], false);
    assert.equal(r.warnings.length, 1);
    assert.match(r.warnings[0], /^s1: .*statistics symbols.*italics/);
  });
}

test("the words are lists the code reads; a word inside another word does not count", () => {
  for (const key of ["italic", "statistics", "negation"]) {
    assert.ok(FORMAT_WORDS[key].length >= 3, key);
  }
  for (const italic of FORMAT_WORDS.italic) {
    for (const stat of FORMAT_WORDS.statistics) {
      const says = `For ${stat}: ${italic}.`;
      assert.deepEqual(one(says).candidates[0].profile, { italics: { statistics: true } }, says);
      assert.equal("profile" in one(`For x${stat}x: ${italic}.`).candidates[0], false);
      assert.equal("profile" in one(`For ${stat}: x${italic}x.`).candidates[0], false);
    }
  }
  for (const no of FORMAT_WORDS.negation) {
    const sentence = `Statistics symbols: ${no} italics.`;
    assert.equal("profile" in one(sentence).candidates[0], false, no);
  }
});

test("a format label of another shape makes the whole reply unusable", () => {
  for (const bad of [
    { s: 1, type: "format" },
    label({ setting: "italics.genes" }),
    label({ setting: "italics.statistics", value: false }),
    label({ setting: "italics.statistics", note: "P and n" }),
    label("italics.statistics"),
    { s: 1, type: "style", rule: STAT },
  ]) {
    assert.deepEqual(
      candidatesFromLabels([SAYS_IT[0]], { sentences: [bad] }),
      { candidates: [], warnings: [GUIDELINE_WARNING] },
      JSON.stringify(bad),
    );
  }
});

test("the proposed setting is the one that turns the statistics rule on", () => {
  const { profile } = one(SAYS_IT[0]).candidates[0];
  const ps = [{ index: 0, runs: [{ text: "P < 0.05", italic: false, sub: false, sup: false }] }];
  assert.deepEqual(findFormatIssues(ps, {}), []);
  assert.deepEqual(
    findFormatIssues(ps, profile).map((f) => f.rule),
    ["stat"],
  );
});
