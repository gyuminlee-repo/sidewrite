// What a count rule must find in its own guideline sentence.
//
// A number alone proves nothing: the sentence must also name what is counted
// (the section, and the counted thing next to the number) and say in which
// direction the number limits it. Each case is one sentence, one label from
// the model, and the check type that must come out. The sentence itself is
// always kept as the rule text.

import { test } from "node:test";
import assert from "node:assert/strict";
import {
  candidatesFromLabels,
  DIRECTION,
  SECTION_TERMS,
  SECTION_PATTERNS,
  UNIT_WORDS,
} from "../daemon/guideline-extract.mjs";

const one = (sentence, rule) => {
  const r = candidatesFromLabels([sentence], { sentences: [{ s: 1, type: "count", rule }] });
  assert.equal(r.candidates.length, 1);
  assert.equal(r.candidates[0].text, sentence);
  assert.equal(r.candidates[0].source.sentence, sentence);
  return r;
};
const kind = (sentence, rule) => one(sentence, rule).candidates[0].check.type;
const words = (section, bounds) => ({ target: "section_words", section, ...bounds });

const SUPPORTED = [
  ["The abstract must not exceed 250 words.", words("abstract", { max: 250 })],
  ["Abstracts are limited to 250 words.", words("abstract", { max: 250 })],
  ["The abstract should be between 150 and 250 words.", words("abstract", { min: 150, max: 250 })],
  ["The abstract should be between 150 and 250 words.", words("abstract", { max: 250 })],
  ["The abstract should be between 150 and 250 words.", words("abstract", { min: 150 })],
  [
    "The Summary has at least 100 and at most 200 words.",
    words("abstract", { min: 100, max: 200 }),
  ],
  ["The maximum length of the introduction is 600 words.", words("introduction", { max: 600 })],
  ["A 250-word limit applies to the abstract.", words("abstract", { max: 250 })],
  ["The Materials and Methods section: no more than 1,500 words.", words("methods", { max: 1500 })],
  ["Provide 4 to 6 keywords.", { target: "keywords", min: 4, max: 6 }],
  ["Provide 4-6 key words.", { target: "keywords", min: 4, max: 6 }],
  ["Give at least 3 keywords.", { target: "keywords", min: 3 }],
  ["Keywords (maximum 6).", { target: "keywords", max: 6 }],
  ["List 6 keywords at most.", { target: "keywords", max: 6 }],
  ["A manuscript may contain at most 8 figures.", { target: "figures", max: 8 }],
  ["The number of figures must not exceed 8.", { target: "figures", max: 8 }],
  ["Up to 8 display figures are allowed.", { target: "figures", max: 8 }],
  ["No more than 3 tables are allowed.", { target: "tables", max: 3 }],
  ["At most 1 table may be included.", { target: "tables", max: 1 }],
  ["Include no fewer than 2 tables.", { target: "tables", min: 2 }],
];

// name: [sentence, label, what the warning must say]
const UNSUPPORTED = {
  "the same number under another section": [
    "The abstract must not exceed 250 words.",
    words("methods", { max: 250 }),
    /methods/,
  ],
  "the same number under another section, as a minimum": [
    "The abstract must not exceed 250 words.",
    words("methods", { min: 250 }),
    /methods/,
  ],
  "a maximum read as a minimum": [
    "The abstract must not exceed 250 words.",
    words("abstract", { min: 250 }),
    /lower limit/,
  ],
  "a minimum read as a maximum": [
    "Give at least 3 keywords.",
    { target: "keywords", max: 3 },
    /upper limit/,
  ],
  "a negated minimum phrase read as a maximum": [
    "Include no fewer than 2 tables.",
    { target: "tables", max: 2 },
    /upper limit/,
  ],
  "the lower end of a range read as the maximum": [
    "The abstract should be between 150 and 250 words.",
    words("abstract", { max: 150 }),
    /upper limit/,
  ],
  "the upper end of a range read as the minimum": [
    "Provide 4 to 6 keywords.",
    { target: "keywords", min: 6 },
    /lower limit/,
  ],
  "the two limits of one sentence swapped": [
    "The Summary has at least 100 and at most 200 words.",
    words("abstract", { max: 100 }),
    /upper limit/,
  ],
  "a number that counts something else": [
    "Each figure must not exceed 6 panels.",
    { target: "figures", max: 6 },
    /does not count figures/,
  ],
  "a number that counts parts of the thing": [
    "Use at most 6 figure panels.",
    { target: "figures", max: 6 },
    /does not count figures/,
  ],
  "characters, not words": [
    "The abstract must not exceed 250 characters.",
    words("abstract", { max: 250 }),
    /does not count words/,
  ],
  "figures and tables counted together": [
    "At most 8 figures and tables are allowed.",
    { target: "figures", max: 8 },
    /together/,
  ],
  "a table limit read as a figure limit": [
    "No more than 3 tables are allowed.",
    { target: "figures", max: 3 },
    /does not count figures/,
  ],
  "a section name inside another word": [
    "An abstraction must not exceed 250 words.",
    words("abstract", { max: 250 }),
    /abstract/,
  ],
  "another section name inside another word": [
    "The resultset must not exceed 250 words.",
    words("results", { max: 250 }),
    /results/,
  ],
  "a counted word inside another word": [
    "A manuscript may contain at most 8 subfigures.",
    { target: "figures", max: 8 },
    /does not count figures/,
  ],
  "a direction word inside another word": [
    "The abstract has unlimited 250 words.",
    words("abstract", { max: 250 }),
    /upper limit/,
  ],
  "no direction at all": [
    "The abstract should be 250 words.",
    words("abstract", { max: 250 }),
    /upper limit/,
  ],
  "a strict limit": [
    "The abstract must be shorter than 250 words.",
    words("abstract", { max: 250 }),
    /upper limit/,
  ],
  "a direction that belongs to the other number": [
    "Provide at least 4 keywords from the list of 30 keywords.",
    { target: "keywords", min: 30 },
    /lower limit/,
  ],
  "a reference limit read as a table limit": [
    "No more than 40 references are allowed.",
    { target: "tables", max: 40 },
    /does not count tables/,
  ],
};

test("the cases: at least five the code must accept and five it must not", () => {
  assert.ok(SUPPORTED.length >= 5 && Object.keys(UNSUPPORTED).length >= 5);
});

for (const [sentence, rule] of SUPPORTED) {
  test(`a count rule stands: ${JSON.stringify(rule)} in ${JSON.stringify(sentence)}`, () => {
    const r = one(sentence, rule);
    assert.deepEqual(r.warnings, []);
    const { section, ...rest } = rule;
    assert.deepEqual(r.candidates[0].check, {
      type: "count",
      ...rest,
      ...(section && { section: SECTION_PATTERNS[section] }),
    });
  });
}

for (const [name, [sentence, rule, why]] of Object.entries(UNSUPPORTED)) {
  test(`no count rule from ${name}: the sentence stays as a model-checked rule with a warning`, () => {
    const r = one(sentence, rule);
    assert.deepEqual(r.candidates[0].check, { type: "llm" });
    assert.equal(r.warnings.length, 1);
    assert.match(r.warnings[0], /^s1: /);
    assert.match(r.warnings[0], why);
  });
}

test("the direction words are a list the code reads; each one works, and only in its direction", () => {
  for (const side of ["max", "min"]) {
    assert.ok(DIRECTION[side].before.length >= 3 && DIRECTION[side].after.length >= 1, side);
  }
  for (const side of ["max", "min"]) {
    const other = side === "max" ? "min" : "max";
    for (const phrase of DIRECTION[side].before) {
      const sentence = `The abstract: ${phrase} 250 words.`;
      assert.equal(kind(sentence, words("abstract", { [side]: 250 })), "count", phrase);
      assert.equal(kind(sentence, words("abstract", { [other]: 250 })), "llm", phrase);
      // In capitals too, and not as a part of another word.
      const upper = `${phrase.toUpperCase()} 250 words in the abstract.`;
      assert.equal(kind(upper, words("abstract", { [side]: 250 })), "count", phrase);
      const glued = `The abstract: x${phrase} 250 words.`;
      assert.equal(kind(glued, words("abstract", { [side]: 250 })), "llm", phrase);
    }
    for (const phrase of DIRECTION[side].after) {
      const sentence = `The abstract has 250 words ${phrase}.`;
      assert.equal(kind(sentence, words("abstract", { [side]: 250 })), "count", phrase);
      assert.equal(kind(sentence, words("abstract", { [other]: 250 })), "llm", phrase);
    }
  }
  // No phrase of one direction ends a phrase of the other ("fewer than" in "no fewer than").
  const all = (side) => [...DIRECTION[side].before, ...DIRECTION[side].after];
  for (const a of all("max")) {
    for (const b of all("min")) assert.ok(!a.endsWith(b) && !b.endsWith(a), `${a} / ${b}`);
  }
});

test("every section and every counted thing has its words in a list the code reads", () => {
  assert.deepEqual(Object.keys(SECTION_TERMS).sort(), Object.keys(SECTION_PATTERNS).sort());
  for (const [section, terms] of Object.entries(SECTION_TERMS)) {
    for (const term of terms) {
      const sentence = `The ${term} must not exceed 300 words.`;
      assert.equal(kind(sentence, words(section, { max: 300 })), "count", term);
      const inside = `The x${term}x must not exceed 300 words.`;
      assert.equal(kind(inside, words(section, { max: 300 })), "llm", term);
    }
  }
  assert.deepEqual(Object.keys(UNIT_WORDS).sort(), [
    "figures",
    "keywords",
    "section_words",
    "tables",
  ]);
});

test("a section order names its sections as whole words", () => {
  const order = ["methods", "results"];
  const rule = { target: "section_order", order };
  assert.equal(kind("Arrange the text as Methods, then Results.", rule), "count");
  assert.equal(kind("Arrange the methodsless text as resultsets.", rule), "llm");
});
