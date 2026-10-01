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
  ADVISORY,
  candidatesFromLabels,
  CLAUSE_BREAKS,
  DIRECTION,
  LIMIT_NEGATION,
  ORDER_DOUBT,
  ORDER_KEEPS,
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

// ---- a sentence that names several sections
//
// The number must stand in the clause of the section the rule names, and that
// clause must name no other section.

const TWO = "The abstract must not exceed 250 words, and the methods must not exceed 1000 words.";
const LISTED = "Abstract: at most 250 words; Introduction: at most 600 words.";
const RANGE_THEN =
  "The abstract should be between 150 and 250 words, but the discussion may have up to 1500 words.";

const TIED = [
  [TWO, words("abstract", { max: 250 })],
  [TWO, words("methods", { max: 1000 })],
  [LISTED, words("abstract", { max: 250 })],
  [LISTED, words("introduction", { max: 600 })],
  [
    "The abstract is limited to 250 words, whereas the Materials and Methods section is limited to 1,500 words.",
    words("methods", { max: 1500 }),
  ],
  [RANGE_THEN, words("abstract", { min: 150, max: 250 })],
  [RANGE_THEN, words("discussion", { max: 1500 })],
  [
    "The introduction has at most 600 words; the results and the discussion are free.",
    words("introduction", { max: 600 }),
  ],
  [
    "Unlike the summary, which is short, the conclusions have at most 400 words.",
    words("conclusion", { max: 400 }),
  ],
  // "resultset" is not a section: only one section is named, as before.
  [
    "The abstract must not exceed 250 words, and the resultset must not exceed 1000 words.",
    words("abstract", { max: 250 }),
  ],
];

const NOT_TIED = {
  "the number of the other section": [TWO, words("abstract", { max: 1000 })],
  "the number of the other section, the other way round": [TWO, words("methods", { max: 250 })],
  "the number of the other entry of a list": [LISTED, words("abstract", { max: 600 })],
  "the upper end of one section's range under another section": [
    RANGE_THEN,
    words("discussion", { max: 250 }),
  ],
  "two sections joined by and, first one": [
    "The abstract and the introduction must not exceed 250 words.",
    words("abstract", { max: 250 }),
  ],
  "two sections joined by and, second one": [
    "The abstract and the introduction must not exceed 250 words.",
    words("introduction", { max: 250 }),
  ],
  "sections counted together": [
    "The results and discussion sections together must not exceed 3000 words.",
    words("results", { max: 3000 }),
  ],
  "a list of sections with one limit": [
    "Introduction, Methods, and Results must not exceed 3000 words.",
    words("results", { max: 3000 }),
  ],
  "a clause with the number that names no section": [
    "The methods must follow the abstract, which must not exceed 250 words.",
    words("methods", { max: 250 }),
  ],
  "a break word inside another word leaves both sections in one clause": [
    "The abstract must not exceed 250 words andor the methods must not exceed 1000 words.",
    words("abstract", { max: 250 }),
  ],
};

test("several sections: at least five cases the code must accept and five it must not", () => {
  assert.ok(TIED.length >= 5 && Object.keys(NOT_TIED).length >= 5);
});

for (const [sentence, rule] of TIED) {
  test(`a count rule stands among several sections: ${JSON.stringify(rule)} in ${JSON.stringify(sentence)}`, () => {
    const r = one(sentence, rule);
    assert.deepEqual(r.warnings, []);
    assert.equal(r.candidates[0].check.type, "count");
    assert.equal(r.candidates[0].check.section, SECTION_PATTERNS[rule.section]);
  });
}

for (const [name, [sentence, rule]] of Object.entries(NOT_TIED)) {
  test(`no count rule from ${name}`, () => {
    const r = one(sentence, rule);
    assert.deepEqual(r.candidates[0].check, { type: "llm" });
    assert.equal(r.warnings.length, 1);
    assert.match(r.warnings[0], /^s1: the sentence names several sections \(/);
  });
}

test("the clause breaks are a list the code reads; each one separates two sections, and only as a whole word", () => {
  assert.ok(CLAUSE_BREAKS.marks.length >= 2 && CLAUSE_BREAKS.words.length >= 3);
  const both = (sep) =>
    `The abstract must not exceed 250 words${sep}the methods must not exceed 1000 words.`;
  for (const mark of CLAUSE_BREAKS.marks) {
    assert.equal(kind(both(`${mark} `), words("abstract", { max: 250 })), "count", mark);
    assert.equal(kind(both(`${mark} `), words("abstract", { max: 1000 })), "llm", mark);
  }
  for (const word of CLAUSE_BREAKS.words) {
    assert.equal(kind(both(` ${word} `), words("abstract", { max: 250 })), "count", word);
    assert.equal(
      kind(both(` ${word.toUpperCase()} `), words("methods", { max: 1000 })),
      "count",
      word,
    );
    assert.equal(kind(both(` ${word} `), words("methods", { max: 250 })), "llm", word);
    assert.equal(kind(both(` x${word}x `), words("abstract", { max: 250 })), "llm", word);
  }
});

// ---- the order of sections
//
// The order the model gives must be the order in which the sentence names the
// sections, and the sentence must have no word after which that order may not
// be the required one.

const order = (...names) => ({ target: "section_order", order: names });
const PLAIN = "Arrange the text as Methods, then Results.";
const FULL =
  "The manuscript should be organised in the following order: Introduction, Materials and Methods, Results, Discussion.";

const IN_ORDER = [
  [PLAIN, order("methods", "results")],
  [FULL, order("introduction", "methods", "results", "discussion")],
  [FULL, order("introduction", "discussion")],
  [
    "The Introduction is followed by the Methods and the Results.",
    order("introduction", "methods", "results"),
  ],
  [
    "Structure the paper as follows: Background, Methods, Results.",
    order("introduction", "methods", "results"),
  ],
  // A section named twice counts where it is named first.
  [
    "Arrange the text as Methods, then Results; the Methods are short.",
    order("methods", "results"),
  ],
  // "hereafter" and "Notes" hold "after" and "not", but not as words.
  ["Arrange the text hereafter as Methods, then Results.", order("methods", "results")],
  ["Notes: Methods, then Results.", order("methods", "results")],
];

const NOT_IN_ORDER = {
  "the reverse of the order in the sentence": [PLAIN, order("results", "methods"), /order/],
  "two of four sections swapped": [FULL, order("introduction", "results", "methods"), /order/],
  "follow, in the order the sentence means": [
    "Results should follow Methods.",
    order("methods", "results"),
    /"follow"/,
  ],
  "follow, in the order of appearance": [
    "Results should follow Methods.",
    order("results", "methods"),
    /"follow"/,
  ],
  "after, as it reads": [
    "The Methods come after the Results.",
    order("methods", "results"),
    /"after"/,
  ],
  "after, as it means": [
    "The Methods come after the Results.",
    order("results", "methods"),
    /"after"/,
  ],
  "preceded by": [
    "The Results are preceded by the Methods.",
    order("results", "methods"),
    /"preceded"/,
  ],
  "before at the start": [
    "Before the Results, describe the Methods.",
    order("results", "methods"),
    /"before"/,
  ],
  "a negation": [
    "The Methods must not come ahead of the Results.",
    order("methods", "results"),
    /"not"/,
  ],
  "first and last": [
    "Place the Results last and the Methods first.",
    order("results", "methods"),
    /"(?:first|last)"/,
  ],
  "section names inside other words": [
    "Arrange the methodsless text as resultsets.",
    order("methods", "results"),
    /not in the sentence/,
  ],
};

test("section order: at least five cases the code must accept and five it must not", () => {
  assert.ok(IN_ORDER.length >= 5 && Object.keys(NOT_IN_ORDER).length >= 5);
});

for (const [sentence, rule] of IN_ORDER) {
  test(`a section order stands: ${rule.order.join(" < ")} in ${JSON.stringify(sentence)}`, () => {
    const r = one(sentence, rule);
    assert.deepEqual(r.warnings, []);
    assert.deepEqual(r.candidates[0].check, { type: "count", ...rule });
  });
}

for (const [name, [sentence, rule, why]] of Object.entries(NOT_IN_ORDER)) {
  test(`no section order from ${name}`, () => {
    const r = one(sentence, rule);
    assert.deepEqual(r.candidates[0].check, { type: "llm" });
    assert.equal(r.warnings.length, 1);
    assert.match(r.warnings[0], /^s1: /);
    assert.match(r.warnings[0], why);
  });
}

test("the words that put the order in doubt are a list the code reads; each one works, and only as a whole word", () => {
  assert.ok(ORDER_DOUBT.length >= 5 && ORDER_KEEPS.length >= 2);
  const rule = order("methods", "results");
  for (const word of ORDER_DOUBT) {
    assert.equal(kind(`Methods, then Results (${word}).`, rule), "llm", word);
    assert.equal(kind(`Methods, then Results (${word.toUpperCase()}).`, rule), "llm", word);
    assert.equal(kind(`Methods, then Results (x${word}x).`, rule), "count", word);
  }
  // A phrase that keeps the order holds a doubt word and is still accepted.
  for (const phrase of ORDER_KEEPS) {
    assert.ok(
      ORDER_DOUBT.some((word) => phrase.split(" ").includes(word)),
      phrase,
    );
    assert.equal(kind(`Methods, then Results (${phrase}).`, rule), "count", phrase);
  }
});

// ---- the same number twice in one sentence
//
// A range confirms a direction only for the number that stands in it, not for
// the same value written somewhere else in the sentence.

const SAME_TWICE =
  "The abstract needs 150 to 250 words and the introduction needs at least 250 words.";
const SAME_TWICE_KEYS = "Provide 3 to 5 keywords and at least 5 tables.";

const OWN_RANGE = [
  [SAME_TWICE, words("abstract", { min: 150, max: 250 })],
  [SAME_TWICE, words("introduction", { min: 250 })],
  [SAME_TWICE_KEYS, { target: "keywords", min: 3, max: 5 }],
  [SAME_TWICE_KEYS, { target: "tables", min: 5 }],
  [
    "The summary has at least 200 words; the conclusions have 100 to 200 words.",
    words("conclusion", { min: 100, max: 200 }),
  ],
];

const OTHER_RANGE = {
  "the upper end of another section's range": [SAME_TWICE, words("introduction", { max: 250 })],
  "the upper end of the keyword range for tables": [SAME_TWICE_KEYS, { target: "tables", max: 5 }],
  "the upper end of a later range": [
    "The summary has at least 200 words; the conclusions have 100 to 200 words.",
    words("abstract", { max: 200 }),
  ],
  "the lower end of an earlier range": [
    "Provide 4 to 6 keywords, and no more than 4 tables.",
    { target: "tables", min: 4 },
  ],
  "the upper end of a range read as a minimum next to another number": [
    "Use 2 to 4 tables and at least 4 keywords.",
    { target: "tables", min: 4 },
  ],
};

test("the same number twice: at least five cases the code must accept and five it must not", () => {
  assert.ok(OWN_RANGE.length >= 5 && Object.keys(OTHER_RANGE).length >= 5);
});

for (const [sentence, rule] of OWN_RANGE) {
  test(`a count rule stands on its own number: ${JSON.stringify(rule)} in ${JSON.stringify(sentence)}`, () => {
    assert.deepEqual(one(sentence, rule).warnings, []);
    assert.equal(kind(sentence, rule), "count");
  });
}

for (const [name, [sentence, rule]] of Object.entries(OTHER_RANGE)) {
  test(`no count rule from ${name}`, () => {
    const r = one(sentence, rule);
    assert.deepEqual(r.candidates[0].check, { type: "llm" });
    assert.match(
      r.warnings[0],
      /^s1: the sentence does not say that \d+ is (?:an upper|a lower) limit/,
    );
  });
}

// ---- a limit that is denied, or a number that is only advice
//
// "No maximum" is not a maximum, and "about 250 words" is not a limit.

const FIRM = [
  [
    "There is no minimum length, but the abstract must not exceed 250 words.",
    words("abstract", { max: 250 }),
  ],
  [
    "The abstract must not exceed 250 words; a structured format is recommended.",
    words("abstract", { max: 250 }),
  ],
  [
    "The maximum length of the abstract is 250 words, not counting the title.",
    words("abstract", { max: 250 }),
  ],
  // "roundabout", "casino" and "notable" hold "about", "no" and "not", but not as words.
  ["The roundabout abstract must not exceed 250 words.", words("abstract", { max: 250 })],
  ["The casino maximum for the abstract is 250 words.", words("abstract", { max: 250 })],
  ["A notable limit of 6 keywords applies.", { target: "keywords", max: 6 }],
];

const NOT_FIRM = {
  "a maximum that is denied, with a suggested number": [
    "The abstract has no maximum length; 250 words is a suggestion.",
    words("abstract", { max: 250 }),
    /suggestion|denies/,
  ],
  "a limit that is denied": [
    "The abstract is not limited to 250 words.",
    words("abstract", { max: 250 }),
    /upper limit/,
  ],
  "no limit of": [
    "There is no limit of 250 words for the abstract.",
    words("abstract", { max: 250 }),
    /upper limit/,
  ],
  "a maximum that is denied, alone": [
    "The abstract has no maximum length, and 250 words is fine.",
    words("abstract", { max: 250 }),
    /upper limit/,
  ],
  "a typical length": [
    "The abstract is typically at most 250 words.",
    words("abstract", { max: 250 }),
    /"typically"/,
  ],
  "a suggested maximum": [
    "We suggest a maximum of 250 words for the abstract.",
    words("abstract", { max: 250 }),
    /"suggest"/,
  ],
  "an approximate number": [
    "Approximately 250 words is the recommended maximum for the abstract.",
    words("abstract", { max: 250 }),
    /"approximately"/,
  ],
  "a minimum that is denied": [
    "Without a minimum, the abstract runs to about 250 words.",
    words("abstract", { min: 250 }),
    /"about"|lower limit/,
  ],
  "an ideal number of keywords": [
    "About 6 keywords at most are ideal.",
    { target: "keywords", max: 6 },
    /"about"/,
  ],
};

test("denied limits and advice: at least five cases the code must accept and five it must not", () => {
  assert.ok(FIRM.length >= 5 && Object.keys(NOT_FIRM).length >= 5);
});

for (const [sentence, rule] of FIRM) {
  test(`a count rule stands as a firm limit: ${JSON.stringify(rule)} in ${JSON.stringify(sentence)}`, () => {
    assert.deepEqual(one(sentence, rule).warnings, []);
    assert.equal(kind(sentence, rule), "count");
  });
}

for (const [name, [sentence, rule, why]] of Object.entries(NOT_FIRM)) {
  test(`no count rule from ${name}`, () => {
    const r = one(sentence, rule);
    assert.deepEqual(r.candidates[0].check, { type: "llm" });
    assert.equal(r.warnings.length, 1);
    assert.match(r.warnings[0], why);
  });
}

test("the negations and the words of advice are lists the code reads; each one works, and only as a whole word", () => {
  assert.ok(LIMIT_NEGATION.length >= 3 && ADVISORY.length >= 6);
  for (const side of ["max", "min"]) {
    for (const phrase of DIRECTION[side].before) {
      for (const neg of LIMIT_NEGATION) {
        const denied = `The abstract: ${neg} ${phrase} 250 words.`;
        assert.equal(kind(denied, words("abstract", { [side]: 250 })), "llm", `${neg} ${phrase}`);
        const article = `The abstract: ${neg.toUpperCase()} a ${phrase} 250 words.`;
        assert.equal(
          kind(article, words("abstract", { [side]: 250 })),
          "llm",
          `${neg} a ${phrase}`,
        );
        const glued = `The abstract: x${neg} ${phrase} 250 words.`;
        assert.equal(kind(glued, words("abstract", { [side]: 250 })), "count", `x${neg} ${phrase}`);
      }
    }
  }
  for (const word of ADVISORY) {
    const rule = words("abstract", { max: 250 });
    assert.equal(kind(`The abstract (${word}): at most 250 words.`, rule), "llm", word);
    assert.equal(
      kind(`The abstract (${word.toUpperCase()}): at most 250 words.`, rule),
      "llm",
      word,
    );
    assert.equal(kind(`The abstract (x${word}x): at most 250 words.`, rule), "count", word);
    // In another clause the word says nothing about this number.
    assert.equal(
      kind(`The abstract: at most 250 words; ${word} is another matter.`, rule),
      "count",
      word,
    );
  }
});
