import { test } from "node:test";
import assert from "node:assert/strict";
import {
  splitSentences,
  extractRequirements,
  requirementPrompt,
  REQUIREMENT_SYSTEM,
  REQUIREMENT_TYPES,
} from "../daemon/review-split.mjs";

const squash = (s) => String(s).replace(/\s+/g, " ").trim();

// ---- sentence splitter -------------------------------------------------------

test("sentences: abbreviations, figure numbers and decimals do not end a sentence", () => {
  assert.deepEqual(
    splitSentences(
      "Controls are missing, e.g. a knockout strain. See Fig. 2 and Figs. 3-4. " +
        "Smith et al. reported p < 0.05 here. Compare A vs. B, i.e. the two runs. Well done.",
    ),
    [
      "Controls are missing, e.g. a knockout strain.",
      "See Fig. 2 and Figs. 3-4.",
      "Smith et al. reported p < 0.05 here.",
      "Compare A vs. B, i.e. the two runs.",
      "Well done.",
    ],
  );
});

test("sentences: 'et al.' before a capital and 'Fig.' before a capital stay whole", () => {
  assert.deepEqual(splitSentences("As Lee et al. Nature showed, see Fig. S1. Add it."), [
    "As Lee et al. Nature showed, see Fig. S1.",
    "Add it.",
  ]);
});

test("sentences: a numbered list head is not cut off as its own sentence", () => {
  assert.deepEqual(
    splitSentences("Two points:\n1. Report n.\n2) State the unit.\n3. Add error bars."),
    ["Two points:", "Report n.", "State the unit.", "Add error bars."],
  );
  assert.deepEqual(splitSentences("Two points: 1. Report n. 2. State the unit."), [
    "Two points: 1. Report n.",
    "State the unit.",
  ]);
});

test("sentences: a wrapped line is one sentence; a blank line ends one", () => {
  assert.deepEqual(splitSentences("Please report the growth\ntemperature used\n\nThanks"), [
    "Please report the growth temperature used",
    "Thanks",
  ]);
});

test("sentences: question and exclamation marks end sentences; no letter, no unit", () => {
  assert.deepEqual(splitSentences("Why 30 °C? Explain! --- 12."), ["Why 30 °C?", "Explain!"]);
  assert.deepEqual(splitSentences("  \n "), []);
});

// ---- extraction --------------------------------------------------------------

const C = (ref, text, requirements = []) => ({
  ref,
  reviewer: ref.startsWith("E") ? 0 : Number(ref.slice(1, ref.indexOf("."))),
  text,
  requirements,
});

// Fake oneshot: label sentences by ref -> [types in order].
const labelled =
  (byRef, calls = []) =>
  async (req) => {
    calls.push(req);
    return {
      ok: true,
      json: {
        sentences: Object.entries(byRef).flatMap(([ref, types]) =>
          types.map((type, i) => ({ ref, s: i + 1, type })),
        ),
      },
    };
  };

test("request sentences become requirements verbatim; praise is left out", async () => {
  const comments = [
    C(
      "R1.1",
      "This is a well written study. Please report the growth temperature.\n" +
        "Could the authors clarify how n was chosen? I enjoyed reading it.",
    ),
    C("R1.2", "Thank you."),
  ];
  const calls = [];
  const r = await extractRequirements(comments, {
    oneshot: labelled(
      { "R1.1": ["other", "request", "request", "other"], "R1.2": ["other"] },
      calls,
    ),
  });
  assert.equal(calls.length, 1);
  assert.equal(calls[0].system, REQUIREMENT_SYSTEM);
  assert.deepEqual(r.comments[0].requirements, [
    { id: "q1", text: "Please report the growth temperature." },
    { id: "q2", text: "Could the authors clarify how n was chosen?" },
  ]);
  assert.deepEqual(r.comments[1].requirements, []);
  assert.deepEqual(r.warnings, []);
  assert.deepEqual(r.requests, { "R1.1": 2, "R1.2": 0 });
  assert.equal(r.asked, true);
  // Comment text and refs are not touched.
  assert.deepEqual(
    r.comments.map((c) => [c.ref, c.text]),
    comments.map((c) => [c.ref, c.text]),
  );
});

test("the prompt lists ref and sentence numbers and asks for types only", () => {
  const p = requirementPrompt([{ ref: "R2.1", sentences: ["Add n.", "Nice."] }]);
  assert.match(p, /\{"sentences":\[\{"ref":"R1\.2","s":1,"type":"request"\}/);
  assert.match(p, /\[R2\.1 s1\] Add n\./);
  assert.match(p, /\[R2\.1 s2\] Nice\./);
  assert.deepEqual(REQUIREMENT_TYPES, ["request", "other"]);
  assert.match(REQUIREMENT_SYSTEM, /Do not return any text/);
});

test("comments that already have requirements are skipped and kept", async () => {
  const own = [{ id: "q3", text: "use Celsius" }];
  const calls = [];
  const r = await extractRequirements([C("R1.1", "Please add units.", own), C("R1.2", "Add n.")], {
    oneshot: labelled({ "R1.2": ["request"] }, calls),
  });
  assert.doesNotMatch(calls[0].prompt, /R1\.1/);
  assert.deepEqual(r.comments[0].requirements, own);
  assert.deepEqual(r.comments[1].requirements, [{ id: "q1", text: "Add n." }]);
  assert.deepEqual(r.requests, { "R1.2": 1 });
});

test("nothing to label asks no model; no oneshot leaves requirements empty", async () => {
  let called = false;
  const oneshot = async () => ((called = true), { ok: true, json: { sentences: [] } });
  const done = await extractRequirements([C("R1.1", "x", [{ id: "q1", text: "x" }])], { oneshot });
  assert.equal(called, false);
  assert.equal(done.asked, false);
  const none = await extractRequirements([C("R1.1", "Add n.")], {});
  assert.deepEqual(none.comments[0].requirements, []);
  assert.deepEqual(none.warnings, []);
  assert.equal(none.asked, false);
});

const TWO = [C("R1.1", "Nice work. Please add n."), C("R1.2", "Report the unit.")];
const WARN = "model could not label requirement sentences; add them by hand";
const BAD = {
  "a sentence missing": {
    sentences: [
      { ref: "R1.1", s: 1, type: "other" },
      { ref: "R1.1", s: 2, type: "request" },
    ],
  },
  "a sentence doubled": {
    sentences: [
      { ref: "R1.1", s: 1, type: "other" },
      { ref: "R1.1", s: 2, type: "request" },
      { ref: "R1.2", s: 1, type: "request" },
      { ref: "R1.2", s: 1, type: "other" },
    ],
  },
  "a type outside the choices": {
    sentences: [
      { ref: "R1.1", s: 1, type: "other" },
      { ref: "R1.1", s: 2, type: "question" },
      { ref: "R1.2", s: 1, type: "request" },
    ],
  },
  "an unknown ref": {
    sentences: [
      { ref: "R1.1", s: 1, type: "other" },
      { ref: "R1.1", s: 2, type: "request" },
      { ref: "R9.9", s: 1, type: "request" },
    ],
  },
  "text instead of labels": { requirements: ["Please add n."] },
};
for (const [name, json] of Object.entries(BAD)) {
  test(`a reply with ${name} leaves requirements empty and warns`, async () => {
    const r = await extractRequirements(TWO, { oneshot: async () => ({ ok: true, json }) });
    assert.deepEqual(
      r.comments.map((c) => [c.ref, c.text, c.requirements]),
      TWO.map((c) => [c.ref, c.text, []]),
    );
    assert.deepEqual(r.warnings, [WARN]);
    assert.equal(r.requests, null);
  });
}

test("a failed model call leaves requirements empty and warns", async () => {
  const r = await extractRequirements(TWO, {
    oneshot: async () => ({ ok: false, error: "timed out" }),
  });
  assert.ok(r.comments.every((c) => c.requirements.length === 0));
  assert.equal(r.warnings.length, 1);
  assert.match(r.warnings[0], /timed out.*add them by hand/);
});

// Property: whatever the labels, every request text is inside its comment
// (whitespace normalized), and the requests are exactly the request labels.
function mulberry32(seed) {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
const PIECES = [
  "Please report n.",
  "e.g. a knockout",
  "see Fig. 2",
  "Lee et al. showed it",
  "p < 0.05",
  "Could the authors clarify this?",
  "Great work!",
  "\n3. Add error bars.",
  "\n\n",
  "\n  ",
  "A vs. B",
  "i.e. the same",
];

test("property: request texts are substrings of their comment text", async () => {
  const rand = mulberry32(20260930);
  for (let run = 0; run < 100; run++) {
    const count = 1 + Math.floor(rand() * 4);
    const comments = [];
    for (let k = 0; k < count; k++) {
      const n = 1 + Math.floor(rand() * 8);
      const text = Array.from({ length: n }, () => PIECES[Math.floor(rand() * PIECES.length)]).join(
        " ",
      );
      comments.push(C(`R1.${k + 1}`, text));
    }
    const byRef = {};
    for (const c of comments) {
      byRef[c.ref] = splitSentences(c.text).map(() => (rand() < 0.5 ? "request" : "other"));
    }
    const r = await extractRequirements(comments, { oneshot: labelled(byRef) });
    comments.forEach((c, i) => {
      const got = r.comments[i].requirements.map((q) => q.text);
      for (const t of got) assert.ok(squash(c.text).includes(squash(t)), `run ${run}: ${t}`);
      const want = [
        ...new Set(splitSentences(c.text).filter((_, j) => byRef[c.ref][j] === "request")),
      ];
      assert.deepEqual(got, want, `run ${run} ${c.ref}`);
    });
  }
});
