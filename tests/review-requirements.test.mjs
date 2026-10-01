import { test } from "node:test";
import assert from "node:assert/strict";
import {
  splitSentences,
  clauseCuts,
  requirementParts,
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
          types.map((type, i) =>
            typeof type === "string" ? { ref, s: i + 1, type } : { ref, s: i + 1, ...type },
          ),
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
  assert.match(REQUIREMENT_SYSTEM, /cuts and types where cut marks are shown/);
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
  // Class audit (Codex review 13): sentence and clause numbers.
  "a sentence number out of range": {
    sentences: [
      { ref: "R1.1", s: 1, type: "other" },
      { ref: "R1.1", s: 3, type: "request" },
      { ref: "R1.2", s: 1, type: "request" },
    ],
  },
  "a sentence number zero": {
    sentences: [
      { ref: "R1.1", s: 0, type: "other" },
      { ref: "R1.1", s: 2, type: "request" },
      { ref: "R1.2", s: 1, type: "request" },
    ],
  },
  "a sentence number as text": {
    sentences: [
      { ref: "R1.1", s: "1", type: "other" },
      { ref: "R1.1", s: 2, type: "request" },
      { ref: "R1.2", s: 1, type: "request" },
    ],
  },
  "cuts on a sentence without cut marks": {
    sentences: [
      { ref: "R1.1", s: 1, type: "other" },
      { ref: "R1.1", s: 2, cuts: ["split"], types: ["request", "request"] },
      { ref: "R1.2", s: 1, type: "request" },
    ],
  },
  "an entry that is not an object": {
    sentences: [{ ref: "R1.1", s: 1, type: "other" }, null, { ref: "R1.2", s: 1, type: "request" }],
  },
  "no sentences list": { sentences: "none" },
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
  "Please state the unit, and describe the method.",
  "Report n; give the spread (mean, and SD) too.",
  "State the mean and also the median.",
  "Name the agent and explain its role.",
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
    const type = () => (rand() < 0.5 ? "request" : "other");
    for (const c of comments) {
      byRef[c.ref] = splitSentences(c.text).map((sentence) => {
        const k = clauseCuts(sentence).length;
        if (!k) return type();
        const cuts = Array.from({ length: k }, () => (rand() < 0.5 ? "split" : "keep"));
        return { cuts, types: [type(), ...cuts.filter((x) => x === "split").map(type)] };
      });
    }
    const r = await extractRequirements(comments, { oneshot: labelled(byRef) });
    comments.forEach((c, i) => {
      const got = r.comments[i].requirements.map((q) => q.text);
      for (const t of got) assert.ok(squash(c.text).includes(squash(t)), `run ${run}: ${t}`);
      const want = [
        ...new Set(
          splitSentences(c.text).flatMap((sentence, j) => {
            const l = byRef[c.ref][j];
            const parts = requirementParts(sentence, typeof l === "string" ? [] : l.cuts);
            const types = typeof l === "string" ? [l] : l.types;
            for (const part of parts) assert.ok(sentence.includes(part), `run ${run}: ${part}`);
            return parts.filter((_, n) => types[n] === "request");
          }),
        ),
      ];
      assert.deepEqual(got, want, `run ${run} ${c.ref}`);
    });
  }
});

// ---- S2-30: more than one request in one sentence ---------------------------

const R13 =
  "In the first paragraph of section 3, please state how growth was followed (the Methods say optical density was read at 600 nm), and describe the synthetic carbon source as a novel substrate.";

test("clause cuts: coordinators at depth 0 only; the parts are pieces of the sentence", () => {
  assert.equal(clauseCuts(R13).length, 1);
  assert.deepEqual(requirementParts(R13, ["split"]), [
    "In the first paragraph of section 3, please state how growth was followed (the Methods say optical density was read at 600 nm)",
    "describe the synthetic carbon source as a novel substrate.",
  ]);
  assert.deepEqual(requirementParts(R13, ["keep"]), [R13]);
  // Kept from the 2026-10-01 acceptance run: one request, no cut point.
  assert.deepEqual(
    clauseCuts(
      "In Materials and Methods, please name the antifoam agent that was used and its concentration.",
    ),
    [],
  );
  assert.deepEqual(clauseCuts("Give the spread (mean, and SD; or range) of each run."), []);
  assert.deepEqual(requirementParts("Report n; state the unit.", ["split"]), [
    "Report n",
    "state the unit.",
  ]);
  assert.deepEqual(requirementParts("Report n and also state the unit.", ["split"]), [
    "Report n",
    "state the unit.",
  ]);
  assert.deepEqual(requirementParts("Name the agent and explain its role.", ["split"]), [
    "Name the agent",
    "explain its role.",
  ]);
  assert.deepEqual(clauseCuts("Name the agent and its supplier."), []);
  // An astral character before a bracket does not shift the depth lookup.
  assert.deepEqual(clauseCuts(`Use ${"\u{1D400}".repeat(20)} (n, and m) here.`), []);
  assert.deepEqual(requirementParts("Use \u{1D400} (n), and state the unit.", ["split"]), [
    "Use \u{1D400} (n)",
    "state the unit.",
  ]);
});

test("S2-30: the R1.3 sentence gives two requirements when the model splits it", async () => {
  const calls = [];
  const r = await extractRequirements([C("R1.3", R13)], {
    oneshot: labelled({ "R1.3": [{ cuts: ["split"], types: ["request", "request"] }] }, calls),
  });
  assert.match(calls[0].prompt, /\[R1\.3 s1\] .*600 nm\)<c1>, and describe/);
  assert.match(calls[0].prompt, /"cuts"/);
  assert.deepEqual(r.warnings, []);
  assert.deepEqual(r.comments[0].requirements, [
    {
      id: "q1",
      text: "In the first paragraph of section 3, please state how growth was followed (the Methods say optical density was read at 600 nm)",
    },
    { id: "q2", text: "describe the synthetic carbon source as a novel substrate." },
  ]);
  assert.deepEqual(r.requests, { "R1.3": 2 });
});

test("S2-30: one request with a cut point stays one when the model keeps it", async () => {
  const one = "Please state the mean, and the spread of the same measurement.";
  assert.equal(clauseCuts(one).length, 1);
  const r = await extractRequirements([C("R1.1", one)], {
    oneshot: labelled({ "R1.1": [{ cuts: ["keep"], types: ["request"] }] }),
  });
  assert.deepEqual(r.comments[0].requirements, [{ id: "q1", text: one }]);
});

for (const [name, label] of Object.entries({
  "no cuts for a sentence that has a cut point": { type: "request" },
  "too few cuts": { cuts: [], types: ["request"] },
  "a cut that is neither split nor keep": { cuts: ["maybe"], types: ["request"] },
  "types that do not match the parts": { cuts: ["split"], types: ["request"] },
})) {
  test(`S2-30: a reply with ${name} leaves requirements empty and warns`, async () => {
    const r = await extractRequirements([C("R1.3", R13)], {
      oneshot: labelled({ "R1.3": [label] }),
    });
    assert.deepEqual(r.comments[0].requirements, []);
    assert.deepEqual(r.warnings, [WARN]);
  });
}

// Class audit (Codex review 13): labels of a sentence with cut marks. One
// cut mark: exactly one decision, and one type per part after the splits.
const CUT_ONE = [C("R1.3", "Please state the unit, and describe the method.")];
const BAD_CUTS = {
  "a type instead of cuts": { type: "request" },
  "too many cut decisions": { cuts: ["split", "keep"], types: ["request", "request"] },
  "no cut decision": { cuts: [], types: ["request"] },
  "a cut decision outside the choices": { cuts: ["maybe"], types: ["request"] },
  "one type too few": { cuts: ["split"], types: ["request"] },
  "one type too many": { cuts: ["keep"], types: ["request", "other"] },
  "a part type outside the choices": { cuts: ["split"], types: ["request", "question"] },
  "cuts that are not a list": { cuts: "split", types: ["request", "request"] },
};
for (const [name, label] of Object.entries(BAD_CUTS)) {
  test(`a cut-mark reply with ${name} leaves requirements empty and warns`, async () => {
    assert.equal(clauseCuts(CUT_ONE[0].text).length, 1);
    const r = await extractRequirements(CUT_ONE, {
      oneshot: async () => ({ ok: true, json: { sentences: [{ ref: "R1.3", s: 1, ...label }] } }),
    });
    assert.deepEqual(r.comments[0].requirements, []);
    assert.deepEqual(r.warnings, [WARN]);
  });
}
