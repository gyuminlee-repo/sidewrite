import { test } from "node:test";
import assert from "node:assert/strict";
import { buildLetterModel, letterPlainText } from "../daemon/response-model.mjs";
import { withDefaults } from "../daemon/journal-profile.mjs";

const profile = withDefaults({ colors: { response_body: "AA0000", response_revised: "00AA00" } });
const rev = {
  comments: [
    {
      ref: "R1.1",
      reviewer: 1,
      text: "Please state the unit.\nUse SI.",
      declined: null,
      changes: [],
      response: {
        text: "We now give the unit.",
        gloss: "",
        declined_requirements: [],
        citations: [
          {
            kind: "revised",
            section: "2. Materials and Methods",
            quote: "The temperature was held at 30 °C.",
            lines: { start: 9, end: 9 },
          },
          { kind: "removed", section: "1. Introduction", quote: "", lines: null },
        ],
      },
    },
    {
      ref: "R2.1",
      reviewer: 2,
      text: "Name the antifoam.",
      declined: { reason: "No antifoam agent was used." },
      changes: [],
      response: null,
    },
  ],
};
const base = (o) => ({
  declined: null,
  changes: [],
  response: { text: "Done.", gloss: "", declined_requirements: [], citations: [] },
  ...o,
});
const rows = (m) => m.paragraphs.map((p) => [p.kind, p.text, p.color]);

test("blocks: reviewer head, comment verbatim, response, location and quote, declined line", () => {
  const m = buildLetterModel(rev, profile);
  assert.deepEqual(rows(m), [
    ["reviewer_head", "Reviewer #1", null],
    ["comment_head", "Comment 1", null],
    ["comment", "Please state the unit.", null],
    ["comment", "Use SI.", null],
    ["response", "We now give the unit.", "AA0000"],
    ["location", "[Revised in 2. Materials and Methods, line 9]", "00AA00"],
    ["quote", "The temperature was held at 30 °C.", "00AA00"],
    ["location", "[Removed from 1. Introduction]", "00AA00"],
    ["reviewer_head", "Reviewer #2", null],
    ["comment_head", "Comment 1", null],
    ["comment", "Name the antifoam.", null],
    ["response", "No antifoam agent was used.", "AA0000"],
    ["declined", "[No change made]", "AA0000"],
  ]);
  // A removal has no quote after it; the docx verifier reads this flag.
  assert.deepEqual(
    m.paragraphs.filter((p) => p.kind === "location").map((p) => p.removed),
    [false, true],
  );
  assert.deepEqual(
    m.paragraphs.map((p) => [p.ref, p.reviewer]).filter(([ref]) => ref === "R2.1").length,
    4,
  );
});

test("plain text per reviewer for a portal box", () => {
  const t = letterPlainText(buildLetterModel(rev, profile), 2);
  assert.match(t, /^Comment 1\nName the antifoam\./);
  assert.ok(!t.includes("Reviewer #1") && !t.includes("state the unit"));
  assert.equal(letterPlainText(buildLetterModel(rev, profile), 7), "");
});

test("editor comments come first as an Editor block with their own numbering", () => {
  const r = {
    comments: [
      base({ ref: "R1.1", reviewer: 1, text: "c1" }),
      base({ ref: "E.1", reviewer: 0, text: "Shorten the title." }),
      base({ ref: "E.2", reviewer: 0, text: "Check the references." }),
    ],
  };
  const m = buildLetterModel(r, profile);
  assert.deepEqual(
    m.paragraphs.slice(0, 7).map((p) => [p.kind, p.text]),
    [
      ["reviewer_head", "Editor"],
      ["comment_head", "Comment 1"],
      ["comment", "Shorten the title."],
      ["response", "Done."],
      ["comment_head", "Comment 2"],
      ["comment", "Check the references."],
      ["response", "Done."],
    ],
  );
  assert.equal(m.paragraphs[7].text, "Reviewer #1");
  assert.equal(m.paragraphs[0].reviewer, 0);
  // The stored order of the comments is not changed.
  assert.deepEqual(
    r.comments.map((c) => c.ref),
    ["R1.1", "E.1", "E.2"],
  );
});

test("a section label is written once where the comment section changes", () => {
  const r = {
    comments: [
      base({ ref: "R1.1", reviewer: 1, text: "a", section: "Major comments" }),
      base({ ref: "R1.2", reviewer: 1, text: "b", section: "Major comments" }),
      base({ ref: "R1.3", reviewer: 1, text: "c", section: "Minor comments" }),
      base({ ref: "R2.1", reviewer: 2, text: "d" }),
    ],
  };
  const heads = buildLetterModel(r, profile)
    .paragraphs.filter((p) => p.kind === "section_head" || p.kind === "reviewer_head")
    .map((p) => p.text);
  assert.deepEqual(heads, ["Reviewer #1", "Major comments", "Minor comments", "Reviewer #2"]);
});

test("the same section name under the next reviewer is written again", () => {
  const r = {
    comments: [
      base({ ref: "R1.1", reviewer: 1, text: "a", section: "Major comments" }),
      base({ ref: "R2.1", reviewer: 2, text: "b", section: "Major comments" }),
    ],
  };
  const kinds = buildLetterModel(r, profile).paragraphs.map((p) => p.kind);
  assert.equal(kinds.filter((k) => k === "section_head").length, 2);
});

test("section_head null in the profile writes no section label; plain text keeps the labels", () => {
  const off = { ...profile, response_letter: { ...profile.response_letter, section_head: null } };
  const r = {
    comments: [base({ ref: "R1.1", reviewer: 1, text: "a", section: "Major comments" })],
  };
  assert.equal(
    buildLetterModel(r, off).paragraphs.some((p) => p.kind === "section_head"),
    false,
  );
  assert.match(
    letterPlainText(buildLetterModel(r, profile), 1),
    /^Major comments\nComment 1\na\nDone\.$/,
  );
  assert.match(
    letterPlainText(
      buildLetterModel({ comments: [base({ ref: "E.1", reviewer: 0, text: "t" })] }, profile),
      0,
    ),
    /^Comment 1\nt/,
  );
});

test("labels, frames and line wording come from the profile", () => {
  const de = withDefaults({
    colors: { response_body: "AA0000", response_revised: "00AA00" },
    response_letter: {
      reviewer_label: "Gutachter {n}",
      comment_label: "Punkt {n}",
      location_line: "[Geändert in {section}, {lines}]",
      line_range: "Zeilen {a}-{b}",
    },
  });
  const r = {
    comments: [
      base({
        ref: "R3.12",
        reviewer: 3,
        text: "x",
        response: {
          text: "Erledigt.",
          citations: [
            { kind: "revised", section: "2. Methoden", quote: "Q.", lines: { start: 4, end: 6 } },
          ],
        },
      }),
    ],
  };
  assert.deepEqual(
    buildLetterModel(r, de).paragraphs.map((p) => p.text),
    ["Gutachter 3", "Punkt 12", "x", "Erledigt.", "[Geändert in 2. Methoden, Zeilen 4-6]", "Q."],
  );
});

test("every line of a response is its own paragraph; blank lines of a comment are dropped", () => {
  const r = {
    comments: [
      base({
        ref: "R1.1",
        reviewer: 1,
        text: "First.\r\n\r\n  Second, indented.",
        response: { text: "One.\n\nTwo.\nThree.", citations: [] },
      }),
    ],
  };
  const m = buildLetterModel(r, profile);
  assert.deepEqual(
    m.paragraphs.filter((p) => p.kind === "comment").map((p) => p.text),
    ["First.", "  Second, indented."],
  );
  assert.deepEqual(
    m.paragraphs.filter((p) => p.kind === "response").map((p) => p.text),
    ["One.", "Two.", "Three."],
  );
  assert.ok(m.paragraphs.every((p) => !/[\r\n]/.test(p.text)));
});

test("a citation that could not be placed is not written; a comment without an answer has none", () => {
  const r = {
    comments: [
      base({
        ref: "R1.1",
        reviewer: 1,
        text: "x",
        response: {
          text: "Done.",
          citations: [
            { kind: "revised", section: null, quote: "", lines: null, problem: "unlocated" },
          ],
        },
      }),
      base({ ref: "R1.2", reviewer: 1, text: "y", response: null }),
    ],
  };
  assert.deepEqual(
    buildLetterModel(r, profile).paragraphs.map((p) => p.kind),
    ["reviewer_head", "comment_head", "comment", "response", "comment_head", "comment"],
  );
});

test("a declined comment with its own answer keeps the answer; with changes it has no 'no change' line", () => {
  const r = {
    comments: [
      base({
        ref: "R1.1",
        reviewer: 1,
        text: "x",
        declined: { reason: "Out of scope." },
        response: { text: "We respectfully keep the wording.", citations: [] },
      }),
      base({
        ref: "R1.2",
        reviewer: 1,
        text: "y",
        declined: { reason: "Partly." },
        response: {
          text: "",
          citations: [{ kind: "revised", section: "S", quote: "Q.", lines: null }],
        },
      }),
    ],
  };
  assert.deepEqual(
    buildLetterModel(r, profile)
      .paragraphs.filter((p) => !["reviewer_head", "comment_head", "comment"].includes(p.kind))
      .map((p) => [p.kind, p.text]),
    [
      ["response", "We respectfully keep the wording."],
      ["response", "Out of scope."],
      ["declined", "[No change made]"],
      ["response", "Partly."],
      ["location", "[Revised in S]"],
      ["quote", "Q."],
    ],
  );
});

test("the reason of each declined request is written after the answer, in the answer colour", () => {
  const comment = (declined_requirements, text = "We now state the unit.") =>
    base({
      ref: "R1.1",
      reviewer: 1,
      text: "State the unit, name the supplier and the lot.",
      requirements: [
        { id: "q1", text: "state the unit" },
        { id: "q2", text: "name the supplier" },
        { id: "q3", text: "name the lot" },
      ],
      response: { text, gloss: "Übersetzung.", declined_requirements, citations: [] },
    });
  const body = (c, opts) =>
    buildLetterModel({ comments: [c] }, profile, opts)
      .paragraphs.slice(3)
      .map((p) => [p.kind, p.text, p.color]);
  // In the order of the requests, not of the stored list.
  const two = comment([
    { id: "q3", reason: "The lot number was not kept." },
    { id: "q2", reason: "The supplier was not recorded." },
  ]);
  assert.deepEqual(body(two), [
    ["response", "We now state the unit.", "AA0000"],
    ["response", "The supplier was not recorded.", "AA0000"],
    ["response", "The lot number was not kept.", "AA0000"],
  ]);
  // The portal text holds them too; no "no change" line: the comment is not declined.
  assert.match(
    letterPlainText(buildLetterModel({ comments: [two] }, profile), 1),
    /not recorded\./,
  );
  // Before the translation paragraph.
  assert.deepEqual(
    body(two, { gloss: true }).map(([kind]) => kind),
    ["response", "response", "response", "gloss"],
  );
  // A reason the answer already gives word for word is not written twice.
  const said = comment(
    [{ id: "q2", reason: "The supplier  was not recorded." }],
    "We now state the unit. The supplier was not recorded.",
  );
  assert.deepEqual(body(said), [
    ["response", "We now state the unit. The supplier was not recorded.", "AA0000"],
  ]);
  // No reason, or a request that is no longer in the list: nothing to write.
  const none = comment([
    { id: "q2", reason: "  " },
    { id: "q9", reason: "Left over from a request that was removed." },
  ]);
  assert.deepEqual(body(none), [["response", "We now state the unit.", "AA0000"]]);
  // With no answer text at all the reasons are the answer.
  const only = comment([{ id: "q2", reason: "The supplier was not recorded." }], "");
  assert.deepEqual(body(only), [["response", "The supplier was not recorded.", "AA0000"]]);
});

test("the translation paragraph is written only when asked for", () => {
  const p = withDefaults({
    colors: { response_body: "AA0000", response_kr: "7F7F7F" },
    response_letter: { gloss_prefix: "[T] " },
  });
  const r = {
    comments: [
      base({
        ref: "R1.1",
        reviewer: 1,
        text: "x",
        response: { text: "Done.", gloss: "Erledigt.", citations: [] },
      }),
    ],
  };
  assert.equal(
    buildLetterModel(r, p).paragraphs.some((x) => x.kind === "gloss"),
    false,
  );
  const g = buildLetterModel(r, p, { gloss: true }).paragraphs.map((x) => [
    x.kind,
    x.text,
    x.color,
  ]);
  assert.deepEqual(g.slice(3), [
    ["response", "Done.", "AA0000"],
    ["gloss", "[T] Erledigt.", "7F7F7F"],
  ]);
});
