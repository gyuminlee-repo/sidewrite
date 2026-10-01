import { test } from "node:test";
import assert from "node:assert/strict";
import {
  citationsFor,
  sectionOf,
  staleReasons,
  snapshotHash,
  keepLines,
  sentenceAt,
  changeOffset,
  viewText,
} from "../daemon/response-quotes.mjs";

const snap = {
  paragraphs: [
    { index: 0, text: "2. Materials and Methods" },
    {
      index: 1,
      text: "Flasks were shaken. The incubation temperature was held at 30 °C. Samples were drawn daily.",
    },
    { index: 2, text: "3. Results" },
    {
      index: 3,
      text: "Growth was followed by optical density at 600 nm. The plateau was reached on day four.",
    },
  ],
};
const heads = [
  { index: 0, level: 1, text: "2. Materials and Methods" },
  { index: 2, level: 1, text: "3. Results" },
];
const ch = (o) => ({
  id: "R1.1-1",
  kind: "replace",
  find: "degrees",
  text: "°C",
  state: "accepted",
  paragraph_index: 1,
  context: { left: "held at 30", right: "Samples were drawn" },
  ...o,
});

test("section is the nearest heading above", () => {
  assert.equal(sectionOf(1, heads), "2. Materials and Methods");
  assert.equal(sectionOf(3, heads), "3. Results");
  assert.equal(sectionOf(0, [{ index: 2, level: 1, text: "3. Results" }]), null);
});

test("section_level keeps to headings of that level or above", () => {
  const nested = [
    { index: 0, level: 1, text: "2. Methods" },
    { index: 4, level: 2, text: "2.1 Strains" },
    { index: 8, level: 3, text: "2.1.1 Media" },
  ];
  assert.equal(sectionOf(9, nested), "2.1.1 Media");
  assert.equal(sectionOf(9, nested, 2), "2.1 Strains");
  assert.equal(sectionOf(9, nested, 1), "2. Methods");
});

test("an accepted change cites the whole sentence it sits in, verbatim", () => {
  const [c] = citationsFor({ changes: [ch({})] }, snap, heads);
  assert.equal(c.quote, "The incubation temperature was held at 30 °C.");
  assert.ok(snap.paragraphs[1].text.includes(c.quote));
  assert.deepEqual(
    [c.kind, c.section, c.change_ids, c.paragraph_index, c.lines, c.problem],
    ["revised", "2. Materials and Methods", ["R1.1-1"], 1, null, null],
  );
});

test("the quote is read from the accepted view, not from the raw paragraph text", () => {
  const pending = structuredClone(snap);
  pending.paragraphs[1].reviewed_text = pending.paragraphs[1].text;
  pending.paragraphs[1].text =
    "Flasks were shaken. The incubation temperature was held at 30 °Cdegrees. Samples were drawn daily.";
  const [c] = citationsFor({ changes: [ch({})] }, pending, heads);
  assert.equal(c.quote, "The incubation temperature was held at 30 °C.");
});

test("pending and rejected changes are not cited; two changes in one sentence give one citation", () => {
  assert.equal(
    citationsFor(
      {
        changes: [ch({ state: "pending" }), ch({ state: "rejected" }), ch({ state: "unlocated" })],
      },
      snap,
      heads,
    ).length,
    0,
  );
  const two = citationsFor(
    {
      changes: [
        ch({}),
        ch({
          id: "R1.1-2",
          find: "kept",
          text: "held",
          context: { left: "temperature was", right: "at 30" },
        }),
      ],
    },
    snap,
    heads,
  );
  assert.deepEqual(
    two.map((c) => c.change_ids),
    [["R1.1-1", "R1.1-2"]],
  );
});

test("a change edited by the author is cited where its context still is", () => {
  const edited = structuredClone(snap);
  edited.paragraphs[1].text = edited.paragraphs[1].text.replace("30 °C", "30 degrees Celsius");
  const [c] = citationsFor({ changes: [ch({ state: "user_edited" })] }, edited, heads);
  assert.equal(c.quote, "The incubation temperature was held at 30 degrees Celsius.");
});

test("a deleted sentence is a removal with no quote; an unlocated change says so", () => {
  const del = ch({
    kind: "delete",
    text: "",
    find: "The flasks came from a supplier.",
    context: { left: "", right: "" },
  });
  assert.equal(citationsFor({ changes: [del] }, snap, heads)[0].kind, "removed");
  assert.equal(
    citationsFor({ changes: [ch({ paragraph_index: null })] }, snap, heads)[0].problem,
    "unlocated",
  );
  assert.equal(
    citationsFor({ changes: [ch({ paragraph_index: 99 })] }, snap, heads)[0].problem,
    "unlocated",
  );
});

test("a whole sentence deleted between two others is a removal, not a quote of its neighbour", () => {
  // Context as the daemon records it: up to three words on each side.
  const del = ch({
    kind: "delete",
    text: "",
    find: "The flasks came from a supplier. ",
    context: { left: "Flasks were shaken.", right: "The incubation temperature" },
  });
  const [c] = citationsFor({ changes: [del] }, snap, heads);
  assert.deepEqual(
    [c.kind, c.quote, c.section, c.problem],
    ["removed", "", "2. Materials and Methods", null],
  );
});

test("words deleted inside a sentence cite the sentence that remains", () => {
  const del = ch({
    kind: "delete",
    text: "",
    find: "carefully ",
    context: { left: "Samples were", right: "drawn daily." },
  });
  const [c] = citationsFor({ changes: [del] }, snap, heads);
  assert.deepEqual([c.kind, c.quote], ["revised", "Samples were drawn daily."]);
});

test("a change that cannot be placed in its paragraph is reported, never guessed", () => {
  const twice = {
    paragraphs: [{ index: 0, text: "The yield was 5 g. Later the yield was 5 g again." }],
  };
  const [c] = citationsFor(
    {
      changes: [ch({ paragraph_index: 0, text: "5 g", context: { left: "yield was", right: "" } })],
    },
    twice,
    [],
  );
  assert.deepEqual([c.problem, c.quote, c.section], ["sentence not found", "", null]);
  assert.equal(changeOffset("a b a", { kind: "replace", text: "a", context: null }), null);
});

// ---- a wrong quote is worse than none ----------------------------------------
//
// A unit change in the first sentence; the same unit also stands in the last.
const UNIT = {
  kind: "replace",
  find: "degrees",
  text: "°C",
  state: "user_edited",
  context: { left: "kept at 30 ", right: " overnight" },
};
const FIRST = "Cells were kept at 30 degrees Celsius overnight.";
const OTHER = "Storage was at 4 °C.";
const where = (text, change) => {
  const off = changeOffset(text, change);
  return off == null ? null : sentenceAt(text, off);
};

test("the new text found once elsewhere is not the change when its context is still in place", () => {
  // The author wrote the unit out again; "°C" now only stands in another sentence.
  assert.equal(where(`${FIRST} ${OTHER}`, UNIT), FIRST);
  const [c] = citationsFor(
    { changes: [{ ...UNIT, id: "R1.1-1", paragraph_index: 0 }] },
    { paragraphs: [{ index: 0, text: `${FIRST} ${OTHER}` }] },
    [],
  );
  assert.deepEqual([c.quote, c.problem], [FIRST, null]);
});

test("with both context sides and the new text in place, a second use elsewhere changes nothing", () => {
  const accepted = "Cells were kept at 30 °C overnight.";
  assert.equal(where(`${accepted} ${OTHER}`, { ...UNIT, state: "accepted" }), accepted);
  assert.equal(where(`${OTHER} ${accepted}`, { ...UNIT, state: "accepted" }), accepted);
});

test("one context side still in place is enough to stay in the sentence of the change", () => {
  const leftOnly = "Cells were kept at 30 degrees Celsius for one night.";
  assert.equal(where(`${leftOnly} ${OTHER}`, UNIT), leftOnly);
  const rightOnly = "Cells stayed near 30 degrees Celsius overnight.";
  assert.equal(where(`${rightOnly} ${OTHER}`, UNIT), rightOnly);
  // The right side alone, where the rewritten text ends its sentence.
  const ends = { ...UNIT, context: { left: "kept at 30 ", right: " Storage was at" } };
  const rewritten = "Cells stayed near 30 degrees Celsius.";
  assert.equal(where(`${rewritten} ${OTHER}`, ends), rewritten);
});

test("a sentence rewritten as a whole cannot be placed, whatever stands elsewhere", () => {
  const text = `The culture was grown warm for one night. ${OTHER}`;
  assert.equal(changeOffset(text, UNIT), null);
  const [c] = citationsFor(
    { changes: [{ ...UNIT, id: "R1.1-1", paragraph_index: 0 }] },
    { paragraphs: [{ index: 0, text }] },
    [],
  );
  assert.deepEqual([c.quote, c.problem], ["", "sentence not found"]);
});

test("context sides that point at two sentences place nothing", () => {
  assert.equal(
    changeOffset("Cells were kept at 30 units. Storage at 4 °C was overnight.", UNIT),
    null,
  );
});

test("the old text back between its context words means the change is not there", () => {
  // Reverted to the letter: "degrees" stands again where "°C" was put.
  assert.equal(changeOffset(`Cells were kept at 30 degrees overnight. ${OTHER}`, UNIT), null);
  // Also when the new form turns up in another sentence.
  assert.equal(
    changeOffset(
      "Cells were kept at 30 degrees overnight. Controls were kept at 30 °C overnight.",
      UNIT,
    ),
    null,
  );
});

test("a context word counts as a whole word only", () => {
  // "at" is inside "separate"; the place of the change is gone.
  const short = { ...UNIT, context: { left: "at ", right: "" } };
  assert.equal(changeOffset("Each separate flask was cooled.", short), null);
  assert.equal(
    where("Each flask was held at room temperature.", short),
    "Each flask was held at room temperature.",
  );
});

test("without recorded context the new text alone counts only for an accepted change", () => {
  const bare = { kind: "replace", find: "degrees", text: "°C", context: null };
  const text = `Cells were grown overnight. ${OTHER}`;
  // accepted without context: the paragraph reads exactly as after the change.
  assert.equal(where(text, { ...bare, state: "accepted" }), OTHER);
  assert.equal(changeOffset(text, { ...bare, state: "user_edited" }), null);
  assert.equal(changeOffset(text, bare), null);
  assert.equal(changeOffset(`Kept at 30 °C. ${OTHER}`, { ...bare, state: "accepted" }), null);
  assert.equal(
    changeOffset(text, { ...bare, state: "accepted", context: { left: "", right: "" } }),
    text.indexOf("°C"),
  );
});

test("a deletion: taken back, in effect, and with one context side left", () => {
  const del = {
    kind: "delete",
    find: "approximately ",
    text: "",
    state: "user_edited",
    context: { left: "were kept at ", right: "30 °C." },
  };
  // Taken back by the author, and the shortened form stands in another sentence.
  assert.equal(
    changeOffset("Cells were kept at approximately 30 °C. Controls were kept at 30 °C.", del),
    null,
  );
  // In effect.
  assert.equal(
    where("Cells were kept at 30 °C. Controls were stored cold.", del),
    "Cells were kept at 30 °C.",
  );
  // The author rewrote what follows; the words before the deletion are still there.
  assert.equal(
    where("Cells were kept at room temperature. Controls were stored cold.", del),
    "Cells were kept at room temperature.",
  );
  // Nothing of it left.
  assert.equal(changeOffset("Cells grew warm. Controls were stored cold.", del), null);
});

test("sentenceAt gives the sentence that holds the offset, or null", () => {
  const t = viewText(snap.paragraphs[1]);
  assert.equal(sentenceAt(t, 0), "Flasks were shaken.");
  assert.equal(sentenceAt(t, t.indexOf("Samples")), "Samples were drawn daily.");
  assert.equal(sentenceAt(t, t.length + 5), null);
});

// ---- typed list numbers and bullets in the manuscript -------------------------

test("a sentence keeps the typed number or bullet in front of it, as it stands in the manuscript", () => {
  const t = "1. Cells were kept at 30 °C overnight. 2. Samples were drawn daily.";
  const first = "1. Cells were kept at 30 °C overnight.";
  const second = "2. Samples were drawn daily.";
  // Every character belongs to a sentence, the number included.
  assert.equal(sentenceAt(t, 0), first);
  assert.equal(sentenceAt(t, 1), first);
  assert.equal(sentenceAt(t, t.indexOf("Cells")), first);
  assert.equal(sentenceAt(t, t.indexOf("2.")), second);
  assert.equal(sentenceAt(t, t.indexOf("daily")), second);
  for (const mark of ["- ", "* ", "• ", "(a) ", "b) ", "(12) "]) {
    const item = `${mark}Cells were kept at 30 °C.`;
    assert.equal(sentenceAt(item, 0), item, mark);
    assert.equal(sentenceAt(item, item.length - 1), item, mark);
  }
  // Without such a mark nothing changes.
  assert.equal(sentenceAt("Cells were kept. Samples were drawn.", 0), "Cells were kept.");
  // A unit without a letter after the last sentence stays with that sentence.
  const tail = "Cells were kept at 30 °C. (2)";
  assert.equal(sentenceAt(tail, tail.length - 2), tail);
  assert.equal(sentenceAt("12.", 0), null);
});

test("a change whose range holds the list number is quoted with it", () => {
  const cite = (text, change) =>
    citationsFor(
      { changes: [{ id: "R1.1-1", paragraph_index: 0, state: "accepted", ...change }] },
      { paragraphs: [{ index: 0, text }] },
      [],
    ).map((c) => [c.kind, c.quote, c.problem]);
  // The whole paragraph was replaced: no context words on either side.
  const whole = "1. Cells were kept at 30 °C.";
  assert.deepEqual(
    cite(whole, {
      kind: "replace",
      find: "1. Cells were kept at 30 degrees.",
      text: whole,
      context: { left: "", right: "" },
    }),
    [["revised", whole, null]],
  );
  // The range starts at the number; the words after it are the context.
  const text = "1. Cultures were kept at 30 °C. 2. Samples were drawn daily.";
  assert.deepEqual(
    cite(text, {
      kind: "replace",
      find: "1. Cells",
      text: "1. Cultures",
      context: { left: "", right: " were kept at" },
    }),
    [["revised", "1. Cultures were kept at 30 °C.", null]],
  );
  // A change inside the second item; the quote is in the paragraph letter for letter.
  const [[, quote]] = cite(text, {
    kind: "replace",
    find: "taken",
    text: "drawn",
    context: { left: "2. Samples were ", right: " daily." },
  });
  assert.equal(quote, "2. Samples were drawn daily.");
  assert.ok(text.includes(quote));
  // The context sides of one item agree on its sentence although one lies in the number.
  assert.equal(
    changeOffset("1. Flasks were kept at 30 °C. 2. Samples were drawn daily.", {
      kind: "replace",
      find: "Cells",
      text: "Cultures",
      state: "user_edited",
      context: { left: "1. ", right: " were kept at" },
    }),
    3,
  );
});

test("a quote with its list number is stale once the number is gone", () => {
  const text = "1. Cells were kept at 30 °C.";
  const c = {
    changes: [
      {
        id: "R1.1-1",
        kind: "replace",
        find: "degrees",
        text: "°C",
        state: "accepted",
        paragraph_index: 0,
        context: { left: "kept at 30 ", right: "" },
      },
    ],
  };
  const snapshot = { paragraphs: [{ index: 0, text }] };
  c.response = { citations: citationsFor(c, snapshot, []) };
  assert.equal(c.response.citations[0].quote, text);
  assert.deepEqual(staleReasons(c, snapshot, null), []);
  assert.deepEqual(
    staleReasons(c, { paragraphs: [{ index: 0, text: "Cells were kept at 30 °C." }] }, null),
    ["quote_changed"],
  );
});

test("keepLines carries a line range over only while the quote is the same", () => {
  const old = [
    {
      kind: "revised",
      quote: "A.",
      paragraph_index: 1,
      lines: { start: 4, end: 4 },
      lines_pdf: "sha",
    },
  ];
  const fresh = [
    { kind: "revised", quote: "A.", paragraph_index: 1, lines: null, lines_pdf: null },
    { kind: "revised", quote: "B.", paragraph_index: 1, lines: null, lines_pdf: null },
  ];
  const out = keepLines(fresh, old);
  assert.deepEqual(
    out.map((c) => [c.lines, c.lines_pdf]),
    [
      [{ start: 4, end: 4 }, "sha"],
      [null, null],
    ],
  );
  assert.deepEqual(keepLines(fresh), fresh);
});

test("stale: quote edited in Word, change no longer accepted, new manuscript text since the PDF", () => {
  const cits = citationsFor({ changes: [ch({})] }, snap, heads);
  const c = { changes: [ch({})], response: { citations: cits, stale: [] } };
  const render = { snapshot_hash: snapshotHash(snap) };
  assert.deepEqual(staleReasons(c, snap, render), []);
  const edited = structuredClone(snap);
  edited.paragraphs[1].text = edited.paragraphs[1].text.replace("held", "kept");
  assert.deepEqual(staleReasons(c, edited, render), ["quote_changed", "lines_old"]);
  assert.deepEqual(staleReasons({ ...c, changes: [ch({ state: "rejected" })] }, snap, render), [
    "change_state",
  ]);
  assert.deepEqual(staleReasons({ changes: [ch({})], response: null }, edited, render), []);
  assert.deepEqual(staleReasons(c, edited, null), ["quote_changed"]);
});

test("snapshotHash follows the accepted text and ignores spacing", () => {
  const spaced = structuredClone(snap);
  spaced.paragraphs[1].text = spaced.paragraphs[1].text.replace("were shaken", "were  shaken");
  assert.equal(snapshotHash(spaced), snapshotHash(snap));
  const moved = { paragraphs: [...snap.paragraphs].reverse() };
  assert.notEqual(snapshotHash(moved), snapshotHash(snap));
});
