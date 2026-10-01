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

test("sentenceAt gives the sentence that holds the offset, or null", () => {
  const t = viewText(snap.paragraphs[1]);
  assert.equal(sentenceAt(t, 0), "Flasks were shaken.");
  assert.equal(sentenceAt(t, t.indexOf("Samples")), "Samples were drawn daily.");
  assert.equal(sentenceAt(t, t.length + 5), null);
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
