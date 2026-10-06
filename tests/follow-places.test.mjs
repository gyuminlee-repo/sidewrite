import { test } from "node:test";
import assert from "node:assert/strict";
import {
  changedIndexes,
  figurePointers,
  followPlaces,
  followState,
} from "../taskpane/shared/follow-places.js";
import { headingAbove } from "../taskpane/shared/paragraph-label.js";
import { paragraphKinds } from "../taskpane/shared/named-terms.js";

// A made-up inspection report of a bridge.
const PARAGRAPHS = [
  "Inspection of the Harbour Road Bridge", // 0 heading
  "Abstract", // 1 heading
  "All bearings were replaced and corrosion was eliminated across the deck.", // 2 body
  "Methods", // 3 heading
  "Bearing wear was graded by ultrasonic gauging at each pier.", // 4 methods
  "A span was scored as eliminated when no rust was visible on it.", // 5 methods
  "Results", // 6 heading
  "Corrosion on the deck", // 7 heading
  "Corrosion was eliminated on every span after recoating (Fig. 2B and Table S1).", // 8 body
  "The coating also held on the northern approach.", // 9 body
  "Cable tension", // 10 heading
  "Cable tension stayed within limits (Figs. 3 and 4).", // 11 body
  "Figure legends", // 12 heading
  "Figure 2. Rust coverage per span before and after recoating.", // 13 caption
  "Panel B marks the spans where corrosion was eliminated.", // 14 body, goes on from 13
  "Fig. 3. Cable tension by year.", // 15 caption
  "Table S1. Coating thickness per span.", // 16 caption
  "Supplementary Figure 4. Position of the tension gauges.", // 17 caption
  "References", // 18 heading
  "Okafor T. Coatings for steel bridges, eliminated rust. J Bridge Maint. 2018.", // 19
];
const IDS = PARAGRAPHS.map((_, i) => `k${i}`);
const change = (id, index, find, text, more = {}) => ({
  id,
  kind: "replace",
  find,
  text,
  state: "pending",
  paragraph_index: index,
  paragraph_id: `k${index}`,
  ...more,
});
const COMMENT = {
  ref: "R1.1",
  text: "Saying corrosion was “eliminated” is too strong for one season of data.",
  changes: [change("R1.1-1", 8, "was eliminated", "was reduced")],
};
const indexes = (places) => places.map((p) => p.index);

test("headingAbove gives the nearest heading above, or -1", () => {
  const kinds = paragraphKinds(PARAGRAPHS);
  assert.equal(headingAbove(kinds, 8), 7);
  assert.equal(headingAbove(kinds, 7), 6);
  assert.equal(headingAbove(kinds, 0), -1);
});

test("figurePointers reads plurals, panel letters, S numbers and ranges", () => {
  assert.deepEqual(figurePointers("as shown (Fig. 5C)"), ["fig:5"]);
  assert.deepEqual(figurePointers("Figs. 2 and 3"), ["fig:2", "fig:3"]);
  assert.deepEqual(figurePointers("Figure 7A–C and figures 9-11"), [
    "fig:7",
    "fig:9",
    "fig:10",
    "fig:11",
  ]);
  assert.deepEqual(figurePointers("Table S5, Fig. S7 and Tables 1, 2"), [
    "tab:S5",
    "fig:S7",
    "tab:1",
    "tab:2",
  ]);
  assert.deepEqual(figurePointers("Supplementary Figure 2; Supplementary Table S3; Scheme 1"), [
    "fig:S2",
    "tab:S3",
    "sch:1",
  ]);
  // Not a pointer: a number that follows no figure word, a figure word alone.
  assert.deepEqual(figurePointers("The figure shows 12 spans; configure 3 gauges."), []);
});

test("changedIndexes: the paragraph of the id, else of the index; live changes only", () => {
  const c = {
    changes: [
      change("a", 2, "x", "y", { paragraph_id: "k8" }),
      change("b", 4, "x", "y", { paragraph_id: "gone" }),
      change("c", 9, "x", "y", { state: "rejected" }),
      change("d", 11, "x", "y", { state: "unlocated" }),
      change("e", 99, "x", "y", { paragraph_id: null }),
    ],
  };
  assert.deepEqual(changedIndexes(c, IDS, PARAGRAPHS.length), [4, 8]);
});

test("followPlaces lists heading, legends, methods and section, in that order", () => {
  const places = followPlaces(COMMENT, PARAGRAPHS, IDS);
  assert.deepEqual(
    places.map((p) => [p.index, p.id, p.kind, p.why]),
    [
      [7, "k7", "heading", "heading above a changed paragraph"],
      [13, "k13", "caption", "legend of Fig. 2"],
      [14, "k14", "body", "continues the legend of Fig. 2"],
      [16, "k16", "caption", "legend of Table S1"],
      [5, "k5", "methods", 'methods; holds "eliminated"'],
      [9, "k9", "body", "same section as a changed paragraph"],
    ],
  );
  // The changed paragraph itself, the abstract (another section) and the
  // reference list are not places; without ids a place is named by its index.
  assert.equal(followPlaces(COMMENT, PARAGRAPHS)[0].id, "p7");
});

test("followPlaces: a figure named only in the comment or in what a change took out", () => {
  const inComment = { ...COMMENT, text: "See Figure 3A: the claim is too strong." };
  assert.ok(indexes(followPlaces(inComment, PARAGRAPHS, IDS)).includes(15));
  const removed = {
    ...COMMENT,
    changes: [change("R1.1-1", 9, "held (Supplementary Fig. 4)", "held")],
  };
  // "Supplementary Figure 4" is the legend of S4.
  const got = followPlaces(removed, PARAGRAPHS, IDS);
  assert.equal(got.find((p) => p.index === 17)?.why, "legend of Fig. S4");
  // Fig. 4 without the S is another figure: the paragraph that cites it gets none.
  const plain = {
    ref: "R1.2",
    text: "Limits?",
    changes: [change("R1.2-1", 11, "within", "under")],
  };
  assert.ok(!indexes(followPlaces(plain, PARAGRAPHS, IDS)).includes(17));
  assert.ok(indexes(followPlaces(plain, PARAGRAPHS, IDS)).includes(15));
});

test("followPlaces: a legend goes on for a few paragraphs, up to the next caption or heading", () => {
  const texts = [
    "Deck",
    "Rust fell after recoating (Fig. 1).",
    "Figure 1. Rust coverage per span.",
    "Filled bars are the spans coated in spring.",
    "",
    "Open bars are the spans left as they were.",
    "Error bars give the range of three surveys.",
    "This fourth paragraph after the caption is body text again.",
    "Figure 2. Another figure.",
    "Its legend goes on here.",
  ];
  const c = { ref: "R1.1", text: "Too strong.", changes: [change("x", 1, "fell", "was lower")] };
  const got = followPlaces(c, texts);
  assert.deepEqual(
    got.filter((p) => /legend/.test(p.why)).map((p) => p.index),
    [2, 3, 5, 6],
  );
  // Paragraphs 7 and 9 are still under the heading "Deck": listed as of the
  // section. The empty paragraph is no place.
  for (const i of [7, 9])
    assert.equal(got.find((p) => p.index === i).why, "same section as a changed paragraph");
  assert.ok(!indexes(got).includes(4));
});

test("followPlaces: methods paragraphs with a word a change took out, unless it is everywhere", () => {
  const texts = [
    "Methods",
    "Joints were sealed with a permanent membrane and checked each spring.",
    "Drains were cleared by hand.",
    "Results",
    "The permanent seal kept the joints dry.",
    "Joints stayed dry in both winters.",
  ];
  const c = {
    ref: "R2.1",
    text: "The durability statement goes too far.",
    changes: [change("R2.1-1", 4, "permanent seal", "seal")],
  };
  assert.deepEqual(
    followPlaces(c, texts).map((p) => [p.index, p.why]),
    [
      [3, "heading above a changed paragraph"],
      [1, 'methods; holds "permanent", which a change took out'],
      [5, "same section as a changed paragraph"],
    ],
  );
  // The word in more than a fifth of the paragraphs (never fewer than 3): vocabulary.
  const many = [...texts, "A permanent anchor.", "Permanent bolts hold.", "No permanent mark."];
  assert.ok(!indexes(followPlaces(c, many)).includes(1));
});

test("followPlaces: no live change, no places; changed paragraphs are never places", () => {
  assert.deepEqual(followPlaces({ ...COMMENT, changes: [] }, PARAGRAPHS, IDS), []);
  const rejected = {
    ...COMMENT,
    changes: [change("R1.1-1", 8, "was eliminated", "was reduced", { state: "rejected" })],
  };
  assert.deepEqual(followPlaces(rejected, PARAGRAPHS, IDS), []);
  const more = {
    ...COMMENT,
    changes: [
      ...COMMENT.changes,
      change("R1.1-2", 9, "also held", "held"),
      change("R1.1-3", 13, "Rust coverage", "Rust cover"),
    ],
  };
  const got = indexes(followPlaces(more, PARAGRAPHS, IDS));
  assert.deepEqual(got, [7, 12, 14, 16, 5]);
});

test("followPlaces: the reference list is skipped, also past an entry that reads as a heading", () => {
  const texts = [
    "Deck",
    "Rust fell after recoating (Fig. 1).",
    "References",
    "Okafor T. Coatings for steel bridges. J Bridge Maint. 2018.",
    "Lindqvist A, Maro P. Rust on decks. doi:10.0000/abc123", // no full stop: a heading by kind
    "Yeo K. Drainage of expansion joints. Road Eng. 2020.",
    "Figure legends",
    "Figure 1. Rust cover per span.",
    "Bars give the range of three surveys.",
  ];
  assert.equal(paragraphKinds(texts)[4], "heading");
  const c = {
    ref: "R1.1",
    text: "Too strong.",
    changes: [change("a", 1, "fell", "was lower"), change("b", 5, "2020", "2021")],
  };
  assert.deepEqual(indexes(followPlaces(c, texts)), [0, 7, 8]);
});

test("followPlaces: each paragraph once, and the list is cut from the end", () => {
  // Paragraph 5 holds the term and stands under a heading of a changed
  // paragraph: listed once, as the methods place.
  const c = { ...COMMENT, changes: [...COMMENT.changes, change("R1.1-2", 4, "graded", "rated")] };
  const got = followPlaces(c, PARAGRAPHS, IDS);
  assert.equal(got.filter((p) => p.index === 5).length, 1);
  assert.equal(got.find((p) => p.index === 5).why, 'methods; holds "eliminated"');
  assert.deepEqual(indexes(followPlaces(COMMENT, PARAGRAPHS, IDS, { maxPlaces: 2 })), [7, 13]);
  const chars = PARAGRAPHS[7].length + PARAGRAPHS[13].length + PARAGRAPHS[14].length;
  assert.deepEqual(
    indexes(followPlaces(COMMENT, PARAGRAPHS, IDS, { maxChars: chars })),
    [7, 13, 14],
  );
  // The defaults: 20 paragraphs.
  const long = ["Deck", "Rust fell.", ...Array.from({ length: 30 }, (_, i) => `Span ${i} held.`)];
  const one = { ref: "R1.1", text: "Why?", changes: [change("x", 1, "fell", "was lower")] };
  assert.equal(followPlaces(one, long).length, 20);
});

test("followState: the button is offered until the check ran, and again after a new change", () => {
  const snapshot = { paragraphs: PARAGRAPHS, ids: IDS };
  assert.deepEqual(followState(COMMENT, snapshot), { offer: 6, line: null });
  assert.deepEqual(followState({ ...COMMENT, changes: [] }, snapshot), { offer: null, line: null });
  const ran = {
    ...COMMENT,
    changes: [...COMMENT.changes, change("R1.1-2", 13, "Rust coverage", "Rust cover")],
    follow_through: {
      at: "2026-01-02T03:04:05.000Z",
      candidate_ids: ["k7", "k13", "k14", "k16", "k5", "k9"],
      change_ids: ["R1.1-2"],
      questions: 2,
      seen_change_ids: ["R1.1-1", "R1.1-2"],
    },
  };
  assert.deepEqual(followState(ran, snapshot), {
    offer: null,
    line: "Other places checked: 6 (1 change, 2 questions).",
  });
  const again = { ...ran, changes: [...ran.changes, change("R1.1-3", 11, "within", "under")] };
  const s = followState(again, snapshot);
  assert.equal(s.line, "Other places checked: 6 (1 change, 2 questions).");
  assert.ok(s.offer > 0);
  // A file from before the field: as if no check ran.
  assert.equal(followState({ ...COMMENT, follow_through: null }, snapshot).offer, 6);
});
