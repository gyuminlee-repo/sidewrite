import { test } from "node:test";
import assert from "node:assert/strict";
import {
  MAX_CHARS,
  MAX_PLACES,
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
  // A comma before "and", a long dash for a range, the short forms of
  // "supplementary".
  assert.deepEqual(figurePointers("Figs. 2, 3, and 4"), ["fig:2", "fig:3", "fig:4"]);
  assert.deepEqual(figurePointers("Figs. 2\u20144 and Tables 6–7"), [
    "fig:2",
    "fig:3",
    "fig:4",
    "tab:6",
    "tab:7",
  ]);
  assert.deepEqual(figurePointers("Supplemental Fig. 2, Supp. Fig. 3 and Suppl. Table 1"), [
    "fig:S2",
    "fig:S3",
    "tab:S1",
  ]);
  // A range too long to be one is read as its two ends.
  assert.deepEqual(figurePointers("Figs. 1-40"), ["fig:1", "fig:40"]);
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
      [5, "k5", "methods", 'methods; also has "eliminated"'],
      [9, "k9", "body", "same section as a changed paragraph"],
    ],
  );
  // The changed paragraph itself, the abstract (another section) and the
  // reference list are not places; without ids a place is named by its index.
  assert.equal(followPlaces(COMMENT, PARAGRAPHS)[0].id, "p7");
});

test("followPlaces: a figure named only in the comment or in what a change took out", () => {
  const inComment = { ...COMMENT, text: "See Figure 3A: the claim is too strong." };
  // What the comment or a change names comes before what only the rest of
  // the changed paragraph cites.
  assert.deepEqual(indexes(followPlaces(inComment, PARAGRAPHS, IDS)).slice(0, 4), [7, 15, 13, 14]);
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

test("followPlaces: a legend goes on only where at most three paragraphs follow its caption", () => {
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
  const c = {
    ref: "R1.1",
    text: "Too strong (see Figure 2).",
    changes: [change("x", 1, "fell", "was lower")],
  };
  const got = followPlaces(c, texts);
  // Fig. 2 (named in the comment) first, with the one paragraph after its
  // caption; four paragraphs follow the caption of Fig. 1: body text.
  assert.deepEqual(
    got.filter((p) => /legend/.test(p.why)).map((p) => [p.index, p.why]),
    [
      [8, "legend of Fig. 2"],
      [9, "continues the legend of Fig. 2"],
      [2, "legend of Fig. 1"],
    ],
  );
  for (const i of [3, 5, 6, 7])
    assert.equal(got.find((p) => p.index === i).why, "same section as a changed paragraph");
  assert.ok(!indexes(got).includes(4));
});

test("followPlaces: the legends a changed paragraph cites do not crowd out the rest", () => {
  const texts = [
    "Deck",
    "Rust fell on every span (Figs. 1-10).",
    "The coating held.",
    ...Array.from({ length: 10 }, (_, i) => `Figure ${i + 1}. Span group ${i + 1}.`),
  ];
  const c = {
    ref: "R1.1",
    text: "Figure 10 does not show this.",
    changes: [change("x", 1, "fell", "was lower")],
  };
  // Fig. 10 (the comment names it), then seven of the others: eight legends.
  assert.deepEqual(indexes(followPlaces(c, texts)), [0, 12, 3, 4, 5, 6, 7, 8, 9, 2]);
});

test("followPlaces: the legends a paragraph cites stop at eight, their continuations counted", () => {
  const texts = [
    "Deck",
    "Rust fell on every span (Figs. 1-6).",
    ...Array.from({ length: 6 }, (_, i) => [
      `Figure ${i + 1}. Span group ${i + 1}.`,
      `Filled bars of group ${i + 1}.`,
      `Open bars of group ${i + 1}.`,
    ]).flat(),
  ];
  const c = { ref: "R1.1", text: "Too strong.", changes: [change("x", 1, "fell", "was lower")] };
  assert.deepEqual(
    followPlaces(c, texts)
      .filter((p) => /legend/.test(p.why))
      .map((p) => p.index),
    [2, 3, 4, 5, 6, 7, 8, 9],
  );
  // What the comment names is not held to that.
  const named = { ...c, text: "Figures 1-6 do not show this." };
  assert.equal(followPlaces(named, texts).filter((p) => /legend/.test(p.why)).length, 18);
});

test("followPlaces: a paragraph without words is no heading, no boundary and no place", () => {
  const texts = [
    "Deck survey",
    "/",
    "0.05",
    "Rust fell after recoating.",
    "12 / 14",
    "The coating held.",
    "Cable tension",
    "Tension stayed within limits.",
  ];
  assert.deepEqual(
    [1, 2, 4].map((i) => paragraphKinds(texts)[i]),
    ["heading", "heading", "heading"],
  );
  const c = { ref: "R1.1", text: "Too strong.", changes: [change("x", 3, "fell", "was lower")] };
  assert.deepEqual(
    followPlaces(c, texts).map((p) => [p.index, p.why]),
    [
      [0, "heading above a changed paragraph"],
      [5, "same section as a changed paragraph"],
    ],
  );
  // Nor does such a line end the reference list.
  const refs = [
    "Deck",
    "Rust fell (Fig. 1).",
    "References",
    "Okafor T. Coatings for steel bridges. 2018.",
    "/",
    "Yeo K. Drainage of expansion joints. 2020.",
    "Figure 1. Rust cover per span.",
  ];
  const d = {
    ref: "R1.1",
    text: "Too strong.",
    changes: [change("a", 1, "fell", "was lower"), change("b", 3, "2018", "2019")],
  };
  assert.deepEqual(indexes(followPlaces(d, refs)), [0, 6]);
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
      [1, 'methods; also has "permanent"'],
      [5, "same section as a changed paragraph"],
    ],
  );
  // The word in more than two body or methods paragraphs: vocabulary.
  const many = [...texts, "A permanent anchor.", "Permanent bolts hold.", "No permanent mark."];
  assert.ok(!indexes(followPlaces(c, many)).includes(1));
});

for (const source of ["named", "claim", "removed"]) {
  test(`followPlaces: ${source} terms reach Methods and uncited figure and table legends`, () => {
    const texts = [
      "Deck survey",
      "The seal held during the inspection.",
      "Methods",
      "A permanent seal was applied around each joint.",
      "Figure legends",
      "Figure 7. Permanent sealing at the west joint.",
      "A permanent seal prevents leaks at the edge.",
      "Table 8. Permanent sealing at the east joint.",
      "Table 9. Coating thickness at each joint.",
      "References",
      "Hale R. Sealing of bridge joints. 2031.",
    ];
    const c = {
      ref: "R1.1",
      text:
        source === "named"
          ? "The term “permanent” is not supported."
          : source === "claim"
            ? "Nothing supports a permanent seal."
            : "The durability claim is not supported.",
      changes: [
        change("R1.1-1", 1, "The permanent seal prevents leaks.", texts[1], {
          rationale: {
            why:
              source === "claim"
                ? "No evidence of permanent sealing."
                : "Only the inspection is reported.",
          },
        }),
      ],
    };
    const places = followPlaces(c, texts);
    assert.ok(places.some((p) => p.index === 3 && p.kind === "methods"));
    for (const index of [5, 6, 7])
      assert.ok(
        places.some((p) => p.index === index),
        `missing legend paragraph ${index}`,
      );
    assert.ok(!places.some((p) => p.index === 8 || p.index === 10));
    assert.ok(places.length <= MAX_PLACES);
  });
}

test("followPlaces: term legends keep the legend cap and do not displace Methods", () => {
  const texts = [
    "Deck survey",
    "The seal held during the inspection.",
    "Methods",
    "A permanent seal was applied around each joint.",
    "Figure legends",
    ...Array.from({ length: 12 }, (_, i) => `Figure ${i + 1}. Permanent sealing at a joint.`),
  ];
  const c = {
    ref: "R1.1",
    text: "The term “permanent” is not supported.",
    changes: [change("R1.1-1", 1, "The permanent seal prevents leaks.", texts[1])],
  };
  const places = followPlaces(c, texts);
  assert.equal(places.filter((p) => p.kind === "caption").length, 8);
  assert.ok(places.some((p) => p.index === 3 && p.kind === "methods"));
});

test("followPlaces: which removed words bring a methods paragraph", () => {
  const filler = Array.from({ length: 150 }, (_, i) => `Span ${i} held firm.`);
  const methods = [
    "Methods",
    "Girders were scanned with ultrasonic probes.", // 1
    "Ultrasonic readings were logged hourly.", // 2
    "The ultrasonic unit was calibrated daily.", // 3
    "Ultrasonic gel was applied across the weld.", // 4
    "Welds were graded by radiographic film.", // 5
    "Loads were applied by hydraulic jacks.", // 6
    "Rivets were set by pneumatic hammers.", // 7
    "A rusty plate with significant pitting was the reference.", // 8
    "Positions came from the GPS-2 logger.", // 9
    "Results", // 10
    "Cracks were absent.", // 11, changed
  ];
  const texts = [...methods, ...filler];
  const of = (find, more = {}) => ({
    ref: "R1.1",
    text: "The testing claim goes too far.",
    changes: [change("x", 11, find, "Cracks were absent", more)],
  });
  const inMethods = (c) =>
    followPlaces(c, texts)
      .filter((p) => p.kind === "methods")
      .map((p) => [p.index, p.why]);
  // Common words of any report ("across", "significant", "showed") and short
  // plain words ("rusty") bring nothing. Of the others at most three, those
  // in the fewest paragraphs first.
  assert.deepEqual(
    inMethods(
      of(
        "Ultrasonic radiographic hydraulic pneumatic scans across rusty spans showed significant cracks",
      ),
    ),
    [
      [5, 'methods; also has "radiographic"'],
      [6, 'methods; also has "hydraulic"'],
      [7, 'methods; also has "pneumatic"'],
    ],
  );
  // One word: at most three methods paragraphs (160 paragraphs: it may be in four).
  assert.deepEqual(inMethods(of("Ultrasonic scans found cracks")), [
    [1, 'methods; also has "ultrasonic"'],
    [2, 'methods; also has "ultrasonic"'],
    [3, 'methods; also has "ultrasonic"'],
  ]);
  // In a short text a word in more than two paragraphs is vocabulary.
  const short = [...methods, "Ultrasonic scans were repeated."];
  assert.deepEqual(
    followPlaces(of("Ultrasonic scans found cracks"), short).filter((p) => p.kind === "methods"),
    [],
  );
  // A short word that is a technical token (a digit, a hyphen, a capital inside).
  assert.deepEqual(inMethods(of("GPS-2 fixes found cracks")), [[9, 'methods; also has "gps-2"']]);
  // A short plain word the comment and the reason of the change both name.
  const claim = of("Rusty plates had cracks", {
    rationale: { what: "w", why: "No plate was rusty.", addresses: "a", requirement_ids: [] },
  });
  claim.text = "Nothing shows the plates were rusty.";
  assert.deepEqual(inMethods(claim), [[8, 'methods; also has "rusty"']]);
  // Such a word too is vocabulary when it is in more than a fifth of the
  // body and methods paragraphs (never fewer than 3), and brings at most
  // three methods paragraphs.
  const rusty = (n, more = []) => [
    "Methods",
    ...Array.from({ length: n }, (_, i) => `Plate ${i} was rusty at the weld.`),
    "Results",
    "Cracks were absent.",
    ...more,
  ];
  const at = (n) => ({
    ...claim,
    changes: [{ ...claim.changes[0], paragraph_index: n + 2, paragraph_id: `k${n + 2}` }],
  });
  const methodsOf = (c, list) =>
    followPlaces(c, list)
      .filter((p) => p.kind === "methods")
      .map((p) => p.index);
  assert.deepEqual(methodsOf(at(5), rusty(5)), []);
  assert.deepEqual(methodsOf(at(4), rusty(4, filler.slice(0, 20))), [1, 2, 3]);
});

// A made-up load test of a footbridge: the comment changes how a quantity
// is measured, in the methods only.
const SURVEY = [
  "Load test of the Mill Lane footbridge", // 0 heading
  "Summary", // 1 heading
  "Midspan deflection stayed small in every load test.", // 2 body
  "Methods", // 3 heading
  "Midspan deflection was read from a calibrated dial gauge under a static load.", // 4 methods
  "Results", // 5 heading
  "Loads were applied in four steps.", // 6 body
  "The midspan deflection reached 4 mm at full load.", // 7 body
  "Midspan deflection per load step is shown in Fig. 2.", // 8 body
  "Figure legends", // 9 heading
  "Figure 2. Deflection per load step.", // 10 caption
  "References", // 11 heading
  "Hale R. Midspan deflection of footbridges. J Span Eng. 2011.", // 12
];
const GAUGE = {
  ref: "R2.1",
  text: "How “midspan deflection” was measured is not clear (Fig. 2).",
  changes: [change("R2.1-1", 4, "a calibrated dial gauge", "a laser level")],
};

test("followPlaces: a change in the methods brings the paragraphs that report the measurement", () => {
  assert.deepEqual(
    followPlaces(GAUGE, SURVEY).map((p) => [p.index, p.kind, p.why]),
    [
      [3, "heading", "heading above a changed paragraph"],
      [10, "caption", "legend of Fig. 2"],
      // The one that cites the figure the comment names first, then the
      // order of the document. The reference list is no place.
      [8, "body", "reports; cites Fig. 2"],
      [2, "body", 'reports; also has "midspan deflection"'],
      [7, "body", 'reports; also has "midspan deflection"'],
    ],
  );
  // A change in the body only: the methods paragraph is listed as before,
  // and no paragraph as one that reports.
  const body = { ...GAUGE, changes: [change("R2.1-1", 7, "reached 4 mm", "stayed under 5 mm")] };
  const places = followPlaces(body, SURVEY);
  assert.deepEqual(indexes(places), [5, 10, 4, 6, 8]);
  assert.ok(places.every((p) => !p.why.startsWith("reports")));
  // The paragraphs this comment changed, and those where the author
  // rejected a change of it, are not listed.
  const more = {
    ...GAUGE,
    changes: [
      ...GAUGE.changes,
      change("R2.1-2", 2, "stayed small", "was small"),
      change("R2.1-3", 7, "reached", "was", { state: "rejected" }),
    ],
  };
  assert.deepEqual(indexes(followPlaces(more, SURVEY)), [1, 3, 10, 8]);
});

test("followPlaces: a withdrawn word of a methods change brings what reports it, unless it is everywhere", () => {
  const texts = [
    "Methods",
    "Sag was read from a calibrated dial gauge.",
    "Results",
    "The calibrated readings agreed within 1 mm.",
    "Sag grew with each load step.",
    "17 / 0.4", // no letters: no place
  ];
  const c = {
    ref: "R2.2",
    text: "Nothing shows that the gauge was calibrated.",
    changes: [
      change("R2.2-1", 1, "a calibrated dial gauge", "a dial gauge", {
        rationale: { why: "No record that it was calibrated.", addresses: "calibrated" },
      }),
    ],
  };
  assert.deepEqual(
    followPlaces(c, texts).map((p) => [p.index, p.why]),
    [
      [0, "heading above a changed paragraph"],
      [3, 'reports; also has "calibrated"'],
    ],
  );
  // The word in more than three body or methods paragraphs: vocabulary.
  const many = [...texts, "A calibrated scale.", "Calibrated rods held.", "No calibrated mark."];
  assert.ok(!indexes(followPlaces(c, many)).includes(3));
});

test("followPlaces: the paragraphs that report stop at six, those that cite the figure first", () => {
  const texts = [
    "Methods",
    "Midspan deflection was read from a dial gauge.",
    "Results",
    ...Array.from({ length: 8 }, (_, i) => `Midspan deflection of span ${i} was small.`),
    "Midspan deflection of the last span is in Table 3.",
  ];
  const c = {
    ref: "R2.3",
    text: "Define “midspan deflection”.",
    changes: [change("R2.3-1", 1, "a dial gauge (Table 3)", "a laser level (Table 3)")],
  };
  const got = followPlaces(c, texts).filter((p) => p.why.startsWith("reports"));
  assert.deepEqual(indexes(got), [11, 3, 4, 5, 6, 7]);
});

// A made-up shelf survey of a library: the comment gets an insert in the
// methods, which takes no word out, and names no term the results hold.
test("followPlaces: a methods change brings the paragraphs that cite the figure it points to", () => {
  const texts = [
    "Methods", // 0 heading
    "Shelf marks were read by two clerks.", // 1 methods
    "Misfiled volumes were counted per aisle (Fig. 5).", // 2 methods
    "Results", // 3 heading
    "Most misfiled volumes stood in the east wing (Fig. 5B).", // 4 body
    "Loans rose over the year (Fig. 4).", // 5 body
    "The reading room was full on weekdays.", // 6 body
    "Aisle counts and loan counts are given in Tables 2 and 3.", // 7 body
    "Figure legends", // 8 heading
    "Figure 5. Misfiled volumes per aisle.", // 9 caption
    "Table 2. Aisle counts.", // 10 caption
  ];
  const insert = (index, find, text) => change("R3.1-1", index, find, text, { kind: "insert" });
  // The changed methods paragraph cites the figure itself.
  const c = {
    ref: "R3.1",
    text: "Say who counted.",
    changes: [insert(2, "were counted", " by one clerk")],
  };
  assert.deepEqual(
    followPlaces(c, texts).map((p) => [p.index, p.why]),
    [
      [0, "heading above a changed paragraph"],
      [9, "legend of Fig. 5"],
      [4, "reports; cites Fig. 5"],
      [1, "same section as a changed paragraph"],
    ],
  );
  // The figure is named by the comment, or by what the insert put in.
  const byComment = {
    ref: "R3.1",
    text: "Say who counted the volumes of Table 2.",
    changes: [insert(1, "two clerks", " in March")],
  };
  assert.deepEqual(
    followPlaces(byComment, texts)
      .filter((p) => p.why.startsWith("reports"))
      .map((p) => [p.index, p.why]),
    [[7, "reports; cites Table 2"]],
  );
  const byInsert = { ...c, changes: [insert(1, "two clerks", " (see Fig. 4)")] };
  assert.deepEqual(
    followPlaces(byInsert, texts)
      .filter((p) => p.why.startsWith("reports"))
      .map((p) => [p.index, p.why]),
    [[5, "reports; cites Fig. 4"]],
  );
  // A change in the body only: no paragraph is listed as one that reports.
  const body = { ...c, changes: [change("R3.1-1", 6, "was full", "was busy (Fig. 5)")] };
  assert.ok(followPlaces(body, texts).every((p) => !p.why.startsWith("reports")));
});

test("followPlaces: direct terms precede incidental methods citations, six in all", () => {
  const texts = [
    "Methods",
    "Misfiled volumes were counted per aisle (Fig. 5).",
    "Results",
    "A “recount” of aisle 1 gave the same number.",
    ...Array.from({ length: 4 }, (_, i) => `Aisle ${i} is in Fig. 5.`),
    "The recount of aisle 9 differed.",
    ...Array.from({ length: 3 }, (_, i) => `Wing ${i} is in Fig. 5 too.`),
  ];
  const c = {
    ref: "R3.2",
    text: "Was there a “recount”?",
    changes: [change("R3.2-1", 1, "were counted", " once", { kind: "insert" })],
  };
  const got = followPlaces(c, texts).filter((p) => p.why.startsWith("reports"));
  assert.deepEqual(indexes(got), [3, 8, 4, 5, 6, 7]);
  // With fewer citations, direct terms still come first.
  const few = followPlaces(c, texts.slice(0, 9)).filter((p) => p.why.startsWith("reports"));
  assert.deepEqual(
    few.map((p) => [p.index, p.why]),
    [
      [3, 'reports; also has "recount"'],
      [8, 'reports; also has "recount"'],
      [4, "reports; cites Fig. 5"],
      [5, "reports; cites Fig. 5"],
      [6, "reports; cites Fig. 5"],
      [7, "reports; cites Fig. 5"],
    ],
  );
});

test("followPlaces: methods citations cannot starve direct measurement reports", () => {
  const texts = [
    "Methods",
    "Midspan deflection was read using a dial gauge (Figs. 1-7).",
    "Results",
    ...Array.from(
      { length: 7 },
      (_, i) => `Coating coverage at pier ${i + 1} is shown in Fig. ${i + 1}.`,
    ),
    "Midspan deflection decreased under the static load.",
    "Midspan deflection remained stable after unloading (Fig. 1).",
    "The load response is shown in Fig. 9.",
  ];
  const c = {
    ref: "R2.4",
    text: "Define “midspan deflection” and clarify Fig. 9.",
    changes: [change("R2.4-1", 1, "a dial gauge", "a calibrated dial gauge")],
  };
  const got = followPlaces(c, texts).filter((p) => p.why.startsWith("reports"));
  assert.deepEqual(
    got.map((p) => [p.index, p.why]),
    [
      [12, "reports; cites Fig. 9"],
      [10, 'reports; also has "midspan deflection"'],
      [11, 'reports; also has "midspan deflection"'],
      [3, "reports; cites Fig. 1"],
      [4, "reports; cites Fig. 2"],
      [5, "reports; cites Fig. 3"],
    ],
  );
});

test("followPlaces: the paragraph an answered question names comes first", () => {
  const q = (id, paragraph_id, answer) => ({ id, question: "?", why: "w", answer, paragraph_id });
  const c = {
    ...COMMENT,
    questions: [
      q("a1", "k11", "Two cables were retensioned."),
      q("a2", "k9", ""), // not answered
      q("a3", "k8", "Yes."), // already checked and changed
      q("a4", "k99", "No."), // not in the manuscript
      { id: "a5", question: "?", why: "w", answer: "Yes." }, // no place
      q("a6", "k11", "Again."), // listed once
    ],
    follow_through: { seen_answer_ids: ["a3"] },
  };
  const got = followPlaces(c, PARAGRAPHS, IDS);
  assert.deepEqual(
    got.slice(0, 2).map((p) => [p.index, p.id, p.kind, p.why]),
    [
      [11, "k11", "body", "the author's answer to a1 goes here"],
      [7, "k7", "heading", "heading above a changed paragraph"],
    ],
  );
  assert.deepEqual(indexes(got), [11, 7, 13, 14, 16, 5, 9]);
  // Without task pane ids a paragraph is named by its index.
  const byIndex = { ...COMMENT, questions: [q("a1", "p11", "Two.")] };
  assert.equal(followPlaces(byIndex, PARAGRAPHS)[0].index, 11);
  // A place the answer names is one also where the author rejected a change
  // of this comment: the answer came after.
  const rejected = {
    ...COMMENT,
    changes: [...COMMENT.changes, change("R1.1-2", 9, "also held", "held", { state: "rejected" })],
    questions: [q("a1", "k9", "It held for one winter.")],
  };
  assert.equal(followPlaces(rejected, PARAGRAPHS, IDS)[0].index, 9);
  // No change in the document: the places the answers name are the only
  // ones with unchecked answers, nothing else is looked for.
  assert.deepEqual(
    followPlaces({ ...c, changes: [] }, PARAGRAPHS, IDS).map((p) => [p.index, p.why, p.answer_ids]),
    [[11, "the author's answer to a1 goes here", ["a1", "a6"]]],
  );
  assert.deepEqual(got[0].answer_ids, ["a1", "a6"]);
  assert.deepEqual(got[1].answer_ids, []);
  // An answer does not go into the reference list; a heading or a legend
  // is a place like any other.
  const list = {
    ...COMMENT,
    changes: [],
    questions: [q("a1", "k19", "Okafor 2019."), q("a2", "k7", "Rust."), q("a3", "k13", "Cover.")],
  };
  assert.deepEqual(indexes(followPlaces(list, PARAGRAPHS, IDS)), [7, 13]);
  assert.ok(
    !indexes(followPlaces({ ...list, changes: COMMENT.changes }, PARAGRAPHS, IDS)).includes(19),
  );
});

test("followState: an answer that names a place offers the check, also without a change and after a check", () => {
  const snapshot = { paragraphs: PARAGRAPHS, ids: IDS };
  const q = (id, paragraph_id, answer) => ({ id, question: "?", why: "w", answer, paragraph_id });
  const bare = { ...COMMENT, changes: [] };
  assert.equal(followState({ ...bare, questions: [q("a1", "k11", "")] }, snapshot).offer, null);
  assert.equal(followState({ ...bare, questions: [q("a1", "k11", "Two.")] }, snapshot).offer, 1);
  // The check ran for the changes there are, and saw the answer a1.
  const ran = {
    ...COMMENT,
    questions: [q("a1", "k11", "Two.")],
    follow_through: {
      at: "2026-01-02T03:04:05.000Z",
      candidate_ids: ["k11", "k7"],
      change_ids: [],
      questions: 0,
      seen_change_ids: ["R1.1-1"],
      seen_answer_ids: ["a1"],
    },
  };
  assert.equal(followState(ran, snapshot).offer, null);
  // A new answer reopens even a changed place, but not references or no place.
  const more = (x) => followState({ ...ran, questions: [...ran.questions, x] }, snapshot).offer;
  assert.ok(more(q("a2", "k9", "One winter.")) > 0);
  assert.ok(more(q("a2", "k8", "Yes.")) > 0);
  assert.equal(more(q("a2", "k19", "Okafor 2019.")), null);
  assert.equal(more({ id: "a2", question: "?", why: "w", answer: "Yes." }), null);
  // A check from before the field saw no answer.
  const { seen_answer_ids, ...old } = ran.follow_through;
  assert.ok(followState({ ...ran, follow_through: old }, snapshot).offer > 0);
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
  // Nor is a paragraph where the author rejected a change of this comment.
  const refused = {
    ...COMMENT,
    changes: [
      ...COMMENT.changes,
      change("R1.1-2", 9, "also held", "held", { state: "rejected" }),
      change("R1.1-3", 13, "Rust coverage", "Rust cover", { state: "rejected" }),
    ],
  };
  assert.deepEqual(indexes(followPlaces(refused, PARAGRAPHS, IDS)), [7, 14, 16, 5]);
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
  assert.equal(got.find((p) => p.index === 5).why, 'methods; also has "eliminated"');
  assert.deepEqual(indexes(followPlaces(COMMENT, PARAGRAPHS, IDS, { maxPlaces: 2 })), [7, 13]);
  const chars = PARAGRAPHS[7].length + PARAGRAPHS[13].length + PARAGRAPHS[14].length;
  assert.deepEqual(
    indexes(followPlaces(COMMENT, PARAGRAPHS, IDS, { maxChars: chars })),
    [7, 13, 14],
  );
  // A paragraph too long for what is left is passed over, not the end of
  // the list; one that alone is over the limit does not leave the list empty.
  const big = ["Deck", "Rust fell.", "x ".repeat(60), "Span A held.", "Span B held."];
  const cut = { ref: "R1.1", text: "Why?", changes: [change("x", 1, "fell", "was lower")] };
  assert.deepEqual(indexes(followPlaces(cut, big, null, { maxChars: 40 })), [0, 3, 4]);
  assert.deepEqual(
    indexes(followPlaces(cut, ["Deck of the long bridge", "Rust fell."], null, { maxChars: 5 })),
    [0],
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
  const nothing = {
    ...ran,
    follow_through: { ...ran.follow_through, change_ids: [], questions: 0 },
  };
  assert.equal(followState(nothing, snapshot).line, "Other places checked: 6 (no change needed).");
  // A file from before the field: as if no check ran.
  assert.equal(followState({ ...COMMENT, follow_through: null }, snapshot).offer, 6);
});

for (const [name, paragraphs, sizes] of [
  ["paragraph limit", Array.from({ length: 41 }, (_, i) => `Span ${i} held firm.`), [20, 20, 1]],
  ["character limit", ["a".repeat(25000) + ".", "b".repeat(25000) + "."], [1, 1]],
]) {
  test(`followPlaces: checked answers do not starve later batches at the ${name}`, () => {
    const ids = paragraphs.map((_, i) => `k${i}`);
    const c = {
      changes: [],
      questions: ids.map((paragraph_id, i) => ({ id: `a${i + 1}`, paragraph_id, answer: "Yes." })),
    };
    const snapshot = { paragraphs, ids };
    const checked = new Set();
    for (const size of sizes) {
      const before = structuredClone(c);
      const places = followPlaces(c, paragraphs, ids);
      assert.equal(places.length, size);
      assert.equal(followState(c, snapshot).offer, size);
      assert.ok(places.length <= MAX_PLACES);
      assert.ok(places.reduce((n, p) => n + paragraphs[p.index].length, 0) <= MAX_CHARS);
      assert.equal(new Set(places.map((p) => p.id)).size, places.length);
      assert.deepEqual(c, before, "selection does not mutate checked history");
      for (const p of places) {
        assert.ok(!checked.has(p.id), "a checked answer must not consume another batch");
        checked.add(p.id);
      }
      c.follow_through = {
        candidate_ids: places.map((p) => p.id),
        change_ids: [],
        questions: 0,
        seen_change_ids: [],
        seen_answer_ids: [
          ...(c.follow_through?.seen_answer_ids ?? []),
          ...places.flatMap((p) => p.answer_ids),
        ],
      };
    }
    assert.equal(checked.size, ids.length);
    assert.deepEqual(followPlaces(c, paragraphs, ids), []);
    assert.equal(followState(c, snapshot).offer, null);
  });
}

test("followPlaces: pending answers on a checked paragraph keep all its answer context once", () => {
  const c = {
    changes: [],
    questions: [
      { id: "a1", paragraph_id: "k2", answer: "One winter." },
      { id: "a2", paragraph_id: "k2", answer: "Coating held." },
      { id: "a3", paragraph_id: "k2", answer: "Inspection completed." },
    ],
    follow_through: { seen_answer_ids: ["a1", "a2"] },
  };
  const before = structuredClone(c);
  const places = followPlaces(c, PARAGRAPHS, IDS, { maxPlaces: 1 });
  assert.deepEqual(
    places.map((p) => [p.id, p.answer_ids]),
    [["k2", ["a1", "a2", "a3"]]],
  );
  assert.deepEqual(c, before);
});

test("followPlaces: answer progress leaves unrelated, unanswered and unavailable places out", () => {
  const c = {
    changes: [],
    questions: [
      { id: "checked", paragraph_id: "k11", answer: "Checked." },
      { id: "empty", paragraph_id: "k2", answer: "" },
      { id: "missing", paragraph_id: "gone", answer: "Missing." },
      { id: "unplaced", answer: "No place." },
      { id: "reference", paragraph_id: "k19", answer: "Reference." },
      { id: "pending", paragraph_id: "k9", answer: "One winter." },
    ],
    follow_through: {
      candidate_ids: ["k11"],
      change_ids: [],
      questions: 0,
      seen_change_ids: [],
      seen_answer_ids: ["checked"],
    },
  };
  assert.deepEqual(indexes(followPlaces(c, PARAGRAPHS, IDS)), [9]);
  const absent = IDS.map((id) => (id === "k9" ? "replacement" : id));
  assert.deepEqual(followPlaces(c, PARAGRAPHS, absent), []);
  assert.equal(followState(c, { paragraphs: PARAGRAPHS, ids: absent }).offer, null);
  assert.equal(followState(c, { paragraphs: PARAGRAPHS, ids: IDS }).offer, 1);
  assert.ok(!c.follow_through.seen_answer_ids.includes("pending"));
});

test("followPlaces: a checked answer can still be context for a live change", () => {
  const c = {
    ...COMMENT,
    questions: [{ id: "a1", paragraph_id: "k9", answer: "One winter." }],
    follow_through: { seen_answer_ids: ["a1"] },
  };
  const places = followPlaces(c, PARAGRAPHS, IDS);
  assert.equal(places[0].id, "k7");
  const section = places.filter((p) => p.id === "k9");
  assert.equal(section.length, 1);
  assert.equal(section[0].why, "same section as a changed paragraph");
  assert.deepEqual(section[0].answer_ids, ["a1"]);
});

test("followPlaces: oversized unanswered checks stay reachable beside smaller change context", () => {
  const paragraphs = ["Results", "A value was measured.", "x".repeat(MAX_CHARS) + "."];
  const ids = ["k0", "k1", "k2"];
  const c = {
    ref: "R1.1",
    text: "Clarify the value.",
    changes: [change("R1.1-1", 1, "value", "measurement")],
    questions: [{ id: "a1", paragraph_id: "k2", answer: "The measurement was verified." }],
    follow_through: {
      candidate_ids: ["k0"],
      change_ids: [],
      questions: 0,
      seen_change_ids: ["R1.1-1"],
      seen_answer_ids: [],
    },
  };
  const snapshot = { paragraphs, ids };
  const before = structuredClone(c);
  assert.deepEqual(
    followPlaces(c, paragraphs, ids).map((p) => p.id),
    ["k2"],
  );
  assert.equal(followState(c, snapshot).offer, 1);
  assert.deepEqual(c, before, "selection does not consume the pending answer");
  c.follow_through.seen_answer_ids = ["a1"];
  assert.equal(followState(c, snapshot).offer, null);
  assert.ok(!followPlaces(c, paragraphs, ids).some((p) => p.id === "k2"));
});

test("followPlaces: mixed answer batches keep each oversized paragraph alone without losing progress", () => {
  for (const lengths of [
    [MAX_CHARS + 1, MAX_CHARS + 2, 20],
    [20, MAX_CHARS + 1, 20],
  ]) {
    const paragraphs = [
      "Results",
      "A value was measured.",
      ...lengths.map((n) => "x".repeat(n - 1) + "."),
    ];
    const ids = paragraphs.map((_, i) => `k${i}`);
    const c = {
      ref: "R1.1",
      text: "Clarify the value.",
      changes: [change("R1.1-1", 1, "value", "measurement")],
      questions: lengths.map((_, i) => ({
        id: `a${i}`,
        paragraph_id: `k${i + 2}`,
        answer: "Verified.",
      })),
    };
    const snapshot = { paragraphs, ids };
    const seen = new Set();
    let batches = 0;
    while (followState(c, snapshot).offer) {
      assert.ok(++batches <= 3, "the remaining answers must make progress");
      const places = followPlaces(c, paragraphs, ids);
      assert.ok(places.length <= MAX_PLACES);
      const oversized = places.filter((p) => paragraphs[p.index].length > MAX_CHARS);
      if (oversized.length) assert.equal(places.length, 1, "oversized means standalone");
      else assert.ok(places.reduce((n, p) => n + paragraphs[p.index].length, 0) <= MAX_CHARS);
      const fresh = places.flatMap((p) => p.answer_ids).filter((id) => !seen.has(id));
      assert.ok(fresh.length, "old change context must not consume an answer batch");
      fresh.forEach((id) => seen.add(id));
      c.follow_through = {
        candidate_ids: places.map((p) => p.id),
        change_ids: [],
        questions: 0,
        seen_change_ids: ["R1.1-1"],
        seen_answer_ids: [...seen],
      };
    }
    assert.equal(seen.size, 3);
  }
});
