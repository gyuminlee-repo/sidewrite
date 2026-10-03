// A phrase the reviewer names stands also where only a word of its family
// is left ("densified network" for "network densification"), and a word the
// whole paper uses does not list the paper. A made-up corpus on urban
// transport planning and bridge engineering, with the places known.
import { test } from "node:test";
import assert from "node:assert/strict";
import { termPlaces, termBlock, familyStem } from "../taskpane/shared/named-terms.js";

const CORPUS = [
  /* 0 */ "Introduction",
  /* 1 */ "Network densification shortened average trips in the northern district.",
  /* 2 */ "The densified network carried more riders during peak hours.",
  /* 3 */ "After densification, transfers between lines became rarer.",
  /* 4 */ "Planners kept densifying the grid until 2019.",
  /* 5 */ "The network was redrawn once more in the final phase.",
  /* 6 */ "Population density stayed constant over the period.",
  /* 7 */ "Dense housing lined the main corridors of the network.",
  /* 8 */ "Condensation on the deck slowed the inspection crews.",
  /* 9 */ "Load redistribution after the cable failure was measured on the main span.",
  /* 10 */ "Loads were redistributed to the adjacent girders within seconds.",
  /* 11 */ "The redistribution pattern matched the model of the deck.",
  /* 12 */ "Traffic distribution across the bridge lanes was uneven.",
  /* 13 */ "Redistricting of the wards changed the bus routes in the network.",
  /* 14 */ "Each station of the network had a new timetable.",
  /* 15 */ "The state funded the bridge retrofit.",
  /* 16 */ "Ridership grew on the network in the second year.",
  /* 17 */ "The network survey covered every line of the network.",
];

const comment = (text) => ({ ref: "R1.1", reviewer: 1, text });
const placesOf = (text) =>
  termPlaces(comment(text), CORPUS).map((r) => ({
    term: r.term,
    exact: r.places.filter((p) => !p.form).map((p) => p.index),
    family: r.places.filter((p) => p.form).map((p) => [p.index, p.form]),
  }));

test("familyStem: one family for the forms of a word, apart from look-alikes", () => {
  const same = [
    ["densification", "densified"],
    ["densification", "densifying"],
    ["redistribution", "redistributed"],
    ["congestion", "congested"],
    ["networks", "network"],
  ];
  for (const [a, b] of same) assert.equal(familyStem(a), familyStem(b), `${a} ~ ${b}`);
  const apart = [
    ["densification", "density"],
    ["densification", "dense"],
    ["densification", "condensation"],
    ["redistribution", "distribution"],
    ["redistribution", "redistricting"],
    ["station", "state"],
  ];
  for (const [a, b] of apart) assert.notEqual(familyStem(a), familyStem(b), `${a} !~ ${b}`);
});

test("corpus: a phrase is found in its family forms, not in look-alikes, and a common word lists nothing", () => {
  const [dens] = placesOf('The phrase "network densification" overstates what was done.');
  const [load] = placesOf('Explain "load redistribution" in the abstract.');
  // Known answers. Positives: 2, 3, 4 (densification); 10, 11 (redistribution).
  // Negatives: 6, 7, 8 (density, dense, condensation); 12, 13 (distribution,
  // redistricting, substring and prefix traps); 5, 14, 16, 17 (only the
  // common word "network").
  assert.deepEqual(dens.exact, [1]);
  assert.deepEqual(dens.family, [
    [2, "densified"],
    [3, "densification"],
    [4, "densifying"],
  ]);
  assert.deepEqual(load.exact, [9]);
  assert.deepEqual(load.family, [
    [10, "redistributed"],
    [11, "redistribution"],
  ]);
  const positives = [2, 3, 4, 10, 11];
  const negatives = [5, 6, 7, 8, 12, 13, 14, 15, 16, 17];
  const found = new Set([...dens.family, ...load.family].map(([i]) => i));
  assert.ok(positives.every((i) => found.has(i)));
  assert.ok(negatives.every((i) => !found.has(i)));
  // "network" stands in more than a fifth of the paragraphs: it is
  // vocabulary of the paper and finds nothing on its own.
  const network = CORPUS.filter((t) => /\bnetwork/i.test(t)).length;
  assert.ok(network > Math.ceil(CORPUS.length / 5));
  console.log(
    `named-terms family corpus: paragraphs=${CORPUS.length} positives=${positives.length} negatives=${negatives.length}` +
      ` exact=${dens.exact.length + load.exact.length} family=${found.size}` +
      ` false_positives=${negatives.filter((i) => found.has(i)).length}` +
      ` missed=${positives.filter((i) => !found.has(i)).length} network_paragraphs=${network}`,
  );
});

test("a single word keeps its own match; a phrase without a family hit reads as before", () => {
  const [single] = placesOf('The word "densification" is vague.');
  assert.deepEqual(single.exact, [1, 3]);
  assert.deepEqual(single.family, []);
  const [none] = placesOf('Avoid "survey covered".');
  assert.deepEqual(none.exact, [17]);
  assert.deepEqual(none.family, []);
});

test("termBlock: a family place names the form found", () => {
  const lines = termBlock(comment('The phrase "network densification" overstates it.'), CORPUS);
  assert.ok(
    lines.includes(
      '- "network densification": p1 (body), p2 (body, as "densified"), p3 (body, as "densification"), p4 (body, as "densifying")',
    ),
    lines.join("\n"),
  );
});
