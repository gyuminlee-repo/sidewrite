// A paragraph rewritten as a whole becomes the tracked edits of the words
// that differ. All texts here are made up.
import { test } from "node:test";
import assert from "node:assert/strict";
import { rewriteSegments, applySegments } from "../daemon/word-diff.mjs";

const BEFORE =
  "Gene A is essential for growth on glycerol. The mutant grew at 0.21 per hour at 30 C. These data prove that gene A is essential.";
const AFTER =
  "Deletion of gene A halved the growth rate on glycerol. The mutant grew at 0.21 per hour at 30 °C. These data show a slower growth of the mutant.";

test("rewriteSegments: only the words that differ, each found once when its turn comes", () => {
  const { segments, error } = rewriteSegments(BEFORE, AFTER);
  assert.equal(error, undefined);
  // The places that changed: the first sentence, the unit, the last sentence.
  assert.equal(segments.length, 3);
  assert.equal(applySegments(BEFORE, segments), AFTER);
  // The unit: widened by the words before it until it is found once.
  assert.deepEqual(segments[1], { find: "30 C", replace: "30 °C" });
  // The sentence in the middle is in no edit.
  assert.equal(
    segments.some((s) => s.find.includes("The mutant grew at 0.21")),
    false,
  );
  // Nothing to do, and white space alone is nothing.
  assert.deepEqual(rewriteSegments(BEFORE, BEFORE), { segments: [] });
  assert.deepEqual(rewriteSegments("a  b", " a b "), { segments: [] });
});

test("rewriteSegments: an insertion and a deletion carry a word next to them; repeated words are widened", () => {
  let r = rewriteSegments("The rate was high.", "The growth rate was high.");
  assert.equal(r.segments.length, 1);
  // The inserted word comes with a word next to it, so Word has a place to find.
  assert.ok(r.segments[0].find.length >= 3 && r.segments[0].replace.includes("growth"));
  assert.equal(applySegments("The rate was high.", r.segments), "The growth rate was high.");
  r = rewriteSegments("The rate was very high.", "The rate was high.");
  assert.equal(r.segments.length, 1);
  assert.equal(applySegments("The rate was very high.", r.segments), "The rate was high.");
  // "the" occurs three times: the edit takes in what tells its place apart.
  const before = "In the first run the yield rose, and in the second run the yield fell.";
  const after = "In the first run the yield rose, and in the second run a lower yield was seen.";
  r = rewriteSegments(before, after);
  assert.equal(applySegments(before, r.segments), after);
  for (const s of r.segments) assert.equal(s.find, s.find.trim());
  // An insertion at the very start takes the word after it.
  r = rewriteSegments("Growth was slow.", "In all lines, growth was slow.");
  assert.equal(applySegments("Growth was slow.", r.segments), "In all lines, growth was slow.");
});

test("rewriteSegments: two changes that cannot be told apart become one edit", () => {
  const before = "up up up down";
  const after = "up x up y down";
  const r = rewriteSegments(before, after);
  assert.equal(r.error, undefined);
  assert.equal(applySegments(before, r.segments), after);
});

test("rewriteSegments: a changed part over the search limit is cut at word boundaries", () => {
  const long = Array.from({ length: 60 }, (_, i) => `word${i}`).join(" "); // 409 characters
  const before = `Start here. ${long}. End here.`;
  const after = "Start here. A short sentence. End here.";
  const r = rewriteSegments(before, after);
  assert.equal(r.error, undefined);
  assert.ok(r.segments.length >= 2);
  for (const s of r.segments) assert.ok(s.find.length <= 255, String(s.find.length));
  // The first piece gets the new text, the others are deleted.
  assert.ok(r.segments[0].replace.includes("A short sentence"));
  assert.deepEqual(
    r.segments.slice(1).map((s) => s.replace),
    r.segments.slice(1).map(() => ""),
  );
  assert.equal(applySegments(before, r.segments), after);
  // A smaller limit, for the same text.
  const small = rewriteSegments(before, after, { max: 60 });
  for (const s of small.segments) assert.ok(s.find.length <= 60);
  assert.equal(applySegments(before, small.segments), after);
});

// Whatever is rewritten: the edits, applied from the last to the first as
// Word applies them, give the new paragraph, and each find is there once.
function mulberry32(seed) {
  let t = seed;
  return () => {
    t += 0x6d2b79f5;
    let x = Math.imul(t ^ (t >>> 15), 1 | t);
    x ^= x + Math.imul(x ^ (x >>> 7), 61 | x);
    return ((x ^ (x >>> 14)) >>> 0) / 4294967296;
  };
}
const WORDS =
  "the gene strain growth rate was is high low and of in on glycerol mutant parent , . ( ) essential shown 30 0.5 % data these".split(
    " ",
  );

test("property: the segments applied back to front give the rewritten paragraph", () => {
  const rnd = mulberry32(20261002);
  const pick = () => WORDS[Math.floor(rnd() * WORDS.length)];
  let made = 0;
  let refused = 0;
  for (let n = 0; n < 400; n++) {
    const a = Array.from({ length: 5 + Math.floor(rnd() * 40) }, pick);
    const b = a.flatMap((w) => {
      const x = rnd();
      if (x < 0.12) return []; // deleted
      if (x < 0.24) return [pick()]; // replaced
      if (x < 0.34) return [w, pick()]; // inserted after
      return [w];
    });
    const before = a.join(" ");
    const after = b.join(" ");
    const r = rewriteSegments(before, after, { max: 40 });
    if (r.error) {
      refused += 1;
      continue;
    }
    made += 1;
    const norm = (s) => s.replace(/\s+/g, " ").trim();
    assert.equal(applySegments(before, r.segments), norm(after), `${before}\n->\n${after}`);
    for (const s of r.segments) {
      assert.ok(s.find.length <= 40 && s.find.length > 0);
      assert.notEqual(s.find, s.replace);
    }
  }
  console.log(`rewrites made ${made}, refused ${refused} of 400`);
  assert.ok(made > 300, `made ${made}`);
});
