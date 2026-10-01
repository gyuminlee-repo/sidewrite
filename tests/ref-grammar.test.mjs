import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { buildRefGrammar, refGroups } from "../daemon/ref-grammar.mjs";

const load = async (name) =>
  JSON.parse(
    await readFile(new URL(`../profiles/journal_profiles/${name}.json`, import.meta.url), "utf8"),
  );
const prof = await load("example-elsevier");
const generic = await load("generic");
const texts = (sentence, g) => refGroups(sentence, g).map((x) => x.text);

test("every corpus sentence gives exactly its expected groups", () => {
  const g = buildRefGrammar(prof);
  assert.ok(prof.supp_ref_corpus.length >= 50);
  assert.ok(prof.supp_ref_corpus.filter(([, want]) => want.length === 0).length >= 5);
  for (const [sentence, want] of prof.supp_ref_corpus)
    assert.deepEqual(texts(sentence, g), want, sentence);
});

test("later panels after a separator stay in the group", () => {
  const g = buildRefGrammar(prof);
  assert.deepEqual(texts("see Figs. 2A, B and 3 for detail", g), ["Figs. 2A, B and 3"]);
});

test("groups carry their place in the text", () => {
  const s = "(Fig. 3A; see Table S1)";
  const got = refGroups(s, buildRefGrammar(prof));
  assert.deepEqual(
    got.map((x) => [x.start, x.end, s.slice(x.start, x.end)]),
    [
      [1, 8, "Fig. 3A"],
      [14, 22, "Table S1"],
    ],
  );
});

test("a caption title is no reference, a reference after it is", () => {
  const g = buildRefGrammar(prof);
  assert.deepEqual(texts("Fig. 2. Growth as in Table S1.", g), ["Table S1"]);
  // Only at the start of the paragraph.
  assert.deepEqual(texts("See Fig. 2. Growth was slow.", g), ["Fig. 2"]);
});

test("the generic profile: both panel cases, no caption rule", () => {
  const g = buildRefGrammar(generic);
  assert.equal(g.caption, null);
  assert.deepEqual(texts("as in Fig. 2a, b and Supplementary Table S3; xFig. 1 is none", g), [
    "Fig. 2a, b and Supplementary Table S3",
  ]);
  assert.deepEqual(texts("Tables of Contents", g), []);
});

test("the grammar can be used twice (no state kept between calls)", () => {
  const g = buildRefGrammar(prof);
  assert.deepEqual(texts("in Table S3, and", g), ["Table S3"]);
  assert.deepEqual(texts("in Table S3, and", g), ["Table S3"]);
});

test("no grammar when the profile has none", () => {
  assert.equal(buildRefGrammar({}), null);
  assert.equal(buildRefGrammar({ ref_head: [], ref_item: "\\d+" }), null);
  assert.equal(buildRefGrammar(null), null);
  assert.deepEqual(refGroups("Fig. 1", null), []);
});
