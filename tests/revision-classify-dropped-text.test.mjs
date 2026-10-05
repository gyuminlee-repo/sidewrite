// A change whose old words Word dropped, or whose new words a sibling of the
// same turn rewrote, and nobody has decided anything: it must read pending.
// Shapes of two replays (2026-10-05) where such changes read accepted and
// rejected although the saved document still held every change tracked.
//   - A deletion inside a sibling's pending insertion: Word removes inserted
//     text outright, so no tracked deletion is left and the paragraph reads
//     as if the deletion had been accepted. The text on both sides of the
//     place is still a tracked insertion (one run, or two runs that meet
//     there).
//   - A replace whose new words a sibling rewrote: only its old words are
//     left to match, and they are all still a tracked deletion; a rejected
//     replace has them back as plain text.
// Each shape has controls that must stay decided. Invented text (bridge
// works).
import { test } from "node:test";
import assert from "node:assert/strict";
import { classifyChange, changeContext, normalizeParaText } from "../daemon/revision-classify.mjs";

// A snapshot paragraph whose tracked runs are given by their text; start is
// where the text stands in the normalized paragraph text.
function para(text, runs) {
  const t = normalizeParaText(text);
  const tracked = runs.map(([kind, s]) => {
    const start = t.indexOf(s.trim());
    assert.notEqual(start, -1, `run text not in paragraph: ${s}`);
    return { kind, text: s.trim(), start, marks: 1 };
  });
  return {
    index: 4,
    id: "P4",
    text,
    w_ins: runs.filter(([k]) => k === "ins").length,
    w_del: runs.filter(([k]) => k === "del").length,
    tracked,
  };
}
const snap = (p) => ({ paragraphs: [p], comments: [] });
const state = (change, p) => classifyChange(change, snap(p)).state;

const change = (kind, before, find, text) => ({
  id: "R3.1-2",
  kind,
  paragraph_index: 4,
  paragraph_id: "P4",
  word_comment_id: null,
  find,
  text,
  before,
  after: before.replace(find, text),
  context: changeContext(before, find),
});

// ---- a deletion inside a sibling's pending insertion ----

const HEAD = "Crews checked the north span daily. ";
const I1 = "New bearings with thick steel plates were fitted.";
const DEL = change("delete", `${HEAD}${I1}`, "thick ", "");
const I1_NOW = I1.replace("thick ", "");

test("a deletion inside a sibling's pending insertion stays pending", () => {
  assert.deepEqual(DEL.context, { left: "New bearings with ", right: "steel plates were" });
  // Word dropped "thick " from the insertion: no tracked deletion is left.
  const p = para(`${HEAD}${I1_NOW}`, [["ins", I1_NOW]]);
  assert.equal(state(DEL, p), "pending");
});

test("a deletion where two runs of a sibling's pending insertion meet stays pending", () => {
  // Word may split the insertion at the place of the dropped text.
  const p = para(`${HEAD}${I1_NOW}`, [
    ["ins", "New bearings with"],
    ["ins", "steel plates were fitted."],
  ]);
  assert.equal(state(DEL, p), "pending");
});

test("controls: a deletion that is decided still reads decided", () => {
  // The sibling accepted as well: nothing tracked, the deletion accepted.
  assert.equal(state(DEL, para(`${HEAD}${I1_NOW}`, [])), "accepted");
  // An accepted deletion right before a pending insertion: tracked text on
  // one side of the place only.
  const BEFORE = "Crews checked the span daily. New bearings were fitted.";
  const del = change("delete", BEFORE, "daily. ", "");
  const accepted = para("Crews checked the span New bearings were fitted.", [
    ["ins", "New bearings were fitted."],
  ]);
  assert.equal(state(del, accepted), "accepted");
  // And right after one.
  const del2 = change("delete", BEFORE, " were fitted", "");
  const accepted2 = para("Crews checked the span daily. New bearings.", [
    ["ins", "Crews checked the span daily. New bearings"],
  ]);
  assert.equal(state(del2, accepted2), "accepted");
  // Pending as a tracked deletion of its own, and rejected: as before.
  const pending = para(`${HEAD}${I1}`, [["del", "thick"]]);
  assert.equal(state(DEL, pending), "pending");
  assert.equal(state(DEL, para(`${HEAD}${I1}`, [])), "rejected");
});

// ---- a replace whose new words a sibling rewrote ----

const BEFORE_R = "Crews checked the span daily. The old bearings were replaced.";
const OLD = "old bearings";
const NEW = "new bearings (with steel plates)";
const REP = change("replace", BEFORE_R, OLD, NEW);
// A sibling of the same turn replaced "steel plates" inside NEW (Word drops
// the old words of a pending insertion), and other siblings rewrote the
// words around the change: only its old words are left to match.
const NEW_NOW = NEW.replace("steel plates", "shims");

test("a replace whose new words a sibling rewrote, its old words still tracked, stays pending", () => {
  assert.deepEqual(REP.context, { left: "span daily. The ", right: " were replaced." });
  const t = `Crews inspected the deck weekly. Its ${NEW_NOW}${OLD} are replaced.`;
  const p = para(t, [
    ["ins", "inspected the deck weekly. Its"],
    ["ins", NEW_NOW],
    ["del", OLD],
    ["ins", "are"],
  ]);
  assert.equal(state(REP, p), "pending");
});

test("controls: a replace that is decided still reads decided", () => {
  // Rejected: the old words are plain text again (the siblings around it
  // accepted, so only the old words match).
  const rejected = para(`Crews inspected the deck weekly. Its ${OLD} are replaced.`, []);
  assert.equal(state(REP, rejected), "rejected");
  // Rejected next to pending siblings: the old words plain, other text tracked.
  const nextTo = para(`Crews inspected the deck weekly. Its ${OLD} are replaced.`, [
    ["ins", "inspected the deck weekly. Its"],
    ["ins", "are"],
  ]);
  assert.equal(state(REP, nextTo), "rejected");
  // Only part of the old words tracked (Word applied a narrower edit): the
  // full form is not pending, it reads rejected as before.
  const narrow = change(
    "replace",
    "Crews carefully checked the span.",
    "carefully checked",
    "checked",
  );
  const part = para("Crews carefully checked the span.", [["del", "carefully"]]);
  assert.equal(state(narrow, part), "rejected");
  // Accepted: the new words stand, nothing tracked.
  assert.equal(
    state(REP, para(`Crews checked the span daily. The ${NEW} were replaced.`, [])),
    "accepted",
  );
});
