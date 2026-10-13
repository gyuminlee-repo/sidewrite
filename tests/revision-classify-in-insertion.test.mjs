// A tool deletion whose words lay inside the author's own pending tracked
// insertion (typed in Word, not a recorded change). Word drops deleted text
// of a pending insertion outright, so the deletion leaves nothing of its
// own; the paragraph reads as an accepted deletion would. Only a recorded
// sibling proved the insertion before (inInsertion), and an author's
// insertion is no recorded change: the deletion read accepted while it
// still waits on the insertion. The task pane now says at apply time
// whether the deleted range lay inside a tracked insertion (in_insertion,
// tools-word.js insertionHolds); that is the missing proof. Adjacency alone
// is not: an accepted plain-text deletion between an independent insertion
// and an independent deletion has the same shape (PR #32), and plain text
// carries no flag. Cases from
// docs/review/external-insertion-evidence-2026-10-13/reproduce.mjs.
// Invented text (bridge works).
import { test } from "node:test";
import assert from "node:assert/strict";
import { trackedRuns, countRevisionMarks } from "../taskpane/shared/review-view.js";
import { changeContext, classifyChange } from "../daemon/revision-classify.mjs";

const rationale = { what: "w", why: "y", addresses: "a" };
const prefix = "Crews installed ";
const inserted = "new steel supports ";
const old = "old piers";
const suffix = " beside the road.";
const before = prefix + inserted + old + suffix;
const after = before.replace(" supports", "");
const make = (kind, find, text, input, id, extra = {}) => ({
  id,
  kind,
  find,
  text,
  before: input,
  after: input.replace(find, text),
  paragraph_index: 4,
  paragraph_id: "P4",
  word_comment_id: null,
  rationale,
  context: changeContext(input, find),
  ...extra,
});
const deletion = (extra) => make("delete", " supports", "", before, "R1.1-2", extra);
const sibling = make("replace", old, inserted, prefix + old + suffix, "R1.1-1");

// The paragraph as the task pane reads it: runs from its OOXML.
const para = (body, text = after, reviewed = prefix + "new steel" + suffix) => {
  const ooxml = `<pkg:package><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body><w:p>${body}</w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
  return {
    paragraphs: [
      { index: 4, id: "P4", text, reviewed_text: reviewed, ...countRevisionMarks(ooxml) },
    ].map((p) => ({ ...p, tracked: trackedRuns(ooxml) })),
    comments: [],
  };
};
const r = (s) => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`;
const ins = (s) => `<w:ins w:id="1" w:author="A" w:date="2026-10-12T09:00:00Z">${r(s)}</w:ins>`;
const del = (s) =>
  `<w:del w:id="2" w:author="A" w:date="2026-10-12T09:00:00Z"><w:r><w:delText>${s}</w:delText></w:r></w:del>`;
// The author replaced "old piers" by "new steel supports " in Word; the
// tool deleted " supports" from its end. Same runs as the independent
// control below: an insertion "new steel " and a deletion "old piers".
const snapshot = para(r(prefix) + ins("new steel ") + del(old) + r(suffix));

test("author-replace tail deletion flagged in_insertion: pending", () => {
  assert.equal(classifyChange(deletion({ in_insertion: true }), snapshot).state, "pending");
});

test("the same deletion without the flag (an older task pane): accepted, as before", () => {
  assert.equal(classifyChange(deletion(), snapshot).state, "accepted");
});

test("controls: a plain-text deletion between an independent insertion and deletion, accepted", () => {
  // The pane read the deleted range outside any tracked insertion.
  assert.equal(classifyChange(deletion({ in_insertion: false }), snapshot).state, "accepted");
  assert.equal(classifyChange(deletion(), snapshot).state, "accepted");
});

test("controls: a recorded replace sibling still proves it, flag or not", () => {
  assert.equal(classifyChange(deletion(), snapshot, [sibling, deletion()]).state, "pending");
  assert.equal(
    classifyChange(deletion({ in_insertion: false }), snapshot, [sibling]).state,
    "pending",
  );
});

test("the flag opens the junction of two insertion runs too", () => {
  // "new steel supports" and "and beams " typed as two insertions with a
  // plain space between them; " supports" was the end of the first one
  // when the tool deleted it. Two runs, and no recorded sibling.
  const text = prefix + "new steel and beams " + old + suffix;
  const two = para(
    r(prefix) + ins("new steel") + r(" ") + ins("and beams ") + r(old + suffix),
    text,
    text,
  );
  const input = prefix + "new steel supports and beams " + old + suffix;
  const d = (extra) => make("delete", " supports", "", input, "R1.1-3", extra);
  assert.equal(classifyChange(d(), two).state, "accepted");
  assert.equal(classifyChange(d({ in_insertion: true }), two).state, "pending");
});

test("the flag never bypasses the insertion on the left or the tracked text on the right", () => {
  // The carrying insertion was accepted: nothing tracked is left of the
  // place, and the deletion is accepted with it.
  const accepted = para(r(prefix + "new steel ") + del(old) + r(suffix));
  assert.equal(classifyChange(deletion({ in_insertion: true }), accepted).state, "accepted");
  // Plain text right of the place: the flag alone opens nothing new there.
  const plainRight = para(r(prefix) + ins("new steel ") + r(old + suffix));
  assert.equal(classifyChange(deletion({ in_insertion: true }), plainRight).state, "accepted");
});

test("same-type sweep: a replace inside the author's insertion needs no flag", () => {
  // The tool replaced "supports" by "beams" inside the author's pending
  // insertion "new steel supports ". Word drops the old words and its new
  // words stand as a tracked insertion (here merged with the author's run):
  // the replace keeps text of its own, so ownTracked proves it pending.
  const input = prefix + inserted + old + suffix;
  const rep = make("replace", "supports", "beams", input, "R1.1-4");
  const text = prefix + "new steel beams " + old + suffix;
  const pending = para(r(prefix) + ins("new steel beams ") + r(old + suffix), text, text);
  assert.equal(classifyChange(rep, pending).state, "pending");
  // The author accepts the insertion: all plain, accepted.
  const accepted = para(r(text), text, text);
  assert.equal(classifyChange(rep, accepted).state, "accepted");
});

test("the author rejects the carrying insertion: the deletion's words are gone with it", () => {
  // Rejecting the replace takes "new steel " out and restores "old piers":
  // the paragraph reads as before the author's edit. The deletion's context
  // ("steel" on its left) is gone, and its own words never come back.
  const text = prefix + old + suffix;
  const rejected = para(r(text), text, text);
  const got = classifyChange(deletion({ in_insertion: true }), rejected).state;
  assert.equal(got, classifyChange(deletion(), rejected).state, "the flag decides nothing here");
  assert.equal(got, "user_edited");
});
