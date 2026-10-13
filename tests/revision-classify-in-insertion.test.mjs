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

// in_insertion with what was left of the carrying insertion on each side
// of the deleted words, as the task pane read it (tools-word.js
// insertionAround). Only a tracked run next to the place that still holds
// that remainder is the carrying insertion.
const held = (left, right) => ({
  in_insertion: true,
  insertion_left: left,
  insertion_right: right,
});

test("author-replace tail deletion flagged in_insertion: pending", () => {
  // "new steel supports " less " supports": "new steel" left, " " right.
  assert.equal(classifyChange(deletion(held("new steel", " ")), snapshot).state, "pending");
});

test("the flag without the remainders proves nothing: accepted, as before", () => {
  assert.equal(classifyChange(deletion({ in_insertion: true }), snapshot).state, "accepted");
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
  assert.equal(classifyChange(d(held("new steel", "")), two).state, "pending");
});

test("the flag needs a tracked insertion next to the place", () => {
  // The carrying insertion was accepted: nothing tracked is left of the
  // place, and the deletion is accepted with it.
  const accepted = para(r(prefix + "new steel ") + del(old) + r(suffix));
  assert.equal(classifyChange(deletion(held("new steel", " ")), accepted).state, "accepted");
  // Plain text right of the place, the insertion on the left: the flag
  // proves the deletion took the end of that insertion; without it nothing
  // changes (an accepted deletion next to an insertion).
  const plainRight = para(r(prefix) + ins("new steel ") + r(old + suffix));
  assert.equal(classifyChange(deletion(held("new steel", " ")), plainRight).state, "pending");
  assert.equal(classifyChange(deletion(), plainRight).state, "accepted");
  assert.equal(classifyChange(deletion({ in_insertion: false }), plainRight).state, "accepted");
});

// The author's own pure insertion "fully checked and " (typed in Word,
// pending) before "painted"; the tool deleted a part of it.
test("a deletion at the end, the start or across the whole of the author's insertion", () => {
  const head = "The deck was ";
  const typed = "fully checked and ";
  const tail = "painted after the repair.";
  const input = head + typed + tail;
  const d = (find, extra) => make("delete", find, "", input, "R2.1-1", extra);
  const read = (find, body, flag) => {
    const text = input.replace(find, "");
    return classifyChange(d(find, flag), para(body, text, text)).state;
  };
  const cases = [
    // End: " and" leaves "fully checked " tracked on the left, plain right.
    [" and", r(head) + ins("fully checked ") + r(tail), held("fully checked", " ")],
    // Start: "fully " leaves plain text left, "checked and " tracked right.
    ["fully ", r(head) + ins("checked and ") + r(tail), held("", "checked and ")],
  ];
  for (const [find, body, flag] of cases) {
    assert.equal(read(find, body, flag), "pending", find);
    // Controls: no flag, or a pane that read plain text: as before.
    assert.equal(read(find, body), "accepted", find);
    assert.equal(read(find, body, { in_insertion: false }), "accepted", find);
  }
  // Whole: nothing of the insertion is left, so nothing is pending in Word
  // either (no revision mark remains to accept or reject). The paragraph
  // reads exactly as an accepted deletion of plain text: accepted, flag or
  // not. Known limit, and the state Word itself is in.
  const whole = r(head + tail);
  assert.equal(read(typed, whole, held("", "")), "accepted");
  assert.equal(read(typed, whole), "accepted");
});

// Codex review: the flag says the deleted words were in SOME insertion, not
// that the insertion next to the place now is that one.
test("an independent insertion next to the place is no proof", () => {
  const head = "The deck was ";
  const typed = "fully checked and ";
  const tail = "painted after the repair.";
  const input = head + typed + tail;
  const read = (find, body, flag) => {
    const text = input.replace(find, "");
    return classifyChange(make("delete", find, "", input, "R2.3-1", flag), para(body, text, text))
      .state;
  };
  // The whole insertion deleted; "painted" an independent pending insertion.
  const painted = r(head) + ins("painted") + r(" after the repair.");
  assert.equal(read(typed, painted, held("", "")), "accepted");
  assert.equal(read(typed, painted, { in_insertion: true }), "accepted");
  // " and" took the end of "fully checked and "; the author has since
  // accepted that insertion ("fully checked" now plain) and "painted" is an
  // independent insertion right of the place: its text is not the recorded
  // remainder, so it proves nothing.
  const since = r(head + "fully checked ") + ins("painted") + r(" after the repair.");
  assert.equal(read(" and", since, held("fully checked", " ")), "accepted");
  // The same on the left: "fully " took the start; an independent
  // insertion "was " left of the place is not "checked and ".
  const left = r("The deck ") + ins("was ") + r("checked and " + tail);
  assert.equal(read("fully ", left, held("", "checked and ")), "accepted");
});

test("the start of the author's insertion at the start of the paragraph", () => {
  // "Checked twice: " typed at the start; the tool deleted "Checked ".
  const input = "Checked twice: the deck was painted.";
  const text = "twice: the deck was painted.";
  const d = (extra) => ({ ...make("delete", "Checked ", "", input, "R2.2-1"), ...extra });
  const snap = para(ins("twice: ") + r("the deck was painted."), text, text);
  assert.equal(classifyChange(d(held("", "twice: ")), snap).state, "pending");
  assert.equal(classifyChange(d(), snap).state, "accepted");
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
  const got = classifyChange(deletion(held("new steel", " ")), rejected).state;
  assert.equal(got, classifyChange(deletion(), rejected).state, "the flag decides nothing here");
  assert.equal(got, "user_edited");
});

// A replace whose old words lay inside the author's insertion now carries
// the flag and remainders too (for the note on its card). Its state is read
// from its own tracked new words, as before: the flag changes nothing.
test("a replace inside the author's insertion: the same state with and without the flag", () => {
  const head = "The deck was ";
  const typed = "fully checked and ";
  const tail = "painted after the repair.";
  const input = head + typed + tail;
  const change = (extra) => make("replace", "checked", "inspected", input, "R2.4-1", extra);
  const flag = held("fully ", " and ");
  const after = head + "fully inspected and " + tail;
  const shapes = [
    // Pending: the author's insertion and the tool's new words, all tracked.
    [para(r(head) + ins("fully inspected and ") + r(tail), after, after), "pending"],
    // Accepted: all plain.
    [para(r(head + "fully inspected and " + tail), after, after), "accepted"],
  ];
  for (const [snap, state] of shapes) {
    assert.equal(classifyChange(change(), snap).state, state);
    assert.equal(classifyChange(change(flag), snap).state, state);
    assert.equal(classifyChange(change({ in_insertion: false }), snap).state, state);
  }
  // The new words rejected: neither form stands; whatever it reads, the
  // flag does not change it.
  const gone = head + "fully  and " + tail;
  const rejected = para(r(head) + ins("fully  and ") + r(tail), gone, gone);
  assert.equal(
    classifyChange(change(flag), rejected).state,
    classifyChange(change(), rejected).state,
  );
});
