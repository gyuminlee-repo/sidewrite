// A request answered partly in this file and partly in another one: the
// change here does not hide the part that belongs elsewhere. The outside
// note carries `part`, and the card shows that part as a row of its own.
// All texts here are made up (a paper on bridge engineering).
import { test } from "node:test";
import assert from "node:assert/strict";
import {
  RevisionSchema,
  emptyRevision,
  replaceComments,
  setRequirementOutside,
  mirrorFromMain,
} from "../daemon/revision-store.mjs";
import { requirementOutside } from "../daemon/review-tools.mjs";
import { outsideNotes } from "../daemon/revision-parts.mjs";
import { buildReviewPrompt } from "../daemon/review-prompt.mjs";
import {
  requirementStates,
  requestStateText,
  unaddressedRequirements,
  commentState,
} from "../taskpane/shared/review-view.js";

const REQS = [
  {
    id: "q1",
    text: "Report the load redistribution after the cable failure in the text and in Table S2",
  },
  { id: "q2", text: "Name the deck material" },
];
const changed = (ids, state = "pending") => ({
  id: `R1.1-${ids.join("")}`,
  state,
  rationale: { what: "w", why: "y", addresses: "a", requirement_ids: ids },
});
const note = (id, extra = {}) => ({ id, why: "Table S2 is in the supplement.", at: "t", ...extra });

test("requirementStates: a changed request with a part elsewhere keeps both rows", () => {
  const c = {
    ref: "R1.1",
    requirements: REQS,
    changes: [changed(["q1"]), changed(["q2"])],
    requirements_outside: [note("q1", { part: "the values in Table S2" })],
  };
  const rows = requirementStates(c);
  assert.deepEqual(
    rows.map((r) => [r.id, r.state, r.part ?? null]),
    [
      ["q1", "changed", null],
      ["q1", "outside_this_file", "the values in Table S2"],
      ["q2", "changed", null],
    ],
  );
  assert.equal(
    requestStateText(rows[1]),
    "This part is in another file: the values in Table S2 (Table S2 is in the supplement.)",
  );
  // Nothing is open, and the comment is not taken as all elsewhere.
  assert.deepEqual(unaddressedRequirements(c), []);
  assert.notEqual(commentState(c).label, "Belongs to another file");
});

test("requirementStates: without a part a note on a changed request reads as before; a part on an untouched request leaves the rest open", () => {
  const whole = {
    requirements: REQS,
    changes: [changed(["q1"])],
    requirements_outside: [note("q1")],
  };
  assert.deepEqual(
    requirementStates(whole).map((r) => [r.id, r.state]),
    [
      ["q1", "changed"],
      ["q2", "open"],
    ],
  );
  const untouched = {
    requirements: REQS,
    changes: [],
    requirements_outside: [note("q1", { part: "Table S2" })],
  };
  assert.deepEqual(
    requirementStates(untouched).map((r) => [r.id, r.state]),
    [
      ["q1", "open"],
      ["q1", "outside_this_file"],
      ["q2", "open"],
    ],
  );
});

test("the store keeps the part: set, parse, a comment-list save, the copy in a supplement", () => {
  let rev = emptyRevision("bridge.docx", new Date("2026-10-01T00:00:00Z"));
  rev = replaceComments(rev, [
    { ref: "R1.1", reviewer: 1, text: "t", requirements: REQS.map((r) => r.text) },
  ]);
  rev = setRequirementOutside(rev, "R1.1", { id: "q1", why: "w", part: "Table S2" });
  rev = setRequirementOutside(rev, "R1.1", { id: "q2", why: "whole" });
  const parsed = RevisionSchema.parse(JSON.parse(JSON.stringify(rev)));
  assert.deepEqual(
    parsed.comments[0].requirements_outside.map((n) => [n.id, n.part]),
    [
      ["q1", "Table S2"],
      ["q2", undefined],
    ],
  );
  assert.ok(!("part" in parsed.comments[0].requirements_outside[1]));
  const saved = replaceComments(parsed, [{ ref: "R1.1", reviewer: 1, text: "t" }]);
  assert.equal(saved.comments[0].requirements_outside[0].part, "Table S2");
  // The supplement keeps its own notes; the requests stay those of the manuscript.
  const part = mirrorFromMain(parsed, parsed);
  assert.equal(part.comments[0].requirements_outside[0].part, "Table S2");
  assert.deepEqual(
    part.comments[0].requirements.map((r) => r.id),
    ["q1", "q2"],
  );
  // The turn of a supplement is told which part was left for it.
  const notes = outsideNotes(parsed, "R1.1");
  assert.deepEqual(notes[0].part, "Table S2");
  assert.ok(!("part" in notes[1]));
  const built = buildReviewPrompt({
    comment: parsed.comments[0],
    part: { name: "supp.docx", main_name: "bridge.docx", outside: notes },
    fence: "F",
  });
  assert.match(
    built.text,
    /- \[q1\] Report the load redistribution[^\n]*\(only this part: Table S2\) \(note: w\)/,
  );
});

test("office_requirement_outside: the part goes with the note", async () => {
  const noted = [];
  const deps = {
    log: () => {},
    getReviewContext: async () => ({ requirements: REQS }),
    onRequirementOutside: async ({ ref, note }) => noted.push([ref, note]),
  };
  const r = await requirementOutside(
    deps,
    { comment_ref: "R1.1", requirement_id: "q1", why: "w", part: "Table S2" },
    { ref: "R1.1" },
  );
  assert.equal(r.ok, true);
  assert.deepEqual(noted, [["R1.1", { id: "q1", why: "w", part: "Table S2" }]]);
  assert.match(r.message, /next to what is changed here/);
});
