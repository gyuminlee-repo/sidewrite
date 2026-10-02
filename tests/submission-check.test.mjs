// The rows of the submission checklist (daemon/submission-check.mjs), one
// fixture per row and state. Pure: an invented record, an invented document
// read and invented file hashes.
import { test } from "node:test";
import assert from "node:assert/strict";
import { submissionRows, READY, NOT_READY, NOT_APPLICABLE } from "../daemon/submission-check.mjs";
import { formatParagraphs } from "../taskpane/shared/format-runs.js";
import {
  FILE,
  SECTION,
  QUOTE,
  liveOf,
  recordOf,
  profileOf,
  docComment,
  rowOf,
} from "./helpers/submission-fixture.mjs";

const rows = ({ mutate, live = liveOf(), profile = profileOf(), files } = {}) =>
  submissionRows({ revision: recordOf(mutate, live), profile, live, files, file: FILE });
const state = (all, id) => rowOf(all, id)?.state;

// The formatting of the two paragraphs of the fixture, as the pane reads it.
const ooxml = (text) =>
  `<pkg:part pkg:name="/word/document.xml"><w:p><w:r><w:t>${text}</w:t></w:r></w:p></pkg:part>`;
const formatOf = (texts) =>
  formatParagraphs(
    texts.map((text) => ({
      ooxml: ooxml(text),
      first_line: 0,
      left: 0,
      is_list: false,
      table_level: 0,
    })),
    [],
  );
const BODY = `Flasks were shaken. ${QUOTE} Samples were drawn daily.`;
const STATS = "The difference was significant (p &lt; 0.05).";
const countRule = (max) => ({
  id: "L1",
  kind: "instruction",
  text: `The methods must not exceed ${max} words.`,
  scope: "manuscript",
  check: { type: "count", target: "section_words", section: "Materials", max },
});

test("every row is about the open file, and a row carries a state, one sentence and a way to the fix", () => {
  const all = rows();
  assert.deepEqual(all.map((r) => r.id).slice(0, 6), [
    "answers",
    "requests",
    "changes_decided",
    "tool_comments",
    "file_clean",
    "format_counts",
  ]);
  for (const r of all) {
    assert.equal(r.file, FILE);
    assert.ok([READY, NOT_READY, NOT_APPLICABLE].includes(r.state), r.id);
    assert.ok(r.text.length > 10 && !r.text.includes("—"), r.id);
    // A row that is not ready says where it is fixed.
    if (r.state === NOT_READY) assert.ok(r.go, `${r.id} has no way to its fix`);
  }
});

test("row 1: ready when every answer passes; a missing answer is not ready and leads to the Response tab", () => {
  const ok = rowOf(rows(), "answers");
  assert.equal(ok.state, READY);
  assert.equal(ok.text, "Every comment has an answer that passes the checks.");

  const bad = rowOf(rows({ mutate: (r) => (r.comments[0].response.text = "") }), "answers");
  assert.equal(bad.state, NOT_READY);
  assert.equal(bad.text, "1 answer is missing or fails a check: R1.1.");
  assert.deepEqual(bad.go, { tab: "response", target: "response-check" });
});

test("row 1: a quote that is no longer in the live document makes it not ready", () => {
  const live = liveOf({ paragraph: { text: "Flasks were shaken. Samples were drawn daily." } });
  assert.equal(state(rows({ live }), "answers"), NOT_READY);
});

test("row 1: rules a model checks do not hold the row back; the sentence says where they run", () => {
  const llm = {
    id: "L9",
    kind: "forbid",
    text: "polite",
    scope: "response",
    check: { type: "llm" },
  };
  const r = rowOf(rows({ mutate: (rev) => rev.ledger.push(llm) }), "answers");
  assert.equal(r.state, READY);
  assert.match(
    r.text,
    /1 rule checked by a model is not run here; run the checks in the Response tab/,
  );
  assert.deepEqual(r.go, { tab: "response", target: "response-check" });
});

test("row 2: a failing `requirements` check is this row, not row 1", () => {
  // The change that answered the request is rejected: the request is open.
  const all = rows({
    mutate: (r) => {
      r.comments[0].changes[0].state = "rejected";
      r.comments[0].response.citations = [];
    },
  });
  const r = rowOf(all, "requests");
  assert.equal(r.state, NOT_READY);
  assert.equal(r.text, "1 request has no accepted change: R1.1.");
  assert.deepEqual(r.go, { tab: "review", target: "review-comments" });
  assert.equal(state(all, "answers"), READY, "an open request is not a failing answer");
  assert.equal(state(rows(), "requests"), READY);
  assert.equal(
    rowOf(rows(), "requests").text,
    "No request is open and no question to you is waiting.",
  );
});

test("row 2: a question to the author without an answer is not ready; answered, it is ready", () => {
  const ask = (answer) => (r) =>
    r.comments[0].questions.push({ id: "a1", question: "Which antifoam?", answer });
  const open = rowOf(rows({ mutate: ask("") }), "requests");
  assert.equal(open.state, NOT_READY);
  assert.equal(open.text, "1 question to you has no answer: R1.1.");
  assert.equal(state(rows({ mutate: ask("Antifoam 204.") }), "requests"), READY);
});

test("row 3: a tracked insertion in the snapshot is not ready and leads to that paragraph in Word", () => {
  const live = liveOf({ paragraph: { w_ins: 1 } });
  const r = rowOf(rows({ live }), "changes_decided");
  assert.equal(r.state, NOT_READY);
  assert.equal(r.text, "1 paragraph still holds tracked changes. Accept or reject them in Word.");
  assert.deepEqual(r.go, { paragraph: 1 });
  const ok = rowOf(rows(), "changes_decided");
  assert.equal(ok.state, READY);
  assert.equal(ok.text, "No tracked change is waiting for a decision in this file.");
});

test("row 4: a comment of the tool is not ready and leads to the cleanup button; a comment of a person is not counted", () => {
  const live = liveOf({ comments: [docComment("1", "[R1.1] unit / asked / q1")] });
  const r = rowOf(rows({ live }), "tool_comments");
  assert.equal(r.state, NOT_READY);
  assert.equal(
    r.text,
    "1 comment of the tool is still in this file. Remove them before you make the marked copy.",
  );
  assert.deepEqual(r.go, { tab: "review", target: "review-clean-tool" });

  const person = liveOf({ comments: [docComment("2", "Check the units [R1.1]")] });
  const ok = rowOf(rows({ live: person }), "tool_comments");
  assert.equal(ok.state, READY);
  assert.equal(ok.text, "The tool's comments are removed.");
});

test("row 5: clean only with no tracked change and no comment at all; it names what is left", () => {
  const ok = rowOf(rows(), "file_clean");
  assert.equal(ok.state, READY);
  assert.equal(ok.text, "This file is clean: no tracked changes and no comments are left.");

  // A comment of a person passes rows 3 and 4 and is caught here only.
  const person = rows({ live: liveOf({ comments: [docComment("2", "Check the units")] }) });
  assert.equal(state(person, "changes_decided"), READY);
  assert.equal(state(person, "tool_comments"), READY);
  const r = rowOf(person, "file_clean");
  assert.equal(r.state, NOT_READY);
  assert.equal(
    r.text,
    "This file is not clean yet. Left in it: 1 comment of yours or the reviewers.",
  );
  assert.deepEqual(r.go, { tab: "review", target: "review-delete-all" });

  const all = rows({
    live: liveOf({
      paragraph: { w_del: 2 },
      comments: [docComment("1", "[R1.1] a / b / c"), docComment("3", "[E.2] a / b / c")],
    }),
  });
  assert.equal(
    rowOf(all, "file_clean").text,
    "This file is not clean yet. Left in it: 1 paragraph with tracked changes, 2 comments of the tool.",
  );
  assert.deepEqual(rowOf(all, "file_clean").go, { tab: "review", target: "review-clean-tool" });
});

test("rows 4 and 5 do not exist when the document read holds no comment list: they are not guessed", () => {
  const live = liveOf();
  delete live.snapshot.comments;
  const all = rows({ live });
  assert.equal(rowOf(all, "tool_comments"), undefined);
  assert.equal(rowOf(all, "file_clean"), undefined);
  assert.equal(state(all, "changes_decided"), READY);
});

test("row 6: not applicable without a format rule in the profile and a count rule in the ledger", () => {
  const r = rowOf(rows(), "format_counts");
  assert.equal(r.state, NOT_APPLICABLE);
  assert.equal(r.text, "No format rule and no count rule is set for this paper.");
  assert.equal(r.go, null);
});

test("row 6: a failing count rule is not ready; the same rule with room is ready", () => {
  const live = liveOf({ format: formatOf([SECTION, BODY]) });
  const bad = rowOf(rows({ live, mutate: (r) => r.ledger.push(countRule(3)) }), "format_counts");
  assert.equal(bad.state, NOT_READY);
  assert.equal(bad.text, "1 count rule fails. Run the checks in the Tools tab to see where.");
  assert.deepEqual(bad.go, { tab: "tools", target: "tools-check" });

  const ok = rowOf(rows({ live, mutate: (r) => r.ledger.push(countRule(100)) }), "format_counts");
  assert.equal(ok.state, READY);
  assert.equal(ok.text, "The format and count checks have no failing rule.");
});

test("row 6: a format finding of a rule the profile turns on is not ready", () => {
  const profile = profileOf({ italics: { statistics: true } });
  const live = liveOf({ format: formatOf([SECTION, STATS]) });
  live.snapshot.paragraphs[1].text = "The difference was significant (p < 0.05).";
  // The letter quotes the old sentence: only the format row is looked at.
  const bad = rowOf(rows({ live, profile }), "format_counts");
  assert.equal(bad.state, NOT_READY);
  assert.equal(bad.text, "1 format finding. Run the checks in the Tools tab to see where.");

  const clean = liveOf({ format: formatOf([SECTION, BODY]) });
  assert.equal(state(rows({ live: clean, profile }), "format_counts"), READY);
});

test("row 6: formatting that was not read, or does not line up with the text, is never ready", () => {
  const profile = profileOf({ italics: { statistics: true } });
  for (const format of [null, formatOf([SECTION])]) {
    const r = rowOf(rows({ live: liveOf({ format }), profile }), "format_counts");
    assert.equal(r.state, NOT_READY);
    assert.match(r.text, /^The formatting of the document could not be read/);
  }
});
