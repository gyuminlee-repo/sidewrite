// The judges of the acceptance stage `parts` (scripts/e2e/judge.mjs) on
// recorded inputs: what the pane driver's view returns, the two review files,
// what word-com.ps1 answers and lines of the daemon log. No Word is run.
// Each judge is given the answer it must pass on, and the answers it must
// fail on with a line that says which pane or file, what was expected and
// what was found.
import { test } from "node:test";
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { createRequire } from "node:module";
import { readFile } from "node:fs/promises";
import {
  JUDGES,
  PARTS_SEED_RULE,
  docStem,
  forbidBlockEvidence,
  foreignLogLines,
  onlyOwnNames,
  partsBeforeJoin,
  partsBoth,
  partsCrossState,
  partsGuardEdit,
  partsJoined,
  partsKeep,
  partsLeft,
  partsMainOutside,
  partsSavedLine,
  partsSessions,
  partsSuppChange,
} from "../scripts/e2e/judge.mjs";
import { splitReview } from "../daemon/review-split.mjs";

const here = (p) => new URL(p, import.meta.url);
const SPEC = JSON.parse(
  execFileSync("python3", [here("../scripts/e2e/make_parts_fixture.py").pathname, "--spec"], {
    encoding: "utf8",
  }),
);
const names = SPEC.names;
const CANDIDATE = "Join the revision of paper.docx, 4 comments";

// ---- recorded shapes ---------------------------------------------------------------

const requests = (ref) => SPEC.requirements[ref].map((text, i) => ({ id: `q${i + 1}`, text }));
const comment = (ref, extra = {}) => ({
  ref,
  reviewer: 1,
  text: `Comment ${ref}.`,
  status: "pending",
  requirements: requests(ref),
  requirements_outside: [],
  questions: [],
  changes: [],
  blocked: [],
  ...extra,
});
const REFS = ["R1.1", "R1.2", "R1.3", "R1.4"];
const RULE = {
  id: "L1",
  kind: "forbid",
  text: "Do not call the strain or the carbon source novel",
  scope: "manuscript",
  check: { type: "regex", pattern: "\\bnovel(?:ty)?\\b", flags: "i", expect: "absent" },
};
const revision = (over = {}, extra = {}) => ({
  comments: REFS.map((ref) => comment(ref, over[ref] ?? {})),
  ledger: [RULE],
  ...extra,
});
const change = (ref, n, state = "pending", ids = ["q1"], more = {}) => ({
  id: `${ref}-${n}`,
  kind: "replace",
  find: "incubated overnight",
  text: "incubated overnight at 30 degrees Celsius",
  state,
  word_comment_id: String(10 + n),
  rationale: { what: "w", why: "y", addresses: "a", requirement_ids: ids },
  ...more,
});
const PART_OF = { manuscript: "paper.docx", label: "SI", joined_at: "2026-10-02T08:00:00.000Z" };
const PARTS = [{ doc: "SI.docx", label: "SI", joined_at: "2026-10-02T08:00:00.000Z" }];
// What the driver's view gives for a pane (the fields these judges read).
const row = (ref, state = "Not started", reqs = null) => ({
  ref,
  state,
  requests: reqs ?? requests(ref).map((r) => ({ id: r.id, state: "open", text: "Not yet" })),
  can_propose: true,
  changes: [],
  blocked: [],
});
const view = (over = {}) => ({
  ok: true,
  message: "",
  message_error: false,
  comments: REFS.map((ref) => row(ref)),
  linked: "",
  link_box: { shown: false, candidates: [] },
  ledger: ["L1 forbid Forbidden wording Do not call the strain or the carbon source novel."],
  ledger_note: "",
  ledger_edit_buttons: 2,
  ...over,
});
// word-com.ps1 state: n tracked revisions, the texts of its comments.
const word = (revisions, comments = []) => ({
  ok: true,
  revisions: Array.from({ length: revisions }, (_, i) => ({ type: 1, text: `r${i}` })),
  comments: comments.map((text) => ({ text })),
});
const failedNames = (v) =>
  Object.entries(v.checks)
    .filter(([, ok]) => !ok)
    .map(([k]) => k);

// ---- the stage and its fixtures --------------------------------------------------

test("the stage runs on paper.docx, and its review splits into the four comments of the spec", async () => {
  assert.equal(docStem("parts"), "paper");
  assert.deepEqual(names, { paper: "paper.docx", si: "SI.docx" });
  const text = await readFile(here("./fixtures/parts_review.txt"), "utf8");
  const split = await splitReview(text, {});
  const refs = split.comments.map((c) => c.ref);
  assert.deepEqual(
    refs.filter((r) => /^R\d+\.\d+$/.test(r)),
    REFS,
  );
  assert.deepEqual(Object.values(SPEC.refs).sort(), REFS);
  // The judges of the steps that prepare the manuscript.
  const s = JUDGES["parts-split"]({ message: "Split into 5 comments (x).", draft_refs: refs });
  assert.equal(s.pass, true, s.got);
  assert.equal(
    JUDGES["parts-split"]({ message: "Split into 3 comments.", draft_refs: ["R1.1", "R1.2"] }).pass,
    false,
  );
  const drafted = {
    comments: REFS.map((ref) =>
      row(
        ref,
        "Not started",
        requests(ref).map((r) => ({ id: r.id })),
      ),
    ),
  };
  const d = JUDGES["parts-draft"](drafted, { requirements: SPEC.requirements });
  assert.equal(d.pass, true, d.got);
  drafted.comments[1].requests.pop();
  assert.equal(JUDGES["parts-draft"](drafted, { requirements: SPEC.requirements }).pass, false);
  assert.equal(JUDGES["parts-ledger"]({ ledger: view().ledger }).pass, true);
  assert.equal(JUDGES["parts-ledger"]({ ledger: [] }).pass, false);
});

test("the guard edit is the one of the spec, under a rule the ledger does not have", () => {
  const edit = partsGuardEdit(SPEC.guard);
  assert.deepEqual(
    [edit.comment_ref, edit.find, edit.replace],
    [SPEC.refs.si_only, SPEC.guard.find, SPEC.guard.replace],
  );
  assert.match(edit.replace, new RegExp(`\\b${SPEC.forbidden_word}\\b`));
  for (const k of ["what", "why", "addresses"]) assert.ok(edit.rationale[k]);
  assert.notEqual(PARTS_SEED_RULE, RULE.id);
  // The judge of the stage-2 guard reads the supplementary file's review file
  // when it is given its comment and its rule.
  const seed = { id: "R1.1-b1", args: edit, blocked_by: [PARTS_SEED_RULE], state: "blocked" };
  const before = revision({ "R1.1": { blocked: [seed] } });
  const after = revision({
    "R1.1": {
      blocked: [seed, { id: "R1.1-b2", args: edit, blocked_by: ["L1"], state: "blocked" }],
    },
  });
  const e = forbidBlockEvidence({
    step: {
      ok: true,
      message_error: true,
      message: "The edit matches forbidden rule(s) L1 and was not applied.",
    },
    log: "[review] blocked R1.1 by=L1\n",
    revision: after,
    before,
    word: word(1),
    wordBefore: word(1),
    seedId: seed.id,
    ref: "R1.1",
    rule: "L1",
  });
  assert.equal(e.pass, true, JSON.stringify(e.checks));
  // The edit went into Word after all: not blocked.
  const applied = forbidBlockEvidence({
    step: { ok: true, message_error: false, message: "" },
    log: "",
    revision: before,
    before,
    word: word(2),
    wordBefore: word(1),
    seedId: seed.id,
    ref: "R1.1",
    rule: "L1",
  });
  assert.deepEqual(
    [applied.pass, applied.checks.refused, applied.checks.word_unchanged],
    [false, false, false],
  );
});

// ---- a: before the join --------------------------------------------------------------

test("a: the join box with the candidate, and no link in either file", () => {
  const before = view({ comments: [], link_box: { shown: true, candidates: [CANDIDATE] } });
  const ok = partsBeforeJoin({ view: before, si: null, paper: revision(), names, comments: 4 });
  assert.equal(ok.pass, true, ok.got);
  assert.match(ok.got, /Join the revision of paper\.docx, 4 comments/);
  // A review file of the supplement that exists and holds no link is fine too.
  assert.equal(
    partsBeforeJoin({ view: before, si: { comments: [] }, paper: revision(), names, comments: 4 })
      .pass,
    true,
  );
  // It joined by itself.
  const auto = partsBeforeJoin({
    view: view({ link_box: { shown: false, candidates: [] } }),
    si: revision({}, { part_of: PART_OF }),
    paper: revision({}, { parts: PARTS }),
    names,
    comments: 4,
  });
  assert.equal(auto.pass, false);
  assert.equal(failedNames(auto).length, 5);
  assert.match(
    auto.got,
    /review file of SI\.docx: no link \(nothing joins by itself\): expected no part_of, found \{"manuscript":"paper\.docx"/,
  );
  assert.match(auto.got, /SI\.docx pane: the join box is shown: expected true, found false/);
  // The pane of the supplementary file was never reached.
  const none = partsBeforeJoin({
    view: { ok: false, error: "pane not reachable on port 9333: no pane page of SI.docx yet" },
    si: null,
    paper: revision(),
    names,
    comments: 4,
  });
  assert.equal(none.pass, false);
  assert.match(none.got, /no pane page of SI\.docx yet/);
  // Another number of comments on the button.
  const three = partsBeforeJoin({
    view: view({
      comments: [],
      link_box: { shown: true, candidates: ["Join the revision of paper.docx, 3 comments"] },
    }),
    si: null,
    paper: revision(),
    names,
    comments: 4,
  });
  assert.deepEqual(failedNames(three), ["SI.docx pane: the candidate"]);
});

// ---- b: joined -------------------------------------------------------------------------

const JOINED_VIEW = view({
  linked: "This file is part of the revision of paper.docx. Leave this revision",
  ledger_note: "These rules are edited in paper.docx.",
  ledger_edit_buttons: 0,
});
const MAIN_LINKED = { ...view({ linked: "Linked files: SI.docx" }), matched: true };

test("b: the rows, the read-only rules, the link at both ends, the same request ids, and the manuscript's line", () => {
  const args = {
    view: JOINED_VIEW,
    mainView: MAIN_LINKED,
    si: revision({}, { part_of: PART_OF }),
    paper: revision({}, { parts: PARTS }),
    names,
  };
  const ok = partsJoined(args);
  assert.equal(ok.pass, true, ok.got);
  // Request ids that differ (numbered anew in the supplement).
  const renumbered = revision({}, { part_of: PART_OF });
  renumbered.comments[1].requirements = [
    { id: "q1", text: requests("R1.2")[1].text },
    { id: "q2", text: requests("R1.2")[0].text },
  ];
  assert.deepEqual(failedNames(partsJoined({ ...args, si: renumbered })), [
    "request ids are the same in both review files",
  ]);
  // The manuscript's pane never showed the line (it would need a reload).
  const stale = partsJoined({ ...args, mainView: { ...view(), matched: false } });
  assert.deepEqual(failedNames(stale), [
    'paper.docx pane: "Linked files: SI.docx" without a reload',
  ]);
  assert.match(stale.got, /expected Linked files: SI\.docx, found $/);
  // The rules can still be edited in the supplement, three rows only, one end of the link.
  const bad = partsJoined({
    ...args,
    view: view({ comments: REFS.slice(0, 3).map((ref) => row(ref)), ledger_edit_buttons: 2 }),
    paper: revision(),
  });
  assert.deepEqual(failedNames(bad), [
    "SI.docx pane: the rows of the manuscript",
    "SI.docx pane: the rules are shown read-only with their note",
    "review file of paper.docx: parts lists SI.docx",
  ]);
  assert.match(bad.got, /found 1 rule\(s\), 2 Edit\/Remove button\(s\), ""/);
  // No Propose changes in the supplement (a pane from before step A2).
  const a1 = partsJoined({
    ...args,
    view: { ...JOINED_VIEW, comments: REFS.map((ref) => ({ ...row(ref), can_propose: false })) },
  });
  assert.deepEqual(failedNames(a1), ["SI.docx pane: Propose changes on every row"]);
  // An absolute path in the link is the same file.
  assert.equal(
    partsJoined({
      ...args,
      si: revision({}, { part_of: { ...PART_OF, manuscript: "D:\\run\\paper.docx" } }),
    }).pass,
    true,
  );
});

// ---- c: the manuscript's turn notes the request as another file's -----------------

test("c: every request of the comment noted as outside, no change, Word as it was", () => {
  const before = revision({}, { parts: PARTS });
  const outside = [{ id: "q1", why: "It is in the Supplementary Methods." }];
  const after = revision({ "R1.1": { requirements_outside: outside } }, { parts: PARTS });
  const args = { before, after, word: word(0), wordBefore: word(0), ref: "R1.1", names };
  const ok = partsMainOutside(args);
  assert.equal(ok.pass, true, ok.got);
  assert.match(ok.got, /It is in the Supplementary Methods/);
  // The turn changed the manuscript instead.
  const changed = partsMainOutside({
    ...args,
    after: revision({ "R1.1": { changes: [change("R1.1", 1)] } }),
    word: word(2),
  });
  assert.deepEqual(failedNames(changed), [
    "review file of paper.docx: R1.1 requests noted as belonging to another file",
    "review file of paper.docx: no change recorded for R1.1",
    "paper.docx in Word: no new tracked change",
  ]);
  assert.match(changed.got, /found outside \[\], questions 0, changes 1/);
  assert.match(changed.got, /expected 0, found 2/);
  // It asked the author instead, and Word could not be read.
  const asked = partsMainOutside({
    ...args,
    after: revision({ "R1.1": { questions: [{ id: "a1", question: "?" }] } }),
    word: { ok: false, error: "bound to another document" },
  });
  assert.equal(asked.pass, false);
  assert.match(asked.got, /questions 1, changes 0/);
  assert.match(asked.got, /bound to another document/);
});

// ---- d: the supplement's turn ------------------------------------------------------

test("d: a tracked change with its rationale comment in the supplement only, the manuscript's text as it was", () => {
  const siBefore = revision({}, { part_of: PART_OF });
  const si = revision({ "R1.1": { changes: [change("R1.1", 1)] } }, { part_of: PART_OF });
  const paper = revision({}, { parts: PARTS });
  const texts = ["Title", "The main cultures were shaken at a constant speed."];
  const args = {
    siBefore,
    si,
    paperBefore: paper,
    paper,
    word: word(2, ["[R1.1] what: temperature stated"]),
    wordBefore: word(0),
    paperTexts: [texts, [...texts]],
    ref: "R1.1",
    names,
  };
  const ok = partsSuppChange(args);
  assert.equal(ok.pass, true, ok.got);
  assert.match(ok.got, /1 change\(s\) in SI\.docx: R1\.1-1/);
  // The change landed in the manuscript's files.
  const wrong = partsSuppChange({
    ...args,
    si: siBefore,
    paper: revision({ "R1.1": { changes: [change("R1.1", 1)] } }),
    word: word(0),
    paperTexts: [texts, ["Title", "The main cultures were shaken at 30 degrees."]],
  });
  assert.deepEqual(failedNames(wrong), [
    "review file of SI.docx: a change recorded for R1.1",
    "SI.docx in Word: a tracked change with a [R1.1] rationale comment",
    "review file of paper.docx: no change recorded for R1.1",
    "paper.docx in Word: its text is as it was",
  ]);
  assert.match(wrong.got, /found differs at paragraph 1/);
  assert.match(wrong.got, /expected 0, found \["R1\.1-1"\]/);
  // Applied without its comment.
  const bare = partsSuppChange({
    ...args,
    si: revision({
      "R1.1": { changes: [change("R1.1", 1, "pending", ["q1"], { word_comment_id: null })] },
    }),
    word: word(2),
  });
  assert.deepEqual(failedNames(bare), [
    "SI.docx in Word: a tracked change with a [R1.1] rationale comment",
  ]);
  assert.match(bare.got, /0 \[R1\.1\] comment\(s\), without a comment: \["R1\.1-1"\]/);
  // The manuscript's text could not be read: not taken for unchanged.
  const unread = partsSuppChange({ ...args, paperTexts: [null, null] });
  assert.deepEqual(failedNames(unread), ["paper.docx in Word: its text is as it was"]);
  assert.match(unread.got, /the text could not be read/);
});

// ---- f: where the request stands, in the manuscript's pane -------------------------

test("f: the manuscript's row and request say the change waits in the supplement", () => {
  const si = revision({ "R1.1": { changes: [change("R1.1", 1)] } }, { part_of: PART_OF });
  const waiting = [
    { id: "q1", state: "waiting_elsewhere", text: "Waiting for your decision in SI.docx" },
  ];
  const shown = {
    ...view({
      comments: REFS.map((ref) =>
        ref === "R1.1" ? row(ref, "1 change to review in SI.docx", waiting) : row(ref),
      ),
    }),
    matched: true,
  };
  const ok = partsCrossState({ view: shown, si, ref: "R1.1", names });
  assert.equal(ok.pass, true, ok.got);
  // The pane was not told: the row is what the manuscript's own turn left.
  const untold = {
    ...view({
      comments: REFS.map((ref) =>
        ref === "R1.1"
          ? row(ref, "Belongs to another file", [
              { id: "q1", state: "outside_this_file", text: "In another file: see SI" },
            ])
          : row(ref),
      ),
    }),
    matched: false,
  };
  const stale = partsCrossState({ view: untold, si, ref: "R1.1", names });
  assert.deepEqual(failedNames(stale), [
    "paper.docx pane: the row of R1.1 says the change waits in SI.docx (no Refresh pressed)",
    'paper.docx pane: the request shows "Waiting for your decision in SI.docx"',
  ]);
  assert.match(stale.got, /found Belongs to another file/);
  assert.match(stale.got, /found \["q1: outside_this_file \\"In another file: see SI\\""\]/);
  // The change of the supplement names no request: no request row can follow it.
  const unnamed = partsCrossState({
    view: shown,
    si: revision({ "R1.1": { changes: [change("R1.1", 1, "pending", [])] } }),
    ref: "R1.1",
    names,
  });
  assert.equal(unnamed.pass, false);
  assert.match(
    unnamed.got,
    /a pending change of R1\.1 names a request: expected at least one request id/,
  );
  // Two changes: "2 changes to review in SI.docx".
  const two = {
    ...shown,
    comments: shown.comments.map((c) =>
      c.ref === "R1.1" ? { ...c, state: "2 changes to review in SI.docx" } : c,
    ),
  };
  assert.equal(partsCrossState({ view: two, si, ref: "R1.1", names }).pass, true);
});

// ---- g: a turn in each pane at the same time ---------------------------------------

const BOTH_LOG = [
  "[review] prompt R1.2 instructions=- forbid=L1 omitted_forbid=- paragraphs=0 manuscript=changed:1",
  "[review] part turn R1.2 file=SI.docx main=paper.docx outside=-",
  "[review] prompt R1.2 instructions=- forbid=L1 omitted_forbid=- paragraphs=0 manuscript=changed:1",
  "[review] change recorded R1.2-1 kind=replace comment=21 paragraph=5",
  "[review] turn done R1.2",
  "[review] change recorded R1.2-1 kind=replace comment=4 paragraph=3",
  "[review] turn done R1.2",
].join("\n");

test("g: each change in its own document and review file, the turns overlapping, nothing unknown or crossed", () => {
  const paperBefore = revision({}, { parts: PARTS });
  const siBefore = revision({ "R1.1": { changes: [change("R1.1", 1)] } }, { part_of: PART_OF });
  const args = {
    paperBefore,
    paper: revision({ "R1.2": { changes: [change("R1.2", 1)] } }, { parts: PARTS }),
    siBefore,
    si: revision(
      {
        "R1.1": { changes: [change("R1.1", 1)] },
        "R1.2": { changes: [change("R1.2", 1, "pending", ["q2"])] },
      },
      { part_of: PART_OF },
    ),
    wordPaper: word(2, ["[R1.2] what: speed"]),
    wordPaperBefore: word(0),
    wordSi: word(4, ["[R1.1] what: temperature", "[R1.2] what: speed"]),
    wordSiBefore: word(2, ["[R1.1] what: temperature"]),
    log: BOTH_LOG,
    steps: { paper: { ok: true }, si: { ok: true } },
    ref: "R1.2",
    names,
  };
  const ok = partsBoth(args);
  assert.equal(ok.pass, true, ok.got);
  assert.equal(ok.got, 'paper.docx: ["R1.2-1"]; SI.docx: ["R1.2-1"]');
  // One after the other: the second turn was queued after the first ended.
  const serial = partsBoth({
    ...args,
    log: BOTH_LOG.split("\n")
      .filter((_, i) => i !== 2)
      .concat("[review] prompt R1.2 instructions=- forbid=L1 omitted_forbid=- paragraphs=0")
      .join("\n"),
  });
  assert.deepEqual(failedNames(serial), ["the two turns ran at the same time"]);
  assert.match(serial.got, /prompt lines at \[0,6\], first turn done at 3/);
  // A preview is no turn.
  const preview = partsBoth({
    ...args,
    log: BOTH_LOG.replace(
      /manuscript=changed:1\n\[review\] part/,
      "manuscript=changed:1 (preview)\n[review] part",
    ),
  });
  assert.deepEqual(failedNames(preview), ["the two turns ran at the same time"]);
  // The supplement's edit was lost: nothing new in its files; Word did not answer one.
  const lost = partsBoth({
    ...args,
    si: siBefore,
    wordSi: word(2, ["[R1.1] what: temperature"]),
    log: `${BOTH_LOG}\n[review] outcome unknown R1.2 code=Timeout id=t9\n[bridge] tool_result t9 from another pane; ignoring`,
    steps: {
      paper: { ok: true },
      si: { ok: false, error: "turn R1.2 still running after 1200000 ms" },
    },
  });
  assert.deepEqual(failedNames(lost), [
    "both propose steps ended without an error",
    "review file of SI.docx: a change recorded for R1.2",
    "SI.docx in Word: its own tracked change(s) with their [R1.2] comments",
    "no edit whose outcome is unknown",
    "no tool result crossed between the panes",
  ]);
  assert.match(lost.got, /SI\.docx: turn R1\.2 still running/);
  assert.match(lost.got, /from another pane/);
  // Both comments landed in one document: more comments there than its changes.
  const crossed = partsBoth({
    ...args,
    wordPaper: word(4, ["[R1.2] what: speed", "[R1.2] what: speed again"]),
    wordSi: word(2, ["[R1.1] what: temperature"]),
  });
  assert.deepEqual(failedNames(crossed), [
    "paper.docx in Word: its own tracked change(s) with their [R1.2] comments",
    "SI.docx in Word: its own tracked change(s) with their [R1.2] comments",
  ]);
});

// ---- h: the manuscript keeps the comment --------------------------------------------

test("h: the save without the comment is refused and the list is as it was", () => {
  const before = revision({}, { parts: PARTS });
  const refused = {
    ...view(),
    message: "Comment R1.1 has recorded change(s) in SI.docx; keep it",
    message_error: true,
  };
  const ok = partsKeep({ step: refused, before, after: before, ref: "R1.1", names });
  assert.equal(ok.pass, true, ok.got);
  // It went through.
  const saved = partsKeep({
    step: { ...view(), message: "", message_error: false },
    before,
    after: { comments: before.comments.filter((c) => c.ref !== "R1.1") },
    ref: "R1.1",
    names,
  });
  assert.deepEqual(failedNames(saved), [
    "paper.docx pane: the save is refused with the reason",
    "review file of paper.docx: the comment list is as it was",
  ]);
  assert.match(saved.got, /found \["R1\.2","R1\.3","R1\.4"\]/);
  // Refused for another reason.
  const other = partsKeep({
    step: { ...view(), message: "Bad ref", message_error: true },
    before,
    after: before,
    ref: "R1.1",
    names,
  });
  assert.deepEqual(failedNames(other), ["paper.docx pane: the save is refused with the reason"]);
  assert.match(other.got, /found Bad ref/);
});

// ---- j: leave ---------------------------------------------------------------------------

test("j: the join box is back, the link is gone at both ends, the supplement keeps its changes", () => {
  const siBefore = revision(
    { "R1.1": { changes: [change("R1.1", 1)] }, "R1.2": { changes: [change("R1.2", 1)] } },
    { part_of: PART_OF },
  );
  const si = { comments: siBefore.comments.filter((c) => c.changes.length), ledger: [] };
  const left = view({
    comments: si.comments.map((c) => row(c.ref)),
    link_box: { shown: true, candidates: [CANDIDATE] },
  });
  const args = { view: left, siBefore, si, paper: revision(), names, comments: 4 };
  const ok = partsLeft(args);
  assert.equal(ok.pass, true, ok.got);
  assert.equal(ok.got, "SI.docx left; it keeps 2 change(s)");
  // The changes went with the copied list, and the manuscript still lists the file.
  const bad = partsLeft({
    ...args,
    view: view({ link_box: { shown: false, candidates: [] } }),
    si: { comments: [], ledger: [] },
    paper: revision({}, { parts: PARTS }),
  });
  assert.deepEqual(failedNames(bad), [
    "SI.docx pane: the join box is shown again, with the manuscript",
    "review file of paper.docx: the link is gone",
    "review file of SI.docx: its own changes are kept",
  ]);
  assert.match(bad.got, /expected \["R1\.1\/R1\.1-1","R1\.2\/R1\.2-1"\], found \[\]/);
  // The leave was refused: the link is still in the supplement's file.
  const still = partsLeft({ ...args, si: siBefore });
  assert.deepEqual(failedNames(still), ["review file of SI.docx: the link is gone"]);
});

// ---- what is told and not judged -------------------------------------------------------

test("i: the line of Document.saved and the file on disk says what was read at each point", () => {
  const line = partsSavedLine([
    {
      at: "opened, before any edit",
      saved: { ok: true, saved: true, raw: true, raw_type: "boolean", error: null },
      file: { mtime: "2026-10-02T08:00:00.000Z", bytes: 14000 },
    },
    {
      at: "after a tracked edit, not saved",
      saved: { ok: true, saved: false, raw: false, raw_type: "boolean", error: null },
      file: { mtime: "2026-10-02T08:00:00.000Z", bytes: 14000 },
    },
    {
      at: "after the save through Word",
      saved: { ok: true, saved: null, raw: null, raw_type: "object", error: "GeneralException" },
      file: null,
    },
    { at: "later", saved: { ok: false, error: "pane not reachable" }, file: null },
  ]);
  assert.equal(
    line,
    "opened, before any edit: Document.saved=true (boolean), sent as true; mtime 2026-10-02T08:00:00.000Z, 14000 bytes | " +
      "after a tracked edit, not saved: Document.saved=false (boolean), sent as false; mtime 2026-10-02T08:00:00.000Z, 14000 bytes | " +
      "after the save through Word: Document.saved=null (object, GeneralException), sent as null; file not read | " +
      "later: not read (pane not reachable); file not read",
  );
});

test("the sessions the panes ran on are read from the daemon log, and two panes on one session fail the row", () => {
  const NAMES = { paper: "paper.docx", si: "SI.docx" };
  const P = "1a2b3c4d-0000-4000-8000-000000000000";
  const S = "9f8e7d6c-0000-4000-8000-000000000000";
  const lines = (...rows) => rows.join("\n");
  // Each pane on a session of its own, the manuscript's second turn in its loop.
  // The daemon names the document by its path (a folder name may hold a
  // space); the row tells the two files apart by their names.
  const good = lines(
    `[agent] turn doc=/mnt/d/my run/paper.docx session=${P} asked=- (started)`,
    "[daemon] Starting session for /mnt/d/my run (new session)",
    `[agent] turn doc=/mnt/d/my run/SI.docx session=${S} asked=- (started)`,
    `[agent] turn doc=/mnt/d/my run/paper.docx session=${P} asked=- (continued)`,
  );
  const s = partsSessions(good, NAMES);
  assert.deepEqual(s.rows[1], { doc: "SI.docx", session: S, asked: null, how: "started" });
  assert.equal(s.rows.length, 3);
  assert.equal(s.pass, true, s.got);
  assert.equal(s.got, "paper.docx: started 1a2b3c4d; SI.docx: started 9f8e7d6c");
  // The pane of the supplement resumed the session of the manuscript.
  const shared = partsSessions(
    lines(
      `[agent] turn doc=paper.docx session=${P} asked=- (started)`,
      `[agent] turn doc=SI.docx session=${P} asked=${P} (resumed)`,
    ),
    NAMES,
  );
  assert.equal(shared.pass, false);
  assert.equal(shared.checks["the two panes share no session"], false);
  assert.match(
    shared.got,
    /1a2b3c4d in both: paper\.docx: started 1a2b3c4d; SI\.docx: resumed 1a2b3c4d/,
  );
  // The same, when the SDK gave the resumed session a new id.
  const forked = partsSessions(
    lines(
      `[agent] turn doc=paper.docx session=${P} asked=- (started)`,
      `[agent] turn doc=SI.docx session=${S} asked=${P} (started)`,
    ),
    NAMES,
  );
  assert.equal(forked.pass, false);
  assert.match(forked.got, /SI\.docx: started 9f8e7d6c \(asked 1a2b3c4d\)/);
  // Each pane resumed its own earlier session: fine.
  const own = partsSessions(
    lines(
      `[agent] turn doc=paper.docx session=${P} asked=${P} (resumed)`,
      `[agent] turn doc=SI.docx session=${S} asked=${S} (resumed)`,
    ),
    NAMES,
  );
  assert.equal(own.pass, true, own.got);
  // A pane without a turn line, and a log without any (an older daemon).
  const one = partsSessions(`[agent] turn doc=paper.docx session=${P} asked=- (started)`, NAMES);
  assert.equal(one.pass, false);
  assert.equal(one.checks["SI.docx: a turn ran on a session"], false);
  const none = partsSessions("", NAMES);
  assert.equal(none.pass, false);
  assert.match(none.got, /no turn line in the log/);
});

test("a turn on a document of the run is the run's own, a turn on any other document is foreign", () => {
  const own = ["D:\\run\\e2e-1", "/mnt/d/run/e2e-1"];
  const turn = (doc) => `[agent] turn doc=${doc} session=9e202b73 asked=- (started)`;
  const mine = [
    "[bridge] hello received; host: word; active doc: D:\\run\\e2e-1\\paper.docx",
    turn("/mnt/d/run/e2e-1/paper.docx"),
    turn("/mnt/d/run/e2e-1/cite/anchor.docx"),
    "[review] refresh changes=4",
    turn("-"),
  ];
  assert.deepEqual(foreignLogLines(mine, own), []);
  // Another Word's pane said hello: that alone is not a tool call.
  const hello = "[bridge] hello received; host: word; active doc: D:\\thesis\\paper.docx";
  assert.deepEqual(foreignLogLines([...mine, hello], own), []);
  // A turn on a document of another folder, of the same file name or not,
  // and one the log names without a folder.
  for (const other of ["/mnt/d/thesis/paper.docx", "/mnt/d/thesis/other.docx", "paper.docx"])
    assert.deepEqual(foreignLogLines([...mine, turn(other)], own), [turn(other)], other);
  // Stage parts: the link lines name the two files without a folder.
  const link = "[review] link join SI.docx -> paper.docx comments=4";
  assert.deepEqual(foreignLogLines([link], own, ["paper.docx", "SI.docx"]), []);
  assert.deepEqual(foreignLogLines([link], own), [link]);
  assert.deepEqual(
    foreignLogLines(["[tool] office_find D:\\other\\paper.docx"], own, ["paper.docx", "SI.docx"]),
    ["[tool] office_find D:\\other\\paper.docx"],
  );
});

test("a log line that names only the files of the run by their bare names is the run's own", () => {
  const own = ["paper.docx", "SI.docx"];
  assert.equal(onlyOwnNames("[review] link join SI.docx -> paper.docx comments=4", own), true);
  assert.equal(
    onlyOwnNames("[review] part turn R1.1 file=SI.docx main=paper.docx outside=q1", own),
    true,
  );
  assert.equal(onlyOwnNames("[review] link leave SI.docx from paper.docx", own), true);
  // A path, or a file of another name, is not.
  assert.equal(onlyOwnNames("[tool] office_find D:\\other\\paper.docx", own), false);
  assert.equal(onlyOwnNames("[review] link join SI.docx -> thesis.docx", own), false);
  assert.equal(onlyOwnNames("[review] refresh changes=1", own), false);
  // Other stages give no names: nothing is the run's own by a bare name.
  assert.equal(onlyOwnNames("[review] link join SI.docx -> paper.docx", []), false);
});

// ---- the driver picks the pane of a document ---------------------------------------------

test("the driver tells the pane of SI.docx from the pane of paper.docx", () => {
  const { isPaneOf } = createRequire(import.meta.url)("../scripts/e2e/pane-driver.cjs");
  const si = { url: "D:\\run\\e2e-1\\SI.docx", file: "" };
  const paper = {
    url: "D:\\run\\e2e-1\\paper.docx",
    file: "Review file: /mnt/d/run/e2e-1/paper.revision.json",
  };
  assert.deepEqual([isPaneOf(si, "SI"), isPaneOf(si, "paper")], [true, false]);
  assert.deepEqual([isPaneOf(paper, "paper"), isPaneOf(paper, "SI")], [true, false]);
  // By the review file line alone (Office gave no url), and by a file URL.
  assert.equal(isPaneOf({ url: "", file: "Review file: /mnt/d/run/SI.revision.json" }, "SI"), true);
  assert.equal(isPaneOf({ url: "file:///D:/run/SI.docx", file: "" }, "SI"), true);
  // A longer name that ends the same is another document; so is a pane that names nothing yet.
  assert.equal(isPaneOf({ url: "D:\\run\\paper_SI.docx", file: "" }, "SI"), false);
  assert.equal(isPaneOf({ url: "D:\\run\\xpaper.docx", file: "" }, "paper"), false);
  assert.equal(isPaneOf({ url: "", file: "" }, "SI"), false);
  assert.equal(isPaneOf(undefined, "SI"), false);
});
