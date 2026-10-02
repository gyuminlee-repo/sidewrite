// A supplementary Word document in the revision of a manuscript, step A2:
// the prompt of a turn in the supplement, where a request stands in the
// other file, the state of the document a refresh writes, the comment a
// running turn of a supplement holds, the author's answers in both files,
// and the figures page. Invented files only.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  RevisionSchema,
  emptyRevision,
  withRevision,
  addChange,
  addQuestion,
  answerQuestion,
  applyStates,
  setRequirementOutside,
  setDocState,
  setPartOf,
} from "../daemon/revision-store.mjs";
import {
  linkedAnswers,
  linkedChanges,
  otherFiles,
  outsideNotes,
} from "../daemon/revision-parts.mjs";
import { createReviewService, docState } from "../daemon/review-service.mjs";
import { applyTrackedEdit } from "../daemon/review-tools.mjs";
import {
  buildReviewPrompt,
  rulesFor,
  REVIEW_RULES,
  POLICY_PACKS,
  OUTSIDE_RULE,
  outsideRuleInPart,
} from "../daemon/review-prompt.mjs";
import { snapshotHash } from "../daemon/response-quotes.mjs";
import {
  changesElsewhere,
  commentState,
  doneBlockers,
  openRequestsText,
  pendingDecisions,
  requestStateText,
  requirementStates,
  unaddressedRequirements,
} from "../taskpane/shared/review-view.js";
import { fakeRunPy, figureWorld } from "./helpers/figure-world.mjs";

const ENV = { CC_OFFICE_WSL_PATHS: "0" };

// ---- the prompt of a turn in a supplement ---------------------------------------

const COMMENT = {
  ref: "R1.2",
  reviewer: 1,
  text: "State the incubation temperature in the Supplementary Methods and in the Methods.",
  requirements: [
    { id: "q1", text: "state the temperature in the Supplementary Methods" },
    { id: "q2", text: "state the temperature in the Methods" },
  ],
  questions: [],
  changes: [],
};
const PART = {
  name: "SI.docx",
  main_name: "paper.docx",
  outside: [
    {
      id: "q1",
      text: "state the temperature in the Supplementary Methods",
      why: "The Supplementary Methods\nare not in this file.",
    },
  ],
};
const prompt = (extra = {}) =>
  buildReviewPrompt({
    comment: COMMENT,
    manuscript: { paragraphs: ["Supplementary Methods", "Cultures were incubated."], ids: null },
    fence: "f",
    ...extra,
  });

test("the prompt of a part: one line names the file, and the manuscript's outside notes are listed", () => {
  const lines = prompt({ part: PART }).text.split("\n");
  const turn = lines.indexOf("[Review turn R1.2]");
  assert.match(
    lines[turn + 1],
    /^You are working in SI\.docx, a supplementary file of paper\.docx\. /,
  );
  assert.equal(lines.filter((l) => l.startsWith("You are working in ")).length, 1);
  const head = lines.findIndex((l) =>
    l.startsWith("Requests the turn in paper.docx noted as belonging to another file"),
  );
  assert.ok(head > turn, "the block is there, after the comment");
  // One line per note, its text on one line.
  assert.equal(
    lines[head + 1],
    "- [q1] state the temperature in the Supplementary Methods (note: The Supplementary Methods are not in this file.)",
  );
  assert.equal(lines[head + 2], "");
  assert.deepEqual(prompt({ part: PART }).outsideIds, ["q1"]);
  // No note of the manuscript: no block; the file line stays.
  const bare = prompt({ part: { ...PART, outside: [] } }).text;
  assert.doesNotMatch(bare, /noted as belonging to another file \(its note/);
  assert.match(bare, /You are working in SI\.docx/);
  // The manuscript block is this file's own text: nothing else is given.
  assert.match(prompt({ part: PART }).text, /\[p1\] Cultures were incubated\./);
});

test("the prompt of a manuscript is what it was: no file line, no block, the same rules", () => {
  const plain = prompt();
  assert.equal(plain.text, prompt({ part: null, linkedAnswers: [] }).text);
  assert.doesNotMatch(plain.text, /You are working in /);
  assert.doesNotMatch(plain.text, /noted as belonging to another file \(its note/);
  assert.deepEqual(plain.outsideIds, []);
  assert.deepEqual(rulesFor({}), REVIEW_RULES);
});

test("the rule about what is not in this file is swapped at its index; the field rule stays last", () => {
  const rules = rulesFor({}, PART);
  assert.equal(rules.length, REVIEW_RULES.length);
  assert.match(REVIEW_RULES[OUTSIDE_RULE], /^Tell apart what is not in this file/);
  rules.forEach((r, i) => {
    if (i === OUTSIDE_RULE) return;
    assert.equal(r, REVIEW_RULES[i], `rule ${i} is the one of a manuscript`);
  });
  const swapped = rules[OUTSIDE_RULE];
  assert.equal(swapped, outsideRuleInPart(PART));
  assert.notEqual(swapped, REVIEW_RULES[OUTSIDE_RULE]);
  // In a supplement the manuscript is one of the other files, and the
  // supplementary material is not: this file is part of it.
  assert.match(swapped, /^Tell apart what is not in this file/);
  assert.match(
    swapped,
    /the main manuscript \(paper\.docx\), a figure file, another supplementary/,
  );
  assert.doesNotMatch(swapped, /a figure file, the supplementary material,/);
  assert.match(swapped, /office_requirement_outside/);
  assert.match(swapped, /office_ask_author/);
  assert.match(rules.at(-1), /^Citations and other fields/);
  // In the prompt: the same number, the swapped text.
  const lines = prompt({ part: PART }).text.split("\n");
  assert.equal(
    lines[lines.indexOf("Rules:") + 1 + OUTSIDE_RULE],
    `${OUTSIDE_RULE + 1}. ${swapped}`,
  );
  assert.match(
    lines[lines.indexOf("Rules:") + REVIEW_RULES.length],
    /^\d+\. Citations and other fields/,
  );
  // A policy pack that is off takes its rules out by index, in a part too.
  const off = rulesFor({ policy_packs: { review_scope: false } }, PART);
  assert.equal(off.length, REVIEW_RULES.length - POLICY_PACKS.review_scope.length);
  assert.ok(off.includes(swapped));
  assert.match(off.at(-1), /^Citations and other fields/);
});

test("author answers of the other files go into the prompt, named with their file", () => {
  const text = prompt({
    linkedAnswers: [
      {
        file: "paper.docx",
        ref: "R1.1",
        id: "a1",
        question: "Which temperature?",
        answer: "30 degrees\nCelsius",
      },
    ],
  }).text;
  assert.match(text, /^Author answers \(what the author said when asked/m);
  assert.match(
    text,
    /^- \[R1\.1 a1, asked in paper\.docx\] Which temperature\? Answer: 30 degrees Celsius$/m,
  );
});

// ---- where a request stands in the other file ---------------------------------

const elsewhere = (state, ids = ["q1"], file = "SI.docx") => [
  { file, ref: "R1.2", id: "R1.2-1", state, requirement_ids: ids },
];
const VIEW = { ...COMMENT, requirements_outside: [], blocked: [], figure_changes: [] };

test("requirementStates: an accepted change in another file answers a request, a pending one waits there, a rejected one counts for nothing", () => {
  const states = (c, e) => requirementStates(c, e).map((r) => [r.id, r.state, r.file ?? null]);
  assert.deepEqual(states(VIEW, []), [
    ["q1", "open", null],
    ["q2", "open", null],
  ]);
  assert.deepEqual(states(VIEW, elsewhere("accepted")), [
    ["q1", "changed_elsewhere", "SI.docx"],
    ["q2", "open", null],
  ]);
  assert.deepEqual(states(VIEW, elsewhere("user_edited"))[0], [
    "q1",
    "changed_elsewhere",
    "SI.docx",
  ]);
  assert.deepEqual(states(VIEW, elsewhere("pending"))[0], ["q1", "waiting_elsewhere", "SI.docx"]);
  assert.deepEqual(states(VIEW, elsewhere("unlocated"))[0], ["q1", "waiting_elsewhere", "SI.docx"]);
  assert.deepEqual(states(VIEW, elsewhere("rejected"))[0], ["q1", "open", null]);
  // The words of the row.
  const row = (e) => requestStateText(requirementStates(VIEW, e)[0]);
  assert.equal(row(elsewhere("accepted")), "Changed in SI.docx");
  assert.equal(row(elsewhere("pending")), "Waiting for your decision in SI.docx");
  assert.equal(row(elsewhere("rejected")), "Not yet");
  // The note of this file's own turn stays; the change elsewhere ranks above it.
  const noted = { ...VIEW, requirements_outside: [{ id: "q1", why: "in the supplement" }] };
  assert.equal(states(noted, [])[0][1], "outside_this_file");
  assert.equal(states(noted, elsewhere("pending"))[0][1], "waiting_elsewhere");
  assert.equal(states(noted, elsewhere("accepted"))[0][1], "changed_elsewhere");
  assert.equal(states(noted, elsewhere("rejected"))[0][1], "outside_this_file");
  // A change of this file ranks first; an accepted one elsewhere beats a pending one.
  const own = {
    ...VIEW,
    changes: [{ id: "R1.2-1", state: "pending", rationale: { requirement_ids: ["q1"] } }],
  };
  assert.equal(states(own, elsewhere("accepted"))[0][1], "changed");
  assert.equal(
    states(VIEW, [...elsewhere("pending", ["q1"], "S2.docx"), ...elsewhere("accepted")])[0][2],
    "SI.docx",
  );
});

test("the card, the blockers of Done and the waiting list follow the other file", () => {
  const state = (e, extra = {}) => {
    const s = commentState(VIEW, { elsewhere: e, ...extra });
    return [s.kind, s.label, s.attention];
  };
  assert.deepEqual(state([]), ["not_started", "Not started", false]);
  assert.deepEqual(state(elsewhere("pending")), [
    "to_review",
    "1 change to review in SI.docx",
    true,
  ]);
  assert.deepEqual(state(elsewhere("accepted")), ["decided", "Changed in SI.docx", false]);
  assert.deepEqual(state(elsewhere("rejected")), ["not_started", "Not started", false]);
  // A turn here that proposed nothing, while the other file has the change.
  assert.deepEqual(state(elsewhere("accepted"), { noChange: true }), [
    "decided",
    "Changed in SI.docx",
    false,
  ]);
  assert.deepEqual(state([], { noChange: true }), ["no_change", "No change proposed", true]);
  // Requests: answered elsewhere is not open.
  const both = elsewhere("accepted", ["q1", "q2"]);
  assert.deepEqual(
    unaddressedRequirements(VIEW, elsewhere("accepted")).map((r) => r.id),
    ["q2"],
  );
  assert.equal(openRequestsText(VIEW, both), "");
  assert.deepEqual(doneBlockers(VIEW, [], both), []);
  assert.deepEqual(doneBlockers(VIEW, [], elsewhere("pending", ["q1", "q2"])), [
    "1 change is not decided in SI.docx",
  ]);
  // The list of what waits for the author: by the link of the last load.
  const revision = { comments: [VIEW], ledger: [] };
  const open = (linked) =>
    pendingDecisions(revision, linked)
      .requirements.filter((x) => x.kind === "request_open")
      .map((x) => x.id);
  assert.deepEqual(open(null), ["q1", "q2"]);
  assert.deepEqual(open({ changes: elsewhere("accepted") }), ["q2"]);
  assert.deepEqual(open({ changes: elsewhere("pending") }), ["q2"]);
  assert.deepEqual(open({ changes: elsewhere("rejected") }), ["q1", "q2"]);
  assert.deepEqual(changesElsewhere({ changes: elsewhere("accepted") }, "R9.9"), []);
  assert.deepEqual(changesElsewhere(null, "R1.2"), []);
});

// ---- the state of the document at a refresh -------------------------------------

const SNAPSHOT = {
  track_changes_on: true,
  paragraphs: [
    { index: 0, text: "Supplementary Methods", w_ins: 0, w_del: 0 },
    { index: 1, text: "Cultures were incubated at 30 C.", w_ins: 1, w_del: 0 },
    { index: 2, text: "Samples were taken hourly.", w_ins: 0, w_del: 2 },
    { index: 3, text: "Both.", w_ins: 1, w_del: 1 },
  ],
  comments: [
    { id: "1", content: "[R1.2] what: unit added" },
    { id: "2", content: "  [E.1] what: style" },
    { id: "3", content: "check this with the lab" },
    { id: "4", content: "R1.2 is mentioned, but this is the author's own" },
  ],
};

test("docState: tracked paragraphs, comments, tool comments, and saved as true, false or not known", () => {
  const now = new Date("2026-10-02T08:00:00.000Z");
  const s = docState(SNAPSHOT, {
    saved: true,
    file: { mtime: "2026-10-02T07:59:00.000Z", bytes: 12345 },
    now,
  });
  assert.deepEqual(s, {
    at: "2026-10-02T08:00:00.000Z",
    text_hash: snapshotHash(SNAPSHOT),
    tracked_paragraphs: 3,
    comments: 4,
    tool_comments: 2,
    track_changes_on: true,
    saved: true,
    file: { mtime: "2026-10-02T07:59:00.000Z", bytes: 12345 },
  });
  assert.equal(docState(SNAPSHOT, { saved: false }).saved, false);
  // Word did not say, or said something that is no boolean: not known.
  for (const saved of [null, undefined, "true", 1, 0]) {
    assert.equal(docState(SNAPSHOT, { saved }).saved, null, String(saved));
  }
  const bare = docState({ paragraphs: [{ text: "x" }] });
  assert.deepEqual(
    [bare.tracked_paragraphs, bare.comments, bare.tool_comments, bare.track_changes_on, bare.file],
    [0, 0, 0, null, null],
  );
  // It is a value the review file takes, and an earlier schema drops it.
  const rev = setDocState(emptyRevision("/work/SI.docx"), s);
  assert.deepEqual(RevisionSchema.parse(rev).doc_state, s);
  assert.equal("doc_state" in RevisionSchema.omit({ doc_state: true }).parse(rev), false);
  assert.equal("doc_state" in RevisionSchema.parse(emptyRevision("/work/SI.docx")), false);
  assert.throws(() => setDocState(rev, { ...s, saved: "yes" }));
});

// ---- the services ---------------------------------------------------------------

const COMMENTS = [
  {
    ref: "R1.1",
    reviewer: 1,
    text: "State the growth medium.",
    requirements: ["state the medium"],
  },
  {
    ref: "R1.2",
    reviewer: 1,
    text: "State the incubation temperature in the Supplementary Methods.",
    requirements: ["state the temperature in the Supplementary Methods"],
  },
  { ref: "R2.1", reviewer: 2, text: "Do not call the method novel." },
];
const CHANGE = {
  word_comment_id: "9",
  paragraph_index: 1,
  kind: "replace",
  find: "incubated",
  text: "incubated at 30 degrees Celsius",
  before: "Cultures were incubated.",
  after: "Cultures were incubated at 30 degrees Celsius.",
  rationale: { what: "unit", why: "asked", addresses: "R1.2", requirement_ids: ["q1"] },
};
// A rationale that names no request.
const PLAIN = { ...CHANGE.rationale, requirement_ids: [] };
const snapOf = (...texts) => ({
  track_changes_on: true,
  paragraphs: texts.map((text, index) => ({ index, id: null, text, w_ins: 0, w_del: 0 })),
  comments: [],
});

// One folder, one daemon, a pane for each document (the pane key is the file
// name). paper.docx holds a review of three comments; SI.docx has joined it.
async function world() {
  const dir = await mkdtemp(join(tmpdir(), "parts-turns-"));
  const docOf = (name) => join(dir, name);
  const events = [];
  const pushed = [];
  const open = new Set(["paper.docx", "SI.docx"]);
  const service = createReviewService({
    getContext: (key) => ({ activeDoc: docOf(key), caps: ["target_hash"] }),
    sendEvent: (key, obj) => events.push({ key, ...obj }),
    callTool: async () => ({}),
    pushTurn: (key, host, text, turn) => pushed.push({ key, text, turn }),
    listPanes: () => [...open].map((key) => ({ key, host: "word", activeDoc: docOf(key) })),
    log: () => {},
    env: ENV,
  });
  const ask = (key, type, payload = {}) =>
    new Promise((resolve) =>
      service.handlers[type]({ type, request_id: "q", ...payload }, resolve, key, "word"),
    );
  const pathOf = (name) => join(dir, name.replace(/\.docx$/, ".revision.json"));
  const file = async (name) => JSON.parse(await readFile(pathOf(name), "utf8"));
  const edit = (name, fn) => withRevision(pathOf(name), docOf(name), fn);
  const told = () => events.filter((e) => e.event === "linked_changed").map((e) => [e.key, e.doc]);
  const w = { dir, docOf, pathOf, ask, file, edit, events, pushed, open, service, told };
  w.cleanup = () => rm(dir, { recursive: true, force: true });
  const saved = await ask("paper.docx", "review_save_comments", { comments: COMMENTS });
  assert.equal(saved.ok, true, saved.error);
  const joined = await ask("SI.docx", "review_link_join", { main: docOf("paper.docx") });
  assert.equal(joined.ok, true, joined.error);
  events.length = 0;
  return w;
}

test("a turn in the supplement: the copy is brought up to date first, and the prompt knows the file and nothing of the manuscript's text", async () => {
  const w = await world();
  try {
    await w.ask("paper.docx", "review_refresh", {
      snapshot: snapOf("Title of the paper", "ONLY IN THE MANUSCRIPT."),
    });
    await w.ask("SI.docx", "review_refresh", {
      snapshot: snapOf("Supplementary Methods", "Cultures were incubated."),
    });
    // The turn of the manuscript noted the request as belonging elsewhere.
    await w.edit("paper.docx", (rev) =>
      setRequirementOutside(rev, "R1.2", { id: "q1", why: "It is in the Supplementary Methods." }),
    );
    // The manuscript gets a comment the supplement has not seen: no load of
    // the supplement in between, the preview has it all the same.
    const more = [...COMMENTS, { ref: "R2.2", reviewer: 2, text: "Add a scale bar." }];
    assert.equal((await w.ask("paper.docx", "review_save_comments", { comments: more })).ok, true);
    assert.equal((await w.file("SI.docx")).comments.length, 3);
    const late = await w.ask("SI.docx", "review_preview", { ref: "R2.2", paragraphs: [] });
    assert.equal(late.ok, true, late.error);
    assert.equal((await w.file("SI.docx")).comments.length, 4, "the copy was brought up to date");

    const r = await w.ask("SI.docx", "review_preview", { ref: "R1.2", paragraphs: [] });
    assert.equal(r.ok, true, r.error);
    assert.match(r.prompt, /You are working in SI\.docx, a supplementary file of paper\.docx\./);
    assert.match(
      r.prompt,
      /- \[q1\] state the temperature in the Supplementary Methods \(note: It is in the Supplementary Methods\.\)/,
    );
    assert.match(r.prompt, /\[p1\] Cultures were incubated\./);
    assert.doesNotMatch(r.prompt, /ONLY IN THE MANUSCRIPT/);
    assert.match(r.prompt, /the main manuscript \(paper\.docx\), a figure file/);
    // The manuscript's own prompt is the one of a manuscript.
    const m = await w.ask("paper.docx", "review_preview", { ref: "R1.2", paragraphs: [] });
    assert.doesNotMatch(m.prompt, /You are working in /);
    assert.match(m.prompt, /ONLY IN THE MANUSCRIPT/);
    assert.doesNotMatch(m.prompt, /Cultures were incubated/);

    // Propose: the turn is queued for the pane of the supplement with that prompt.
    const p = await w.ask("SI.docx", "review_propose", { ref: "R1.2", paragraphs: [] });
    assert.equal(p.ok, true, p.error);
    assert.deepEqual(
      w.pushed.map((x) => [x.key, x.turn.ref]),
      [["SI.docx", "R1.2"]],
    );
    assert.match(w.pushed[0].text, /You are working in SI\.docx/);
    assert.equal(
      (await w.file("SI.docx")).comments.find((c) => c.ref === "R1.2").status,
      "in_progress",
    );
    assert.equal(
      (await w.file("paper.docx")).comments.find((c) => c.ref === "R1.2").status,
      "pending",
      "the status of a comment is each file's own",
    );
  } finally {
    await w.cleanup();
  }
});

test("where a request stands in the other file is told to both panes, from the stored state", async () => {
  const w = await world();
  try {
    await w.edit("SI.docx", (rev) => addChange(rev, "R1.2", CHANGE).revision);
    const states = async (key) =>
      (await w.ask(key, "review_load")).linked.changes.map((x) => [
        x.file,
        x.ref,
        x.id,
        x.state,
        x.requirement_ids,
      ]);
    assert.deepEqual(await states("paper.docx"), [
      ["SI.docx", "R1.2", "R1.2-1", "pending", ["q1"]],
    ]);
    assert.deepEqual(await states("SI.docx"), [], "the manuscript has no change");
    for (const state of ["accepted", "rejected"]) {
      await w.edit("SI.docx", (rev) =>
        applyStates(rev, [{ ref: "R1.2", change_id: "R1.2-1", state, paragraph_index: 1 }]),
      );
      assert.equal((await states("paper.docx"))[0][3], state);
    }
    // The other way, and with a second supplement: each sees the others.
    await w.edit(
      "paper.docx",
      (rev) => addChange(rev, "R1.1", { ...CHANGE, rationale: PLAIN }).revision,
    );
    assert.deepEqual(await states("SI.docx"), [["paper.docx", "R1.1", "R1.1-1", "pending", []]]);
    assert.equal(
      (await w.ask("S2.docx", "review_link_join", { main: w.docOf("paper.docx") })).ok,
      true,
    );
    assert.deepEqual(
      (await states("S2.docx")).map((x) => x[0]),
      ["paper.docx", "SI.docx"],
    );
    assert.deepEqual(
      (await otherFiles(w.docOf("paper.docx"), await w.file("paper.docx"))).map((f) => f.name),
      ["SI.docx", "S2.docx"],
    );
    // A file that is in no revision has no other files.
    assert.deepEqual(await otherFiles(w.docOf("x.docx"), emptyRevision("x")), []);
    assert.deepEqual((await w.ask("other.docx", "review_load")).linked.changes, []);
    // The pure readers.
    assert.deepEqual(linkedChanges([{ name: "gone.docx", revision: null }]), []);
    assert.deepEqual(outsideNotes(null, "R1.2"), []);
  } finally {
    await w.cleanup();
  }
});

test("a change recorded or decided in one file is told to the other open pane; a refresh that moves nothing is not", async () => {
  const w = await world();
  try {
    const texts = ["Supplementary Methods", "Cultures were incubated."];
    await w.ask("SI.docx", "review_refresh", { snapshot: snapOf(...texts) });
    assert.deepEqual(w.told(), [], "nothing moved");
    // Recorded (the path of a tracked edit that Word applied).
    const deps = w.service.trackedEditDeps("SI.docx");
    await deps.onEdit({ ref: "R1.2", change: CHANGE });
    assert.deepEqual(w.told(), [["paper.docx", "SI.docx"]]);
    w.events.length = 0;
    // Decided in Word: the refresh finds the new text accepted.
    const accepted = snapOf(texts[0], CHANGE.after);
    const r = await w.ask("SI.docx", "review_refresh", { snapshot: accepted });
    assert.equal(r.ok, true, r.error);
    const state = (await w.file("SI.docx")).comments.find((c) => c.ref === "R1.2").changes[0].state;
    assert.notEqual(state, "pending");
    assert.deepEqual(w.told(), [["paper.docx", "SI.docx"]]);
    w.events.length = 0;
    // The same document again: nothing to tell.
    await w.ask("SI.docx", "review_refresh", { snapshot: accepted });
    assert.deepEqual(w.told(), []);
    // In the manuscript: its supplement's pane is told.
    await w.service
      .trackedEditDeps("paper.docx")
      .onEdit({ ref: "R1.1", change: { ...CHANGE, rationale: PLAIN } });
    assert.deepEqual(w.told(), [["SI.docx", "paper.docx"]]);
    // The pane that is not open gets nothing, and nothing fails.
    w.events.length = 0;
    w.open.delete("paper.docx");
    await deps.onEdit({ ref: "R1.1", change: CHANGE });
    assert.deepEqual(w.told(), []);
  } finally {
    await w.cleanup();
  }
});

test("a refresh writes the state of the document into its own review file", async () => {
  const w = await world();
  try {
    await writeFile(w.docOf("SI.docx"), "not a real docx, 31 bytes long.");
    const onDisk = await stat(w.docOf("SI.docx"));
    const r = await w.ask("SI.docx", "review_refresh", { snapshot: SNAPSHOT, saved: false });
    assert.equal(r.ok, true, r.error);
    const s = (await w.file("SI.docx")).doc_state;
    assert.deepEqual(
      { ...s, at: null },
      {
        at: null,
        text_hash: snapshotHash(SNAPSHOT),
        tracked_paragraphs: 3,
        comments: 4,
        tool_comments: 2,
        track_changes_on: true,
        saved: false,
        file: { mtime: onDisk.mtime.toISOString(), bytes: 31 },
      },
    );
    assert.equal(s.at, (await w.file("SI.docx")).snapshot.taken_at);
    assert.equal(r.revision.doc_state.saved, false);
    // Saved, and a pane that sends nothing about it.
    await w.ask("SI.docx", "review_refresh", { snapshot: SNAPSHOT, saved: true });
    assert.equal((await w.file("SI.docx")).doc_state.saved, true);
    await w.ask("SI.docx", "review_refresh", { snapshot: SNAPSHOT });
    assert.equal((await w.file("SI.docx")).doc_state.saved, null);
    // A document that is not on disk where the daemon looks: no file facts.
    await w.ask("paper.docx", "review_refresh", { snapshot: snapOf("Title"), saved: true });
    const main = (await w.file("paper.docx")).doc_state;
    assert.deepEqual([main.file, main.saved, main.tracked_paragraphs], [null, true, 0]);
    // The state of one file is never written into the other.
    assert.equal(main.text_hash === s.text_hash, false);
  } finally {
    await w.cleanup();
  }
});

test("the manuscript keeps a comment that a running turn of a supplement holds", async () => {
  const w = await world();
  try {
    await w.ask("SI.docx", "review_refresh", { snapshot: snapOf("Supplementary Methods") });
    const without = COMMENTS.filter((c) => c.ref !== "R1.2");
    const p = await w.ask("SI.docx", "review_propose", { ref: "R1.2", paragraphs: [] });
    assert.equal(p.ok, true, p.error);
    const before = await readFile(w.pathOf("paper.docx"), "utf8");
    const refused = await w.ask("paper.docx", "review_save_comments", { comments: without });
    assert.equal(refused.ok, false);
    assert.equal(refused.error, "Comment R1.2 is being worked on in SI.docx; keep it");
    assert.equal(await readFile(w.pathOf("paper.docx"), "utf8"), before, "nothing was written");
    // A comment no turn holds can still go while that turn runs.
    const other = await w.ask("paper.docx", "review_save_comments", {
      comments: COMMENTS.filter((c) => c.ref !== "R2.1"),
    });
    assert.equal(other.ok, true, other.error);
    // The turn is over: nothing holds the comment any more.
    w.service.turnEnded("SI.docx", { ref: "R1.2" });
    const after = await w.ask("paper.docx", "review_save_comments", {
      comments: COMMENTS.filter((c) => c.ref === "R1.1"),
    });
    assert.equal(after.ok, true, after.error);
    assert.deepEqual(
      (await w.file("paper.docx")).comments.map((c) => c.ref),
      ["R1.1"],
    );
  } finally {
    await w.cleanup();
  }
});

test("author answers are shared read-only: a turn in either file gets what was answered in the other", async () => {
  const w = await world();
  try {
    const answered = (rev, ref, question, answer) => {
      const q = addQuestion(rev, ref, { question, why: "the file does not say" });
      return answerQuestion(q.revision, ref, q.question.id, answer);
    };
    await w.edit("SI.docx", (rev) => answered(rev, "R1.2", "Which temperature?", "30 degrees"));
    await w.edit("paper.docx", (rev) => answered(rev, "R1.1", "Which medium?", "LB medium"));
    const prompts = async (key, ref) =>
      (await w.ask(key, "review_preview", { ref, paragraphs: [] })).prompt;
    // In the manuscript: its own answer as before, the supplement's with its file.
    const main = await prompts("paper.docx", "R1.1");
    assert.match(main, /^- \[R1\.1 \S+\] Which medium\? Answer: LB medium$/m);
    assert.match(
      main,
      /^- \[R1\.2 \S+, asked in SI\.docx\] Which temperature\? Answer: 30 degrees$/m,
    );
    // In the supplement: the other way round.
    const supp = await prompts("SI.docx", "R1.2");
    assert.match(supp, /^- \[R1\.2 \S+\] Which temperature\? Answer: 30 degrees$/m);
    assert.match(supp, /^- \[R1\.1 \S+, asked in paper\.docx\] Which medium\? Answer: LB medium$/m);
    // A number in an answer is a source for an edit in the other file too.
    const texts = async (key, ref) =>
      (await w.service.trackedEditDeps(key).getReviewContext(ref)).author_texts;
    assert.deepEqual(await texts("paper.docx", "R1.1"), ["LB medium", "30 degrees"]);
    assert.deepEqual(await texts("SI.docx", "R1.2"), ["30 degrees", "LB medium"]);
    // Read-only: the question and its answer stay in the file they were asked in.
    assert.equal((await w.file("paper.docx")).comments[1].questions.length, 0);
    assert.equal((await w.file("SI.docx")).comments[0].questions.length, 0);
    assert.deepEqual(
      linkedAnswers([{ name: "SI.docx", revision: await w.file("SI.docx") }]).map((q) => [
        q.file,
        q.ref,
        q.answer,
      ]),
      [["SI.docx", "R1.2", "30 degrees"]],
    );
    // A file in no revision asks nobody.
    assert.deepEqual(await texts("other.docx", "R1.1"), []);
  } finally {
    await w.cleanup();
  }
});

test("the forbidden list is enforced before the edit in a supplement, by the pre-check of the manuscript", async () => {
  const w = await world();
  try {
    // The rule is the manuscript's; the supplement copies it on its next load.
    const added = await w.ask("paper.docx", "ledger_add", {
      item: {
        kind: "forbid",
        text: "do not write novel",
        check: { type: "regex", pattern: String.raw`\bnovel\b`, flags: "i" },
      },
    });
    assert.equal(added.ok, true, added.error);
    assert.equal((await w.ask("SI.docx", "review_load")).revision.ledger[0].id, "L1");
    const sentence = "The invented carbon source was added to each flask.";
    for (const key of ["paper.docx", "SI.docx"]) {
      await w.ask(key, "review_refresh", { snapshot: snapOf("Methods", sentence) });
    }
    const calls = [];
    const edit = (key, replace) =>
      applyTrackedEdit(
        {
          ...w.service.trackedEditDeps(key),
          call: async (name) => {
            calls.push([key, name]);
            return {
              ok: true,
              word_comment_id: 5,
              paragraph_index: 1,
              before: sentence,
              after: sentence.replace("invented carbon source", replace),
            };
          },
        },
        {
          comment_ref: "R1.2",
          find: "invented carbon source",
          replace,
          rationale: { what: "w", why: "y", addresses: "a", requirement_ids: ["q1"] },
        },
        { ref: "R1.2" },
      );
    const blockedOf = async (name) =>
      (await w.file(name)).comments
        .find((c) => c.ref === "R1.2")
        .blocked.map((x) => [x.state, x.blocked_by, x.args.replace]);
    // In the supplement: refused by L1, Word is never called, and the
    // blocked edit is on record in the supplement's own review file.
    const inSupp = await edit("SI.docx", "novel carbon source");
    assert.equal(inSupp.ok, false);
    assert.deepEqual(inSupp.blocked_by, ["L1"]);
    assert.deepEqual(calls, [], "Word was not called");
    assert.deepEqual(await blockedOf("SI.docx"), [["blocked", ["L1"], "novel carbon source"]]);
    assert.deepEqual(await blockedOf("paper.docx"), []);
    assert.equal((await w.file("SI.docx")).comments[1].changes.length, 0);
    // The same edit in the manuscript gets the same answer from the same check.
    const inMain = await edit("paper.docx", "novel carbon source");
    assert.deepEqual([inMain.ok, inMain.code, inMain.blocked_by], [false, inSupp.code, ["L1"]]);
    assert.deepEqual(calls, []);
    // An edit without the word goes to Word, from the pane of the supplement
    // only, and is recorded in its review file.
    const fine = await edit("SI.docx", "defined carbon source");
    assert.equal(fine.ok, true, JSON.stringify(fine));
    assert.deepEqual(calls, [["SI.docx", "office_tracked_edit"]]);
    assert.equal((await w.file("SI.docx")).comments[1].changes.length, 1);
    assert.equal((await w.file("paper.docx")).comments[1].changes.length, 0);
  } finally {
    await w.cleanup();
  }
});

test("a supplement that left with changes of its own is offered its manuscript again, and can join it", async () => {
  const w = await world();
  try {
    await w.edit("SI.docx", (rev) => addChange(rev, "R1.2", CHANGE).revision);
    const covers = async (key) =>
      (await w.ask(key, "review_link_candidates")).candidates.map((c) => [c.name, c.covers]);
    // A second paper in the folder, with a review of its own.
    await w.ask("other.docx", "review_save_comments", {
      comments: [{ ref: "R1.2", reviewer: 1, text: "Another paper, another comment." }],
    });
    assert.equal((await w.ask("SI.docx", "review_link_leave")).ok, true);
    const left = await w.file("SI.docx");
    assert.deepEqual(
      left.comments.map((c) => [c.ref, c.changes.length]),
      [["R1.2", 1]],
      "the comment with its change stays",
    );
    assert.deepEqual(await covers("SI.docx"), [
      ["other.docx", false],
      ["paper.docx", true],
    ]);
    // The other paper holds comments no manuscript of the folder lists, and
    // a file with no comment is covered by nothing.
    // (The file that left is a review file with comments like any other.)
    assert.deepEqual(await covers("other.docx"), [
      ["SI.docx", false],
      ["paper.docx", false],
    ]);
    assert.deepEqual(
      (await covers("empty.docx")).map((c) => c[1]),
      [false, false, false],
    );
    const again = await w.ask("SI.docx", "review_link_join", { main: w.docOf("paper.docx") });
    assert.equal(again.ok, true, again.error);
    assert.equal(again.revision.comments.length, 3);
    assert.equal(again.revision.comments.find((c) => c.ref === "R1.2").changes.length, 1);
  } finally {
    await w.cleanup();
  }
});

// ---- the figures page -------------------------------------------------------------

test("the figures page takes no new record under a supplementary file: it says where the record goes", async () => {
  const w = await figureWorld({ runPy: fakeRunPy() });
  try {
    // The document of the pane as a manuscript: the record is taken.
    const ok = await w.propose({});
    assert.equal(ok.status, 200, JSON.stringify(ok.json));
    // The same document as a supplementary file of main.docx.
    const path = join(w.dir, "paper.revision.json");
    await withRevision(path, w.doc, (rev) =>
      setPartOf(rev, { manuscript: "main.docx", label: "SI", joined_at: "t" }),
    );
    const refused = await w.propose({ figure: "Figure 3" });
    assert.equal(refused.status, 409);
    assert.deepEqual(refused.json, { ok: false, error: "Record this under main.docx." });
    assert.equal((await w.rev()).comments[0].figure_changes.length, 1, "nothing was recorded");
    // What it recorded before it joined can still be decided.
    const rejected = await w.call("reject", {
      body: { doc: w.doc, ref: "R1.1", id: ok.json.id },
    });
    assert.equal(rejected.status, 200, JSON.stringify(rejected.json));
    assert.equal((await w.call("state", { query: { doc: w.doc } })).status, 200);
  } finally {
    await w.cleanup();
  }
});
