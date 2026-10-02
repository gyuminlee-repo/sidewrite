// The stretch between "Word applied the edit" and "the review file says so"
// (H1). A tracked edit is written to the file as an attempt BEFORE Word is
// called, with the form it would have in the document. Whatever stops the
// daemon or the save afterwards, the attempt is on file and a refresh
// settles it. The record of the change and the removal of its attempt are
// one save.
//
// Saves are watched and failed at the rename of the atomic write
// (revision-store.mjs writeJsonAtomic): node:module syncBuiltinESMExports
// makes the patched fs.promises.rename the one the daemon imports.
import { plainPane } from "./helpers/plain-pane.mjs";
import { test } from "node:test";
import assert from "node:assert/strict";
import fsp from "node:fs/promises";
import { syncBuiltinESMExports } from "node:module";
import { tmpdir } from "node:os";
import { join } from "node:path";

const realRename = fsp.rename;
// Every save of a revision file, in order: [changes, blocked] of each comment.
const saves = [];
// (save number, 1-based) => true to fail that save like a full disk.
let failSave = () => false;
fsp.rename = async (from, to) => {
  if (!String(to).endsWith(".revision.json")) return realRename(from, to);
  const rev = JSON.parse(await fsp.readFile(from, "utf8"));
  if (failSave(saves.length + 1)) {
    saves.push({ failed: true });
    throw Object.assign(new Error("ENOSPC: no space left on device, rename"), { code: "ENOSPC" });
  }
  saves.push({
    comments: rev.comments.map((c) => [
      c.changes.length,
      c.blocked.length,
      c.blocked.map((b) => b.state).join(","),
    ]),
  });
  return realRename(from, to);
};
syncBuiltinESMExports();

const { createReviewService } = await import("../daemon/review-service.mjs");
const { applyTrackedEdit } = await import("../daemon/review-tools.mjs");

const ENV = { CC_OFFICE_WSL_PATHS: "0" };
const R = { what: "w", why: "y", addresses: "a" };
const BEFORE = "Cells grew on the substrate at 30 C.";
const AFTER = "Cells grew on the medium at 30 C.";
const COMMENT = "[R1.1] w / y / a";
const EDIT = { comment_ref: "R1.1", find: "the substrate", replace: "the medium", rationale: R };
const APPLIED = { ok: true, word_comment_id: 12, paragraph_index: 1, before: BEFORE, after: AFTER };
const snap = (text, comments = []) => ({
  paragraphs: [
    { index: 0, id: "p0", text: "Title", w_ins: 0, w_del: 0 },
    { index: 1, id: "p1", text, w_ins: 0, w_del: 0 },
  ],
  comments,
});
const noAnswer = (id, code = "ToolTimeout") =>
  Object.assign(new Error("Tool office_tracked_edit timed out"), { code, toolCallId: id });

// One daemon on a folder. call(name, args, n): the answer of the task pane.
function service(dir, call = async () => APPLIED) {
  const doc = join(dir, "paper.docx");
  const calls = [];
  const events = [];
  const lines = [];
  const callTool = async (name, args) => {
    calls.push([name, args]);
    return call(name, args, calls.length);
  };
  const svc = createReviewService({
    getContext: () => ({ activeDoc: doc, caps: ["target_hash"] }),
    callTool,
    pushTurn: () => {},
    sendEvent: (key, obj) => events.push(obj),
    log: (s) => lines.push(s),
    env: ENV,
  });
  const ask = (type, payload = {}) =>
    new Promise((resolve) =>
      svc.handlers[type]({ type, request_id: "q", ...plainPane(payload) }, resolve, "k", "word"),
    );
  const edit = (args = EDIT) =>
    applyTrackedEdit({ ...svc.trackedEditDeps("k"), call: (name, a) => callTool(name, a) }, args, {
      ref: args.comment_ref,
    });
  const file = async () =>
    JSON.parse(await fsp.readFile(join(dir, "paper.revision.json"), "utf8")).comments[0];
  return { svc, ask, edit, file, calls, events, lines, doc };
}

async function world(call) {
  failSave = () => false;
  const dir = await fsp.mkdtemp(join(tmpdir(), "edit-attempt-"));
  const t = service(dir, call);
  await t.ask("review_save_comments", {
    comments: [{ ref: "R1.1", reviewer: 1, text: "Name the medium." }],
  });
  await t.ask("review_refresh", { snapshot: snap(BEFORE) });
  saves.length = 0;
  return { ...t, dir, cleanup: () => fsp.rm(dir, { recursive: true, force: true }) };
}

// ---- applyTrackedEdit with fake deps ------------------------------------------

function fakeDeps(over = {}) {
  const d = {
    order: [],
    getLedger: async () => [],
    onAttempt: async (x) => {
      d.order.push(["attempt", x]);
      return { id: "R1.1-b1", end: () => d.order.push(["end"]) };
    },
    call: async (name, args) => {
      d.order.push(["call", args]);
      return APPLIED;
    },
    onEdit: async (x) => {
      d.order.push(["edit", x]);
      return { id: "R1.1-1" };
    },
    onNotApplied: async (x) => d.order.push(["not_applied", x]),
    onOutcomeUnknown: async (x) => d.order.push(["unknown", x]),
    onRecordFailed: async (x) => d.order.push(["record_failed", x]),
    ...over,
  };
  return d;
}
const steps = (d) => d.order.map((o) => o[0]);

test("the attempt is handed over before Word is called, and the change that settles it carries its id", async () => {
  const d = fakeDeps();
  const r = await applyTrackedEdit(d, EDIT, { ref: "R1.1" });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(steps(d), ["attempt", "call", "edit", "end"]);
  const pre = d.order[0][1];
  assert.deepEqual(
    [pre.ref, pre.args.find, pre.args.replace],
    ["R1.1", "the substrate", "the medium"],
  );
  assert.deepEqual(
    [pre.attempt.tool_call_id, pre.attempt.comment_text, pre.attempt.change.kind],
    [null, COMMENT, "replace"],
  );
  assert.deepEqual(
    [pre.attempt.change.find, pre.attempt.change.text],
    ["the substrate", "the medium"],
  );
  assert.equal(d.order[2][1].attempt_id, "R1.1-b1");
});

test("an attempt that cannot be written: Word is not called", async () => {
  const d = fakeDeps({
    onAttempt: async () => {
      throw new Error("ENOSPC: no space left on device");
    },
  });
  const r = await applyTrackedEdit(d, EDIT, { ref: "R1.1" });
  assert.deepEqual([r.ok, r.code], [false, "AttemptNotRecorded"]);
  assert.match(r.message, /ENOSPC/);
  assert.match(r.message, /not applied|nothing was sent/i);
  assert.deepEqual(steps(d), []);
});

test("applied, then the record fails: told apart from a plain error, and the attempt is left for the refresh", async () => {
  const d = fakeDeps({
    onEdit: async () => {
      d.order.push(["edit"]);
      throw new Error("ENOSPC: no space left on device");
    },
  });
  const r = await applyTrackedEdit(d, EDIT, { ref: "R1.1" });
  assert.deepEqual(
    [r.ok, r.applied, r.code, r.attempt_id],
    [false, true, "AppliedNotRecorded", "R1.1-b1"],
  );
  assert.deepEqual([r.before, r.after, r.word_comment_id], [BEFORE, AFTER, "12"]);
  assert.match(r.message, /do not repeat/i);
  assert.match(r.message, /Refresh/);
  // Not thrown, and the attempt is not taken back. The task pane is told
  // once the call is over, so its reload finds the attempt unowned.
  assert.deepEqual(steps(d), ["attempt", "call", "edit", "end", "record_failed", "end"]);
  assert.deepEqual([d.order[4][1].ref, d.order[4][1].attempt_id], ["R1.1", "R1.1-b1"]);
});

test("not applied: the attempt is taken back (an answer of Word, or an error it sent); no answer keeps it with the call id", async () => {
  let d = fakeDeps({
    call: async () => {
      d.order.push(["call"]);
      return { ok: false, code: "NotFound", message: "not found" };
    },
  });
  let r = await applyTrackedEdit(d, EDIT, { ref: "R1.1" });
  assert.deepEqual([r.ok, r.code], [false, "NotFound"]);
  assert.deepEqual(steps(d), ["attempt", "call", "not_applied", "end"]);
  assert.equal(d.order[2][1].attempt_id, "R1.1-b1");

  d = fakeDeps({
    call: async () => {
      d.order.push(["call"]);
      throw new Error("GeneralException");
    },
  });
  await assert.rejects(applyTrackedEdit(d, EDIT, { ref: "R1.1" }), /GeneralException/);
  assert.deepEqual(steps(d), ["attempt", "call", "not_applied", "end"]);

  d = fakeDeps({
    call: async () => {
      d.order.push(["call"]);
      throw noAnswer("x1");
    },
  });
  r = await applyTrackedEdit(d, EDIT, { ref: "R1.1" });
  assert.equal(r.code, "OutcomeUnknown");
  assert.deepEqual(steps(d), ["attempt", "call", "unknown", "end"]);
  assert.deepEqual(
    [d.order[2][1].attempt_id, d.order[2][1].attempt.tool_call_id],
    ["R1.1-b1", "x1"],
  );
});

// ---- the review service: what is in the file, and when ---------------------------

test("service: the attempt is in the file while Word works; the record and its removal are one save", async () => {
  let during = null;
  const w = await world(async () => {
    during = await w.file();
    return APPLIED;
  });
  try {
    const r = await w.edit();
    assert.equal(r.ok, true, JSON.stringify(r));
    const b = during.blocked[0];
    assert.deepEqual(
      [during.changes.length, b.id, b.state, b.blocked_by],
      [0, "R1.1-b1", "applying", []],
    );
    assert.deepEqual([b.args.find, b.args.replace], ["the substrate", "the medium"]);
    assert.deepEqual(
      [
        b.attempt.tool_call_id,
        b.attempt.comment_text,
        b.attempt.change.before,
        b.attempt.change.after,
      ],
      [null, COMMENT, BEFORE, AFTER],
    );
    // Two saves: the attempt, then the change with the attempt gone.
    assert.deepEqual(
      saves.map((s) => s.comments[0]),
      [
        [0, 1, "applying"],
        [1, 0, ""],
      ],
    );
    assert.deepEqual(
      w.events.map((e) => e.event),
      ["change_recorded"],
    );
    // The id is spent: the next blocked edit or attempt gets the next number.
    assert.equal((await w.file()).blocked_seq, 1);
  } finally {
    await w.cleanup();
  }
});

test("service: the attempt cannot be saved: Word is not called and nothing is left", async () => {
  const w = await world();
  try {
    failSave = (n) => n === 1;
    const r = await w.edit();
    assert.deepEqual([r.ok, r.code, w.calls.length], [false, "AttemptNotRecorded", 0]);
    const c = await w.file();
    assert.deepEqual([c.changes.length, c.blocked.length], [0, 0]);
  } finally {
    await w.cleanup();
  }
});

test("service: Word applied, the record cannot be saved: the attempt stays and the next refresh records the change", async () => {
  const w = await world();
  try {
    failSave = (n) => n === 2; // 1: the attempt, 2: the change
    const r = await w.edit();
    assert.deepEqual(
      [r.ok, r.applied, r.code, r.attempt_id, w.calls.length],
      [false, true, "AppliedNotRecorded", "R1.1-b1", 1],
    );
    let c = await w.file();
    assert.deepEqual([c.changes.length, c.blocked.map((b) => b.state)], [0, ["applying"]]);
    // The task pane is told, for its status line.
    assert.deepEqual(
      w.events.map((e) => [e.event, e.ref, e.attempt_id]),
      [["record_failed", "R1.1", "R1.1-b1"]],
    );
    // A daemon started after that has the attempt only (the daemon that got
    // Word's answer saves the change itself at its next load: see below).
    // Nothing runs it any more: on load it reads like an unanswered edit.
    const d2 = service(w.dir);
    c = (await d2.ask("review_load")).revision.comments[0];
    assert.deepEqual(
      c.blocked.map((b) => b.state),
      ["outcome_unknown"],
    );
    const done = await d2.ask("review_refresh", {
      snapshot: snap(AFTER, [{ id: "12", paragraph_index: 1, content: COMMENT }]),
    });
    c = done.revision.comments[0];
    assert.deepEqual([c.changes.length, c.blocked.length], [1, 0]);
    assert.deepEqual([c.changes[0].word_comment_id, c.changes[0].text], ["12", "the medium"]);
    assert.deepEqual(done.external, [], "not reported as an edit from outside");
  } finally {
    await w.cleanup();
  }
});

test("service: a daemon that stopped while Word worked: the next daemon finds the attempt and a refresh settles it both ways", async () => {
  for (const [text, comments, changes] of [
    [AFTER, [{ id: "12", paragraph_index: 1, content: COMMENT }], 1],
    [BEFORE, [], 0],
  ]) {
    const w = await world(() => new Promise(() => {})); // Word never answers this daemon
    try {
      w.edit(); // left running: the daemon is gone before any answer
      await until(async () => (await w.file()).blocked.length === 1, "attempt on file");
      const d2 = service(w.dir);
      const loaded = (await d2.ask("review_load")).revision.comments[0];
      assert.deepEqual(
        loaded.blocked.map((b) => [b.id, b.state]),
        [["R1.1-b1", "outcome_unknown"]],
      );
      // Not applied again without the author saying so.
      const plain = await d2.ask("review_override", { ref: "R1.1", blocked_id: "R1.1-b1" });
      assert.equal(plain.ok, false);
      assert.match(plain.error, /Refresh status/);
      assert.equal(d2.calls.length, 0);
      const r = await d2.ask("review_refresh", { snapshot: snap(text, comments) });
      const c = r.revision.comments[0];
      assert.deepEqual([c.changes.length, c.blocked.length], [changes, 0]);
    } finally {
      await w.cleanup();
    }
  }
});

test("service: an attempt Word is working on cannot be applied a second time from the Review tab", async () => {
  let answer = null;
  const w = await world(() => new Promise((resolve) => (answer = resolve)));
  try {
    const running = w.edit();
    await until(() => answer, "the call to Word");
    // A load in between does not take the running attempt for a stale one.
    const c = (await w.ask("review_load")).revision.comments[0];
    assert.deepEqual(
      c.blocked.map((b) => b.state),
      ["applying"],
    );
    for (const force of [false, true]) {
      const again = await w.ask("review_override", { ref: "R1.1", blocked_id: "R1.1-b1", force });
      assert.equal(again.ok, false);
      assert.match(again.error, /already being applied/);
    }
    answer(APPLIED);
    assert.equal((await running).ok, true);
    assert.equal(w.calls.length, 1);
    const end = await w.file();
    assert.deepEqual([end.changes.length, end.blocked.length], [1, 0]);
  } finally {
    await w.cleanup();
  }
});

test("service: Word answers that nothing was applied, or its call throws: the attempt leaves the file", async () => {
  for (const call of [
    async () => ({ ok: false, code: "NotFound", message: "not found" }),
    async () => {
      throw new Error("GeneralException");
    },
  ]) {
    const w = await world(call);
    try {
      await w.edit().catch(() => {});
      const c = await w.file();
      assert.deepEqual([c.changes.length, c.blocked.length], [0, 0]);
      assert.deepEqual(w.events, []);
    } finally {
      await w.cleanup();
    }
  }
});

// (b) A late answer used to save the change and then, in a second save,
// remove the unanswered edit: a reader in between saw both.
const neverBoth = (what) => {
  const both = saves.filter((s) => s.comments?.[0][0] >= 1 && s.comments[0][1] >= 1);
  assert.deepEqual(both, [], `${what}: a save held the change and its attempt together`);
};

test("service: a late answer records the change and removes its attempt in one save; a second copy records nothing", async () => {
  const w = await world(async () => {
    throw noAnswer("x1");
  });
  try {
    const r = await w.edit();
    assert.equal(r.code, "OutcomeUnknown");
    const kept = await w.file();
    assert.deepEqual(
      kept.blocked.map((b) => [b.id, b.state, b.attempt.tool_call_id]),
      [["R1.1-b1", "outcome_unknown", "x1"]],
    );
    saves.length = 0;
    const late = { id: "x1", name: "office_tracked_edit", ok: true, result: APPLIED };
    await w.svc.lateToolResult("k", late);
    await w.svc.lateToolResult("k", late); // a second copy
    assert.deepEqual(
      saves.map((s) => s.comments[0]),
      [[1, 0, ""]],
    );
    neverBoth("late answer");
    const c = await w.file();
    assert.deepEqual([c.changes.length, c.blocked.length], [1, 0]);
  } finally {
    await w.cleanup();
  }
});

test("service: an approved exception is settled in one save too, and its attempt is on file before Word is called", async () => {
  let during = null;
  const w = await world(async () => {
    during = await w.file();
    return APPLIED;
  });
  try {
    await w.ask("ledger_add", {
      item: { kind: "forbid", text: "no medium", check: { type: "regex", pattern: "medium" } },
    });
    const blocked = await w.edit();
    assert.deepEqual(blocked.blocked_by, ["L1"]);
    saves.length = 0;
    const r = await w.ask("review_override", { ref: "R1.1", blocked_id: blocked.blocked_id });
    assert.deepEqual([r.ok, r.applied], [true, true], JSON.stringify(r));
    assert.deepEqual(
      [during.blocked[0].state, during.blocked[0].attempt?.change.after],
      ["applying", AFTER],
    );
    neverBoth("override");
    const c = r.revision.comments[0];
    assert.deepEqual([c.changes.length, c.blocked.length], [1, 0]);
  } finally {
    await w.cleanup();
  }
});

test("service: an exception Word refuses goes back as it was, without the attempt of this try", async () => {
  const w = await world(async () => ({ ok: false, code: "NotFound", message: "not found" }));
  try {
    await w.ask("ledger_add", {
      item: { kind: "forbid", text: "no medium", check: { type: "regex", pattern: "medium" } },
    });
    const blocked = await w.edit();
    const r = await w.ask("review_override", { ref: "R1.1", blocked_id: blocked.blocked_id });
    assert.deepEqual([r.ok, r.applied], [true, false]);
    const b = r.revision.comments[0].blocked[0];
    assert.deepEqual([b.state, b.attempt], ["blocked", null]);
  } finally {
    await w.cleanup();
  }
});

// Back to "blocked" would offer Allow exception again for an edit that is
// in the document: a second press would apply it twice.
test("service: an exception Word applied whose record cannot be saved does not go back to blocked", async () => {
  const w = await world();
  try {
    await w.ask("ledger_add", {
      item: { kind: "forbid", text: "no medium", check: { type: "regex", pattern: "medium" } },
    });
    const blocked = await w.edit();
    saves.length = 0;
    failSave = (n) => n === 3; // 1: the claim, 2: its attempt, 3: the change
    const r = await w.ask("review_override", { ref: "R1.1", blocked_id: blocked.blocked_id });
    assert.deepEqual(
      [r.ok, r.applied, r.outcome_unknown, r.result?.code],
      [true, false, true, "AppliedNotRecorded"],
      JSON.stringify(r),
    );
    const b = r.revision.comments[0].blocked[0];
    assert.deepEqual(
      [b.state, b.blocked_by, b.attempt?.change.after],
      ["outcome_unknown", ["L1"], AFTER],
    );
    assert.equal(r.revision.comments[0].changes.length, 0);
    const again = await w.ask("review_override", { ref: "R1.1", blocked_id: blocked.blocked_id });
    assert.equal(again.ok, false);
    assert.match(again.error, /Refresh status/);
    assert.equal(w.calls.length, 1, "Word was called once");
    const done = await w.ask("review_refresh", {
      snapshot: snap(AFTER, [
        { id: "12", paragraph_index: 1, content: "[R1.1] w / y (금지 규칙 L1 예외 승인) / a" },
      ]),
    });
    const c = done.revision.comments[0];
    assert.deepEqual([c.changes.length, c.blocked.length], [1, 0]);
    assert.deepEqual([c.changes[0].override_of, c.changes[0].word_comment_id], [["L1"], "12"]);
  } finally {
    await w.cleanup();
  }
});

// Codex review of the recovery path (P2): Word may apply less than the call
// asked for. "grew carefully on the substrate" -> "grew on the substrate" goes
// in as the tracked deletion of "carefully " (narrow). When the save of that
// change failed, only the attempt written before the call was left, and it
// holds the full replace: a refresh read the paragraph (old words still there,
// one of them tracked as deleted) as "the replace was rejected", and the
// deletion in the document had no change on record.
const B2 = "Cells grew carefully on the substrate at 30 C.";
const A2 = "Cells grew on the substrate at 30 C.";
const EDIT2 = {
  comment_ref: "R1.1",
  find: "grew carefully on the substrate",
  replace: "grew on the substrate",
  rationale: R,
};
const NARROWED = {
  ok: true,
  word_comment_id: 12,
  paragraph_index: 1,
  paragraph_id: "p1",
  before: B2,
  after: A2,
  narrowed: true,
  target_checked: true,
};
// The paragraph with "carefully " tracked as deleted: Word's text still has it.
const pendingDeletion = (comments = []) => ({
  paragraphs: [
    { index: 0, id: "p0", text: "Title", w_ins: 0, w_del: 0, tracked: [] },
    {
      index: 1,
      id: "p1",
      text: B2,
      w_ins: 0,
      w_del: 1,
      tracked: [{ kind: "del", text: "carefully ", start: 11, marks: 1 }],
    },
  ],
  comments,
});
const untouched = () => ({
  paragraphs: [
    { index: 0, id: "p0", text: "Title", w_ins: 0, w_del: 0, tracked: [] },
    { index: 1, id: "p1", text: B2, w_ins: 0, w_del: 0, tracked: [] },
  ],
  comments: [],
});
// Word narrows the edit, and the save of its change fails.
async function narrowedUnsaved() {
  const w = await world(async () => NARROWED);
  await w.ask("review_refresh", { snapshot: untouched() });
  saves.length = 0;
  failSave = (n) => n === 2; // 1: the attempt, 2: the change
  const r = await w.edit(EDIT2);
  assert.deepEqual(
    [r.ok, r.applied, r.code, r.kind],
    [false, true, "AppliedNotRecorded", "delete"],
  );
  assert.deepEqual(w.calls[0][1].narrow, { find: "carefully ", replace: "" });
  failSave = () => false;
  return w;
}
const narrowedForm = (ch) => [
  ch.kind,
  ch.find,
  ch.text,
  ch.narrowed_from,
  ch.context,
  ch.word_comment_id,
  ch.paragraph_id,
];
const DELETION = [
  "delete",
  "carefully ",
  "",
  { find: EDIT2.find, text: EDIT2.replace },
  { left: "Cells grew ", right: "on the substrate" },
  "12",
  "p1",
];

test("service: a narrowed edit whose record failed is recorded by the same daemon as Word applied it", async () => {
  const w = await narrowedUnsaved();
  try {
    // The disk is still full: the load fails, the applied result is kept.
    failSave = () => true;
    assert.equal((await w.ask("review_load")).ok, false);
    assert.equal((await w.file()).changes.length, 0);
    failSave = () => false;
    // The next load records what Word answered, not what the call asked for.
    const c = (await w.ask("review_load")).revision.comments[0];
    assert.deepEqual([c.changes.length, c.blocked.length], [1, 0], JSON.stringify(c));
    assert.deepEqual(narrowedForm(c.changes[0]), DELETION);
    assert.deepEqual([c.changes[0].before, c.changes[0].after], [B2, A2]);
    assert.ok(
      w.lines.some((l) => /change recorded R1\.1-1 .*kept in memory/.test(l)),
      w.lines.join("\n"),
    );
    // Once: a second load adds nothing, and the refresh reads it pending.
    assert.equal((await w.ask("review_load")).revision.comments[0].changes.length, 1);
    const done = await w.ask("review_refresh", {
      snapshot: pendingDeletion([{ id: "12", paragraph_index: 1, content: COMMENT }]),
    });
    assert.deepEqual(
      done.rows.map((x) => [x.change_id, x.state]),
      [["R1.1-1", "pending"]],
    );
    assert.equal(done.revision.comments[0].changes.length, 1);
    assert.deepEqual(done.external, []);
  } finally {
    await w.cleanup();
  }
});

test("service: after a restart a narrowed edit in the document is not read as a rejected replace", async () => {
  // With its rationale comment in the document: recorded in the form Word applied.
  let w = await narrowedUnsaved();
  try {
    const file = await w.file();
    const d2 = service(w.dir);
    assert.deepEqual(
      (await d2.ask("review_load")).revision.comments[0].blocked.map((b) => b.state),
      ["outcome_unknown"],
    );
    const done = await d2.ask("review_refresh", {
      snapshot: pendingDeletion([{ id: "12", paragraph_index: 1, content: COMMENT }]),
    });
    const c = done.revision.comments[0];
    assert.deepEqual([c.changes.length, c.blocked.length], [1, 0], JSON.stringify(c));
    const attempt = file.blocked[0].attempt.change;
    assert.deepEqual(
      [attempt.kind, attempt.find, attempt.narrow],
      ["replace", EDIT2.find, { find: "carefully ", replace: "" }],
      "the attempt on file names the narrower form Word may apply",
    );
    assert.deepEqual(narrowedForm(c.changes[0]), DELETION);
    assert.equal("narrow" in c.changes[0], false);
    assert.deepEqual(
      done.rows.map((x) => [x.change_id, x.state]),
      [["R1.1-1", "pending"]],
    );
  } finally {
    await w.cleanup();
  }

  // Without the comment a tracked deletion of its words proves nothing (the
  // author may have deleted them): not recorded, and not dropped as "not
  // applied" either. It stays in the list as unanswered.
  w = await narrowedUnsaved();
  try {
    const d2 = service(w.dir);
    await d2.ask("review_load");
    const c = (await d2.ask("review_refresh", { snapshot: pendingDeletion() })).revision
      .comments[0];
    assert.deepEqual(
      [c.changes.length, c.blocked.map((b) => [b.id, b.state])],
      [0, [["R1.1-b1", "outcome_unknown"]]],
    );
    assert.equal(
      d2.lines.some((l) => /not applied/.test(l)),
      false,
      d2.lines.join("\n"),
    );
  } finally {
    await w.cleanup();
  }

  // The paragraph as it was, nothing tracked, no comment: neither form is in
  // the document, and the attempt goes.
  w = await narrowedUnsaved();
  try {
    const d2 = service(w.dir);
    await d2.ask("review_load");
    const c = (await d2.ask("review_refresh", { snapshot: untouched() })).revision.comments[0];
    assert.deepEqual([c.changes.length, c.blocked.length], [0, 0]);
  } finally {
    await w.cleanup();
  }
});

// Which of the two forms a refresh records, for a daemon that has the
// attempt only. The full call is looked for first; the narrower one counts
// without its comment only as Word's tracked replace (new words, then old).
test("service: after a restart the form found in the document is the one recorded", async () => {
  const para = (text, tracked) => ({
    paragraphs: [
      { index: 0, id: "p0", text: "Title", w_ins: 0, w_del: 0, tracked: [] },
      {
        index: 1,
        id: "p1",
        text,
        w_ins: tracked.filter((r) => r.kind === "ins").length,
        w_del: tracked.filter((r) => r.kind === "del").length,
        tracked,
      },
    ],
    comments: [],
  });
  const settle = async (w, snapshot) => {
    const d2 = service(w.dir);
    await d2.ask("review_load");
    return (await d2.ask("review_refresh", { snapshot })).revision.comments[0];
  };

  // Word could not narrow and replaced the whole find; no comment. The old
  // words are tracked as deleted after the new ones: the full replace, and
  // not "the deletion of carefully was accepted".
  let w = await narrowedUnsaved();
  try {
    const c = await settle(
      w,
      para("Cells grew on the substrategrew carefully on the substrate at 30 C.", [
        { kind: "ins", text: "grew on the substrate", start: 6, marks: 1 },
        { kind: "del", text: "grew carefully on the substrate", start: 27, marks: 1 },
      ]),
    );
    assert.deepEqual([c.changes.length, c.blocked.length], [1, 0], JSON.stringify(c));
    assert.deepEqual(
      [c.changes[0].kind, c.changes[0].find, c.changes[0].narrowed_from, c.changes[0].state],
      ["replace", EDIT2.find, null, "pending"],
    );
  } finally {
    await w.cleanup();
  }

  // A narrowed replace, still tracked, without its comment: "°C" next to the
  // deleted "degrees" is Word's tracked replace of exactly those words.
  const B3 = "It was held at 30 degrees for 5 days.";
  const EDIT3 = {
    comment_ref: "R1.1",
    find: "held at 30 degrees for",
    replace: "held at 30 °C for",
    rationale: R,
  };
  const unsaved3 = async () => {
    const x = await world(async () => ({
      ...NARROWED,
      before: B3,
      after: "It was held at 30 °C for 5 days.",
    }));
    await x.ask("review_refresh", { snapshot: para(B3, []) });
    saves.length = 0;
    failSave = (n) => n === 2;
    const r = await x.edit(EDIT3);
    assert.equal(r.code, "AppliedNotRecorded", JSON.stringify(r));
    assert.deepEqual(x.calls[0][1].narrow, { find: "degrees", replace: "°C" });
    failSave = () => false;
    return x;
  };
  w = await unsaved3();
  try {
    const c = await settle(
      w,
      para("It was held at 30 °Cdegrees for 5 days.", [
        { kind: "ins", text: "°C", start: 18, marks: 1 },
        { kind: "del", text: "degrees", start: 20, marks: 1 },
      ]),
    );
    assert.deepEqual([c.changes.length, c.blocked.length], [1, 0], JSON.stringify(c));
    assert.deepEqual(
      [c.changes[0].kind, c.changes[0].find, c.changes[0].text, c.changes[0].narrowed_from],
      ["replace", "degrees", "°C", { find: EDIT3.find, text: EDIT3.replace }],
    );
    assert.deepEqual(
      [c.changes[0].comment_missing, c.changes[0].state],
      ["CommentNotFound", "pending"],
    );
  } finally {
    await w.cleanup();
  }

  // "°C" somewhere else in the paragraph, the old words in place and another
  // tracked change of the author's: the new words alone say nothing, and
  // both forms read the old words in their place. Not recorded.
  w = await unsaved3();
  try {
    const c = await settle(
      w,
      para("It was held at 30 degrees for 5 days, not at 4 °C as before.", [
        { kind: "ins", text: "as before", start: 50, marks: 1 },
      ]),
    );
    assert.deepEqual([c.changes.length, c.blocked.map((b) => b.state)], [0, []], JSON.stringify(c));
  } finally {
    await w.cleanup();
  }
});

async function until(fn, what, ms = 3000) {
  const end = Date.now() + ms;
  for (;;) {
    const v = await fn();
    if (v) return v;
    if (Date.now() > end) throw new Error(`timed out waiting for ${what}`);
    await new Promise((r) => setTimeout(r, 5));
  }
}
