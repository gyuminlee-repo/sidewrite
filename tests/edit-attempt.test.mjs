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
    getContext: () => ({ activeDoc: doc }),
    callTool,
    pushTurn: () => {},
    sendEvent: (key, obj) => events.push(obj),
    log: (s) => lines.push(s),
    env: ENV,
  });
  const ask = (type, payload = {}) =>
    new Promise((resolve) =>
      svc.handlers[type]({ type, request_id: "q", ...payload }, resolve, "k", "word"),
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
  // Not thrown, and the attempt is not taken back.
  assert.deepEqual(steps(d), ["attempt", "call", "edit", "end"]);
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
    // Nothing runs it any more: on load it reads like an unanswered edit.
    c = (await w.ask("review_load")).revision.comments[0];
    assert.deepEqual(
      c.blocked.map((b) => b.state),
      ["outcome_unknown"],
    );
    const done = await w.ask("review_refresh", {
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

async function until(fn, what, ms = 3000) {
  const end = Date.now() + ms;
  for (;;) {
    const v = await fn();
    if (v) return v;
    if (Date.now() > end) throw new Error(`timed out waiting for ${what}`);
    await new Promise((r) => setTimeout(r, 5));
  }
}
