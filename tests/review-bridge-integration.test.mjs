// The review service on the REAL bridge (daemon/bridge.mjs) with a real
// WebSocket client in the role of the task pane. Codex review 15: the unit
// tests fake the bridge, so they missed that a dropped socket removes the
// pane's context; the service then re-read the document path from that
// context after the tool call, threw, and lost a tracked edit Word had
// applied ("late result dropped: no pending tracked edit").
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, access } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { WebSocket } from "ws";
import { createBridge } from "../daemon/bridge.mjs";
import { createReviewService } from "../daemon/review-service.mjs";
import { applyTrackedEdit } from "../daemon/review-tools.mjs";
import { PANE_CAPS } from "../taskpane/shared/tools-word.js";

const ENV = { CC_OFFICE_WSL_PATHS: "0" };
const TOKEN = "t";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
async function until(fn, what, ms = 3000) {
  const end = Date.now() + ms;
  for (;;) {
    const v = await fn();
    if (v) return v;
    if (Date.now() > end) throw new Error(`timed out waiting for ${what}`);
    await sleep(10);
  }
}

// A daemon: the bridge and the review service wired as in daemon/index.mjs.
async function daemon() {
  const lines = [];
  const handlers = {};
  const keys = new Map(); // document -> pane key, as the bridge made it
  let review = null;
  const bridge = createBridge({
    port: 0,
    token: TOKEN,
    allowedOrigins: [],
    extraHandlers: handlers,
    onHello: (key, host, doc) => keys.set(doc, key),
    onLateToolResult: (e) => review.lateToolResult(e.key, e),
  });
  review = createReviewService({
    getContext: bridge.getContext,
    callTool: (name, args, key) => bridge.callTaskpaneTool(name, args, key),
    pushTurn: () => {},
    sendEvent: (key, obj) => bridge.sendToTaskpane(obj, key),
    log: (s) => lines.push(s),
    env: ENV,
  });
  Object.assign(handlers, review.handlers);
  await until(() => bridge.address(), "bridge listening");
  // A tracked edit of a review turn, as office-tools.mjs runs it.
  const edit = (key, args, ref) =>
    applyTrackedEdit(
      {
        ...review.trackedEditDeps(key),
        call: (name, a) => bridge.callTaskpaneTool(name, a, key),
      },
      { comment_ref: ref, rationale: { what: "w", why: "y", addresses: "a" }, ...args },
      { ref },
    );
  return { bridge, review, lines, edit, keys, port: () => bridge.address().port };
}

// The task pane of one document: hello, requests with request_id, and the
// tool calls it was sent (it answers only when the test says so).
// caps: what the hello says the pane can do. The real task pane sends
// PANE_CAPS (taskpane.js sendHello); [] is one from before caps.
async function pane(port, doc, caps = PANE_CAPS) {
  const ws = new WebSocket(`ws://127.0.0.1:${port}`);
  const frames = [];
  const waiters = new Map();
  ws.on("message", (raw) => {
    const m = JSON.parse(raw.toString());
    frames.push(m);
    if (m.request_id && waiters.has(m.request_id)) {
      waiters.get(m.request_id)(m);
      waiters.delete(m.request_id);
    }
  });
  await new Promise((resolve, reject) => {
    ws.on("open", resolve);
    ws.on("error", reject);
  });
  ws.send(JSON.stringify({ type: "hello", token: TOKEN, host: "word", active_doc: doc, caps }));
  await until(() => frames.some((f) => f.type === "welcome"), "welcome");
  let n = 0;
  return {
    ws,
    frames,
    toolCalls: () => frames.filter((f) => f.type === "tool_call"),
    ask(type, payload = {}) {
      const request_id = `r${++n}`;
      return new Promise((resolve) => {
        waiters.set(request_id, resolve);
        ws.send(JSON.stringify({ type, request_id, ...payload }));
      });
    },
    send: (obj) => ws.send(JSON.stringify(obj)),
    async close() {
      const closed = new Promise((r) => ws.on("close", r));
      ws.close();
      await closed;
    },
  };
}

const BEFORE = "Cells grew on the substrate at 30 C.";
const AFTER = "Cells grew on the medium at 30 C.";
const COMMENT = "[R1.1] w / y / a";
const EDIT = { find: "the substrate", replace: "the medium" };
const APPLIED = { ok: true, word_comment_id: 12, paragraph_index: 1, before: BEFORE, after: AFTER };
const snap = (text, comments = []) => ({
  paragraphs: [
    { index: 0, id: "p0", text: "Title", w_ins: 0, w_del: 0 },
    { index: 1, id: "p1", text, w_ins: 0, w_del: 0 },
  ],
  comments,
});
const revisionOf = async (doc) =>
  JSON.parse(await readFile(doc.replace(/\.docx$/, ".revision.json"), "utf8"));

async function world(caps) {
  const dir = await mkdtemp(join(tmpdir(), "review-bridge-"));
  const doc = join(dir, "paper.docx");
  const d = await daemon();
  const p = await pane(d.port(), doc, caps);
  const saved = await p.ask("review_save_comments", {
    comments: [{ ref: "R1.1", reviewer: 1, text: "Name the medium." }],
  });
  assert.equal(saved.ok, true, saved.error);
  assert.equal((await p.ask("review_refresh", { snapshot: snap(BEFORE) })).ok, true);
  return { dir, doc, d, p, cleanup: () => rm(dir, { recursive: true, force: true }) };
}

// An edit in flight whose pane drops before it answers.
async function cutEdit(w) {
  const running = w.d.edit(w.d.keys.get(w.doc), EDIT, "R1.1");
  const call = await until(() => w.p.toolCalls()[0], "tool_call");
  await w.p.close();
  const r = await running;
  assert.equal(r.code, "OutcomeUnknown", JSON.stringify(r));
  return call;
}

test("real bridge: the socket drops with an edit in flight, the pane reconnects and delivers the result; recorded once", async () => {
  const w = await world();
  let p2 = null;
  try {
    const call = await cutEdit(w);
    assert.equal(call.name, "office_tracked_edit");
    // Written to the file although the pane (and its context) is gone.
    const cut = (await revisionOf(w.doc)).comments[0];
    assert.deepEqual([cut.changes.length, cut.blocked[0]?.state], [0, "outcome_unknown"]);
    assert.equal(cut.blocked[0].attempt.tool_call_id, call.id);

    p2 = await pane(w.d.port(), w.doc);
    p2.send({ type: "tool_result", id: call.id, ok: true, result: APPLIED });
    // One save records the change and removes the unanswered edit: the
    // first read that shows the change shows no blocked edit (two saves used
    // to leave [1, 1] in between, and a read after a fixed wait hit it).
    const first = await until(async () => {
      const x = (await revisionOf(w.doc)).comments[0];
      return x.changes.length === 1 && x;
    }, "change");
    assert.deepEqual([first.changes.length, first.blocked.length], [1, 0]);
    p2.send({ type: "tool_result", id: call.id, ok: true, result: APPLIED }); // a second copy
    // Answered after the second copy: same socket, and the same file chain.
    const c = (await p2.ask("review_load")).revision.comments[0];
    assert.deepEqual([c.changes.length, c.blocked.length], [1, 0]);
    assert.equal(c.changes[0].word_comment_id, "12");
    assert.ok(w.d.lines.includes(`[review] late result recorded ${call.id}`));
    assert.ok(!w.d.lines.some((l) => /no pending tracked edit/.test(l)));
    // The reconnected pane is told, so it reloads.
    await until(
      () => p2.frames.some((f) => f.event === "late_result_recorded"),
      "late_result_recorded",
    );
  } finally {
    p2?.ws.terminate();
    await w.d.bridge.close();
    await w.cleanup();
  }
});

test("real bridge: the socket drops and no result ever comes; the attempt is in the file and a refresh settles it both ways", async () => {
  for (const [text, comments, changes] of [
    [AFTER, [{ id: "12", paragraph_index: 1, content: COMMENT }], 1],
    [BEFORE, [], 0],
  ]) {
    const w = await world();
    let p2 = null;
    try {
      await cutEdit(w);
      assert.equal((await revisionOf(w.doc)).comments[0].blocked[0].state, "outcome_unknown");
      p2 = await pane(w.d.port(), w.doc);
      const r = await p2.ask("review_refresh", { snapshot: snap(text, comments) });
      assert.equal(r.ok, true, r.error);
      const c = r.revision.comments[0];
      assert.deepEqual([c.changes.length, c.blocked.length], [changes, 0]);
    } finally {
      p2?.ws.terminate();
      await w.d.bridge.close();
      await w.cleanup();
    }
  }
});

test("real bridge: the socket drops and the daemon restarts before the pane is back; the attempt is settled on refresh", async () => {
  const w = await world();
  let d2 = null;
  let p2 = null;
  try {
    const call = await cutEdit(w);
    await w.d.bridge.close();
    d2 = await daemon(); // nothing of the first daemon in memory
    p2 = await pane(d2.port(), w.doc);
    // The answer of the old call means nothing to this daemon.
    p2.send({ type: "tool_result", id: call.id, ok: true, result: APPLIED });
    const loaded = await p2.ask("review_load");
    assert.equal(loaded.revision.comments[0].blocked[0].state, "outcome_unknown");
    const r = await p2.ask("review_refresh", {
      snapshot: snap(AFTER, [{ id: "12", paragraph_index: 1, content: COMMENT }]),
    });
    const c = r.revision.comments[0];
    assert.deepEqual([c.changes.length, c.blocked.length], [1, 0]);
    assert.equal(c.changes[0].word_comment_id, "12");
  } finally {
    p2?.ws.terminate();
    await d2?.bridge.close();
    await w.cleanup();
  }
});

test("real bridge: the pane of another document that connects in between neither gets nor consumes the result", async () => {
  const w = await world();
  let other = null;
  let p2 = null;
  try {
    const call = await cutEdit(w);
    const doc2 = join(w.dir, "other.docx");
    other = await pane(w.d.port(), doc2);
    // It answers with the id of the other document's call: not accepted.
    other.send({ type: "tool_result", id: call.id, ok: true, result: APPLIED });
    await sleep(50);
    assert.equal((await revisionOf(w.doc)).comments[0].changes.length, 0);
    // The pane of the document itself still delivers it.
    p2 = await pane(w.d.port(), w.doc);
    p2.send({ type: "tool_result", id: call.id, ok: true, result: APPLIED });
    await until(async () => (await revisionOf(w.doc)).comments[0].changes.length === 1, "change");
    await sleep(50);
    // Nothing of it reached the other pane or its folder.
    assert.deepEqual(
      other.frames.filter((f) => f.type === "tool_call" || f.type === "review_event"),
      [],
    );
    await assert.rejects(access(doc2.replace(/\.docx$/, ".revision.json")));
  } finally {
    other?.ws.terminate();
    p2?.ws.terminate();
    await w.d.bridge.close();
    await w.cleanup();
  }
});

// The same pane key can come to name another file (the document was saved
// under a new name while the edit ran): the answer still belongs to the
// document the edit was made in.
test("real bridge: an edit answered after the pane's document changed is recorded in the document it was made in", async () => {
  const w = await world();
  try {
    const running = w.d.edit(w.d.keys.get(w.doc), EDIT, "R1.1");
    const call = await until(() => w.p.toolCalls()[0], "tool_call");
    const doc2 = join(w.dir, "renamed.docx");
    w.p.send({ type: "context_update", active_doc: doc2 });
    await sleep(30);
    w.p.send({ type: "tool_result", id: call.id, ok: true, result: APPLIED });
    const r = await running;
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.equal((await revisionOf(w.doc)).comments[0].changes.length, 1);
    await assert.rejects(access(doc2.replace(/\.docx$/, ".revision.json")));
  } finally {
    w.p.ws.terminate();
    await w.d.bridge.close();
    await w.cleanup();
  }
});

// A task pane loaded before the daemon was updated: its hello names no
// capability, so it would apply a tracked edit without comparing the
// paragraph first. It is sent no tool call; one that reloads gets the edit.
test("real bridge: a pane whose hello has no caps is sent no tracked edit; after a reload it is", async () => {
  const w = await world([]);
  let p2 = null;
  try {
    const key = w.d.keys.get(w.doc);
    const r = await w.d.edit(key, EDIT, "R1.1");
    assert.deepEqual([r.ok, r.code], [false, "TaskPaneOutdated"], JSON.stringify(r));
    assert.match(r.message, /reload the task pane/i);
    // Answered after anything the daemon sent before it: no tool call came.
    const loaded = await w.p.ask("review_load");
    assert.deepEqual(w.p.toolCalls(), []);
    const c = loaded.revision.comments[0];
    assert.deepEqual([c.changes.length, c.blocked.length], [0, 0]);

    await w.p.close();
    p2 = await pane(w.d.port(), w.doc); // the reloaded task pane
    const running = w.d.edit(w.d.keys.get(w.doc), EDIT, "R1.1");
    const call = await until(() => p2.toolCalls()[0], "tool_call");
    assert.equal(typeof call.args.expected_before_hash, "string");
    p2.send({ type: "tool_result", id: call.id, ok: true, result: APPLIED });
    assert.equal((await running).ok, true);
  } finally {
    w.p.ws.terminate();
    p2?.ws.terminate();
    await w.d.bridge.close();
    await w.cleanup();
  }
});
