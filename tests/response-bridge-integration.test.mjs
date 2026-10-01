// The response service on the REAL bridge (daemon/bridge.mjs, port 0) with a
// real WebSocket client in the role of the task pane, wired as in
// daemon/index.mjs. The unit tests call the handlers directly; this checks
// what goes over the wire: the reply to a plain request, a job accepted and
// ended by a response_event, and a draft whose pane drops while the model is
// still writing.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { WebSocket } from "ws";
import { createBridge } from "../daemon/bridge.mjs";
import { createResponseService } from "../daemon/response-service.mjs";
import { emptyRevision, loadRevision, saveRevision } from "../daemon/revision-store.mjs";

const ENV = { CC_OFFICE_WSL_PATHS: "0", CC_OFFICE_CONFIG_DIR: "/nonexistent-config" };
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

const SECTION = "2. Materials and Methods";
const BEFORE = "Flasks were shaken. The temperature was held at 30 degrees. Samples were drawn.";
const AFTER = "Flasks were shaken. The temperature was held at 30 °C. Samples were drawn.";
const SNAPSHOT = {
  paragraphs: [
    { index: 0, id: "P0", text: SECTION, w_ins: 0, w_del: 0 },
    { index: 1, id: "P1", text: AFTER, w_ins: 0, w_del: 0 },
  ],
  comments: [],
};
const HEADINGS = [{ index: 0, level: 1, text: SECTION }];
const COMMENTS = [
  {
    ref: "R1.1",
    reviewer: 1,
    text: "Please state the unit.",
    requirements: [{ id: "q1", text: "state the unit" }],
    changes: [
      {
        id: "R1.1-1",
        word_comment_id: null,
        paragraph_index: 1,
        paragraph_id: "P1",
        kind: "replace",
        find: "30 degrees",
        text: "30 °C",
        before: BEFORE,
        after: AFTER,
        rationale: { what: "Unit.", why: "Asked.", addresses: "unit", requirement_ids: ["q1"] },
        context: { left: "was held at ", right: ". Samples were" },
        created_at: "2026-01-01T00:00:00.000Z",
      },
    ],
  },
];

async function world() {
  const dir = await mkdtemp(join(tmpdir(), "response-bridge-"));
  const doc = join(dir, "paper.docx");
  await saveRevision(join(dir, "paper.revision.json"), {
    ...emptyRevision(doc),
    comments: COMMENTS,
  });
  const lines = [];
  const handlers = {};
  const pinned = [];
  let release = null; // set while the model is "writing"
  const bridge = createBridge({
    port: 0,
    token: TOKEN,
    allowedOrigins: [],
    extraHandlers: handlers,
  });
  const response = createResponseService({
    getContext: (key) => bridge.getContext(key),
    oneshotJob: async (key, options) => {
      pinned.push(options ?? null);
      return {
        model: "fake:model",
        run: () => new Promise((resolve) => (release = resolve)),
      };
    },
    sendEvent: (key, obj) => bridge.sendToTaskpane(obj, key),
    runPy: async () => {
      throw new Error("no Python in this test");
    },
    log: (s) => lines.push(s),
    env: ENV,
  });
  Object.assign(handlers, response.handlers);
  await until(() => bridge.address(), "bridge listening");

  const ws = new WebSocket(`ws://127.0.0.1:${bridge.address().port}`);
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
  ws.send(JSON.stringify({ type: "hello", token: TOKEN, host: "word", active_doc: doc }));
  await until(() => frames.some((f) => f.type === "welcome"), "welcome");
  let n = 0;
  const ask = (type, payload = {}) => {
    const request_id = `r${++n}`;
    return new Promise((resolve) => {
      waiters.set(request_id, resolve);
      ws.send(JSON.stringify({ type, request_id, ...payload }));
    });
  };
  return {
    doc,
    lines,
    frames,
    pinned,
    ask,
    release: (v) => release(v),
    writing: () => until(() => release, "the model call"),
    rev: () => loadRevision(join(dir, "paper.revision.json"), doc),
    async closePane() {
      const closed = new Promise((r) => ws.on("close", r));
      ws.close();
      await closed;
    },
    async cleanup() {
      if (ws.readyState === WebSocket.OPEN) ws.close();
      await bridge.close?.();
      await rm(dir, { recursive: true, force: true });
    },
  };
}

test("real bridge: a plain request is answered, a draft is accepted and ends in a response_event", async () => {
  const w = await world();
  try {
    const state = await w.ask("response_state", { snapshot: SNAPSHOT, headings: HEADINGS });
    assert.equal(state.type, "response_state_result");
    assert.equal(state.ok, true, state.error);
    const c = state.revision.comments[0];
    assert.equal(c.changes[0].state, "accepted");
    assert.equal(c.response.citations[0].section, SECTION);

    const accepted = await w.ask("response_draft", {
      ref: "R1.1",
      snapshot: SNAPSHOT,
      headings: HEADINGS,
    });
    assert.deepEqual([accepted.ok, accepted.accepted], [true, true]);
    await w.writing();
    assert.deepEqual(w.pinned, [{ role: "writing" }]);
    w.release({ ok: true, json: { response: "We now state the unit." } });
    const done = await until(
      () => w.frames.find((f) => f.type === "response_event" && f.job_id === accepted.job_id),
      "job_done",
    );
    assert.equal(done.event, "job_done");
    assert.equal(done.result.ok, true, done.result.error);
    assert.equal(done.result.revision.comments[0].response.text, "We now state the unit.");
  } finally {
    await w.cleanup();
  }
});

test("real bridge: the pane drops while the model writes; the draft is saved to the document it was asked for", async () => {
  const w = await world();
  try {
    const accepted = await w.ask("response_draft", {
      ref: "R1.1",
      snapshot: SNAPSHOT,
      headings: HEADINGS,
    });
    assert.equal(accepted.accepted, true);
    await w.writing();
    await w.closePane(); // the bridge forgets the pane's context
    w.release({ ok: true, json: { response: "We now state the unit." } });
    const rev = await until(async () => {
      const r = await w.rev();
      return r.comments[0].response?.text ? r : null;
    }, "the saved draft");
    assert.equal(rev.comments[0].response.text, "We now state the unit.");
    assert.equal(rev.comments[0].response.source, "model");
    // Nothing thrown into the void: the undelivered event is at most a log line.
    assert.equal(
      w.lines.some((l) => /no document|not connected.*draft/i.test(l)),
      false,
    );
  } finally {
    await w.cleanup();
  }
});
