// A manuscript and its supplementary document, each in a pane of its own, on
// the REAL bridge (daemon/bridge.mjs) with WebSocket clients in the role of
// the task panes (as tests/review-bridge-integration.test.mjs): a change of
// the link reaches the other pane as a review_event, and a tracked edit of
// one document is neither sent to nor taken from the pane of the other.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
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

// The bridge and the review service wired as in daemon/index.mjs.
async function daemon() {
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
    listPanes: () => bridge.listPanes(),
    log: () => {},
    env: ENV,
  });
  Object.assign(handlers, review.handlers);
  await until(() => bridge.address(), "bridge listening");
  const edit = (key, args, ref) =>
    applyTrackedEdit(
      {
        ...review.trackedEditDeps(key),
        call: (name, a) => bridge.callTaskpaneTool(name, a, key),
      },
      { comment_ref: ref, rationale: { what: "w", why: "y", addresses: "a" }, ...args },
      { ref },
    );
  return { bridge, edit, keys, port: () => bridge.address().port };
}

async function pane(port, doc) {
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
  ws.send(
    JSON.stringify({ type: "hello", token: TOKEN, host: "word", active_doc: doc, caps: PANE_CAPS }),
  );
  await until(() => frames.some((f) => f.type === "welcome"), "welcome");
  let n = 0;
  return {
    ws,
    frames,
    of: (type) => frames.filter((f) => f.type === type),
    ask(type, payload = {}) {
      const request_id = `r${++n}`;
      return new Promise((resolve) => {
        waiters.set(request_id, resolve);
        ws.send(JSON.stringify({ type, request_id, ...payload }));
      });
    },
    send: (obj) => ws.send(JSON.stringify(obj)),
  };
}

const BEFORE = "Cells grew on the substrate at 30 C.";
const AFTER = "Cells grew on the medium at 30 C.";
const snap = (text) => ({
  paragraphs: [
    { index: 0, id: "p0", text: "Title", w_ins: 0, w_del: 0 },
    { index: 1, id: "p1", text, w_ins: 0, w_del: 0 },
  ],
  comments: [],
});
const revisionOf = async (doc) =>
  JSON.parse(await readFile(doc.replace(/\.docx$/, ".revision.json"), "utf8"));

test("real bridge, two panes: a join in the supplement reaches the manuscript's pane; tool results do not cross", async () => {
  const dir = await mkdtemp(join(tmpdir(), "parts-bridge-"));
  const paper = join(dir, "paper.docx");
  const si = join(dir, "SI.docx");
  const d = await daemon();
  let main = null;
  let supp = null;
  let third = null;
  try {
    main = await pane(d.port(), paper);
    supp = await pane(d.port(), si);
    third = await pane(d.port(), join(dir, "other.docx"));
    const saved = await main.ask("review_save_comments", {
      comments: [{ ref: "R1.1", reviewer: 1, text: "Name the medium." }],
    });
    assert.equal(saved.ok, true, saved.error);
    assert.equal((await main.ask("review_refresh", { snapshot: snap(BEFORE) })).ok, true);
    // Nothing is linked by a load.
    assert.equal((await supp.ask("review_load")).linked.part_of, null);
    assert.deepEqual(
      (await supp.ask("review_link_candidates")).candidates.map((c) => [c.name, c.comments]),
      [["paper.docx", 1]],
    );

    const joined = await supp.ask("review_link_join", { main: paper });
    assert.equal(joined.ok, true, joined.error);
    assert.deepEqual(
      joined.revision.comments.map((c) => c.ref),
      ["R1.1"],
    );
    await until(() => main.of("review_event").length, "the event in the manuscript's pane");
    await sleep(50);
    assert.deepEqual(
      main.of("review_event").map((e) => [e.event, e.doc]),
      [["linked_changed", "SI.docx"]],
    );
    assert.deepEqual(supp.of("review_event"), [], "the pane that joined gets its answer, no event");
    assert.deepEqual(third.of("review_event"), [], "a pane of another paper gets nothing");
    // The manuscript's pane loads again and sees the linked file.
    assert.deepEqual(
      (await main.ask("review_load")).linked.parts.map((p) => p.name),
      ["SI.docx"],
    );

    // A tracked edit of the manuscript: its call goes to the manuscript's
    // pane only, and the supplement's pane cannot answer it.
    const running = d.edit(
      d.keys.get(paper),
      { find: "the substrate", replace: "the medium" },
      "R1.1",
    );
    const call = await until(() => main.of("tool_call")[0], "tool_call in the manuscript's pane");
    assert.deepEqual(supp.of("tool_call"), []);
    const applied = {
      ok: true,
      word_comment_id: 12,
      paragraph_index: 1,
      before: BEFORE,
      after: AFTER,
    };
    supp.send({ type: "tool_result", id: call.id, ok: true, result: applied });
    await sleep(50);
    assert.equal(
      (await revisionOf(paper)).comments[0].changes.length,
      0,
      "not taken from the other pane",
    );
    main.send({ type: "tool_result", id: call.id, ok: true, result: applied });
    const r = await running;
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.equal((await revisionOf(paper)).comments[0].changes.length, 1);
    assert.equal(
      (await revisionOf(si)).comments[0].changes.length,
      0,
      "the change is the manuscript's",
    );
    // The supplement's copy still follows the manuscript after it.
    const again = await supp.ask("review_load");
    assert.deepEqual(
      [again.revision.comments[0].changes.length, again.linked.part_of.name],
      [0, "paper.docx"],
    );

    // A change of the comment list in the manuscript reaches the supplement's pane.
    const more = await main.ask("review_save_comments", {
      comments: [
        { ref: "R1.1", reviewer: 1, text: "Name the medium." },
        { ref: "R1.2", reviewer: 1, text: "Add a scale bar." },
      ],
    });
    assert.equal(more.ok, true, more.error);
    await until(() => supp.of("review_event").length, "the event in the supplement's pane");
    assert.deepEqual(
      supp.of("review_event").map((e) => [e.event, e.doc]),
      [["linked_changed", "paper.docx"]],
    );
    assert.deepEqual(
      (await supp.ask("review_load")).revision.comments.map((c) => c.ref),
      ["R1.1", "R1.2"],
    );
  } finally {
    for (const p of [main, supp, third]) p?.ws.terminate();
    await d.bridge.close();
    await rm(dir, { recursive: true, force: true });
  }
});
