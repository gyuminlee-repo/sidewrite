// Unit test for daemon/bridge.mjs network binding.
//
// Regression guard for the WS loopback-bind security fix: the `ws`
// library binds 0.0.0.0 when constructed with only `port`, exposing the
// agent bridge to the local network. createBridge must pass
// host: "127.0.0.1" so the bridge is loopback-only.

import { test } from "node:test";
import assert from "node:assert/strict";

import { WebSocket } from "ws";
import { createBridge } from "../daemon/bridge.mjs";

// Wait for the ws server to finish binding (it binds asynchronously).
async function waitForListening(bridge, timeoutMs = 3000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (bridge.address()) return;
    await new Promise((r) => setTimeout(r, 20));
  }
  throw new Error("bridge did not start listening within timeout");
}

test("WS bridge binds to loopback (127.0.0.1), not all interfaces", async () => {
  // port: 0 → ephemeral free port, so the test never collides with a
  // running daemon or another test.
  const bridge = createBridge({ port: 0, token: "test-token", allowedOrigins: [] });
  try {
    await waitForListening(bridge);
    const addr = bridge.address();
    assert.ok(addr && typeof addr === "object", "address() should return the bound socket info");
    assert.equal(addr.address, "127.0.0.1", "WS bridge must bind loopback, not 0.0.0.0");
  } finally {
    await bridge.close();
  }
});

test("createBridge requires a token", () => {
  assert.throws(() => createBridge({ port: 0 }), /requires a token/);
});

test("pushUserMessage queues a daemon-originated turn with the pane context", async () => {
  const bridge = createBridge({ port: 0, token: "t", allowedOrigins: [] });
  try {
    await waitForListening(bridge);
    const key = "word";
    bridge.pushUserMessage("[Review turn R1.1]", key);
    const m = await bridge.nextUserMessage(key);
    assert.equal(m.text, "[Review turn R1.1]");
    assert.equal(m.context.host, "word");
  } finally {
    await bridge.close();
  }
});

// A pane that says hello and then never answers tool calls; returns the pane
// key the bridge bound it to.
async function silentPane(bridge, keys) {
  const ws = new WebSocket(`ws://127.0.0.1:${bridge.address().port}`);
  await new Promise((r, j) => (ws.once("open", r), ws.once("error", j)));
  ws.send(JSON.stringify({ type: "hello", token: "t", host: "word", active_doc: "/d/a.docx" }));
  const deadline = Date.now() + 3000;
  while (!keys.length && Date.now() < deadline) await new Promise((r) => setTimeout(r, 10));
  return ws;
}

// The daemon cannot know whether a write it got no answer for was applied:
// office_tracked_edit tells the model so (OutcomeUnknown) by these codes.
test("a tool call with no answer rejects with code ToolTimeout or TaskpaneDisconnected", async () => {
  const keys = [];
  const bridge = createBridge({
    port: 0,
    token: "t",
    allowedOrigins: [],
    toolTimeoutMs: 50,
    onHello: (key) => keys.push(key),
  });
  let ws = null;
  try {
    await waitForListening(bridge);
    ws = await silentPane(bridge, keys);
    await assert.rejects(bridge.callTaskpaneTool("office_tracked_edit", {}, keys[0]), (e) => {
      assert.equal(e.code, "ToolTimeout");
      return true;
    });
    ws.once("message", () => ws.close());
    await assert.rejects(bridge.callTaskpaneTool("office_tracked_edit", {}, keys[0]), (e) => {
      assert.equal(e.code, "TaskpaneDisconnected");
      return true;
    });
  } finally {
    ws?.terminate();
    await bridge.close();
  }
});

// A pane that captures tool_call frames and answers only when told to.
async function recordingPane(bridge, keys, frames) {
  const ws = new WebSocket(`ws://127.0.0.1:${bridge.address().port}`);
  await new Promise((r, j) => (ws.once("open", r), ws.once("error", j)));
  ws.on("message", (raw) => {
    const m = JSON.parse(raw.toString());
    if (m.type === "tool_call") frames.push(m);
  });
  const before = keys.length;
  ws.send(JSON.stringify({ type: "hello", token: "t", host: "word", active_doc: "/d/a.docx" }));
  const deadline = Date.now() + 3000;
  while (keys.length === before && Date.now() < deadline)
    await new Promise((r) => setTimeout(r, 10));
  return ws;
}

const until = async (cond, ms = 1000) => {
  const deadline = Date.now() + ms;
  while (!cond() && Date.now() < deadline) await new Promise((r) => setTimeout(r, 10));
};

// A tool_result that arrives after its call timed out may still be an edit
// Word applied. For office_tracked_edit the bridge hands it to the daemon
// (onLateToolResult) instead of dropping it; other tools keep the old drop.
test("a late office_tracked_edit result goes to onLateToolResult; other tools and expired ids do not", async () => {
  const keys = [];
  const late = [];
  const frames = [];
  const bridge = createBridge({
    port: 0,
    token: "t",
    allowedOrigins: [],
    toolTimeoutMs: 40,
    lateResultTtlMs: 300,
    onHello: (key) => keys.push(key),
    onLateToolResult: (e) => late.push(e),
  });
  let ws = null;
  try {
    await waitForListening(bridge);
    ws = await recordingPane(bridge, keys, frames);
    const args = { find: "a", replace: "b" };
    const err = await bridge.callTaskpaneTool("office_tracked_edit", args, keys[0]).catch((e) => e);
    assert.equal(err.code, "ToolTimeout");
    assert.equal(err.toolCallId, frames[0].id);
    const result = { ok: true, word_comment_id: 7 };
    ws.send(JSON.stringify({ type: "tool_result", id: frames[0].id, ok: true, result }));
    await until(() => late.length === 1);
    assert.equal(late.length, 1);
    assert.deepEqual(
      { id: late[0].id, name: late[0].name, key: late[0].key, ok: late[0].ok },
      { id: frames[0].id, name: "office_tracked_edit", key: keys[0], ok: true },
    );
    assert.deepEqual(late[0].args, args);
    assert.deepEqual(late[0].result, result);
    // A second copy of the same id is not handed over again.
    ws.send(JSON.stringify({ type: "tool_result", id: frames[0].id, ok: true, result }));

    // Another tool: dropped as before.
    await bridge.callTaskpaneTool("office_read_paragraphs", {}, keys[0]).catch(() => {});
    ws.send(JSON.stringify({ type: "tool_result", id: frames[1].id, ok: true, result: {} }));

    // Past the TTL: dropped.
    await bridge.callTaskpaneTool("office_tracked_edit", args, keys[0]).catch(() => {});
    await new Promise((r) => setTimeout(r, 400));
    ws.send(JSON.stringify({ type: "tool_result", id: frames[2].id, ok: true, result }));
    await new Promise((r) => setTimeout(r, 100));
    assert.equal(late.length, 1);
  } finally {
    ws?.terminate();
    await bridge.close();
  }
});

// The task pane keeps running across a WS blip and sends the answer on its
// new socket (same pane key).
test("a result for a call cut by a disconnect is handed over when it comes on the reconnected pane", async () => {
  const keys = [];
  const late = [];
  const frames = [];
  const bridge = createBridge({
    port: 0,
    token: "t",
    allowedOrigins: [],
    onHello: (key) => keys.push(key),
    onLateToolResult: (e) => late.push(e),
  });
  let ws1 = null;
  let ws2 = null;
  try {
    await waitForListening(bridge);
    ws1 = await recordingPane(bridge, keys, frames);
    const p = bridge
      .callTaskpaneTool("office_tracked_edit", { find: "a" }, keys[0])
      .catch((e) => e);
    await until(() => frames.length === 1);
    ws1.close();
    assert.equal((await p).code, "TaskpaneDisconnected");
    ws2 = await recordingPane(bridge, keys, frames);
    assert.equal(keys[1], keys[0]);
    const result = { ok: false, applied: true, code: "CommentInsertFailed" };
    ws2.send(JSON.stringify({ type: "tool_result", id: frames[0].id, ok: true, result }));
    await until(() => late.length === 1);
    assert.equal(late.length, 1);
    assert.equal(late[0].key, keys[0]);
    assert.deepEqual(late[0].result, result);
  } finally {
    ws1?.terminate();
    ws2?.terminate();
    await bridge.close();
  }
});
