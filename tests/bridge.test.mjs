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
