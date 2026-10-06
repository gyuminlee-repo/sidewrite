// Run in a child process so an uncaught bridge exception is a test failure,
// not an exception intercepted by the test runner on the daemon's behalf.
import assert from "node:assert/strict";
import { once } from "node:events";
import { WebSocket } from "ws";
import { createBridge } from "../../daemon/bridge.mjs";

const scenario = process.argv[2];
const key = "word\0/fixture.docx";
const calls = [];
const failure = () => {
  calls.push("hook");
  if (scenario.endsWith("async")) return Promise.reject(new Error("fixture hook failure"));
  throw new Error("fixture hook failure");
};
const options = {};
if (scenario.startsWith("hello-")) options.onHello = failure;
if (scenario.startsWith("message-")) options.onUserMessage = failure;
if (scenario.startsWith("close-")) options.onClose = failure;
if (scenario === "ordering") {
  options.onUserMessage = () => {
    calls.push("hook");
    // The hook must run synchronously before this turn enters the queue.
    assert.equal(bridge.clearUserMessages(key), 0);
  };
}
const bridge = createBridge({ port: 0, token: "fixture-token", ...options });
let ws;
try {
  while (!bridge.address()) await new Promise((resolve) => setTimeout(resolve, 5));
  const connect = async () => {
    const pane = new WebSocket(`ws://127.0.0.1:${bridge.address().port}`);
    await once(pane, "open");
    return pane;
  };
  const hello = async (pane) => {
    const welcome = once(pane, "message");
    pane.send(
      JSON.stringify({
        type: "hello",
        token: "fixture-token",
        host: "word",
        active_doc: "/fixture.docx",
      }),
    );
    assert.equal(JSON.parse((await welcome)[0]).type, "welcome");
  };
  const ping = async (pane) => {
    const pong = once(pane, "message");
    pane.send(JSON.stringify({ type: "ping" }));
    assert.equal(JSON.parse((await pong)[0]).type, "pong");
  };
  ws = await connect();
  if (scenario === "invalid-preauth") {
    const closed = once(ws, "close");
    ws.send("null");
    assert.equal((await closed)[0], 4001);
    ws = await connect();
  }
  await hello(ws);
  if (scenario === "invalid-authed") {
    for (const value of [null, [], false, 1, "hello", {}, { type: null }, { type: {} }]) {
      ws.send(JSON.stringify(value));
      await ping(ws);
    }
  }
  if (scenario.startsWith("message-") || scenario === "ordering") {
    ws.send(JSON.stringify({ type: "user_message", text: "synthetic turn" }));
    await ping(ws);
    assert.equal((await bridge.nextUserMessage(key)).text, "synthetic turn");
  }
  await ping(ws);
  let pending;
  if (scenario.startsWith("close-")) {
    const toolCall = once(ws, "message");
    pending = bridge.callTaskpaneTool("office_read_paragraphs", {}, key).catch((error) => error);
    assert.equal(JSON.parse((await toolCall)[0]).type, "tool_call");
  }
  const closed = once(ws, "close");
  ws.close();
  await closed;
  if (pending) assert.equal((await pending).code, "TaskpaneDisconnected");
  // Give rejected fire-and-forget callbacks an event-loop turn to surface.
  await new Promise((resolve) => setTimeout(resolve, 10));
  if (/^(hello|message|close)-/.test(scenario) || scenario === "ordering")
    assert.deepEqual(calls, ["hook"]);
  if (scenario.startsWith("close-")) {
    ws = await connect();
    await hello(ws);
    await ping(ws);
  }
} finally {
  ws?.terminate();
  await bridge.close();
}
