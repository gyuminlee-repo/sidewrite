import { test } from "node:test";
import assert from "node:assert/strict";
import { createBridge } from "../daemon/bridge.mjs";
import { summarizeSessionJsonl, runProbe } from "../scripts/probe-hi.mjs";

const ORIGIN = "http://127.0.0.1:47834";

async function waitForListening(bridge, timeoutMs = 3000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (bridge.address()) return;
    await new Promise((r) => setTimeout(r, 20));
  }
  throw new Error("bridge did not start listening");
}

test("summarizeSessionJsonl sums the first assistant usage and counts hook lines", () => {
  const lines = [
    { type: "queue-operation", version: "2.1.141" },
    { type: "attachment", attachment: { type: "hook_success", hookEvent: "SessionStart" } },
    { type: "user", message: { role: "user", content: "hi" } },
    {
      type: "assistant",
      message: {
        model: "claude-sonnet-4-6",
        usage: { input_tokens: 3, cache_creation_input_tokens: 50128, cache_read_input_tokens: 0, output_tokens: 45 },
      },
    },
    { type: "assistant", message: { model: "x", usage: { input_tokens: 999 } } },
  ].map((o) => JSON.stringify(o));
  const s = summarizeSessionJsonl(lines.join("\n") + "\nnot json\n");
  assert.equal(s.total, 50131);
  assert.equal(s.output, 45);
  assert.equal(s.hookEvents, 1);
  assert.equal(s.version, "2.1.141");
  assert.equal(s.model, "claude-sonnet-4-6");
  assert.equal(s.hasAssistant, true);
});

test("runProbe speaks the taskpane protocol against the real bridge", async () => {
  const seen = {};
  let bridge;
  bridge = createBridge({
    port: 0,
    token: "t",
    allowedOrigins: [ORIGIN],
    onHello: (key) =>
      bridge.sendToTaskpane({ type: "transcript_replay", session_id: null, truncated: false, events: [] }, key),
    onUserMessage: (key) => {
      setImmediate(async () => {
        seen.message = await bridge.nextUserMessage(key);
        bridge.sendAssistantEvent({ event: "session_init", session_id: "s-1", model: "claude-x" }, key);
        await bridge.callTaskpaneTool("office_get_doc_info", {}, key).catch((e) => {
          seen.toolError = e.message;
        });
        bridge.sendAssistantText("hello", key);
        bridge.sendAssistantEvent({ event: "turn_complete", subtype: "success" }, key);
      });
    },
    extraHandlers: {
      set_model: async (msg, reply) => {
        seen.model = msg.model;
        reply({ type: "set_model_result", ok: true, model: msg.model, request_id: msg.request_id });
      },
    },
  });
  try {
    await waitForListening(bridge);
    const r = await runProbe({
      url: `ws://127.0.0.1:${bridge.address().port}`,
      origin: ORIGIN,
      token: "t",
      doc: "/tmp/probe-x/probe.docx",
      timeoutMs: 5000,
    });
    assert.equal(r.sessionId, "s-1");
    assert.equal(r.model, "claude-x");
    assert.equal(r.toolCalls, 1);
    assert.equal(r.text, "hello");
    assert.equal(seen.message.text, "hi");
    assert.equal(seen.message.context.activeDoc, "/tmp/probe-x/probe.docx");
    assert.equal(seen.model, "sonnet");
    assert.match(seen.toolError, /probe-hi/);
  } finally {
    await bridge.close();
  }
});

test("runProbe rejects when the token is wrong", async () => {
  const bridge = createBridge({ port: 0, token: "t", allowedOrigins: [ORIGIN] });
  try {
    await waitForListening(bridge);
    await assert.rejects(
      runProbe({ url: `ws://127.0.0.1:${bridge.address().port}`, origin: ORIGIN, token: "bad", doc: "/tmp/p/x.docx", timeoutMs: 5000 }),
      /4002/,
    );
  } finally {
    await bridge.close();
  }
});
