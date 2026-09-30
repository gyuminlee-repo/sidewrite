import { test } from "node:test";
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { mkdtemp, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { createBridge } from "../daemon/bridge.mjs";
import { summarizeSessionJsonl, runProbe, parseArgs } from "../scripts/probe-hi.mjs";

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
        usage: {
          input_tokens: 3,
          cache_creation_input_tokens: 50128,
          cache_read_input_tokens: 0,
          output_tokens: 45,
        },
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

test("summarizeSessionJsonl ignores the word hookEvent inside user text", () => {
  const lines = [
    { type: "attachment", attachment: { type: "hook_success", hookEvent: "SessionStart" } },
    { type: "user", message: { role: "user", content: 'the field is "hookEvent": "Stop"' } },
    {
      type: "assistant",
      message: { content: [{ type: "text", text: '{"hookEvent":"x"}' }], usage: {} },
    },
  ].map((o) => JSON.stringify(o));
  assert.equal(summarizeSessionJsonl(lines.join("\n")).hookEvents, 1);
});

const SCRIPT = fileURLToPath(new URL("../scripts/probe-hi.mjs", import.meta.url));

function runScript(args) {
  return new Promise((resolve) => {
    execFile(process.execPath, [SCRIPT, ...args], (err, stdout, stderr) =>
      resolve({ code: err ? err.code : 0, stdout, stderr }),
    );
  });
}

// --jsonl uses the same rule as the live path (waitForSummary): a session
// without assistant usage is a failed probe, not a total of 0.
test("--jsonl fails when the file has no assistant usage", async () => {
  const dir = await mkdtemp(join(tmpdir(), "probe-test-"));
  const empty = join(dir, "empty.jsonl");
  const userOnly = join(dir, "user-only.jsonl");
  await writeFile(empty, "");
  await writeFile(
    userOnly,
    [
      { type: "attachment", attachment: { type: "hook_success", hookEvent: "SessionStart" } },
      { type: "user", message: { role: "user", content: "hi" } },
    ]
      .map((o) => JSON.stringify(o))
      .join("\n"),
  );
  for (const file of [empty, userOnly]) {
    const r = await runScript(["--jsonl", file, "--max-total", "100", "--max-hooks", "5"]);
    assert.notEqual(r.code, 0, file);
    assert.match(r.stderr, /no assistant usage found in /);
  }
  const ok = join(dir, "ok.jsonl");
  await writeFile(
    ok,
    JSON.stringify({ type: "assistant", message: { usage: { input_tokens: 7 } } }) + "\n",
  );
  const r = await runScript(["--jsonl", ok, "--max-total", "100"]);
  assert.equal(r.code, 0);
  assert.match(r.stdout, /total=7 /);
});

test("parseArgs rejects missing values and non-numeric numbers", () => {
  assert.deepEqual(parseArgs(["--timeout", "5", "--max-hooks", "0"]), {
    text: "hi",
    model: "sonnet",
    timeout: 5,
    jsonl: null,
    maxTotal: null,
    maxHooks: 0,
  });
  assert.throws(() => parseArgs(["--timeout", "abc"]), /--timeout needs a number/);
  assert.throws(() => parseArgs(["--timeout"]), /--timeout needs a value/);
  assert.throws(() => parseArgs(["--jsonl"]), /--jsonl needs a value/);
  assert.throws(() => parseArgs(["--max-total"]), /--max-total needs a value/);
  assert.throws(() => parseArgs(["--max-hooks", "Infinity"]), /needs a number/);
  assert.throws(() => parseArgs(["--bogus"]), /unknown argument --bogus/);
});

test("runProbe sends stop_agent before closing on timeout", async () => {
  const seen = { stop: 0 };
  let bridge;
  bridge = createBridge({
    port: 0,
    token: "t",
    allowedOrigins: [ORIGIN],
    onHello: (key) =>
      bridge.sendToTaskpane(
        { type: "transcript_replay", session_id: null, truncated: false, events: [] },
        key,
      ),
    extraHandlers: {
      set_model: async (msg, reply) =>
        reply({ type: "set_model_result", ok: true, model: msg.model, request_id: msg.request_id }),
      stop_agent: async () => {
        seen.stop += 1;
      },
    },
  });
  try {
    await waitForListening(bridge);
    await assert.rejects(
      runProbe({
        url: `ws://127.0.0.1:${bridge.address().port}`,
        origin: ORIGIN,
        token: "t",
        doc: "/tmp/p/x.docx",
        timeoutMs: 300,
      }),
      /timed out/,
    );
    const deadline = Date.now() + 2000;
    while (seen.stop === 0 && Date.now() < deadline) await new Promise((r) => setTimeout(r, 20));
    assert.equal(seen.stop, 1);
  } finally {
    await bridge.close();
  }
});

test("runProbe speaks the taskpane protocol against the real bridge", async () => {
  const seen = {};
  let bridge;
  bridge = createBridge({
    port: 0,
    token: "t",
    allowedOrigins: [ORIGIN],
    onHello: (key) =>
      bridge.sendToTaskpane(
        { type: "transcript_replay", session_id: null, truncated: false, events: [] },
        key,
      ),
    onUserMessage: (key) => {
      setImmediate(async () => {
        seen.message = await bridge.nextUserMessage(key);
        bridge.sendAssistantEvent(
          { event: "session_init", session_id: "s-1", model: "claude-x" },
          key,
        );
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
      runProbe({
        url: `ws://127.0.0.1:${bridge.address().port}`,
        origin: ORIGIN,
        token: "bad",
        doc: "/tmp/p/x.docx",
        timeoutMs: 5000,
      }),
      /4002/,
    );
  } finally {
    await bridge.close();
  }
});
