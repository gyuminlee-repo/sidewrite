import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import {
  agyOneshotArgs,
  agyTurnArgs,
  parseAgyResult,
  parseAgyLine,
  agyMcpAddArgs,
  agyOneshot,
} from "../daemon/providers/agy.mjs";

// Captured from agy 1.2.14 on 2026-10-01 with a trivial prompt (conversation
// ids and the working directory replaced): --output-format json and
// --output-format stream-json.
const read = (name) => readFileSync(new URL(`./fixtures/${name}`, import.meta.url), "utf8");
const PRINT_JSON = read("agy_print_json.txt");
const STREAM_LINES = read("agy_print_stream.jsonl").split("\n").filter(Boolean);

const BYPASS = /dangerously|skip-permissions|accept-edits|bypass|yolo/i;

test("args: print mode, json, model, timeout, no slash commands; flags before name for mcp add", () => {
  const a = agyOneshotArgs({ prompt: "P", model: "gemini-3.8-flash-high", timeoutS: 120 });
  assert.deepEqual(a.slice(0, 2), ["-p", "P"]);
  assert.ok(a.join(" ").includes("--output-format json"));
  assert.ok(a.join(" ").includes("--model gemini-3.8-flash-high"));
  assert.ok(a.join(" ").includes("--print-timeout 120s"));
  assert.ok(a.includes("--disable-slash-commands"));
  assert.ok(a.includes("--sandbox"));
  assert.equal(agyOneshotArgs({ prompt: "P", timeoutS: 5 }).includes("--model"), false);
  assert.ok(
    agyTurnArgs({ prompt: "P", timeoutS: 900 }).join(" ").includes("--output-format stream-json"),
  );
  assert.deepEqual(agyMcpAddArgs({ url: "http://127.0.0.1:47834/mcp/agy", token: "T" }), [
    "mcp",
    "add",
    "--type",
    "http",
    "--header",
    "Authorization: Bearer T",
    "office",
    "http://127.0.0.1:47834/mcp/agy",
  ]);
});

// Hard rule of this stage: no flag that approves tool calls or edits unseen.
test("no permission-bypass flag in any argument list", () => {
  for (const args of [
    agyOneshotArgs({ prompt: "P", model: "m", timeoutS: 120 }),
    agyTurnArgs({ prompt: "P", model: "m", timeoutS: 900 }),
  ]) {
    for (const a of args.slice(2)) assert.equal(BYPASS.test(a), false, a);
    assert.equal(args.includes("--mode"), false);
  }
});

test("result: the captured print output, tolerant field names, a failed status, is_error", () => {
  assert.deepEqual(parseAgyResult(PRINT_JSON), { ok: true, text: '{"ok":true}' });
  assert.deepEqual(parseAgyResult('{"result":"{\\"a\\":1}"}'), { ok: true, text: '{"a":1}' });
  assert.deepEqual(parseAgyResult('{"response":"x"}'), { ok: true, text: "x" });
  assert.equal(parseAgyResult('{"result":"quota","is_error":true}').ok, false);
  const failed = parseAgyResult('{"status":"ERROR","response":"","error":"quota exceeded"}');
  assert.deepEqual(failed, { ok: false, error: "agy ended with status ERROR: quota exceeded" });
  assert.equal(parseAgyResult("garbage").ok, false);
  assert.equal(parseAgyResult('{"status":"SUCCESS"}').ok, false, "no answer text");
});

test("stream lines: the captured run", () => {
  const ev = STREAM_LINES.map(parseAgyLine);
  assert.deepEqual(ev[0], [{ kind: "session", id: "00000000-0000-0000-0000-000000000000" }]);
  assert.deepEqual(ev[1], []); // the user's own input step
  assert.deepEqual(ev[2], [{ kind: "text", text: '{"ok":true}\n' }]);
  assert.deepEqual(ev[3], [{ kind: "done", text: '{"ok":true}' }]);
});

test("stream lines: another step type, a failed result, junk", () => {
  // A tool step was not in the capture: its step_type and state are passed on
  // as they are, for the turn runner to show.
  assert.deepEqual(
    parseAgyLine(
      '{"event":"step_update","step_update":{"step_index":2,"state":"RUNNING","step_type":"call_mcp_tool"}}',
    ),
    [{ kind: "step", step_type: "call_mcp_tool", state: "RUNNING" }],
  );
  assert.deepEqual(
    parseAgyLine('{"event":"result","result":{"status":"ERROR","error":"rate limit reached"}}'),
    [{ kind: "error", message: "agy ended with status ERROR: rate limit reached", limited: true }],
  );
  assert.deepEqual(parseAgyLine("x"), []);
  assert.deepEqual(parseAgyLine('{"event":"something_new"}'), []);
  assert.deepEqual(parseAgyLine("null"), []);
});

test("oneshot: prompt as one argument, empty cwd, JSON parsed, the same result shape as Claude", async () => {
  const run = async (cmd, args, opts) => {
    run.seen = { cmd, args, ...opts };
    return { code: 0, stdout: PRINT_JSON, stderr: "" };
  };
  const r = await agyOneshot(
    { system: "S", prompt: "P", model: "gemini-3.8-flash-high", timeoutMs: 60_000 },
    { runCli: run, bin: "agy", emptyDir: async () => "/tmp/empty" },
  );
  assert.deepEqual(r, { ok: true, text: '{"ok":true}', json: { ok: true } });
  assert.deepEqual(
    run.seen.args,
    agyOneshotArgs({ prompt: "S\n\nP", model: "gemini-3.8-flash-high", timeoutS: 60 }),
  );
  assert.equal(run.seen.cwd, "/tmp/empty");
  assert.equal(run.seen.timeoutMs, 70_000, "the CLI's own timeout comes first");
  const text = await agyOneshot(
    { system: "", prompt: "", expectJson: false },
    { runCli: run, emptyDir: async () => "/tmp/empty" },
  );
  assert.deepEqual(text, { ok: true, text: '{"ok":true}' });
});

test("oneshot failures: exit code, not installed, a limit, not JSON, a prompt too long for one argument", async () => {
  const d = (r) => ({ runCli: async () => r, emptyDir: async () => "/tmp/empty" });
  const req = { system: "", prompt: "" };
  const bad = await agyOneshot(req, d({ code: 1, stdout: "", stderr: "boom" }));
  assert.match(bad.error, /agy exited 1: boom/);
  const missing = await agyOneshot(req, d({ code: -1, stdout: "", stderr: "not installed" }));
  assert.match(missing.error, /Antigravity CLI \(agy\) is not installed/);
  const limit = await agyOneshot(
    req,
    d({ code: 0, stdout: '{"status":"ERROR","error":"quota exceeded"}' }),
  );
  assert.deepEqual([limit.ok, limit.limited.kind], [false, "rate_limit"]);
  const notJson = await agyOneshot(
    req,
    d({ code: 0, stdout: '{"status":"SUCCESS","response":"hm"}' }),
  );
  assert.deepEqual([notJson.ok, notJson.error, notJson.text], [false, "reply was not JSON", "hm"]);
  let ran = false;
  const long = await agyOneshot(
    { system: "", prompt: "x".repeat(130_000) },
    { runCli: async () => ((ran = true), { code: 0, stdout: "" }), emptyDir: async () => "/t" },
  );
  assert.match(long.error, /too long/);
  assert.equal(ran, false);
});
