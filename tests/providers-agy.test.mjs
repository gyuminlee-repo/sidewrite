import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import {
  agyOneshotArgs,
  parseAgyResult,
  parseAgyLine,
  agyOneshot,
} from "../daemon/providers/agy.mjs";

// Captured from agy 1.2.14 on 2026-10-01 with a trivial prompt (conversation
// ids and the working directory replaced): --output-format json and
// --output-format stream-json.
const read = (name) => readFileSync(new URL(`./fixtures/${name}`, import.meta.url), "utf8");
const PRINT_JSON = read("agy_print_json.txt");
const STREAM_LINES = read("agy_print_stream.jsonl").split("\n").filter(Boolean);

const BYPASS = /dangerously|skip-permissions|accept-edits|bypass|yolo/i;

// The stdout of a stream-json run: init (permission mode request-review),
// then the given lines.
const INIT =
  '{"event":"init","conversation_id":"c","init":{"cwd":"/t","tools":[],"permission_mode":"request-review"}}';
const result = (o) => JSON.stringify({ event: "result", result: o });
const step = (o) => JSON.stringify({ event: "step_update", step_update: o });
const out = (...lines) => ({ code: 0, stdout: [INIT, ...lines].join("\n"), stderr: "" });

test("args: print mode, json, model, timeout, no slash commands", () => {
  const a = agyOneshotArgs({ prompt: "P", model: "gemini-3.8-flash-high", timeoutS: 120 });
  assert.deepEqual(a.slice(0, 2), ["-p", "P"]);
  // A judgment is read as a stream too, so that a tool step is seen.
  assert.ok(a.join(" ").includes("--output-format stream-json"));
  assert.ok(a.join(" ").includes("--model gemini-3.8-flash-high"));
  assert.ok(a.join(" ").includes("--print-timeout 120s"));
  assert.ok(a.includes("--disable-slash-commands"));
  assert.ok(a.includes("--sandbox"));
  assert.equal(agyOneshotArgs({ prompt: "P", timeoutS: 5 }).includes("--model"), false);
});

// Hard rule of this stage: no flag that approves tool calls or edits unseen.
test("no permission-bypass flag in any argument list", () => {
  for (const args of [agyOneshotArgs({ prompt: "P", model: "m", timeoutS: 120 })]) {
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
  // A step type not seen in a capture is passed on as it is.
  assert.deepEqual(
    parseAgyLine(
      '{"event":"step_update","step_update":{"step_index":2,"state":"ACTIVE","step_type":"planning"}}',
    ),
    [{ kind: "step", step_type: "planning", state: "ACTIVE" }],
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
    return { code: 0, stdout: STREAM_LINES.join("\n"), stderr: "" };
  };
  const r = await agyOneshot(
    { system: "S", prompt: "P", model: "gemini-3.8-flash-high", timeoutMs: 60_000 },
    { runCli: run, bin: "agy", runDir: (fn) => fn("/tmp/empty") },
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
    { runCli: run, runDir: (fn) => fn("/tmp/empty") },
  );
  assert.deepEqual(text, { ok: true, text: '{"ok":true}' });
});

test("oneshot failures: exit code, not installed, a limit, not JSON, a prompt too long for one argument", async () => {
  const d = (r) => ({ runCli: async () => r, runDir: (fn) => fn("/tmp/empty") });
  const req = { system: "", prompt: "" };
  const bad = await agyOneshot(req, d({ code: 1, stdout: "", stderr: "boom" }));
  assert.match(bad.error, /agy exited 1: boom/);
  const missing = await agyOneshot(req, d({ code: -1, stdout: "", stderr: "not installed" }));
  assert.match(missing.error, /Antigravity CLI \(agy\) is not installed/);
  const limit = await agyOneshot(req, d(out(result({ status: "ERROR", error: "quota exceeded" }))));
  assert.deepEqual([limit.ok, limit.limited.kind], [false, "rate_limit"]);
  const notJson = await agyOneshot(req, d(out(result({ status: "SUCCESS", response: "hm" }))));
  assert.deepEqual([notJson.ok, notJson.error, notJson.text], [false, "reply was not JSON", "hm"]);
  let ran = false;
  const long = await agyOneshot(
    { system: "", prompt: "x".repeat(130_000) },
    { runCli: async () => ((ran = true), { code: 0, stdout: "" }), runDir: (fn) => fn("/t") },
  );
  assert.match(long.error, /too long/);
  assert.equal(ran, false);
});

// The real run in which agy tried an office tool (print mode, --sandbox, no
// permission flag): the call was denied and never reached the daemon.
test("stream lines: the captured tool step and its denial", () => {
  const lines = read("agy_turn_denied.jsonl").split("\n").filter(Boolean);
  const ev = lines.flatMap(parseAgyLine);
  const tools = ev.filter((e) => e.kind === "tool");
  assert.deepEqual(tools, [
    { kind: "tool", name: "mcp__office__office_read_paragraphs", input: {} },
  ]);
  assert.equal(ev.filter((e) => e.kind === "tool_done").length, 1);
  const denied = ev.find((e) => e.kind === "denied");
  assert.deepEqual(denied, { kind: "denied", actions: ["CallMcpTool"] });
  assert.equal(ev.at(-1).kind, "done");
  assert.ok(
    ev
      .filter((e) => e.kind === "text")
      .map((e) => e.text)
      .join("")
      .includes("office_tracked_edit"),
    "the text deltas join to the answer",
  );
});

test("stream lines: a tool of agy's own keeps its name", () => {
  assert.deepEqual(
    parseAgyLine(
      '{"event":"step_update","step_update":{"state":"ACTIVE","step_type":"tool","tool_name":"view_file","tool_info":{"name":"view_file","parameters":{"Path":"/x"}}}}',
    ),
    [{ kind: "tool", name: "agy_view_file", input: { Path: "/x" } }],
  );
});

// ---- A judgment is watched: a tool step, a denied action or another
// permission mode ends the run and its reply is discarded ----

const streaming = (lines) => {
  const run = async (cmd, args, opts) => {
    run.seen = opts;
    run.handed = [];
    let killed = false;
    for (const line of lines) {
      if (killed) break;
      run.handed.push(line);
      opts.onLine?.(line, () => (killed = true));
    }
    run.killed = killed;
    return {
      code: killed ? -1 : 0,
      stdout: run.handed.join("\n"),
      stderr: "",
      ...(killed && { killed }),
    };
  };
  return run;
};
const USED = (name) =>
  `Antigravity CLI (agy) used a tool (${name}) in a run that must not use any. Its reply was discarded.`;

test("oneshot: the captured runs that used a tool are ended at the tool step", async () => {
  const logs = [];
  for (const [file, name] of [
    ["agy_print_tool_probe.jsonl", "agy_run_command"],
    ["agy_turn_denied.jsonl", "mcp__office__office_read_paragraphs"],
  ]) {
    const lines = read(file).split("\n").filter(Boolean);
    const run = streaming(lines);
    const r = await agyOneshot(
      { system: "S", prompt: "P", expectJson: false },
      { runCli: run, bin: "agy", runDir: (fn) => fn("/tmp/run"), log: (l) => logs.push(l) },
    );
    assert.deepEqual(r, { ok: false, tool_event: name, error: USED(name) }, file);
    assert.equal(run.killed, true);
    assert.ok(run.handed.length < lines.length, "ended before the run was over");
  }
  assert.deepEqual(logs, [
    "[providers] tool event provider=antigravity name=agy_run_command",
    "[providers] tool event provider=antigravity name=mcp__office__office_read_paragraphs",
  ]);
});

test("oneshot: only the user's input and the answer are steps a judgment may have", async () => {
  const d = (lines) => ({
    runCli: streaming(lines),
    runDir: (fn) => fn("/tmp/run"),
    log: () => {},
  });
  const req = { system: "", prompt: "", expectJson: false };
  const ok = result({ status: "SUCCESS", response: "fine" });
  const answer = step({ state: "DONE", step_type: "agent_response", text_delta: "fine" });

  const good = await agyOneshot(
    req,
    d([INIT, step({ state: "DONE", step_type: "user_input" }), answer, ok]),
  );
  assert.deepEqual(good, { ok: true, text: "fine" });

  // A step type this code has never seen is a tool until it is known not to be.
  const planning = await agyOneshot(
    req,
    d([INIT, step({ state: "ACTIVE", step_type: "planning" }), ok]),
  );
  assert.deepEqual(planning, { ok: false, tool_event: "planning", error: USED("planning") });

  // A denied action with no tool step shown.
  const denied = await agyOneshot(
    req,
    d([
      INIT,
      answer,
      result({
        status: "SUCCESS",
        response: "fine",
        denied_actions: [{ action: "read_url", display_name: "ReadUrlContent" }],
      }),
    ]),
  );
  assert.deepEqual(denied, {
    ok: false,
    tool_event: "ReadUrlContent",
    error: USED("ReadUrlContent"),
  });

  // Another permission mode: ended at the first line, before any step.
  const mode = '{"event":"init","init":{"permission_mode":"accept-edits"}}';
  const run = streaming([mode, answer, ok]);
  const other = await agyOneshot(req, { runCli: run, runDir: (fn) => fn("/t"), log: () => {} });
  assert.deepEqual(other, {
    ok: false,
    error:
      'agy started in permission mode "accept-edits", not "request-review". The run was ended and its reply discarded.',
  });
  assert.deepEqual([run.killed, run.handed.length], [true, 1]);

  // No init line: the permission mode is not known, so the reply is not taken.
  const blind = await agyOneshot(req, d([answer, ok]));
  assert.deepEqual(blind, {
    ok: false,
    error: "agy did not say which permission mode it ran in. Its reply was discarded.",
  });
});

test("oneshot: its directory, an environment without keys, and a signal that ends it", async () => {
  const run = streaming(STREAM_LINES);
  const ac = new AbortController();
  const r = await agyOneshot(
    { system: "S", prompt: "P", signal: ac.signal },
    {
      runCli: run,
      runDir: (fn) => fn("/tmp/run-7"),
      env: { PATH: "/bin", HOME: "/h", GEMINI_API_KEY: "k", CC_OFFICE_AGY_BIN: "/opt/agy" },
    },
  );
  assert.deepEqual(r.json, { ok: true });
  assert.equal(run.seen.cwd, "/tmp/run-7");
  assert.deepEqual(run.seen.env, { PATH: "/bin", HOME: "/h" });
  assert.equal(run.seen.signal, ac.signal);
  const stopped = await agyOneshot(
    { system: "", prompt: "" },
    {
      runCli: async () => ({ code: -1, stdout: "", stderr: "", aborted: true }),
      runDir: (fn) => fn("/t"),
    },
  );
  assert.deepEqual(stopped, { ok: false, aborted: true, error: "stopped" });
});
