import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import {
  codexTurnArgs,
  codexOneshotArgs,
  parseCodexLine,
  codexOneshot,
} from "../daemon/providers/codex.mjs";

// Captured from codex-cli 0.159.2 on 2026-10-01 with a trivial prompt (ids
// and the home path replaced): a run that answered, and one that failed.
const fixture = (name) =>
  readFileSync(new URL(`./fixtures/${name}`, import.meta.url), "utf8")
    .split("\n")
    .filter(Boolean);
const OK_LINES = fixture("codex_exec_oneshot.jsonl");
const FAILED_LINES = fixture("codex_exec_failed.jsonl");

const BYPASS = /approval|bypass|dangerously|full-auto|yolo|danger-full-access/i;

test("turn args: json, read-only, ephemeral, per-turn MCP url, token by env, stdin prompt", () => {
  const a = codexTurnArgs({
    model: "gpt-6.1-sol",
    cwd: "/tmp/t1",
    mcpUrl: "http://127.0.0.1:47834/mcp/r1",
  });
  assert.deepEqual(a.slice(0, 1), ["exec"]);
  for (const f of ["--json", "--ephemeral", "--skip-git-repo-check"]) assert.ok(a.includes(f), f);
  assert.deepEqual(a.slice(a.indexOf("-s"), a.indexOf("-s") + 2), ["-s", "read-only"]);
  assert.deepEqual(a.slice(a.indexOf("-C"), a.indexOf("-C") + 2), ["-C", "/tmp/t1"]);
  assert.deepEqual(a.slice(a.indexOf("-m"), a.indexOf("-m") + 2), ["-m", "gpt-6.1-sol"]);
  assert.ok(a.includes('mcp_servers.office.url="http://127.0.0.1:47834/mcp/r1"'));
  assert.ok(a.includes('mcp_servers.office.bearer_token_env_var="CC_OFFICE_MCP_TOKEN"'));
  assert.equal(a.at(-1), "-");
  assert.equal(codexTurnArgs({ cwd: "/t", mcpUrl: "u" }).includes("-m"), false);
});

// Hard rule of this stage: the model that edits a manuscript is never run
// with a flag that turns approvals or the sandbox off.
test("no permission-bypass flag in any argument list; the user's config is kept", () => {
  for (const args of [
    codexTurnArgs({ model: "m", cwd: "/t", mcpUrl: "http://127.0.0.1:1/mcp/r" }),
    codexOneshotArgs({ model: "m", cwd: "/t" }),
    codexOneshotArgs({ cwd: "/t" }),
  ]) {
    for (const a of args) assert.equal(BYPASS.test(a), false, a);
    assert.deepEqual(args.slice(args.indexOf("-s"), args.indexOf("-s") + 2), ["-s", "read-only"]);
    // Without the user's config.toml the sign-in has no model provider and
    // the run fails ("out of credits", the failed fixture).
    assert.equal(args.includes("--ignore-user-config"), false);
  }
  const one = codexOneshotArgs({ model: "m", cwd: "/t" });
  assert.equal(one.at(-1), "-");
  assert.equal(
    one.some((a) => a.startsWith("mcp_servers.")),
    false,
    "a judgment gets no office tools",
  );
});

test("parse: the captured run that answered", () => {
  const ev = OK_LINES.map(parseCodexLine);
  assert.deepEqual(ev[0], { kind: "session", id: "00000000-0000-0000-0000-000000000000" });
  // A config warning comes as an item of type "error"; it does not fail the turn.
  assert.equal(ev[1].kind, "notice");
  assert.match(ev[1].message, /unrecognized configuration setting/);
  assert.equal(ev[2].kind, "notice");
  assert.equal(ev[3], null); // turn.started
  assert.deepEqual(ev[4], { kind: "text", text: '{"ok":true}' });
  assert.equal(ev[5].kind, "done");
  assert.equal(ev[5].usage.output_tokens, 9);
});

test("parse: the captured run that failed", () => {
  const ev = FAILED_LINES.map(parseCodexLine).filter(Boolean);
  assert.deepEqual(
    ev.map((e) => e.kind),
    ["session", "error", "error"],
  );
  assert.match(ev[2].message, /out of credits/);
  assert.equal(ev[2].limited, true);
});

test("parse: text, MCP tool, done, failure, junk", () => {
  assert.deepEqual(
    parseCodexLine('{"type":"item.completed","item":{"type":"agent_message","text":"hi"}}'),
    { kind: "text", text: "hi" },
  );
  assert.deepEqual(
    parseCodexLine(
      '{"type":"item.started","item":{"type":"mcp_tool_call","server":"office","tool":"office_find","arguments":{"query":"x"}}}',
    ),
    { kind: "tool", name: "mcp__office__office_find", input: { query: "x" } },
  );
  assert.equal(parseCodexLine('{"type":"turn.completed","usage":{"input_tokens":3}}').kind, "done");
  assert.deepEqual(
    parseCodexLine('{"type":"turn.failed","error":{"message":"You have hit your usage limit"}}'),
    { kind: "error", message: "You have hit your usage limit", limited: true },
  );
  assert.equal(parseCodexLine('{"type":"error","message":"stream closed"}').limited, false);
  assert.equal(parseCodexLine("not json"), null);
  assert.equal(parseCodexLine(""), null);
  assert.equal(parseCodexLine("null"), null);
});

test("oneshot: last agent message parsed as JSON, exit code kept on failure", async () => {
  const lines = [
    '{"type":"thread.started","thread_id":"t"}',
    '{"type":"item.completed","item":{"type":"agent_message","text":"{\\"a\\":1}"}}',
    '{"type":"turn.completed"}',
  ];
  const run = async (cmd, args, opts) => {
    run.seen = { cmd, args, ...opts };
    return { code: 0, stdout: lines.join("\n") };
  };
  const r = await codexOneshot(
    { system: "S", prompt: "P", model: "gpt-6.1-sol" },
    { runCli: run, bin: "codex", runDir: (fn) => fn("/tmp/empty") },
  );
  assert.deepEqual([r.ok, r.json], [true, { a: 1 }]);
  assert.equal(run.seen.stdin, "S\n\nP");
  assert.equal(run.seen.cmd, "codex");
  assert.deepEqual(run.seen.args, codexOneshotArgs({ model: "gpt-6.1-sol", cwd: "/tmp/empty" }));
  assert.equal(run.seen.timeoutMs, 120_000);
  const bad = await codexOneshot(
    { system: "", prompt: "" },
    {
      runCli: async () => ({ code: 2, stdout: "", stderr: "boom" }),
      runDir: (fn) => fn("/tmp/empty"),
    },
  );
  assert.match(bad.error, /codex exited 2: boom/);
});

test("oneshot on the captured runs: a config warning is not a failure; out of credits is a limit", async () => {
  const deps = (lines, code) => ({
    runCli: async () => ({ code, stdout: lines.join("\n"), stderr: "" }),
    runDir: (fn) => fn("/tmp/empty"),
  });
  const ok = await codexOneshot({ system: "S", prompt: "P" }, deps(OK_LINES, 0));
  assert.deepEqual([ok.ok, ok.json], [true, { ok: true }]);
  const failed = await codexOneshot({ system: "S", prompt: "P" }, deps(FAILED_LINES, 1));
  assert.equal(failed.ok, false);
  assert.match(failed.error, /out of credits/);
  assert.equal(failed.limited.kind, "rate_limit");
});

test("oneshot: text mode, not JSON, no answer, timeout, not installed", async () => {
  const d = (r) => ({ runCli: async () => r, runDir: (fn) => fn("/tmp/empty") });
  const say = (t) =>
    JSON.stringify({ type: "item.completed", item: { type: "agent_message", text: t } });
  const text = await codexOneshot(
    { system: "", prompt: "", expectJson: false },
    d({ code: 0, stdout: say("plain") }),
  );
  assert.deepEqual(text, { ok: true, text: "plain" });
  const notJson = await codexOneshot({ system: "", prompt: "" }, d({ code: 0, stdout: say("x") }));
  assert.deepEqual([notJson.ok, notJson.error, notJson.text], [false, "reply was not JSON", "x"]);
  const none = await codexOneshot({ system: "", prompt: "" }, d({ code: 0, stdout: "" }));
  assert.deepEqual([none.ok, none.error], [false, "codex gave no answer"]);
  const slow = await codexOneshot(
    { system: "", prompt: "", timeoutMs: 50 },
    d({ code: -1, stdout: "", stderr: "timed out after 50 ms", timedOut: true }),
  );
  assert.equal(slow.error, "timed out after 50 ms");
  const missing = await codexOneshot(
    { system: "", prompt: "" },
    d({ code: -1, stdout: "", stderr: "not installed" }),
  );
  assert.match(missing.error, /Codex CLI is not installed/);
});

// The one real run in which a tool was called (model gpt-5.5, read-only
// sandbox, no approval flag): Codex itself refused the call.
const APPROVAL_LINES = fixture("codex_exec_turn_approval.jsonl");

test("parse: the captured tool call and Codex's refusal of it", () => {
  const ev = APPROVAL_LINES.map(parseCodexLine).filter(Boolean);
  const started = ev.find((e) => e.kind === "tool");
  assert.equal(started.name, "mcp__office__office_tracked_edit");
  assert.deepEqual(started.input.comment_ref, "R1.1");
  const failed = ev.find((e) => e.kind === "tool_error");
  assert.deepEqual(failed, {
    kind: "tool_error",
    name: "mcp__office__office_tracked_edit",
    message: "MCP tool call requires approval, but approval policy is never",
    approval: true,
  });
  assert.equal(ev.at(-1).kind, "done");
});

test("parse: a finished tool call, and a tool error that is not about approval", () => {
  assert.deepEqual(
    parseCodexLine(
      '{"type":"item.completed","item":{"type":"mcp_tool_call","server":"office","tool":"office_find","arguments":{},"result":{"content":[]},"error":null,"status":"completed"}}',
    ),
    { kind: "tool_done", name: "mcp__office__office_find" },
  );
  const blocked = parseCodexLine(
    '{"type":"item.completed","item":{"type":"mcp_tool_call","server":"office","tool":"office_tracked_edit","arguments":{},"result":null,"error":{"message":"Blocked by L1"},"status":"failed"}}',
  );
  assert.deepEqual([blocked.kind, blocked.approval], ["tool_error", false]);
});

// ---- A judgment runs with Codex's tools off, and a run that shows a tool
// item anyway is ended and its reply discarded ----

const ISOLATION_OFF = [
  "plugins",
  "apps",
  "hooks",
  "shell_tool",
  "unified_exec",
  "browser_use",
  "computer_use",
  "multi_agent",
  "image_generation",
  "skill_search",
  "tool_suggest",
  "view_image",
  "memories",
  "goals",
  "sleep_tool",
  "shell_snapshot",
];

test("judgment args: every tool feature off, no exec rules, no web search", () => {
  const a = codexOneshotArgs({ model: "m", cwd: "/t" });
  const off = a.flatMap((x, i) => (x === "--disable" ? [a[i + 1]] : []));
  assert.deepEqual(off, ISOLATION_OFF);
  assert.ok(a.includes("--ignore-rules"));
  assert.deepEqual(a.slice(a.indexOf("-c"), a.indexOf("-c") + 2), ["-c", 'web_search="disabled"']);
  assert.equal(a.includes("--enable"), false);
  assert.equal(a.at(-1), "-");
  // The switches Codex has for a tool are all there: the request it builds
  // with them (read without a model call) has no shell, MCP, web or image tool.
  const seen = JSON.parse(
    readFileSync(new URL("./fixtures/codex_request_tools.json", import.meta.url), "utf8"),
  );
  assert.ok(seen.today.nested_in_exec.includes("exec_command"));
  assert.deepEqual(seen.final_set.nested_in_exec, ["apply_patch", "clock__curr_time"]);
});

test("toolEventOf: only an answer and a notice are items a judgment may have", async () => {
  const { toolEventOf } = await import("../daemon/providers/codex.mjs");
  const item = (type, extra = {}, ev = "item.started") =>
    JSON.stringify({ type: ev, item: { id: "i", type, ...extra } });
  assert.equal(toolEventOf(item("agent_message", { text: "x" }, "item.completed")), null);
  assert.equal(toolEventOf(item("error", { message: "notice" }, "item.completed")), null);
  assert.equal(toolEventOf('{"type":"turn.completed"}'), null);
  assert.equal(toolEventOf("not json"), null);
  assert.equal(
    toolEventOf(item("mcp_tool_call", { server: "office", tool: "office_find" })),
    "mcp__office__office_find",
  );
  assert.equal(toolEventOf(item("command_execution", { command: "echo hi" })), "codex_shell");
  // A type this code has never seen is a tool until it is known not to be.
  assert.equal(toolEventOf(item("web_search", { query: "q" })), "web_search");
  assert.equal(toolEventOf(item("file_change", {}, "item.updated")), "file_change");
  assert.equal(toolEventOf(item("collab_tool_call", {})), "collab_tool_call");
  // Output of the model itself, not a tool: a reasoning summary and its to-do list.
  assert.equal(toolEventOf(item("reasoning", { text: "t" }, "item.completed")), null);
  assert.equal(toolEventOf(item("todo_list", { items: [] }, "item.updated")), null);
  assert.equal(toolEventOf('{"type":"item.started","item":{}}'), "unknown item");
});

// A runner that hands the lines over one by one, as runCli does, and stops
// when kill() is called.
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

test("oneshot: a run that uses a tool is ended at that item and its reply is discarded", async () => {
  const logs = [];
  for (const [file, name] of [
    ["codex_exec_turn_approval.jsonl", "mcp__office__office_tracked_edit"],
    ["codex_exec_tool_probe.jsonl", "codex_shell"],
  ]) {
    const lines = fixture(file);
    const run = streaming(lines);
    const r = await codexOneshot(
      { system: "S", prompt: "SECRET PARAGRAPH", expectJson: false },
      { runCli: run, bin: "codex", runDir: (fn) => fn("/tmp/run"), log: (l) => logs.push(l) },
    );
    assert.deepEqual(r, {
      ok: false,
      tool_event: name,
      error: `Codex CLI used a tool (${name}) in a run that must not use any. Its reply was discarded.`,
    });
    assert.equal(run.killed, true, file);
    assert.ok(run.handed.length < lines.length, "ended before the run was over");
  }
  assert.deepEqual(logs, [
    "[providers] tool event provider=codex name=mcp__office__office_tracked_edit",
    "[providers] tool event provider=codex name=codex_shell",
  ]);

  // A runner that gives the output only at the end: the reply is discarded all the same.
  const late = await codexOneshot(
    { system: "", prompt: "", expectJson: false },
    {
      runCli: async () => ({
        code: 0,
        stdout: fixture("codex_exec_tool_probe.jsonl").join("\n"),
        stderr: "",
      }),
      runDir: (fn) => fn("/tmp/run"),
      log: () => {},
    },
  );
  assert.deepEqual([late.ok, late.tool_event], [false, "codex_shell"]);
});

test("oneshot: the captured run with the tools off answers; its directory, environment and signal", async () => {
  const run = streaming(fixture("codex_exec_isolated.jsonl"));
  const ac = new AbortController();
  const r = await codexOneshot(
    { system: "S", prompt: "P", signal: ac.signal },
    {
      runCli: run,
      bin: "codex",
      runDir: (fn) => fn("/tmp/run-1"),
      env: { PATH: "/bin", HOME: "/h", OPENAI_API_KEY: "sk", CC_OFFICE_KEY_OPENROUTER: "k" },
    },
  );
  assert.deepEqual(r.json, {
    labels: [
      { i: 1, type: "request" },
      { i: 2, type: "other" },
    ],
  });
  assert.equal(run.killed, false);
  assert.equal(run.seen.cwd, "/tmp/run-1");
  assert.deepEqual(run.seen.env, { PATH: "/bin", HOME: "/h" });
  assert.equal(run.seen.signal, ac.signal);

  const stopped = await codexOneshot(
    { system: "", prompt: "" },
    {
      runCli: async () => ({ code: -1, stdout: "", stderr: "", aborted: true }),
      runDir: (fn) => fn("/tmp/run"),
    },
  );
  assert.deepEqual(stopped, { ok: false, aborted: true, error: "stopped" });
});

// Codex review of cli-isolation (P1): a model that prints a reasoning summary
// is not using a tool. The kinds of item `codex exec --json` has were read
// from the installed binary (codex.mjs says where); the captured run is real
// (made-up puzzle, reasoning effort high, summary detailed).
test("oneshot: a reasoning summary before the answer is not a tool; a tool after it still is", async () => {
  const lines = fixture("codex_exec_reasoning.jsonl");
  assert.ok(lines.some((l) => JSON.parse(l).item?.type === "reasoning"));
  const d = (run) => ({ runCli: run, runDir: (fn) => fn("/tmp/run"), log: () => {} });
  const run = streaming(lines);
  const r = await codexOneshot({ system: "S", prompt: "P" }, d(run));
  assert.deepEqual([r.ok, r.json.pours], [true, 2]);
  assert.equal(run.killed, false);

  const reasoning =
    '{"type":"item.completed","item":{"id":"i1","type":"reasoning","text":"**x**"}}';
  const plan =
    '{"type":"item.started","item":{"id":"i2","type":"todo_list","items":[{"text":"answer","completed":false}]}}';
  const answer = '{"type":"item.completed","item":{"id":"i9","type":"agent_message","text":"{}"}}';
  const quiet = await codexOneshot(
    { system: "", prompt: "" },
    d(streaming([reasoning, plan, answer, '{"type":"turn.completed"}'])),
  );
  assert.deepEqual([quiet.ok, quiet.json], [true, {}]);

  for (const [tool, name] of [
    [
      '{"type":"item.started","item":{"id":"i3","type":"command_execution","command":"ls"}}',
      "codex_shell",
    ],
    ['{"type":"item.started","item":{"id":"i3","type":"web_search","query":"q"}}', "web_search"],
    ['{"type":"item.started","item":{"id":"i3","type":"file_change","changes":[]}}', "file_change"],
    ['{"type":"item.started","item":{"id":"i3","type":"something_new"}}', "something_new"],
  ]) {
    const after = streaming([reasoning, tool, answer]);
    const bad = await codexOneshot({ system: "", prompt: "" }, d(after));
    assert.deepEqual([bad.ok, bad.tool_event, after.killed], [false, name, true], name);
    assert.equal(after.handed.length, 2, "ended at the tool, before the answer");
  }
});

// A proposal turn gives the form of its reply as a JSON Schema: the file is
// written into the directory of that run and named with --output-schema.
test("oneshot: a schema is written into the run directory and handed to the CLI", async () => {
  const written = [];
  const run = streaming(fixture("codex_exec_isolated.jsonl"));
  let args = null;
  const schema = { type: "object", properties: {}, additionalProperties: false };
  await codexOneshot(
    { system: "S", prompt: "P", schema },
    {
      runCli: (bin, a, opts) => ((args = a), run(bin, a, opts)),
      runDir: (fn) => fn("/tmp/run-9"),
      writeFile: async (file, text) => void written.push([file, text]),
    },
  );
  assert.deepEqual(written, [["/tmp/run-9/schema.json", JSON.stringify(schema)]]);
  const at = args.indexOf("--output-schema");
  assert.equal(args[at + 1], "/tmp/run-9/schema.json");
  assert.equal(args.at(-1), "-");
  assert.deepEqual(
    args.filter((a, i) => i !== at && i !== at + 1),
    codexOneshotArgs({ cwd: "/tmp/run-9" }),
    "nothing else changes",
  );
  // Without a schema nothing is written and the arguments are as before.
  written.length = 0;
  await codexOneshot(
    { system: "S", prompt: "P" },
    {
      runCli: (bin, a, opts) => ((args = a), run(bin, a, opts)),
      runDir: (fn) => fn("/tmp/run-9"),
      writeFile: async (file, text) => void written.push([file, text]),
    },
  );
  assert.deepEqual([written.length, args.includes("--output-schema")], [0, false]);
});
