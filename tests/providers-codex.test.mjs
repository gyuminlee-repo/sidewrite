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
    { runCli: run, bin: "codex", emptyDir: async () => "/tmp/empty" },
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
      emptyDir: async () => "/tmp/empty",
    },
  );
  assert.match(bad.error, /codex exited 2: boom/);
});

test("oneshot on the captured runs: a config warning is not a failure; out of credits is a limit", async () => {
  const deps = (lines, code) => ({
    runCli: async () => ({ code, stdout: lines.join("\n"), stderr: "" }),
    emptyDir: async () => "/tmp/empty",
  });
  const ok = await codexOneshot({ system: "S", prompt: "P" }, deps(OK_LINES, 0));
  assert.deepEqual([ok.ok, ok.json], [true, { ok: true }]);
  const failed = await codexOneshot({ system: "S", prompt: "P" }, deps(FAILED_LINES, 1));
  assert.equal(failed.ok, false);
  assert.match(failed.error, /out of credits/);
  assert.equal(failed.limited.kind, "rate_limit");
});

test("oneshot: text mode, not JSON, no answer, timeout, not installed", async () => {
  const d = (r) => ({ runCli: async () => r, emptyDir: async () => "/tmp/empty" });
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
