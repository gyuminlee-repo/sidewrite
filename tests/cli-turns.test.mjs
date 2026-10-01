// Review turns run by a CLI provider (Codex, Antigravity): a fake spawn
// plays the CLI's stdout, a fake MCP endpoint records the bindings.
import { test } from "node:test";
import assert from "node:assert/strict";
import { EventEmitter } from "node:events";
import { PassThrough } from "node:stream";
import { readFileSync } from "node:fs";
import { createCliTurns } from "../daemon/cli-turns.mjs";

const fixture = (name) =>
  readFileSync(new URL(`./fixtures/${name}`, import.meta.url), "utf8")
    .split("\n")
    .filter(Boolean);

function fakeSpawn(lines, { code = 0, hold = false, error = null } = {}) {
  const seen = { spawns: 0 };
  const fn = (cmd, args, opts) => {
    seen.spawns += 1;
    const child = new EventEmitter();
    child.stdout = new PassThrough();
    child.stderr = new PassThrough();
    child.stdin = new PassThrough();
    child.stdin.on("data", (d) => (seen.stdin = (seen.stdin ?? "") + d));
    child.stdin.on("end", () => (seen.stdinEnded = true));
    child.kill = (sig) => {
      (seen.signals ??= []).push(sig);
      child.stdout.end();
      setImmediate(() => child.emit("close", null, sig));
      return true;
    };
    Object.assign(seen, { cmd, args, opts });
    setImmediate(() => {
      if (error) return child.emit("error", error);
      for (const l of lines) child.stdout.write(l + "\n");
      if (!hold) {
        child.stdout.end();
        setImmediate(() => child.emit("close", code, null));
      }
    });
    return child;
  };
  return { fn, seen };
}

function harness(spawn, over = {}) {
  const out = [];
  const ended = [];
  const binds = [];
  const cli = [];
  const lines = [];
  const removed = [];
  // endTurn is the last thing a turn does: turnEnded() resolves when it has
  // been called (again), however long the steps before it took.
  const waiters = [];
  let endedSeen = 0;
  const turnEnded = () =>
    new Promise((resolve, reject) => {
      if (ended.length > endedSeen) return resolve(void (endedSeen = ended.length));
      const timer = setTimeout(() => reject(new Error("the turn did not end within 10 s")), 10_000);
      waiters.push(() => {
        clearTimeout(timer);
        endedSeen = ended.length;
        resolve();
      });
    });
  const turns = createCliTurns({
    spawnFn: spawn.fn,
    mcpHttp: {
      bind: (route, b) => (binds.push([route, b]), () => binds.push(["unbind", route])),
    },
    baseUrl: "http://127.0.0.1:47834",
    sendText: (t, k) => out.push(["text", t, k]),
    sendEvent: (e, k) => out.push(["event", e.event, e, k]),
    endTurn: (k, outcome) => {
      ended.push([k, outcome]);
      waiters.shift()?.();
    },
    log: (s) => lines.push(s),
    mkTurnDir: async () => "/tmp/turn",
    rmTurnDir: async (d) => removed.push(d),
    // `agy mcp add` and `agy mcp remove`, run through runCli.
    runCliFn: async (cmd, args) => (cli.push([cmd?.cmd ?? cmd, ...args]), { code: 0, stdout: "" }),
    env: { PATH: "/usr/bin", HOME: "/home/u" },
    ...over,
  });
  return { turns, out, ended, binds, cli, lines, removed, turnEnded };
}
const kinds = (h) => h.out.map((x) => x[1]);
const BYPASS = /approval|bypass|dangerously|skip-permissions|accept-edits|yolo|full-access/i;

test("codex turn: per-turn route and token by env, events mapped, ends the turn once", async () => {
  const sp = fakeSpawn([
    '{"type":"item.started","item":{"type":"mcp_tool_call","server":"office","tool":"office_tracked_edit","arguments":{"find":"a"},"status":"in_progress"}}',
    '{"type":"item.completed","item":{"type":"mcp_tool_call","server":"office","tool":"office_tracked_edit","arguments":{"find":"a"},"result":{},"error":null,"status":"completed"}}',
    '{"type":"item.completed","item":{"type":"agent_message","text":"Done."}}',
    '{"type":"turn.completed"}',
  ]);
  const h = harness(sp);
  await h.turns.start(
    "k",
    "word",
    "PROMPT",
    { id: "codex", model: "gpt-6.1-sol" },
    { ref: "R1.1" },
  );
  await h.turnEnded();
  const [route, binding] = h.binds[0];
  assert.match(route, /^[0-9a-f]{32}$/);
  assert.deepEqual([binding.key, binding.ref], ["k", "R1.1"]);
  assert.match(binding.token, /^[0-9a-f]{64}$/);
  assert.equal(sp.seen.opts.env.CC_OFFICE_MCP_TOKEN, binding.token);
  assert.equal(sp.seen.opts.env.PATH, "/usr/bin");
  assert.equal(sp.seen.opts.shell, false);
  assert.equal(sp.seen.opts.cwd, "/tmp/turn");
  assert.ok(sp.seen.args.includes(`mcp_servers.office.url="http://127.0.0.1:47834/mcp/${route}"`));
  assert.deepEqual(sp.seen.args.slice(sp.seen.args.indexOf("-s"), sp.seen.args.indexOf("-s") + 2), [
    "-s",
    "read-only",
  ]);
  assert.deepEqual(sp.seen.args.slice(sp.seen.args.indexOf("-C"), sp.seen.args.indexOf("-C") + 2), [
    "-C",
    "/tmp/turn",
  ]);
  // The token is in the child's environment only: no argument holds it.
  for (const a of sp.seen.args) {
    assert.equal(a.includes(binding.token), false);
    assert.equal(BYPASS.test(a), false, a);
  }
  assert.deepEqual([sp.seen.stdin, sp.seen.stdinEnded], ["PROMPT", true]);
  assert.deepEqual(kinds(h), ["tool_use_announce", "Done.\n", "turn_complete"]);
  assert.deepEqual(h.out[0][2], {
    event: "tool_use_announce",
    tool: "mcp__office__office_tracked_edit",
    input: { find: "a" },
  });
  assert.deepEqual(h.out.at(-1)[2], { event: "turn_complete", subtype: "success" });
  assert.deepEqual(h.ended, [["k", { ok: true }]]);
  assert.deepEqual(h.binds.at(-1), ["unbind", route]);
  assert.deepEqual(h.removed, ["/tmp/turn"]);
  assert.ok(h.lines.includes("[cli-turn] done provider=codex code=0 tools=1"));
  assert.equal(h.lines.join("\n").includes(binding.token), false, "the token is never logged");
  assert.equal(h.cli.length, 0, "Codex needs no registration");
  assert.equal(h.turns.running("k"), false);
});

test("failure: error event before turn_complete; the limit text is the outcome", async () => {
  const h = harness(
    fakeSpawn(['{"type":"turn.failed","error":{"message":"usage limit reached"}}'], { code: 1 }),
  );
  await h.turns.start("k", "word", "P", { id: "codex" }, { ref: "R1.1" });
  await h.turnEnded();
  assert.deepEqual(kinds(h), ["error", "turn_complete"]);
  assert.match(h.out[0][2].error, /usage limit/);
  assert.deepEqual(h.out[1][2], { event: "turn_complete", subtype: "error" });
  assert.deepEqual(h.ended, [["k", { ok: false, error: "usage limit reached" }]]);
});

test("the captured Codex run: the refused tool call fails the turn with Codex's own message", async () => {
  const h = harness(fakeSpawn(fixture("codex_exec_turn_approval.jsonl")));
  await h.turns.start("k", "word", "P", { id: "codex", model: "gpt-5.5" }, { ref: "R1.1" });
  await h.turnEnded();
  assert.equal(h.ended.length, 1);
  const [, outcome] = h.ended[0];
  assert.equal(outcome.ok, false);
  assert.match(
    outcome.error,
    /Codex CLI refused the office tool call: MCP tool call requires approval, but approval policy is never/,
  );
  assert.match(outcome.error, /No change was made/);
  assert.equal(kinds(h).at(-2), "error");
  assert.deepEqual(h.out.at(-1)[2], { event: "turn_complete", subtype: "error" });
});

test("a tool error that is not about approval (a blocked edit) does not fail the turn", async () => {
  const h = harness(
    fakeSpawn([
      '{"type":"item.started","item":{"type":"mcp_tool_call","server":"office","tool":"office_tracked_edit","arguments":{}}}',
      '{"type":"item.completed","item":{"type":"mcp_tool_call","server":"office","tool":"office_tracked_edit","arguments":{},"error":{"message":"Blocked by L1"},"status":"failed"}}',
      '{"type":"item.completed","item":{"type":"agent_message","text":"The edit was blocked."}}',
      '{"type":"turn.completed"}',
    ]),
  );
  await h.turns.start("k", "word", "P", { id: "codex" }, { ref: "R1.1" });
  await h.turnEnded();
  assert.deepEqual(h.ended, [["k", { ok: true }]]);
});

test("agy turn: the office server is registered for the turn and removed after it, token and all", async () => {
  const sp = fakeSpawn(fixture("agy_print_stream.jsonl"));
  const h = harness(sp);
  await h.turns.start("k", "word", "PROMPT", { id: "antigravity", model: "g" }, { ref: "R2.1" });
  await h.turnEnded();
  const [route, binding] = h.binds[0];
  assert.deepEqual(h.cli[0], [
    "agy",
    "mcp",
    "add",
    "--type",
    "http",
    "--header",
    `Authorization: Bearer ${binding.token}`,
    "office",
    `http://127.0.0.1:47834/mcp/${route}`,
  ]);
  assert.deepEqual(h.cli.at(-1), ["agy", "mcp", "remove", "office"]);
  assert.deepEqual(sp.seen.args.slice(0, 2), ["-p", "PROMPT"]);
  assert.ok(sp.seen.args.includes("--sandbox"));
  assert.ok(sp.seen.args.join(" ").includes("--output-format stream-json"));
  for (const a of sp.seen.args.slice(2)) assert.equal(BYPASS.test(a), false, a);
  assert.equal(sp.seen.args.join(" ").includes(binding.token), false);
  assert.equal("CC_OFFICE_MCP_TOKEN" in sp.seen.opts.env, false);
  assert.equal(sp.seen.opts.cwd, "/tmp/turn");
  assert.deepEqual(h.ended, [["k", { ok: true }]]);
  assert.deepEqual(kinds(h), ['{"ok":true}\n', "turn_complete"]);
  assert.equal(h.lines.join("\n").includes(binding.token), false);
});

test("the captured agy run: the denied tool call fails the turn with what agy reported", async () => {
  const h = harness(fakeSpawn(fixture("agy_turn_denied.jsonl")));
  await h.turns.start("k", "word", "P", { id: "antigravity" }, { ref: "R1.1" });
  await h.turnEnded();
  const [, outcome] = h.ended[0];
  assert.equal(outcome.ok, false);
  assert.match(outcome.error, /agy denied the office tool call \(CallMcpTool\)/);
  assert.ok(kinds(h).includes("tool_use_announce"));
  assert.deepEqual(h.cli.at(-1), ["agy", "mcp", "remove", "office"]);
});

test("agy: a failed registration ends the turn and starts nothing; one agy turn at a time", async () => {
  const sp = fakeSpawn([]);
  const h = harness(sp, {
    runCliFn: async (cmd, args) =>
      args[1] === "add" ? { code: 1, stdout: "", stderr: "config is read-only" } : { code: 0 },
  });
  await h.turns.start("k", "word", "P", { id: "antigravity" }, { ref: "R1.1" });
  await h.turnEnded();
  assert.equal(sp.seen.spawns, 0);
  assert.match(h.ended[0][1].error, /could not register the office tools with agy.*read-only/i);
  assert.deepEqual(h.binds.at(-1)[0], "unbind");

  const held = harness(fakeSpawn([], { hold: true }));
  await held.turns.start("a", "word", "P", { id: "antigravity" }, { ref: "R1.1" });
  await assert.rejects(
    held.turns.start("b", "word", "P", { id: "antigravity" }, { ref: "R1.1" }),
    /one Antigravity turn at a time/,
  );
  held.turns.stop("a");
  await held.turnEnded();
});

test("stop kills the child and reports interrupted; a second start on the pane is refused", async () => {
  const sp = fakeSpawn([], { hold: true });
  const h = harness(sp);
  await h.turns.start("k", "word", "P", { id: "codex" }, { ref: "R1.1" });
  assert.equal(h.turns.running("k"), true);
  await assert.rejects(
    h.turns.start("k", "word", "P", { id: "codex" }, { ref: "R1.1" }),
    /already running/,
  );
  assert.equal(h.turns.stop("k"), true);
  await h.turnEnded();
  assert.deepEqual(sp.seen.signals, ["SIGTERM"]);
  assert.equal(h.out.at(-1)[2].interrupted, true);
  assert.deepEqual(h.ended, [["k", { ok: false, error: "The turn was stopped." }]]);
  assert.equal(h.turns.stop("k"), false);
});

test("a CLI that is not installed, a non-zero exit and a timeout are failed turns", async () => {
  const missing = harness(
    fakeSpawn([], { error: Object.assign(new Error("spawn codex ENOENT"), { code: "ENOENT" }) }),
  );
  await missing.turns.start("k", "word", "P", { id: "codex" }, { ref: "R1.1" });
  await missing.turnEnded();
  assert.match(missing.ended[0][1].error, /Codex CLI is not installed/);
  assert.deepEqual(missing.binds.at(-1)[0], "unbind");

  const exit = harness(fakeSpawn(["not json"], { code: 3 }));
  await exit.turns.start("k", "word", "P", { id: "codex" }, { ref: "R1.1" });
  await exit.turnEnded();
  assert.match(exit.ended[0][1].error, /Codex CLI exited with code 3/);

  const sp = fakeSpawn([], { hold: true });
  const slow = harness(sp, { timeoutMs: 20 });
  await slow.turns.start("k", "word", "P", { id: "codex" }, { ref: "R1.1" });
  await slow.turnEnded();
  assert.deepEqual(sp.seen.signals, ["SIGTERM"]);
  assert.match(slow.ended[0][1].error, /timed out after 20 ms/);
});

test("only Codex and Antigravity run here; a turn needs its comment ref", async () => {
  const h = harness(fakeSpawn([]));
  await assert.rejects(
    h.turns.start("k", "word", "P", { id: "ollama" }, { ref: "R1.1" }),
    /cannot run review turns/,
  );
  await assert.rejects(h.turns.start("k", "word", "P", { id: "codex" }, null), /comment/);
  assert.equal(h.binds.length, 0);
});

test("shutdown stops every running turn and removes the agy registration", async () => {
  const h = harness(fakeSpawn([], { hold: true }));
  await h.turns.start("a", "word", "P", { id: "antigravity" }, { ref: "R1.1" });
  await h.turns.shutdown();
  await h.turnEnded();
  assert.equal(h.turns.running("a"), false);
  assert.deepEqual(h.cli.at(-1), ["agy", "mcp", "remove", "office"]);
  assert.equal(h.binds.at(-1)[0], "unbind");
});
