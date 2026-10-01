import { test } from "node:test";
import assert from "node:assert/strict";
import { EventEmitter } from "node:events";
import { runCli } from "../daemon/providers/cli-exec.mjs";

// A child process stand-in: stdout and stderr emit "data", stdin records
// what was written, kill() records the signal.
function fakeSpawn(script) {
  const calls = [];
  const spawnFn = (cmd, args, opts) => {
    const child = new EventEmitter();
    child.stdout = new EventEmitter();
    child.stderr = new EventEmitter();
    child.written = "";
    child.stdin = {
      end: (s) => {
        child.written += s ?? "";
        child.stdinEnded = true;
      },
      on() {},
    };
    child.signals = [];
    child.kill = (sig) => {
      child.signals.push(sig);
      return true;
    };
    calls.push({ cmd, args, opts, child });
    queueMicrotask(() => script(child));
    return child;
  };
  return { spawnFn, calls };
}

test("stdin is sent and closed, output and exit code come back, no shell", async () => {
  const f = fakeSpawn((c) => {
    c.stdout.emit("data", Buffer.from("out1 "));
    c.stdout.emit("data", Buffer.from("out2"));
    c.stderr.emit("data", Buffer.from("warn"));
    c.emit("close", 3, null);
  });
  const r = await runCli("codex", ["exec", "-"], {
    stdin: "PROMPT",
    cwd: "/tmp/x",
    env: { A: "1" },
    spawnFn: f.spawnFn,
  });
  assert.deepEqual(r, { code: 3, stdout: "out1 out2", stderr: "warn", timedOut: false });
  const { opts, child, cmd, args } = f.calls[0];
  assert.deepEqual([cmd, args], ["codex", ["exec", "-"]]);
  assert.deepEqual([child.written, child.stdinEnded], ["PROMPT", true]);
  assert.equal(opts.cwd, "/tmp/x");
  assert.equal(opts.env.A, "1");
  assert.equal(opts.shell, false);
});

test("without stdin text the pipe is still closed, so a CLI that reads it does not wait", async () => {
  const f = fakeSpawn((c) => c.emit("close", 0, null));
  const r = await runCli("agy", ["-p", "x"], { spawnFn: f.spawnFn });
  assert.equal(r.code, 0);
  assert.equal(f.calls[0].child.stdinEnded, true);
});

test("timeout: SIGTERM, then SIGKILL when it does not end, and the result says so", async () => {
  const f = fakeSpawn(() => {}); // never ends by itself
  const run = runCli("codex", [], { timeoutMs: 20, killGraceMs: 20, spawnFn: f.spawnFn });
  await new Promise((r) => setTimeout(r, 30));
  const { child } = f.calls[0];
  assert.deepEqual(child.signals, ["SIGTERM"]);
  await new Promise((r) => setTimeout(r, 30));
  assert.deepEqual(child.signals, ["SIGTERM", "SIGKILL"]);
  child.emit("close", null, "SIGKILL");
  const r = await run;
  assert.deepEqual([r.code, r.timedOut], [-1, true]);
  assert.match(r.stderr, /timed out after 20 ms/);
});

test("not installed and other spawn errors are results, not throws", async () => {
  const enoent = fakeSpawn((c) =>
    c.emit("error", Object.assign(new Error("x"), { code: "ENOENT" })),
  );
  assert.deepEqual(await runCli("nope", [], { spawnFn: enoent.spawnFn }), {
    code: -1,
    stdout: "",
    stderr: "not installed",
    timedOut: false,
  });
  const eacces = fakeSpawn((c) => c.emit("error", new Error("EACCES: permission denied")));
  const r = await runCli("nope", [], { spawnFn: eacces.spawnFn });
  assert.deepEqual([r.code, r.stderr], [-1, "EACCES: permission denied"]);
  const thrower = await runCli("nope", [], {
    spawnFn: () => {
      throw new Error("bad options");
    },
  });
  assert.deepEqual([thrower.code, thrower.stderr], [-1, "bad options"]);
});

test("a real process: stdin through cat, exit code of false", async () => {
  const cat = await runCli("cat", [], { stdin: "hello" });
  assert.deepEqual([cat.code, cat.stdout], [0, "hello"]);
  const no = await runCli("false", []);
  assert.equal(no.code, 1);
  const missing = await runCli("definitely-not-a-command-xyz", []);
  assert.equal(missing.stderr, "not installed");
});
