import { test } from "node:test";
import assert from "node:assert/strict";
import { EventEmitter } from "node:events";
import { runCli, resolveCliBin } from "../daemon/providers/cli-exec.mjs";

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

// Windows: the daemon also runs natively (tray app). With shell:false an npm
// ".cmd" shim cannot be started by name, so the executable behind it is found.
function winFs(files) {
  const norm = (p) => p.toLowerCase();
  const map = new Map(Object.entries(files).map(([k, v]) => [norm(k), v]));
  return {
    exists: (p) => map.has(norm(p)),
    readText: (p) => map.get(norm(p)),
  };
}
const WIN_ENV = {
  PATH: "C:\\Windows\\system32;C:\\Users\\u\\AppData\\Roaming\\npm;C:\\Tools",
  PATHEXT: ".COM;.EXE;.BAT;.CMD",
};
const NPM_SHIM = [
  "@ECHO off",
  "GOTO start",
  ":find_dp0",
  "SET dp0=%~dp0",
  ":start",
  'IF EXIST "%dp0%\\node.exe" (',
  '  SET "_prog=%dp0%\\node.exe"',
  ") ELSE (",
  '  SET "_prog=node"',
  ")",
  'endLocal & goto #_undefined_# 2>NUL || title %COMSPEC% & "%_prog%"  "%dp0%\\node_modules\\@openai\\codex\\bin\\codex.js" %*',
].join("\r\n");

test("resolveCliBin: other platforms use the name; the environment variable always wins", () => {
  assert.deepEqual(resolveCliBin("codex", { platform: "linux", env: {} }), {
    cmd: "codex",
    prefixArgs: [],
  });
  assert.deepEqual(
    resolveCliBin("codex", {
      platform: "win32",
      env: { ...WIN_ENV, CC_OFFICE_CODEX_BIN: "D:\\bin\\codex.exe" },
      envVar: "CC_OFFICE_CODEX_BIN",
      ...winFs({}),
    }),
    { cmd: "D:\\bin\\codex.exe", prefixArgs: [] },
  );
  assert.deepEqual(
    resolveCliBin("agy", {
      platform: "linux",
      env: { CC_OFFICE_AGY_BIN: "/opt/agy" },
      envVar: "CC_OFFICE_AGY_BIN",
    }),
    { cmd: "/opt/agy", prefixArgs: [] },
  );
});

test("resolveCliBin on Windows: an .exe on PATH is used as it is, in PATH and PATHEXT order", () => {
  const fs = winFs({ "C:\\Tools\\agy.exe": "", "C:\\Tools\\agy.cmd": "x" });
  assert.deepEqual(resolveCliBin("agy", { platform: "win32", env: WIN_ENV, ...fs }), {
    cmd: "C:\\Tools\\agy.exe",
    prefixArgs: [],
  });
});

test("resolveCliBin on Windows: an npm .cmd shim becomes node plus the script it runs", () => {
  const dir = "C:\\Users\\u\\AppData\\Roaming\\npm";
  const fs = winFs({ [`${dir}\\codex.cmd`]: NPM_SHIM });
  assert.deepEqual(
    resolveCliBin("codex", { platform: "win32", env: WIN_ENV, node: "C:\\node\\node.exe", ...fs }),
    {
      cmd: "C:\\node\\node.exe",
      prefixArgs: [`${dir}\\node_modules\\@openai\\codex\\bin\\codex.js`],
    },
  );
  // The shim's own node.exe next to it is preferred, as the shim does.
  const own = winFs({ [`${dir}\\codex.cmd`]: NPM_SHIM, [`${dir}\\node.exe`]: "" });
  assert.equal(
    resolveCliBin("codex", { platform: "win32", env: WIN_ENV, node: "C:\\node\\node.exe", ...own })
      .cmd,
    `${dir}\\node.exe`,
  );
});

test("resolveCliBin on Windows: a batch file that is no node shim, or nothing on PATH, is said plainly", () => {
  const other = winFs({ "C:\\Tools\\codex.bat": "@echo off\r\npython %~dp0\\codex.py %*\r\n" });
  const r = resolveCliBin("codex", {
    platform: "win32",
    env: WIN_ENV,
    envVar: "CC_OFFICE_CODEX_BIN",
    ...other,
  });
  assert.equal(r.cmd, null);
  assert.match(r.error, /C:\\Tools\\codex\.bat is a batch file.*CC_OFFICE_CODEX_BIN/);
  const none = resolveCliBin("codex", { platform: "win32", env: WIN_ENV, ...winFs({}) });
  assert.deepEqual(
    none,
    { cmd: "codex", prefixArgs: [] },
    "left to spawn, which reports not installed",
  );
});

test("runCli takes a resolved command: prefix arguments come first, an unresolvable one is a result", async () => {
  const f = fakeSpawn((c) => c.emit("close", 0, null));
  await runCli({ cmd: "node.exe", prefixArgs: ["codex.js"] }, ["exec", "-"], {
    spawnFn: f.spawnFn,
  });
  assert.deepEqual([f.calls[0].cmd, f.calls[0].args], ["node.exe", ["codex.js", "exec", "-"]]);
  const bad = await runCli({ cmd: null, error: "codex.bat is a batch file" }, [], {
    spawnFn: f.spawnFn,
  });
  assert.deepEqual([bad.code, bad.stderr], [-1, "codex.bat is a batch file"]);
  assert.equal(f.calls.length, 1);
});
