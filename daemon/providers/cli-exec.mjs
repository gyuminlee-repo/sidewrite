// Run a provider CLI (codex, agy) once and collect what it printed (P14).
//
// runCli never throws and never rejects:
//   { code, stdout, stderr, timedOut }
// code is the exit code, or -1 when the process could not be started, was
// ended by a signal, or timed out. A command that is not installed gives
// stderr "not installed". No shell: the arguments reach the CLI as given, so
// a prompt passed as an argument is never read as shell syntax.
//
// stdin is always closed (after `stdin` is written when given): a CLI that
// reads a piped stdin would otherwise wait for it forever.

import { spawn } from "node:child_process";
import { mkdir } from "node:fs/promises";
import { tmpdir } from "node:os";
import { existsSync, readFileSync } from "node:fs";
import { join, win32 } from "node:path";

const MAX_OUTPUT_BYTES = 32 * 1024 * 1024;

// An empty directory to run a judgment in: the CLI finds nothing to read
// there and no project instructions.
export async function emptyWorkDir() {
  const dir = join(tmpdir(), "cc-office-oneshot-cwd");
  await mkdir(dir, { recursive: true });
  return dir;
}

// The command to start for a CLI name: { cmd, prefixArgs }, or
// { cmd: null, error } when it is known that it cannot be started.
//   - env[envVar] (CC_OFFICE_CODEX_BIN, CC_OFFICE_AGY_BIN) wins, as given.
//   - Not Windows: the name; the system finds it on PATH.
//   - Windows: spawn without a shell cannot start a ".cmd" or ".bat" by name
//     (npm installs CLIs as such shims), so PATH and PATHEXT are searched
//     here. An .exe or .com is used as it is. An npm shim is replaced by
//     node plus the script the shim runs. Another batch file is an error
//     that names the environment variable to set. Nothing found: the name,
//     so the spawn reports "not installed".
export function resolveCliBin(
  name,
  {
    env = process.env,
    envVar = null,
    platform = process.platform,
    node = process.execPath,
    exists = existsSync,
    readText = (p) => readFileSync(p, "utf8"),
  } = {},
) {
  const override = envVar ? String(env[envVar] ?? "").trim() : "";
  if (override) return { cmd: override, prefixArgs: [] };
  if (platform !== "win32") return { cmd: name, prefixArgs: [] };
  const get = (k) => env[Object.keys(env).find((x) => x.toUpperCase() === k) ?? k] ?? "";
  const dirs = String(get("PATH")).split(";").filter(Boolean);
  const exts = String(get("PATHEXT") || ".COM;.EXE;.BAT;.CMD")
    .split(";")
    .filter(Boolean);
  for (const dir of dirs) {
    for (const ext of exts) {
      const file = win32.join(dir, `${name}${ext.toLowerCase()}`);
      if (!exists(file)) continue;
      if (!/\.(cmd|bat)$/i.test(file)) return { cmd: file, prefixArgs: [] };
      // npm's shim: ... "%_prog%"  "%dp0%\node_modules\<pkg>\bin\<x>.js" %*
      const m = /"%(?:~)?dp0%\\([^"]+\.(?:js|cjs|mjs))"/i.exec(String(readText(file) ?? ""));
      if (m) {
        const own = win32.join(dir, "node.exe");
        return { cmd: exists(own) ? own : node, prefixArgs: [win32.join(dir, m[1])] };
      }
      return {
        cmd: null,
        prefixArgs: [],
        error: `${file} is a batch file, which cannot be started directly.${envVar ? ` Set ${envVar} to the program it runs.` : ""}`,
      };
    }
  }
  return { cmd: name, prefixArgs: [] };
}

// cmd: a command name or path, or what resolveCliBin returned.
export function runCli(
  cmd,
  args,
  { stdin = null, timeoutMs = 120_000, killGraceMs = 5_000, cwd, env, spawnFn = spawn } = {},
) {
  const bin = typeof cmd === "string" ? { cmd, prefixArgs: [] } : cmd;
  return new Promise((resolve) => {
    if (!bin?.cmd) {
      resolve({
        code: -1,
        stdout: "",
        stderr: String(bin?.error ?? "no command"),
        timedOut: false,
      });
      return;
    }
    let child;
    try {
      child = spawnFn(bin.cmd, [...(bin.prefixArgs ?? []), ...args], {
        cwd,
        env,
        shell: false,
        windowsHide: true,
        stdio: ["pipe", "pipe", "pipe"],
      });
    } catch (e) {
      resolve({ code: -1, stdout: "", stderr: String(e?.message ?? e), timedOut: false });
      return;
    }
    const out = [];
    const err = [];
    let size = 0;
    let timedOut = false;
    let settled = false;
    let killTimer = null;
    const keep = (list) => (chunk) => {
      size += chunk.length;
      if (size <= MAX_OUTPUT_BYTES) list.push(Buffer.from(chunk));
    };
    const timer = setTimeout(() => {
      timedOut = true;
      child.kill("SIGTERM");
      killTimer = setTimeout(() => child.kill("SIGKILL"), killGraceMs);
    }, timeoutMs);
    const finish = (result) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      clearTimeout(killTimer);
      resolve(result);
    };
    child.stdout.on("data", keep(out));
    child.stderr.on("data", keep(err));
    child.on("error", (e) =>
      finish({
        code: -1,
        stdout: "",
        stderr: e?.code === "ENOENT" ? "not installed" : String(e?.message ?? e),
        timedOut: false,
      }),
    );
    child.on("close", (code) => {
      const stderr = Buffer.concat(err).toString("utf8");
      finish({
        code: timedOut || typeof code !== "number" ? -1 : code,
        stdout: Buffer.concat(out).toString("utf8"),
        stderr: timedOut ? `timed out after ${timeoutMs} ms${stderr ? `: ${stderr}` : ""}` : stderr,
        timedOut,
      });
    });
    // The CLI may end before it reads its input (EPIPE): the exit says why.
    child.stdin.on("error", () => undefined);
    child.stdin.end(stdin ?? undefined);
  });
}
