// Run a provider CLI (codex, agy) once and collect what it printed (P14).
//
// runCli never throws and never rejects:
//   { code, stdout, stderr, timedOut }  plus  aborted: true  or  killed: true
// code is the exit code, or -1 when the process could not be started, was
// ended by a signal, or timed out. A command that is not installed gives
// stderr "not installed". No shell: the arguments reach the CLI as given, so
// a prompt passed as an argument is never read as shell syntax.
//
// A run can be ended three ways, each with SIGTERM and, when the process is
// still there after killGraceMs, SIGKILL: the timeout (timedOut), the
// caller's `signal` (aborted), and kill() from onLine (killed).
// onLine(line, kill) gets every line of stdout as it comes; after kill() no
// further line is handed over.
//
// stdin is always closed (after `stdin` is written when given): a CLI that
// reads a piped stdin would otherwise wait for it forever.

import { spawn } from "node:child_process";
import { StringDecoder } from "node:string_decoder";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { existsSync, readFileSync } from "node:fs";
import { basename, dirname, join, resolve as resolvePath, win32 } from "node:path";

const MAX_OUTPUT_BYTES = 32 * 1024 * 1024;

// A directory of its own for one run: the CLI finds nothing to read there,
// and nothing another process put into a shared place. It is removed when
// fn has ended, with whatever the run left in it (a hook of the user may
// write there).
export const RUN_DIR_PREFIX = "cc-office-run-";

// Only a directory withRunDir made: directly under the temp folder, with
// the prefix. Any other path is left alone.
export async function removeRunDir(dir) {
  const full = resolvePath(String(dir));
  if (dirname(full) !== resolvePath(tmpdir())) return;
  if (!basename(full).startsWith(RUN_DIR_PREFIX)) return;
  await rm(full, { recursive: true, force: true });
}

export async function withRunDir(fn) {
  const dir = await mkdtemp(join(tmpdir(), RUN_DIR_PREFIX));
  try {
    return await fn(dir);
  } finally {
    await removeRunDir(dir);
  }
}

// The environment for a CLI that is given text to judge: the daemon's own,
// without what a CLI and the hooks it starts have no business seeing: the
// provider keys and settings of this add-in (CC_OFFICE_*) and every
// variable named like a secret (*_KEY, *_TOKEN, *_SECRET). A CLI that signs
// in through such a variable instead of its own sign-in loses it here.
const SECRET_NAME = /^CC_OFFICE_|(_KEY|_TOKEN|_SECRET)$/i;
export function childEnv(env = process.env) {
  return Object.fromEntries(Object.entries(env).filter(([k]) => !SECRET_NAME.test(k)));
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
  {
    stdin = null,
    timeoutMs = 120_000,
    killGraceMs = 5_000,
    cwd,
    env,
    signal = null,
    onLine = null,
    spawnFn = spawn,
  } = {},
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
    let aborted = false;
    let killed = false;
    let settled = false;
    let killTimer = null;
    const keep = (list) => (chunk) => {
      size += chunk.length;
      if (size <= MAX_OUTPUT_BYTES) list.push(Buffer.from(chunk));
    };
    const end = () => {
      if (killTimer) return; // being ended already
      child.kill("SIGTERM");
      killTimer = setTimeout(() => child.kill("SIGKILL"), killGraceMs);
    };
    const timer = setTimeout(() => {
      timedOut = true;
      end();
    }, timeoutMs);
    const onAbort = () => {
      aborted = true;
      end();
    };
    const finish = (result) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      clearTimeout(killTimer);
      signal?.removeEventListener("abort", onAbort);
      resolve({ ...result, ...(aborted && { aborted }), ...(killed && { killed }) });
    };
    if (signal?.aborted) onAbort();
    else signal?.addEventListener("abort", onAbort, { once: true });
    // Lines of stdout for onLine: a chunk can end in the middle of one.
    let carry = "";
    const decoder = new StringDecoder("utf8"); // a character can be cut by a chunk too
    const kill = () => {
      killed = true;
      end();
    };
    const lines = (text, last) => {
      if (!onLine) return;
      const parts = (carry + text).split("\n");
      carry = last ? "" : parts.pop();
      for (const part of parts) {
        if (killed) return;
        const line = part.replace(/\r$/, "");
        if (line || !last) onLine(line, kill);
      }
    };
    child.stdout.on("data", (chunk) => {
      keep(out)(chunk);
      lines(decoder.write(Buffer.from(chunk)), false);
    });
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
      lines("", true); // a last line without a newline
      const stderr = Buffer.concat(err).toString("utf8");
      finish({
        code: timedOut || aborted || killed || typeof code !== "number" ? -1 : code,
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
