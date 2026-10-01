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
import { join } from "node:path";

const MAX_OUTPUT_BYTES = 32 * 1024 * 1024;

// An empty directory to run a judgment in: the CLI finds nothing to read
// there and no project instructions.
export async function emptyWorkDir() {
  const dir = join(tmpdir(), "cc-office-oneshot-cwd");
  await mkdir(dir, { recursive: true });
  return dir;
}

export function runCli(
  cmd,
  args,
  { stdin = null, timeoutMs = 120_000, killGraceMs = 5_000, cwd, env, spawnFn = spawn } = {},
) {
  return new Promise((resolve) => {
    let child;
    try {
      child = spawnFn(cmd, args, {
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
