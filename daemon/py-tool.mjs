// Run a Python tool of this repository. Never throws: the result says what
// happened.
//
//   { ok, code, json, stderr, error }
//   ok     exit code 0
//   code   exit code, or null when the interpreter did not start or was stopped
//   json   the last line of stdout parsed as JSON (the tools print one JSON
//          line), undefined when it is not JSON
//   error  null on exit 0; else why: interpreter not found, timed out, or the
//          last line of stderr

import { execFile } from "node:child_process";

export function runPy(
  script,
  args = [],
  { python = process.env.CC_OFFICE_PYTHON || "python3", timeoutMs = 120_000 } = {},
) {
  return new Promise((resolve) => {
    execFile(
      python,
      [script, ...args],
      { timeout: timeoutMs, maxBuffer: 64 << 20, encoding: "utf8" },
      (err, stdout, stderr) => {
        if (err?.code === "ENOENT")
          return resolve({ ok: false, code: null, error: `${python} not found` });
        let json;
        try {
          json = JSON.parse(String(stdout).trim().split("\n").pop());
        } catch {
          json = undefined;
        }
        const code = err ? (typeof err.code === "number" ? err.code : null) : 0;
        const lastErr = String(stderr).trim().split("\n").pop();
        resolve({
          ok: code === 0,
          code,
          json,
          stderr: String(stderr).slice(-2000),
          error: !err
            ? null
            : err.killed
              ? `timed out after ${timeoutMs} ms`
              : lastErr || err.message,
        });
      },
    );
  });
}
