// Says which link between Word and the daemon is broken and how to mend it.
//
//   node scripts/install/doctor.mjs
//
// Reads only: node, the stored Claude login (that the file is there, never
// what is in it), the systemd user unit, the two ports, /health, and from
// Windows whether it reaches the port, whether the add-in is registered and
// whether the logon task is there. The exit code is the number of problems.
//
// A command that is missing or fails is itself a fact (null, false): the
// judge names it as the broken link, so nothing here stops on an error.

import { execFileSync } from "node:child_process";
import { existsSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { judgeDoctor, failures, report } from "./doctor-judge.mjs";
import { HEALTH_URL, unitName } from "./unit.mjs";
import { DEVELOPER_KEY, TASK_NAME, registerPlan } from "./win-args.mjs";

const PORTS = { ws: 47833, http: 47834 };

// One PowerShell command, on one line, that prints the four Windows facts as
// JSON. It connects to the port, reads one registry key, tests one file and
// looks up one task; it changes nothing.
export function windowsProbe() {
  const key = DEVELOPER_KEY.replace(/^HKCU\\/, "HKCU:\\");
  const rel = registerPlan({ appData: "" }).path.replace(/^\\/, "");
  return [
    `$p = Join-Path $env:APPDATA '${rel}'`,
    `$k = Get-Item -LiteralPath '${key}' -ErrorAction SilentlyContinue`,
    `$c = New-Object Net.Sockets.TcpClient`,
    `$reach = $false`,
    `try { $reach = $c.ConnectAsync('127.0.0.1', ${PORTS.http}).Wait(2000) } catch { $reach = $false }`,
    `$c.Close()`,
    `$t = Get-ScheduledTask -TaskName '${TASK_NAME}' -ErrorAction SilentlyContinue`,
    `@{ reach = [bool]$reach; registered = [bool]($k -and ($null -ne $k.GetValue($p, $null))); manifest = [bool](Test-Path -LiteralPath $p); task = [bool]$t } | ConvertTo-Json -Compress`,
  ].join("; ");
}

const realDeps = {
  run: (cmd, args) => execFileSync(cmd, args, { encoding: "utf8", timeout: 30_000, stdio: "pipe" }),
  fetchJson: async (url) => {
    const r = await fetch(url, { signal: AbortSignal.timeout(3000) });
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    return r.json();
  },
  exists: existsSync,
  env: process.env,
  // null: ask the node on PATH (the one the unit is rendered with).
  nodeVersion: null,
};

export async function collectFacts(deps = realDeps) {
  const { run, fetchJson, exists, env } = deps;
  // The trimmed output of a command, also of one that exits non-zero
  // (`systemctl is-active` prints "inactive" or "failed" and exits 3); null
  // when the command is missing or printed nothing.
  const clean = (s) =>
    String(s ?? "")
      .replace(/\r/g, "")
      .trim();
  const out = (cmd, args) => {
    try {
      return clean(run(cmd, args));
    } catch (e) {
      return clean(e?.stdout) || null;
    }
  };
  // The four Windows facts; null when Windows gave no answer to read.
  const askWindows = () => {
    try {
      const w = JSON.parse(out("powershell.exe", ["-NoProfile", "-Command", windowsProbe()]));
      return {
        reach: !!w.reach,
        registered: !!w.registered,
        manifest: !!w.manifest,
        task: !!w.task,
      };
    } catch {
      return null;
    }
  };
  const listening = out("ss", ["-ltn"]) ?? "";
  const open = (port) => new RegExp(`127\\.0\\.0\\.1:${port}\\s`).test(listening);
  return {
    node: deps.nodeVersion ?? out("node", ["--version"]),
    // Where Claude Code keeps the login on Linux. Only that it is there.
    credentials: exists(join(env.HOME ?? "", ".claude", ".credentials.json")),
    systemd: out("ps", ["-p", "1", "-o", "comm="]) === "systemd",
    // is-active prints the state and exits non-zero unless it is active.
    unit: out("systemctl", ["--user", "is-active", `${unitName}.service`]),
    linger: out("loginctl", ["show-user", env.USER ?? "", "-p", "Linger", "--value"]) === "yes",
    ports: { ws: open(PORTS.ws), http: open(PORTS.http) },
    health: await fetchJson(HEALTH_URL).catch(() => null),
    windows: env.WSL_DISTRO_NAME ? askWindows() : null,
  };
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const items = judgeDoctor(await collectFacts());
  console.log(report(items));
  process.exit(failures(items));
}
