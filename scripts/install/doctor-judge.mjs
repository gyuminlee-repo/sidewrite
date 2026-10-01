// Which link between Word and the daemon is broken, and the command that
// mends it. Pure: scripts/install/doctor.mjs collects the facts.
//
// facts
//   node         "v20.11.1" | null     the node on PATH
//   credentials  boolean               a Claude login is stored for the user
//   systemd      boolean               PID 1 is systemd
//   unit         "active" | "inactive" | "failed" | ... | null
//   linger       boolean               the user manager runs without a login
//   ports        { ws, http }          something listens on 47833 / 47834
//   health       the /health reply | null
//   windows      { reach, registered, manifest, task } | null (not WSL, or
//                Windows did not answer)
//
// judgeDoctor gives the items in the order of the chain, each
//   { name, ok: true | false | null, detail, fix: string | null }
// ok null: not judged (its precondition failed, or there was no way to ask);
// it is not counted as a problem.

const SERVICE = "bash scripts/install/daemon-service.sh";
const INSTALL = "bash scripts/install/install.sh";
// CI runs the tests on Node 20 and 22; 18 is past its end of life.
const NODE_MIN = 20;

export function judgeDoctor(facts) {
  const items = [];
  const add = (name, ok, detail, fix) => items.push({ name, ok, detail, fix: ok ? null : fix });
  const skip = (name, detail) => items.push({ name, ok: null, detail, fix: null });

  const major = Number(/^v(\d+)/.exec(facts.node ?? "")?.[1]);
  add(
    "node",
    major >= NODE_MIN,
    facts.node ?? "not found",
    `Install Node ${NODE_MIN} or newer in WSL, then: ${SERVICE} install`,
  );

  add(
    "claude login",
    facts.credentials === true,
    facts.credentials ? "a login is stored" : "no login found in ~/.claude",
    "Run `claude` in a WSL terminal and log in (/login). The daemon uses that login.",
  );

  const running = facts.unit === "active";
  add(
    "service",
    running,
    facts.systemd ? `systemd user unit: ${facts.unit ?? "unknown"}` : "no systemd",
    !facts.systemd
      ? "Add [boot] systemd=true to /etc/wsl.conf, run `wsl --shutdown` in Windows, open WSL again."
      : facts.unit === "failed"
        ? `See why it stopped: ${SERVICE} logs  (then: ${SERVICE} install)`
        : `${SERVICE} install`,
  );

  add(
    "linger",
    facts.linger === true,
    facts.linger ? "on" : "off: the daemon stops when the last WSL terminal closes",
    `${SERVICE} install`,
  );

  const listening = facts.ports?.ws === true && facts.ports?.http === true;
  add(
    "ports",
    listening,
    `47833 ${facts.ports?.ws ? "open" : "closed"}, 47834 ${facts.ports?.http ? "open" : "closed"}`,
    running ? `The service runs but does not listen: ${SERVICE} logs` : `${SERVICE} install`,
  );

  const up = facts.health?.ok === true;
  add(
    "health",
    up,
    up ? `version ${facts.health.version}, up ${facts.health.uptime_s} s` : "no answer on /health",
    listening && !running
      ? "The ports are held by another program (a daemon started by hand, or the tray app). Stop it, then: " +
          `${SERVICE} install`
      : `${SERVICE} logs`,
  );

  const win = facts.windows;
  if (!win) {
    for (const name of ["windows reach", "add-in registered", "wsl autostart"])
      skip(name, "Windows was not asked (not WSL, or powershell.exe gave no answer)");
    return items;
  }
  if (!up) skip("windows reach", "judged once the daemon answers in WSL");
  else
    add(
      "windows reach",
      win.reach === true,
      win.reach ? "Windows reaches 127.0.0.1:47834" : "open in WSL, closed from Windows",
      "In %UserProfile%\\.wslconfig under [wsl2] set localhostForwarding=true (or networkingMode=mirrored), then run `wsl --shutdown` in Windows and open WSL again.",
    );
  add(
    "add-in registered",
    win.registered === true && win.manifest === true,
    win.registered
      ? win.manifest
        ? "the manifest is registered for Word"
        : "registered, but the manifest file is gone"
      : "no value under WEF\\Developer",
    `${INSTALL}  (then restart Word)`,
  );
  add(
    "wsl autostart",
    win.task === true,
    win.task ? "the logon task is there" : "no logon task: WSL is not started with Windows",
    INSTALL,
  );
  return items;
}

export const failures = (items) => items.filter((x) => x.ok === false).length;

export function report(items) {
  const lines = [];
  for (const x of items) {
    const mark = x.ok === true ? "ok  " : x.ok === false ? "FAIL" : "skip";
    lines.push(`${mark}  ${x.name.padEnd(18)} ${x.detail}`);
    if (x.ok === false) lines.push(`      fix: ${x.fix}`);
  }
  const n = failures(items);
  lines.push(n ? `${n} problem(s).` : "Everything is in place.");
  return lines.join("\n");
}
