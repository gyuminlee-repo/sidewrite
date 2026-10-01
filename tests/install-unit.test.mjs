import { test } from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { readFile, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { renderUnit, unitName, unitPath } from "../scripts/install/unit.mjs";
import { fakeBin } from "./fixtures/fake-bin.mjs";

const REPO = join(dirname(fileURLToPath(import.meta.url)), "..");
const SCRIPT = join(REPO, "scripts/install/daemon-service.sh");

test("unit: restarts on failure, starts with the user manager, absolute paths, no session env", () => {
  const u = renderUnit({
    node: "/usr/local/bin/node",
    repo: "/opt/app",
    logFile: "/home/u/.claude/office-addins/daemon.log",
    extraPath: ["/home/u/.local/bin", "/usr/local/bin"],
  });
  assert.match(u, /^ExecStart=\/usr\/local\/bin\/node \/opt\/app\/daemon\/index\.mjs$/m);
  assert.match(u, /^WorkingDirectory=\/opt\/app$/m);
  assert.match(u, /^Restart=on-failure$/m);
  assert.match(u, /^RestartSec=3$/m);
  assert.match(u, /^WantedBy=default\.target$/m);
  // The start limit belongs to [Unit]; in [Service] systemd ignores it.
  const [unit, service] = [u.indexOf("[Unit]"), u.indexOf("[Service]")];
  for (const key of ["StartLimitIntervalSec=60", "StartLimitBurst=5"]) {
    const at = u.indexOf(key);
    assert.ok(unit < at && at < service, key);
  }
  // PATH: the directory of node first, the given ones, the system ones; each once.
  assert.match(u, /^Environment=PATH=\/usr\/local\/bin:\/home\/u\/\.local\/bin:\/usr\/bin:\/bin$/m);
  assert.match(u, /^StandardOutput=append:\/home\/u\/\.claude\/office-addins\/daemon\.log$/m);
  assert.match(u, /^StandardError=append:\/home\/u\/\.claude\/office-addins\/daemon\.log$/m);
  // PATH is the only variable set: nothing of a terminal or an agent session.
  assert.equal(u.match(/^Environment=/gm).length, 1);
  assert.equal(/CLAUDECODE|ELECTRON/.test(u), false);
  assert.match(unitName, /^[a-z-]+$/);
});

test("unit: every path is absolute and holds no blank; PATH has each directory once", () => {
  const ok = { node: "/n/node", repo: "/r", logFile: "/l/d.log" };
  for (const bad of [
    { node: "node" },
    { repo: "relative/repo" },
    { repo: "/has space/repo" },
    { logFile: "~/d.log" },
    { extraPath: ["bin"] },
    { extraPath: ["/a:b"] },
  ])
    assert.throws(() => renderUnit({ ...ok, ...bad }), /absolute path/, JSON.stringify(bad));
  assert.equal(
    unitPath("/n/node", ["/x", "/usr/bin", "/x", "/n"]),
    "/n:/x:/usr/bin:/usr/local/bin:/bin",
  );
});

// ---- daemon-service.sh with stand-ins for systemctl, loginctl, curl ----------------

const FAKES = {
  ps: 'echo "${FAKE_PID1:-systemd}"',
  systemctl: "true",
  loginctl: 'case "$*" in *show-user*) echo "${FAKE_LINGER:-no}";; esac',
  curl: '[ "${FAKE_HEALTH:-ok}" = ok ] && echo \'{"ok":true}\'',
  sleep: "true",
  tail: "true",
};
const run = (fake, args, env = {}) =>
  spawnSync("bash", [SCRIPT, ...args], { env: fake.env(env), encoding: "utf8" });
const unitFile = (fake) => join(fake.home, ".config/systemd/user", `${unitName}.service`);

test("install: writes the unit, enables it, turns linger on once, waits for /health", async () => {
  const fake = await fakeBin(FAKES);
  try {
    let r = run(fake, ["install"]);
    assert.equal(r.status, 0, r.stderr + r.stdout);
    const unit = await readFile(unitFile(fake), "utf8");
    assert.match(
      unit,
      new RegExp(`^ExecStart=${process.execPath} ${REPO}/daemon/index\\.mjs$`, "m"),
    );
    assert.match(unit, new RegExp(`^WorkingDirectory=${REPO}$`, "m"));
    assert.match(unit, new RegExp(`append:${fake.home}/\\.claude/office-addins/daemon\\.log`));
    assert.ok(existsSync(join(fake.home, ".claude/office-addins")), "the log folder is made");
    const svc = `${unitName}.service`;
    assert.deepEqual(await fake.calls(), [
      "ps -p 1 -o comm=",
      "systemctl --user daemon-reload",
      `systemctl --user enable ${svc}`,
      `systemctl --user restart ${svc}`,
      "loginctl show-user tester -p Linger --value",
      "loginctl enable-linger tester",
      "curl -fsS http://127.0.0.1:47834/health",
    ]);
    assert.match(r.stdout, /daemon is up/);

    // Again, everything in place: nothing is rewritten or restarted, linger is
    // not turned on twice. (enable --now starts a unit that was stopped.)
    await fake.reset();
    r = run(fake, ["install"], { FAKE_LINGER: "yes" });
    assert.equal(r.status, 0, r.stderr);
    assert.deepEqual(await fake.calls(), [
      "ps -p 1 -o comm=",
      `systemctl --user enable --now ${svc}`,
      "loginctl show-user tester -p Linger --value",
      "curl -fsS http://127.0.0.1:47834/health",
    ]);

    // uninstall: the unit off and gone; linger off because this install turned it on.
    await fake.reset();
    r = run(fake, ["uninstall"]);
    assert.equal(r.status, 0, r.stderr);
    assert.equal(existsSync(unitFile(fake)), false);
    assert.deepEqual(await fake.calls(), [
      `systemctl --user disable --now ${svc}`,
      "systemctl --user daemon-reload",
      "loginctl disable-linger tester",
    ]);
    // Nothing installed: nothing to undo, and no error.
    await fake.reset();
    r = run(fake, ["uninstall"]);
    assert.equal(r.status, 0, r.stderr);
    assert.deepEqual(await fake.calls(), [
      `systemctl --user disable --now ${svc}`,
      "systemctl --user daemon-reload",
    ]);
  } finally {
    await fake.cleanup();
  }
});

test("install: linger that was on before is left on by uninstall; a changed unit restarts the daemon", async () => {
  const fake = await fakeBin(FAKES);
  try {
    const svc = `${unitName}.service`;
    let r = run(fake, ["install"], { FAKE_LINGER: "yes" });
    assert.equal(r.status, 0, r.stderr);
    assert.ok(!(await fake.calls()).some((c) => c.includes("enable-linger")));
    // The unit on disk is not the one this checkout renders (an older install).
    await writeFile(unitFile(fake), "[Unit]\nDescription=old\n");
    await fake.reset();
    r = run(fake, ["install"], { FAKE_LINGER: "yes" });
    assert.equal(r.status, 0, r.stderr);
    assert.deepEqual((await fake.calls()).slice(0, 4), [
      "ps -p 1 -o comm=",
      "systemctl --user daemon-reload",
      `systemctl --user enable ${svc}`,
      `systemctl --user restart ${svc}`,
    ]);
    assert.doesNotMatch(await readFile(unitFile(fake), "utf8"), /Description=old/);
    await fake.reset();
    r = run(fake, ["uninstall"]);
    assert.ok(!(await fake.calls()).some((c) => c.includes("disable-linger")));
  } finally {
    await fake.cleanup();
  }
});

test("install: no systemd, and a daemon that does not come up", async () => {
  const fake = await fakeBin(FAKES);
  try {
    let r = run(fake, ["install"], { FAKE_PID1: "init" });
    assert.equal(r.status, 2);
    assert.match(r.stderr, /systemd=true/);
    assert.match(r.stderr, /wsl --shutdown/);
    assert.deepEqual(await fake.calls(), ["ps -p 1 -o comm="]);
    assert.equal(existsSync(unitFile(fake)), false);

    await fake.reset();
    r = run(fake, ["install"], { FAKE_HEALTH: "down" });
    assert.equal(r.status, 1);
    assert.match(r.stderr, /did not answer/);
    assert.match(r.stderr, /daemon-service\.sh logs/);
    assert.equal((await fake.calls()).filter((c) => c.startsWith("curl ")).length, 20);
  } finally {
    await fake.cleanup();
  }
});

test("status and logs read; an unknown word is refused", async () => {
  const fake = await fakeBin(FAKES);
  try {
    const svc = `${unitName}.service`;
    let r = run(fake, ["status"]);
    assert.equal(r.status, 0, r.stderr);
    assert.deepEqual(await fake.calls(), [
      `systemctl --user --no-pager status ${svc}`,
      "curl -fsS http://127.0.0.1:47834/health",
    ]);
    await fake.reset();
    r = run(fake, ["logs"]);
    assert.deepEqual(await fake.calls(), [
      `tail -n 100 -f ${fake.home}/.claude/office-addins/daemon.log`,
    ]);
    r = run(fake, ["start"]);
    assert.equal(r.status, 64);
    assert.match(r.stderr, /install\|uninstall\|status\|logs/);
  } finally {
    await fake.cleanup();
  }
});
