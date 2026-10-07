import { test } from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { readFile, writeFile, mkdir, symlink } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import {
  renderUnit,
  unitName,
  unitPath,
  isLegacyUnit,
  isLegacyExecStart,
  legacyUnitName,
} from "../scripts/install/unit.mjs";
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
  systemctl: await readFile(new URL("./fixtures/fake-systemctl.sh", import.meta.url), "utf8"),
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

const legacyFile = (fake) => join(fake.home, ".config/systemd/user", `${legacyUnitName}.service`);
const legacyUnit = (fake) =>
  renderUnit({
    node: "/old/bin/node",
    repo: "/old/checkout",
    logFile: `${fake.home}/.claude/office-addins/daemon.log`,
    extraPath: ["/old/tools"],
  }).replace("Description=Sidewrite daemon", "Description=Draftspect daemon");

test("legacy ownership: only the exact generated unit, not a custom command or setting", () => {
  const text = legacyUnit({ home: "/home/u" });
  assert.equal(legacyUnitName, "draftspect-daemon");
  assert.equal(isLegacyUnit(text), true);
  for (const custom of [
    text.replace("Description=Draftspect", "Description=Custom"),
    text.replace("Restart=on-failure", "Restart=always"),
    text + "ExecStartPost=/custom/command\n",
    text.replace("/old/checkout/daemon/index.mjs", "/custom/app.mjs"),
    "[Unit]\nDescription=Draftspect daemon (bridge between Word and Claude)\n",
  ])
    assert.equal(isLegacyUnit(custom), false, custom);
});

test("legacy ownership: effective ExecStart must be the same single daemon command", () => {
  const text = legacyUnit({ home: "/home/u" });
  const effective =
    "{ path=/old/bin/node ; argv[]=/old/bin/node /old/checkout/daemon/index.mjs ; ignore_errors=no ; start_time=[n/a] ; pid=1 ; }";
  assert.equal(isLegacyExecStart(text, effective), true);
  for (const custom of [
    effective.replace("/old/checkout/daemon/index.mjs", "/custom/app.mjs"),
    effective + " " + effective,
    effective.replace("ignore_errors=no", "ignore_errors=yes"),
    "",
    "unrecognized output",
  ])
    assert.equal(isLegacyExecStart(text, custom), false);
});

test("legacy install: stop and remove the owned unit before starting Sidewrite, repeat safely", async () => {
  const fake = await fakeBin(FAKES);
  try {
    await mkdir(dirname(legacyFile(fake)), { recursive: true });
    await writeFile(legacyFile(fake), legacyUnit(fake));
    const unrelated = join(dirname(legacyFile(fake)), "my-service.service");
    await writeFile(unrelated, "custom service\n");
    const state = join(fake.home, ".claude/office-addins");
    await mkdir(state, { recursive: true });
    await writeFile(join(state, "bridge-token"), "keep-token");
    await writeFile(join(state, "linger-set-by-install"), "");
    let r = run(fake, ["install"], { FAKE_LINGER: "yes" });
    assert.equal(r.status, 0, r.stderr);
    assert.equal(existsSync(legacyFile(fake)), false);
    const calls = await fake.calls();
    const stop = calls.indexOf("systemctl --user stop draftspect-daemon.service");
    const start = calls.indexOf(`systemctl --user restart ${unitName}.service`);
    assert.ok(stop >= 0 && stop < start);
    assert.ok(
      calls.indexOf("systemctl --user disable --now draftspect-daemon.service") >
        calls.indexOf("curl -fsS http://127.0.0.1:47834/health"),
      "retire only after health",
    );
    assert.equal(await readFile(unrelated, "utf8"), "custom service\n");
    assert.equal(await readFile(join(state, "bridge-token"), "utf8"), "keep-token");
    assert.ok(existsSync(join(state, "linger-set-by-install")));
    assert.ok(!calls.some((c) => c.includes("disable-linger")));
    await fake.reset();
    r = run(fake, ["install"], { FAKE_LINGER: "yes" });
    assert.equal(r.status, 0, r.stderr);
    assert.ok(!(await fake.calls()).some((c) => c.includes("draftspect")));
  } finally {
    await fake.cleanup();
  }
});

test("legacy uninstall: handles an old-only install and leaves logs/user files alone", async () => {
  const fake = await fakeBin(FAKES);
  try {
    await mkdir(dirname(legacyFile(fake)), { recursive: true });
    await writeFile(legacyFile(fake), legacyUnit(fake));
    const log = join(fake.home, ".claude/office-addins/daemon.log");
    await mkdir(dirname(log), { recursive: true });
    await writeFile(log, "keep log\n");
    for (let i = 0; i < 2; i++) {
      const r = run(fake, ["uninstall"]);
      assert.equal(r.status, 0, r.stderr);
      assert.equal(existsSync(legacyFile(fake)), false);
      assert.equal(await readFile(log, "utf8"), "keep log\n");
    }
    assert.equal(
      (await fake.calls()).filter(
        (c) => c === "systemctl --user disable --now draftspect-daemon.service",
      ).length,
      1,
    );
  } finally {
    await fake.cleanup();
  }
});

test("legacy cleanup preserves custom units, symlinks and drop-ins", async () => {
  for (const kind of ["custom", "symlink", "drop-in"]) {
    const fake = await fakeBin(FAKES);
    try {
      const path = legacyFile(fake);
      await mkdir(dirname(path), { recursive: true });
      if (kind === "symlink") {
        await writeFile(join(fake.home, "owned-elsewhere"), legacyUnit(fake));
        await symlink(join(fake.home, "owned-elsewhere"), path);
      } else {
        await writeFile(path, legacyUnit(fake) + (kind === "custom" ? "# custom\n" : ""));
        if (kind === "drop-in") await mkdir(`${path}.d`);
      }
      const before = await readFile(path, "utf8");
      for (const command of ["install", "uninstall"]) {
        const r = run(fake, [command]);
        assert.equal(r.status, 0, r.stderr);
        assert.match(r.stderr, /preserved custom service/);
        assert.equal(await readFile(path, "utf8"), before);
      }
      assert.ok(!(await fake.calls()).some((c) => c.includes("draftspect")));
    } finally {
      await fake.cleanup();
    }
  }
});

test("legacy cleanup does not delete or start a second daemon when stopping the old one fails", async () => {
  const fake = await fakeBin(FAKES);
  try {
    await mkdir(dirname(legacyFile(fake)), { recursive: true });
    await writeFile(legacyFile(fake), legacyUnit(fake));
    const r = run(fake, ["install"], {
      FAKE_SYSTEMCTL_FAIL: "--user stop draftspect-daemon.service",
    });
    assert.notEqual(r.status, 0);
    assert.ok(existsSync(legacyFile(fake)));
    assert.equal(existsSync(unitFile(fake)), false);
    assert.ok(!(await fake.calls()).some((c) => /enable sidewrite|restart sidewrite/.test(c)));
  } finally {
    await fake.cleanup();
  }
});

test("legacy cleanup preserves effective drop-ins outside the unit directory", async () => {
  const fake = await fakeBin(FAKES);
  try {
    await mkdir(dirname(legacyFile(fake)), { recursive: true });
    const before = legacyUnit(fake);
    await writeFile(legacyFile(fake), before);
    for (const command of ["install", "uninstall"]) {
      const r = run(fake, [command], {
        FAKE_LEGACY_DROPINS: "/etc/systemd/user/draftspect-daemon.service.d/custom.conf",
      });
      assert.equal(r.status, 0, r.stderr);
      assert.match(r.stderr, /cannot rule out drop-ins/);
      assert.equal(await readFile(legacyFile(fake), "utf8"), before);
    }
    assert.ok(!(await fake.calls()).some((c) => /(?:stop|disable|start) draftspect/.test(c)));
  } finally {
    await fake.cleanup();
  }
});

test("legacy cleanup preserves another effective fragment or a stale cached command", async () => {
  for (const env of [
    { FAKE_LEGACY_FRAGMENT: "/etc/systemd/user/draftspect-daemon.service" },
    {
      FAKE_LEGACY_EXEC:
        "{ path=/custom/node ; argv[]=/custom/node /custom/app.mjs ; ignore_errors=no ; pid=1 ; }",
    },
    { FAKE_SYSTEMCTL_FAIL: "--user show draftspect-daemon.service -p ExecStart --value" },
  ]) {
    const fake = await fakeBin(FAKES);
    try {
      await mkdir(dirname(legacyFile(fake)), { recursive: true });
      const before = legacyUnit(fake);
      await writeFile(legacyFile(fake), before);
      for (const command of ["install", "uninstall"]) {
        const r = run(fake, [command], env);
        assert.equal(r.status, 0, r.stderr);
        assert.match(r.stderr, /effective unit does not match/);
        assert.equal(await readFile(legacyFile(fake), "utf8"), before);
      }
      assert.ok(!(await fake.calls()).some((c) => /(?:stop|disable|start) draftspect/.test(c)));
    } finally {
      await fake.cleanup();
    }
  }
});

test("legacy handover restores its unit and enabled/running state after new enable, start or health failure", async () => {
  for (const failure of ["enable", "restart", "health"]) {
    for (const enabled of ["enabled", "disabled"]) {
      for (const active of ["active", "inactive"]) {
        const fake = await fakeBin(FAKES);
        try {
          await mkdir(dirname(legacyFile(fake)), { recursive: true });
          const before = legacyUnit(fake);
          await writeFile(legacyFile(fake), before);
          const env = {
            FAKE_LEGACY_ENABLED: enabled,
            FAKE_LEGACY_ACTIVE: active,
            FAKE_LINGER: "yes",
          };
          if (failure === "health") env.FAKE_HEALTH = "down";
          else env.FAKE_SYSTEMCTL_FAIL = `--user ${failure} sidewrite-daemon.service`;
          const r = run(fake, ["install"], env);
          assert.notEqual(r.status, 0, `${failure}/${enabled}/${active}`);
          assert.match(r.stderr, /restoring the previous legacy service/);
          assert.equal(await readFile(legacyFile(fake), "utf8"), before);
          const calls = await fake.calls();
          assert.ok(calls.includes("systemctl --user disable --now sidewrite-daemon.service"));
          assert.ok(
            calls.includes(
              `systemctl --user ${enabled === "enabled" ? "enable" : "disable"} draftspect-daemon.service`,
            ),
          );
          assert.equal(
            calls.includes("systemctl --user start draftspect-daemon.service"),
            active === "active",
          );
          assert.ok(!calls.includes("systemctl --user disable --now draftspect-daemon.service"));
          // A retry can finish the migration after the transient failure.
          await fake.reset();
          const retry = run(fake, ["install"], { FAKE_LINGER: "yes" });
          assert.equal(retry.status, 0, retry.stderr);
          assert.equal(existsSync(legacyFile(fake)), false);
        } finally {
          await fake.cleanup();
        }
      }
    }
  }
});

test("legacy handover rejects a zero-exit stop or new start that had no effect", async () => {
  for (const noop of ["stop draftspect-daemon.service", "restart sidewrite-daemon.service"]) {
    const fake = await fakeBin(FAKES);
    try {
      await mkdir(dirname(legacyFile(fake)), { recursive: true });
      const before = legacyUnit(fake);
      await writeFile(legacyFile(fake), before);
      const r = run(fake, ["install"], {
        FAKE_SYSTEMCTL_NOOP: `--user ${noop}`,
        FAKE_LINGER: "yes",
      });
      assert.notEqual(r.status, 0);
      assert.equal(await readFile(legacyFile(fake), "utf8"), before);
      assert.equal(
        await readFile(join(fake.home, "legacy-service-state"), "utf8"),
        "active enabled\n",
      );
      if (noop.startsWith("stop")) {
        assert.match(r.stderr, /Legacy service did not stop/);
        assert.equal(existsSync(unitFile(fake)), false);
      } else {
        assert.match(r.stderr, /healthy port alone does not verify/);
        assert.equal(
          await readFile(join(fake.home, "sidewrite-service-state"), "utf8"),
          "inactive disabled\n",
        );
      }
    } finally {
      await fake.cleanup();
    }
  }
});

test("legacy rollback withholds restart when failed Sidewrite cannot be confirmed stopped", async () => {
  for (const failure of ["FAKE_SYSTEMCTL_FAIL", "FAKE_SYSTEMCTL_NOOP"]) {
    const fake = await fakeBin(FAKES);
    try {
      await mkdir(dirname(legacyFile(fake)), { recursive: true });
      const before = legacyUnit(fake);
      await writeFile(legacyFile(fake), before);
      const r = run(fake, ["install"], {
        FAKE_HEALTH: "down",
        FAKE_LINGER: "yes",
        [failure]: "--user disable --now sidewrite-daemon.service",
      });
      assert.notEqual(r.status, 0);
      assert.match(r.stderr, /legacy restart was withheld/);
      assert.match(r.stderr, /rollback could not finish/);
      assert.equal(await readFile(legacyFile(fake), "utf8"), before);
      assert.equal(
        await readFile(join(fake.home, "legacy-service-state"), "utf8"),
        "inactive enabled\n",
      );
      assert.equal(
        await readFile(join(fake.home, "sidewrite-service-state"), "utf8"),
        "active enabled\n",
      );
      assert.ok(!(await fake.calls()).includes("systemctl --user start draftspect-daemon.service"));
    } finally {
      await fake.cleanup();
    }
  }
});
