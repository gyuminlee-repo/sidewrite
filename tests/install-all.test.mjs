import { test } from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { unitName } from "../scripts/install/unit.mjs";
import { fakeBin } from "./fixtures/fake-bin.mjs";

const REPO = join(dirname(fileURLToPath(import.meta.url)), "..");
const HERE = join(REPO, "scripts/install");

// Stand-ins for everything the two scripts and the doctor call. Nothing of
// the real system, of Windows or of the registry is touched.
const FAKES = {
  ps: "echo systemd",
  systemctl: 'case "$*" in *is-active*) echo active;; esac',
  loginctl: 'case "$*" in *show-user*) echo yes;; esac',
  curl: "echo '{\"ok\":true}'",
  sleep: "true",
  ss: "true",
  wslpath: 'echo "W:$2"',
  "powershell.exe":
    'case "$*" in *-Command*) echo \'{"reach":true,"registered":true,"manifest":true,"task":true}\';; *) printf "done\\r\\n";; esac',
};
const run = (fake, script, env = {}) =>
  spawnSync("bash", [join(HERE, script)], {
    env: fake.env({ WSL_DISTRO_NAME: "Ubuntu", ...env }),
    encoding: "utf8",
  });
const ps1 = (name, ...args) =>
  ["powershell.exe -NoProfile -ExecutionPolicy Bypass -File", `W:${HERE}/${name}`, ...args].join(
    " ",
  );

test("install.sh: the service, the add-in, the logon task, then the doctor; its exit is the doctor's count", async () => {
  const fake = await fakeBin(FAKES);
  try {
    const r = run(fake, "install.sh");
    const calls = await fake.calls();
    const svc = `${unitName}.service`;
    // The order of the steps, up to the doctor.
    const steps = calls.filter(
      (c) => /^systemctl --user (enable|restart|daemon-reload)/.test(c) || /-File/.test(c),
    );
    assert.deepEqual(steps, [
      "systemctl --user daemon-reload",
      `systemctl --user enable ${svc}`,
      `systemctl --user restart ${svc}`,
      ps1("register-addin.ps1", "-Manifest", `W:${REPO}/manifests/word.xml`),
      ps1("autostart.ps1", "-Distro", "Ubuntu"),
    ]);
    // Carriage returns of PowerShell are taken out of what is shown.
    assert.doesNotMatch(r.stdout, /\r/);
    assert.match(r.stdout, /== 1\/4[\s\S]*== 2\/4[\s\S]*== 3\/4[\s\S]*== 4\/4/);
    // The doctor ran last and its count is the exit code.
    assert.ok(calls.some((c) => c.startsWith("powershell.exe -NoProfile -Command ")));
    const last = r.stdout.trim().split("\n").at(-1);
    const count = last === "Everything is in place." ? 0 : Number(/^(\d+) problem/.exec(last)?.[1]);
    assert.equal(r.status, count, r.stdout + r.stderr);
  } finally {
    await fake.cleanup();
  }
});

test("install.sh: a step that fails stops the install; outside WSL the Windows steps are skipped", async () => {
  const fake = await fakeBin({ ...FAKES, ps: "echo init" });
  try {
    let r = run(fake, "install.sh");
    assert.equal(r.status, 2);
    assert.ok(!(await fake.calls()).some((c) => c.startsWith("powershell.exe")));
  } finally {
    await fake.cleanup();
  }
  const linux = await fakeBin(FAKES);
  try {
    const r = run(linux, "install.sh", { WSL_DISTRO_NAME: "" });
    assert.match(r.stdout, /2\/4, 3\/4 skipped/);
    assert.ok(!(await linux.calls()).some((c) => c.startsWith("powershell.exe")));
  } finally {
    await linux.cleanup();
  }
});

test("uninstall.sh: the logon task, the registration, the service, in that order", async () => {
  const fake = await fakeBin(FAKES);
  try {
    const r = run(fake, "uninstall.sh");
    assert.equal(r.status, 0, r.stderr);
    const svc = `${unitName}.service`;
    assert.deepEqual(await fake.calls(), [
      `wslpath -w ${HERE}/autostart.ps1`,
      ps1("autostart.ps1", "-Remove"),
      `wslpath -w ${REPO}/manifests/word.xml`,
      `wslpath -w ${HERE}/register-addin.ps1`,
      ps1("register-addin.ps1", "-Manifest", `W:${REPO}/manifests/word.xml`, "-Unregister"),
      `systemctl --user disable --now ${svc}`,
      "systemctl --user daemon-reload",
    ]);
  } finally {
    await fake.cleanup();
  }
});
