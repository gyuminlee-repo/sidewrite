import { test } from "node:test";
import assert from "node:assert/strict";
import { judgeDoctor, failures, report } from "../scripts/install/doctor-judge.mjs";
import { collectFacts, windowsProbe } from "../scripts/install/doctor.mjs";
import { DEVELOPER_KEY, TASK_NAME } from "../scripts/install/win-args.mjs";
import { unitName } from "../scripts/install/unit.mjs";

// Everything in order on a WSL machine.
const GOOD = {
  node: "v20.11.1",
  credentials: true,
  systemd: true,
  unit: "active",
  linger: true,
  ports: { ws: true, http: true },
  health: { ok: true, version: "0.1.0", uptime_s: 12, ports: { ws: 47833, http: 47834 } },
  windows: { reach: true, registered: true, manifest: true, task: true },
};
const item = (facts, name) => judgeDoctor(facts).find((x) => x.name === name);

test("judgeDoctor: every link in order", () => {
  const items = judgeDoctor(GOOD);
  assert.deepEqual(
    items.map((x) => [x.name, x.ok]),
    [
      ["node", true],
      ["claude login", true],
      ["service", true],
      ["linger", true],
      ["ports", true],
      ["health", true],
      ["windows reach", true],
      ["add-in registered", true],
      ["wsl autostart", true],
    ],
  );
  assert.ok(items.every((x) => x.fix === null));
  assert.equal(failures(items), 0);
});

test("judgeDoctor: each broken link names its fix", () => {
  let x = item({ ...GOOD, node: "v18.19.0" }, "node");
  assert.deepEqual([x.ok, /Node 20/.test(x.fix), /v18\.19\.0/.test(x.detail)], [false, true, true]);
  x = item({ ...GOOD, node: null }, "node");
  assert.equal(x.ok, false);

  x = item({ ...GOOD, credentials: false }, "claude login");
  assert.deepEqual([x.ok, /run `claude`.*log in/i.test(x.fix)], [false, true]);

  x = item({ ...GOOD, unit: "inactive" }, "service");
  assert.deepEqual([x.ok, /daemon-service\.sh install/.test(x.fix)], [false, true]);
  x = item({ ...GOOD, unit: "failed" }, "service");
  assert.match(x.fix, /daemon-service\.sh logs/);
  x = item({ ...GOOD, systemd: false, unit: null }, "service");
  assert.match(x.fix, /systemd=true/);

  x = item({ ...GOOD, linger: false }, "linger");
  assert.deepEqual([x.ok, /daemon-service\.sh install/.test(x.fix)], [false, true]);

  // Nothing listens, and the service is not running: start it.
  let down = { ...GOOD, unit: "inactive", ports: { ws: false, http: false }, health: null };
  assert.match(item(down, "ports").fix, /daemon-service\.sh install/);
  // The service runs and still nothing listens: its log says why.
  down = { ...GOOD, ports: { ws: false, http: false }, health: null };
  assert.match(item(down, "ports").fix, /daemon-service\.sh logs/);
  // The ports are taken and it is not this daemon that answers.
  x = item({ ...GOOD, unit: "inactive", health: null }, "health");
  assert.deepEqual([x.ok, /another program/.test(x.fix)], [false, true]);

  // Open in WSL, closed from Windows.
  x = item({ ...GOOD, windows: { ...GOOD.windows, reach: false } }, "windows reach");
  assert.equal(x.ok, false);
  assert.match(x.fix, /\.wslconfig/);
  assert.match(x.fix, /localhostForwarding=true/);
  assert.match(x.fix, /networkingMode=mirrored/);
  assert.match(x.fix, /wsl --shutdown/);

  x = item({ ...GOOD, windows: { ...GOOD.windows, registered: false } }, "add-in registered");
  assert.deepEqual([x.ok, /install\.sh/.test(x.fix)], [false, true]);
  x = item({ ...GOOD, windows: { ...GOOD.windows, manifest: false } }, "add-in registered");
  assert.equal(x.ok, false);
  x = item({ ...GOOD, windows: { ...GOOD.windows, task: false } }, "wsl autostart");
  assert.deepEqual([x.ok, /install\.sh/.test(x.fix)], [false, true]);
});

test("judgeDoctor: what cannot be told is skipped, not counted", () => {
  // The daemon is down: whether Windows reaches it is not a second failure.
  const down = { ...GOOD, unit: "inactive", ports: { ws: false, http: false }, health: null };
  const items = judgeDoctor({ ...down, windows: { ...GOOD.windows, reach: false } });
  assert.equal(items.find((x) => x.name === "windows reach").ok, null);
  assert.equal(failures(items), 3); // service, ports, health
  // Not in WSL, or Windows did not answer: its three items are skipped.
  const alone = judgeDoctor({ ...GOOD, windows: null });
  assert.deepEqual(
    alone.filter((x) => x.ok === null).map((x) => x.name),
    ["windows reach", "add-in registered", "wsl autostart"],
  );
  assert.equal(failures(alone), 0);
});

test("report: one line per item, the fix under a failure, the count last", () => {
  const text = report(judgeDoctor({ ...GOOD, node: "v18.0.0", windows: null }));
  const lines = text.split("\n");
  assert.match(lines[0], /^FAIL  node\s+v18\.0\.0/);
  assert.match(lines[1], /^ {6}fix: .*Node 20/);
  assert.ok(lines.some((l) => /^ok {4}service/.test(l)));
  assert.ok(lines.some((l) => /^skip {2}windows reach/.test(l)));
  assert.equal(lines.at(-1), "1 problem(s).");
  assert.equal(report(judgeDoctor(GOOD)).split("\n").at(-1), "Everything is in place.");
});

// ---- collecting the facts, with stand-ins for every command --------------------

function world(over = {}) {
  const calls = [];
  const answers = {
    "ps -p 1 -o comm=": "systemd\n",
    [`systemctl --user is-active ${unitName}.service`]: "active\n",
    "loginctl show-user tester -p Linger --value": "yes\n",
    "ss -ltn": [
      "State  Recv-Q Send-Q Local Address:Port  Peer Address:Port",
      "LISTEN 0      511        127.0.0.1:47833      0.0.0.0:*",
      "LISTEN 0      511        127.0.0.1:47834      0.0.0.0:*",
    ].join("\n"),
    "powershell.exe": '{"reach":true,"registered":true,"manifest":true,"task":true}\r\n',
    ...over.answers,
  };
  return {
    calls,
    deps: {
      run: (cmd, args) => {
        const line = [cmd, ...args].join(" ");
        calls.push(line);
        const out = cmd === "powershell.exe" ? answers[cmd] : answers[line];
        if (out === undefined || out === null) throw new Error(`${cmd}: not found`);
        return out;
      },
      fetchJson: async (url) => {
        calls.push(`GET ${url}`);
        if (over.health === null) throw new Error("ECONNREFUSED");
        return over.health ?? GOOD.health;
      },
      exists: (path) => {
        calls.push(`exists ${path}`);
        return over.credentials ?? true;
      },
      env: { USER: "tester", HOME: "/home/tester", WSL_DISTRO_NAME: "Ubuntu", ...over.env },
      nodeVersion: "v20.11.1",
    },
  };
}

test("collectFacts: reads each fact with one command and judges clean on a good machine", async () => {
  const w = world();
  const facts = await collectFacts(w.deps);
  assert.deepEqual(facts, GOOD);
  assert.equal(failures(judgeDoctor(facts)), 0);
  assert.ok(w.calls.includes("exists /home/tester/.claude/.credentials.json"));
  assert.ok(w.calls.includes("GET http://127.0.0.1:47834/health"));
  const ps = w.calls.find((c) => c.startsWith("powershell.exe"));
  assert.match(ps, /^powershell\.exe -NoProfile -Command /);
});

test("collectFacts: a command that is missing or fails is a fact, never a crash", async () => {
  // Not WSL: Windows is not asked at all.
  let w = world({ env: { WSL_DISTRO_NAME: undefined } });
  let facts = await collectFacts(w.deps);
  assert.equal(facts.windows, null);
  assert.ok(!w.calls.some((c) => c.startsWith("powershell.exe")));

  // Nothing works.
  w = world({
    answers: {
      "ps -p 1 -o comm=": "init\n",
      [`systemctl --user is-active ${unitName}.service`]: null,
      "loginctl show-user tester -p Linger --value": null,
      "ss -ltn": null,
      "powershell.exe": "not json",
    },
    health: null,
    credentials: false,
  });
  facts = await collectFacts(w.deps);
  assert.deepEqual(facts, {
    node: "v20.11.1",
    credentials: false,
    systemd: false,
    unit: null,
    linger: false,
    ports: { ws: false, http: false },
    health: null,
    windows: null,
  });
  // A port that only looks alike is not ours.
  w = world({ answers: { "ss -ltn": "LISTEN 0 511 127.0.0.1:478340 0.0.0.0:*\n" } });
  assert.deepEqual((await collectFacts(w.deps)).ports, { ws: false, http: false });

  // `systemctl is-active` exits 3 for a unit that is not running and still
  // prints its state: that state is the fact.
  w = world();
  const run = w.deps.run;
  w.deps.run = (cmd, args) => {
    if (cmd !== "systemctl") return run(cmd, args);
    throw Object.assign(new Error("exit 3"), { status: 3, stdout: "failed\n" });
  };
  facts = await collectFacts(w.deps);
  assert.equal(facts.unit, "failed");
  assert.match(judgeDoctor(facts).find((x) => x.name === "service").fix, /daemon-service\.sh logs/);
});

test("windowsProbe: one read-only PowerShell command with the names of the install", () => {
  const ps = windowsProbe();
  assert.doesNotMatch(ps, /[^\x20-\x7e]/, "one ASCII line");
  assert.ok(ps.includes(DEVELOPER_KEY.replace(/^HKCU\\/, "HKCU:\\")));
  assert.ok(ps.includes("Draftspect\\manifests\\word.xml"));
  assert.ok(ps.includes(`'${TASK_NAME}'`));
  assert.ok(ps.includes("47834"));
  assert.match(ps, /ConvertTo-Json -Compress$/);
  // It reads; it writes nothing and starts nothing.
  assert.doesNotMatch(ps, /Set-|New-Item|Remove-|Register-|Start-|reg(\.exe)? add/);
});
