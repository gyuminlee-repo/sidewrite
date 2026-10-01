import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { pauseResident } from "../scripts/e2e/resident-unit.mjs";
import { unitName } from "../scripts/install/unit.mjs";

const SVC = `${unitName}.service`;
// A systemctl that answers is-active with `state` (and exits non-zero unless active).
function systemctl(state, { failStart = false } = {}) {
  const calls = [];
  const run = (cmd, args) => {
    calls.push([cmd, ...args].join(" "));
    if (args.includes("is-active")) {
      if (state === "active") return "active\n";
      throw Object.assign(new Error("exit 3"), { stdout: `${state}\n` });
    }
    if (args.includes("start") && failStart) throw new Error("unit not found");
    return "";
  };
  return { run, calls };
}

test("an active resident unit is stopped for the run and started again once", () => {
  const s = systemctl("active");
  const lines = [];
  const resident = pauseResident({ run: s.run, log: (l) => lines.push(l) });
  assert.equal(resident.paused, true);
  assert.deepEqual(s.calls, [`systemctl --user is-active ${SVC}`, `systemctl --user stop ${SVC}`]);
  // The end of the run, then the exit handler, then a signal: one start.
  resident.resume();
  resident.resume();
  resident.resume();
  assert.deepEqual(s.calls.slice(2), [`systemctl --user start ${SVC}`]);
  assert.match(lines.join("\n"), /stopped .* for this run/);
  assert.match(lines.join("\n"), /started .* again/);
});

test("no unit, a unit that is not running, no systemctl: nothing is stopped and nothing started", () => {
  for (const state of ["inactive", "failed", "unknown"]) {
    const s = systemctl(state);
    const resident = pauseResident({ run: s.run, log: () => {} });
    assert.equal(resident.paused, false, state);
    resident.resume();
    assert.deepEqual(s.calls, [`systemctl --user is-active ${SVC}`], state);
  }
  const none = pauseResident({
    run: () => {
      throw Object.assign(new Error("spawn systemctl ENOENT"), { code: "ENOENT" });
    },
    log: () => {},
  });
  assert.equal(none.paused, false);
  none.resume();
});

test("a unit that cannot be started again is said, with the command, and does not throw", () => {
  const s = systemctl("active", { failStart: true });
  const lines = [];
  const resident = pauseResident({ run: s.run, log: (l) => lines.push(l) });
  resident.resume();
  assert.match(lines.join("\n"), /could not be started again: unit not found/);
  assert.match(lines.join("\n"), new RegExp(`systemctl --user start ${SVC}`));
});

test("the acceptance run pauses the resident unit before it looks at the ports and resumes on every exit", async () => {
  const src = await readFile(new URL("../scripts/e2e/run-acceptance.mjs", import.meta.url), "utf8");
  assert.match(src, /import \{ pauseResident \} from "\.\/resident-unit\.mjs";/);
  const pause = src.indexOf("resident = pauseResident(");
  assert.ok(pause > 0);
  assert.ok(pause < src.indexOf("is in use; stop that daemon first"), "before the port check");
  // exit covers the end of the run, process.exit and an uncaught error.
  assert.match(src, /process\.on\("exit", \(\) => resident\.resume\(\)\);/);
  // A signal ends the process through process.exit, so the exit handler runs.
  const signals = src.slice(
    src.indexOf('process.on("exit"'),
    src.indexOf("async function startDaemon"),
  );
  assert.match(signals, /\["SIGINT", 130\]/);
  assert.match(signals, /\["SIGTERM", 143\]/);
  // The run's own daemon is stopped (and its ports are free) before that.
  assert.match(
    signals,
    /process\.on\(signal, \(\) => \{[\s\S]*stopDaemon\(\)[\s\S]*\.finally\(\(\) => process\.exit\(code\)\);/,
  );
});
