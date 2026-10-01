// The installed daemon (scripts/install, a systemd user unit) listens on the
// two ports an acceptance run needs for a daemon of its own. A run stops the
// unit when it is active and starts it again when the run ends, however it
// ends. Where no unit is installed, or it is not running, nothing is done.
//
//   const resident = pauseResident();        // at the start of a run
//   process.on("exit", resident.resume);     // sync, so it runs on any exit
//
// run(cmd, args) -> stdout; the default runs the command and throws on a
// non-zero exit, as `systemctl is-active` gives for a unit that is not active.

import { execFileSync } from "node:child_process";
import { unitName } from "../install/unit.mjs";

const realRun = (cmd, args) => execFileSync(cmd, args, { encoding: "utf8", stdio: "pipe" });

export function pauseResident({ run = realRun, log = console.log } = {}) {
  const service = `${unitName}.service`;
  let active;
  try {
    active = String(run("systemctl", ["--user", "is-active", service])).trim() === "active";
  } catch {
    // No systemd, no such unit, or the unit is not running: nothing to pause.
    active = false;
  }
  if (!active) return { paused: false, resume: () => {} };
  run("systemctl", ["--user", "stop", service]);
  log(`[run] stopped ${service} for this run; it is started again at the end`);
  let resumed = false;
  return {
    paused: true,
    // Once, whoever calls it first (the end of the run, an exit, a signal).
    resume: () => {
      if (resumed) return;
      resumed = true;
      try {
        run("systemctl", ["--user", "start", service]);
        log(`[run] started ${service} again`);
      } catch (e) {
        log(
          `[run] ${service} was stopped for this run and could not be started again: ${e.message}`,
        );
        log(`[run] start it by hand: systemctl --user start ${service}`);
      }
    },
  };
}
