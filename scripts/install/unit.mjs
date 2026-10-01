// The systemd user unit that keeps the daemon running without a terminal.
// Pure rendering here; scripts/install/daemon-service.sh writes the file and
// talks to systemctl.
//
//   node scripts/install/unit.mjs name
//   node scripts/install/unit.mjs render <log file> [<directory for PATH>...]
//
// `render` takes node from the node that runs it and the repo from where this
// file is, so the unit starts the checkout the install was run from.

import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

export const unitName = "draftspect-daemon";
export const HEALTH_URL = "http://127.0.0.1:47834/health";

const SYSTEM_PATH = ["/usr/local/bin", "/usr/bin", "/bin"];

// systemd reads a blank as the end of a word and ":" as the end of a PATH
// entry; an install under such a path is refused instead of quoted.
function absolute(path, what) {
  if (typeof path !== "string" || !/^\/[^\s:"'\\]*$/.test(path))
    throw new Error(`${what} must be an absolute path without blanks or colons: ${path}`);
  return path;
}

// PATH of the unit: the directory of node, the given ones (where the command
// line tools the daemon starts were found at install time), the system ones.
export function unitPath(node, extra = []) {
  const dirs = [dirname(absolute(node, "node")), ...extra, ...SYSTEM_PATH];
  return [...new Set(dirs.map((d) => absolute(d, "a PATH directory")))].join(":");
}

export function renderUnit({ node, repo, logFile, extraPath = [] }) {
  absolute(repo, "the repo");
  absolute(logFile, "the log file");
  return `# Written by scripts/install/daemon-service.sh; run it again to change this.
#
# What the daemon runs with:
#   HOME   set by the user manager. The Claude login in $HOME/.claude is the
#          one a terminal uses; nothing is copied and no key is kept here.
#   PATH   the one variable set below: the node that starts the daemon, and
#          the directories where claude, codex and python3 were found when
#          this was installed. Claude turns run the CLI bundled with the
#          daemon unless CC_OFFICE_CLAUDE_PATH names another.
# No variable of a terminal or of an agent session reaches the daemon.
[Unit]
Description=Draftspect daemon (bridge between Word and Claude)
StartLimitIntervalSec=60
StartLimitBurst=5

[Service]
Type=simple
WorkingDirectory=${repo}
ExecStart=${absolute(node, "node")} ${repo}/daemon/index.mjs
Environment=PATH=${unitPath(node, extraPath)}
Restart=on-failure
RestartSec=3
StandardOutput=append:${logFile}
StandardError=append:${logFile}

[Install]
WantedBy=default.target
`;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const [cmd, logFile, ...extraPath] = process.argv.slice(2);
  if (cmd === "name") console.log(unitName);
  else if (cmd === "render" && logFile) {
    const repo = join(dirname(fileURLToPath(import.meta.url)), "..", "..");
    process.stdout.write(renderUnit({ node: process.execPath, repo, logFile, extraPath }));
  } else {
    console.error("usage: unit.mjs name | render <log file> [<PATH directory>...]");
    process.exit(64);
  }
}
