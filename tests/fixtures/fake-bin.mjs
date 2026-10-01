// Stand-ins for the system commands the install scripts call, so a test never
// touches systemd, the registry or the network. Each one appends
// "<name> <args>" to the file $FAKE_CALLS and answers from a FAKE_* variable.
//
//   const fake = await fakeBin({ systemctl: 'echo active', ... });
//   execFileSync("bash", [script], { env: fake.env({ FAKE_LINGER: "no" }) });
//   fake.calls()  -> ["systemctl --user daemon-reload", ...]

import { chmod, mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";

// bodies: { <command name>: bash lines run after the call is logged }.
export async function fakeBin(bodies) {
  const dir = await mkdtemp(join(tmpdir(), "fake-bin-"));
  const bin = join(dir, "bin");
  const home = join(dir, "home");
  const calls = join(dir, "calls.txt");
  await mkdir(bin);
  await mkdir(home);
  for (const [name, body] of Object.entries(bodies)) {
    const path = join(bin, name);
    await writeFile(path, `#!/usr/bin/env bash\necho "${name} $*" >> "$FAKE_CALLS"\n${body}\n`);
    await chmod(path, 0o755);
  }
  return {
    dir,
    bin,
    home,
    // The fakes first, then node, then the plain system tools (dirname, id, ...).
    env: (more = {}) => ({
      PATH: `${bin}:${dirname(process.execPath)}:/usr/bin:/bin`,
      HOME: home,
      USER: "tester",
      FAKE_CALLS: calls,
      ...more,
    }),
    calls: async () =>
      existsSync(calls) ? (await readFile(calls, "utf8")).trim().split("\n").filter(Boolean) : [],
    reset: () => rm(calls, { force: true }),
    cleanup: () => rm(dir, { recursive: true, force: true }),
  };
}
