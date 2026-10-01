// GET /health: what a script needs to tell that the daemon is up (the
// install waits for it, the doctor reads it). Answered without the bridge
// token, to anyone on the loopback, so it holds no token and no path.

import { readFile } from "node:fs/promises";

export const VERSION = JSON.parse(
  await readFile(new URL("../package.json", import.meta.url), "utf8"),
).version;

// startedAt, now: milliseconds. ws, http: the two ports the daemon listens on.
export function healthBody({ startedAt, now, ws, http }) {
  return {
    ok: true,
    uptime_s: Math.max(0, Math.floor((now - startedAt) / 1000)),
    version: VERSION,
    ports: { ws, http },
  };
}
