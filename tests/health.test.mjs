import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { healthBody, VERSION } from "../daemon/health.mjs";

test("healthBody: up, how long, which version, which ports, and nothing else", async () => {
  const body = healthBody({ startedAt: 1_000, now: 61_900, ws: 47833, http: 47834 });
  assert.deepEqual(body, {
    ok: true,
    uptime_s: 60,
    version: VERSION,
    ports: { ws: 47833, http: 47834 },
  });
  const pkg = JSON.parse(await readFile(new URL("../package.json", import.meta.url), "utf8"));
  assert.equal(VERSION, pkg.version);
  // No path and no token: the reply is given to anyone who asks on the loopback.
  const text = JSON.stringify(body);
  assert.doesNotMatch(text, /[\\/]/);
  assert.doesNotMatch(text, /[0-9a-f]{32}/);
  // A clock that went back does not give a negative time.
  assert.equal(healthBody({ startedAt: 5_000, now: 1_000, ws: 1, http: 2 }).uptime_s, 0);
});

test("the daemon answers /health before the token and the files, as JSON", async () => {
  const src = await readFile(new URL("../daemon/index.mjs", import.meta.url), "utf8");
  const health = src.indexOf('urlPath === "/health"');
  assert.ok(health > 0, "no /health route");
  assert.ok(health < src.indexOf('urlPath === "/bridge-token"'));
  const route = src.slice(health, src.indexOf('urlPath === "/bridge-token"'));
  assert.match(route, /application\/json/);
  assert.match(
    route,
    /healthBody\(\{ startedAt: STARTED_AT, now: Date\.now\(\), ws: WS_PORT, http: HTTP_PORT \}\)/,
  );
  assert.doesNotMatch(route, /BRIDGE_TOKEN|TOKEN_FILE/);
});
