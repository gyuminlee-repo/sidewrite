import { test } from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

for (const scenario of [
  "invalid-preauth",
  "invalid-authed",
  "hello-sync",
  "hello-async",
  "message-sync",
  "message-async",
  "close-sync",
  "close-async",
  "ordering",
]) {
  test(`bridge survives ${scenario} without losing subsequent messages`, () => {
    const child = spawnSync(
      process.execPath,
      [fileURLToPath(new URL("./fixtures/bridge-resilience.mjs", import.meta.url)), scenario],
      {
        encoding: "utf8",
        timeout: 5000,
      },
    );
    assert.equal(child.error, undefined, child.error?.message);
    assert.equal(child.status, 0, child.stderr + child.stdout);
  });
}
