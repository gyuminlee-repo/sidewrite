import { test } from "node:test";
import assert from "node:assert/strict";
import { createAutoRefresh } from "../taskpane/shared/auto-refresh.js";

test("skips while busy, coalesces overlapping triggers, honors the interval", async () => {
  let t = 0;
  let busy = false;
  let runs = 0;
  let release;
  const ar = createAutoRefresh({
    run: () => new Promise((r) => (runs++, (release = r))),
    isBusy: () => busy,
    minIntervalMs: 1000,
    now: () => t,
  });
  busy = true;
  assert.equal(ar.trigger("focus"), "busy");
  busy = false;
  assert.equal(ar.trigger("focus"), "started");
  assert.equal(ar.trigger("turn"), "queued");
  release();
  await new Promise((r) => setImmediate(r));
  assert.equal(runs, 2); // the queued one ran after the first ended
  release();
  await new Promise((r) => setImmediate(r));
  t = 500;
  assert.equal(ar.trigger("focus"), "too_soon");
  t = 3000;
  assert.equal(ar.trigger("focus"), "started");
});

test("a focus during a run joins it; a turn end ignores the interval; a failed run does not stick", async () => {
  let t = 0;
  let runs = 0;
  let settle;
  const ar = createAutoRefresh({
    run: () => new Promise((res, rej) => (runs++, (settle = { res, rej }))),
    minIntervalMs: 1000,
    now: () => t,
  });
  assert.equal(ar.trigger("focus"), "started");
  assert.equal(ar.trigger("focus"), "running"); // focus and visibilitychange come as a pair
  assert.equal(ar.trigger("turn"), "queued");
  assert.equal(ar.trigger("turn"), "queued"); // once only
  settle.rej(new Error("Word did not answer"));
  await new Promise((r) => setImmediate(r));
  assert.equal(runs, 2);
  settle.res();
  await new Promise((r) => setImmediate(r));
  assert.equal(runs, 2);
  t = 10;
  assert.equal(ar.trigger("turn"), "started"); // 10 ms after the last start
  settle.res();
});

test("a queued run is dropped when the pane got busy meanwhile; a throwing run is a failed run", async () => {
  let busy = false;
  let runs = 0;
  let release;
  const ar = createAutoRefresh({
    run: () => new Promise((r) => (runs++, (release = r))),
    isBusy: () => busy,
  });
  ar.trigger("focus");
  ar.trigger("turn");
  busy = true;
  release();
  await new Promise((r) => setImmediate(r));
  assert.equal(runs, 1);
  const thrower = createAutoRefresh({
    run: () => {
      throw new Error("sync");
    },
  });
  assert.equal(thrower.trigger("turn"), "started");
  await new Promise((r) => setImmediate(r));
  assert.equal(thrower.trigger("turn"), "started");
});
