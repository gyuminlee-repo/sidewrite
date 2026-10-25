// Requests of the task pane waiting for an answer
// (taskpane/shared/pending-requests.js): a closed connection ends every one
// at once with an error; a slow answer on a live connection still comes
// within the time the request asked for (review file work: 120 s).
import { test, mock } from "node:test";
import assert from "node:assert/strict";
import { createPendingRequests } from "../taskpane/shared/pending-requests.js";

const settled = async (p) => {
  let state = "pending";
  p.then(
    () => (state = "resolved"),
    () => (state = "rejected"),
  );
  await new Promise((r) => setImmediate(r));
  return state;
};

test("a closed connection rejects every waiting request at once", async () => {
  mock.timers.enable({ apis: ["setTimeout"] });
  try {
    const pending = createPendingRequests();
    const a = pending.add("a", "review_refresh", 120_000);
    const b = pending.add("b", "get_context", 10_000);
    a.catch(() => {});
    b.catch(() => {});
    const lost = new Error("The connection to the daemon was lost. Run it again.");
    assert.equal(pending.rejectAll(lost), 2);
    await assert.rejects(a, lost);
    await assert.rejects(b, lost);
    assert.equal(pending.size, 0);
    // The answer of the old connection, if it ever came, finds no request.
    assert.equal(pending.resolve("a", { ok: true }), false);
  } finally {
    mock.timers.reset();
  }
});

test("on a live connection a slow request waits beyond 10 s for its answer", async () => {
  mock.timers.enable({ apis: ["setTimeout"] });
  try {
    const pending = createPendingRequests();
    const p = pending.add("r", "review_refresh", 120_000);
    mock.timers.tick(10_001);
    assert.equal(await settled(p), "pending");
    mock.timers.tick(50_000);
    assert.equal(pending.resolve("r", { ok: true, slow: true }), true);
    assert.deepEqual(await p, { ok: true, slow: true });
  } finally {
    mock.timers.reset();
  }
});

test("a request with no answer ends after its own time", async () => {
  mock.timers.enable({ apis: ["setTimeout"] });
  try {
    const pending = createPendingRequests();
    const p = pending.add("r", "review_refresh", 120_000);
    p.catch(() => {});
    mock.timers.tick(119_999);
    assert.equal(await settled(p), "pending");
    mock.timers.tick(1);
    await assert.rejects(p, /Request "review_refresh" timed out/);
    assert.equal(pending.size, 0);
  } finally {
    mock.timers.reset();
  }
});

test("taskpane.js ends the waiting requests when the socket closes", async () => {
  const { readFile } = await import("node:fs/promises");
  const src = await readFile(new URL("../taskpane/shared/taskpane.js", import.meta.url), "utf8");
  const close = /ws\.onclose = \(\) => \{[\s\S]*?\n {2}\};/.exec(src)[0];
  assert.match(close, /pendingRequests\.rejectAll\(new Error\(CONNECTION_LOST\)\)/);
  // The answer dispatcher and sendRequest go through the same table.
  assert.match(src, /pendingRequests\.resolve\(msg\.request_id, msg\)/);
  assert.match(src, /pendingRequests\.add\(request_id, type, timeoutMs\)/);
});
