// Tool results the task pane produced while its socket was down
// (taskpane/shared/result-queue.js): kept for the next connection, so the
// daemon learns that an edit was applied (Codex review 14).
import { test } from "node:test";
import assert from "node:assert/strict";
import { createResultQueue } from "../taskpane/shared/result-queue.js";

test("results wait for the next connection, oldest first, once", () => {
  const q = createResultQueue();
  q.push({ type: "tool_result", id: "a", ok: true });
  q.push({ type: "tool_result", id: "b", ok: false });
  assert.deepEqual(
    q.drain().map((x) => x.id),
    ["a", "b"],
  );
  assert.deepEqual(q.drain(), []);
});

test("the queue is bounded and old results are not sent", () => {
  let now = 0;
  const q = createResultQueue({ max: 3, ttlMs: 1000, now: () => now });
  for (const id of ["a", "b", "c", "d"]) q.push({ id });
  assert.deepEqual(
    q.drain().map((x) => x.id),
    ["b", "c", "d"],
  );
  q.push({ id: "old" });
  now = 500;
  q.push({ id: "new" });
  now = 1200;
  assert.deepEqual(
    q.drain().map((x) => x.id),
    ["new"],
  );
});

test("taskpane.js queues tool results while disconnected and sends them after the hello", async () => {
  const { readFile } = await import("node:fs/promises");
  const src = await readFile(new URL("../taskpane/shared/taskpane.js", import.meta.url), "utf8");
  const send = /function wsSend\(obj\) \{[\s\S]*?\n\}/.exec(src)[0];
  assert.match(
    send,
    /if \(!wsReady\) \{\s*if \(obj\?\.type === "tool_result"\) unsentResults\.push\(obj\);/,
  );
  const open = /ws\.onopen = \(\) => \{[\s\S]*?\n {2}\};/.exec(src)[0];
  assert.ok(open.indexOf("sendHello()") < open.indexOf("unsentResults.drain()"));
});
