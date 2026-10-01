import { test } from "node:test";
import assert from "node:assert/strict";
import { createProposeQueue } from "../taskpane/shared/propose-queue.js";
import { queueRefs } from "../taskpane/shared/review-view.js";

test("queueRefs: not done, in list order", () => {
  const rev = {
    comments: [
      { ref: "R1.1", status: "done" },
      { ref: "R1.2", status: "pending" },
      { ref: "R2.1", status: "in_progress" },
    ],
  };
  assert.deepEqual(queueRefs(rev), ["R1.2", "R2.1"]);
  assert.deepEqual(queueRefs(null), []);
});

test("runs one at a time, stops on error, stop() ends after the current one", async () => {
  const order = [];
  let active = 0;
  const progress = [];
  const q = createProposeQueue({
    proposeOne: async (ref) => {
      active += 1;
      assert.equal(active, 1);
      order.push(ref);
      await new Promise((r) => setTimeout(r, 5));
      active -= 1;
      return ref === "R1.3" ? { ok: false, error: "usage limit" } : { ok: true };
    },
    onProgress: (p) => progress.push(`${p.index + 1}/${p.total} ${p.ref}`),
  });
  const r = await q.run(["R1.1", "R1.2", "R1.3", "R1.4"]);
  assert.deepEqual(order, ["R1.1", "R1.2", "R1.3"]);
  assert.deepEqual([r.failed, r.error], ["R1.3", "usage limit"]);
  assert.deepEqual(progress, ["1/4 R1.1", "2/4 R1.2", "3/4 R1.3"]);
  const q2 = createProposeQueue({
    proposeOne: async (ref) => (ref === "A" && q2.stop(), { ok: true }),
  });
  const r2 = await q2.run(["A", "B"]);
  assert.deepEqual([r2.done.map((d) => d.ref), r2.stopped, r2.next], [["A"], true, "B"]);
  assert.equal(q2.running, false);
});

test("a second run while one is running is refused; a throw ends the run as a failure", async () => {
  let release;
  const q = createProposeQueue({ proposeOne: () => new Promise((r) => (release = r)) });
  const first = q.run(["A"]);
  await assert.rejects(q.run(["B"]), /already running/);
  release({ ok: true });
  assert.deepEqual((await first).done, [{ ref: "A", ok: true }]);
  const q2 = createProposeQueue({
    proposeOne: async () => {
      throw new Error("Not connected");
    },
  });
  const r = await q2.run(["A", "B"]);
  assert.deepEqual([r.failed, r.error, q2.running], ["A", "Not connected", false]);
});
