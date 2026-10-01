import { test } from "node:test";
import assert from "node:assert/strict";
import { wrapHandlers } from "../daemon/handler-wrap.mjs";

function harness(eventType, sendEvent) {
  const events = [],
    replies = [],
    lines = [];
  const h = wrapHandlers(
    {
      a: async (msg, key, host) => ({ x: 1, key, host }),
      j: async () => ({ y: 2 }),
      bad: async () => {
        throw new Error("no");
      },
      jbad: async () => {
        throw new Error("late");
      },
      jsync: () => {
        throw new Error("at once");
      },
    },
    {
      jobTypes: new Set(["j", "jbad", "jsync"]),
      sendEvent: sendEvent ?? ((key, e) => events.push([key, e])),
      eventType,
      log: (s) => lines.push(s),
    },
  );
  return { h, events, replies, lines, reply: (r) => replies.push(r) };
}
const settled = () => new Promise((r) => setImmediate(r));

test("plain request: <type>_result with the request id; an error becomes ok:false", async () => {
  const t = harness("response_event");
  await t.h.a({ request_id: "q1" }, t.reply, "k", "word");
  await t.h.bad({ request_id: "q2" }, t.reply, "k", "word");
  assert.deepEqual(t.replies, [
    { type: "a_result", request_id: "q1", ok: true, x: 1, key: "k", host: "word" },
    { type: "bad_result", request_id: "q2", ok: false, error: "no" },
  ]);
  assert.deepEqual(t.events, []);
});

test("job: accepted at once, result in a job_done event of the given type", async () => {
  const t = harness("response_event");
  await t.h.j({ request_id: "q3" }, t.reply, "k", "word");
  await t.h.jbad({ request_id: "q4" }, t.reply, "k", "word");
  await settled();
  assert.deepEqual(Object.keys(t.replies[0]).sort(), [
    "accepted",
    "job_id",
    "ok",
    "request_id",
    "type",
  ]);
  assert.deepEqual(
    [t.replies[0].type, t.replies[0].ok, t.replies[0].accepted, t.replies[0].request_id],
    ["j_result", true, true, "q3"],
  );
  const [k, e] = t.events[0];
  assert.deepEqual(
    [k, e.type, e.event, e.job_id, e.result],
    ["k", "response_event", "job_done", t.replies[0].job_id, { ok: true, y: 2 }],
  );
  assert.deepEqual(t.events[1][1].result, { ok: false, error: "late" });
  assert.equal(t.events[1][1].job_id, t.replies[1].job_id);
  assert.notEqual(t.replies[0].job_id, t.replies[1].job_id);
});

test("eventType defaults to review_event", async () => {
  const t = harness(undefined);
  await t.h.j({ request_id: "q" }, t.reply, "k", "word");
  await settled();
  assert.equal(t.events[0][1].type, "review_event");
});

test("a job that throws before its first await still ends in job_done", async () => {
  const t = harness("response_event");
  await t.h.jsync({ request_id: "q" }, t.reply, "k", "word");
  await settled();
  assert.equal(t.replies[0].accepted, true);
  assert.deepEqual(t.events[0][1].result, { ok: false, error: "at once" });
});

test("a job_done that cannot be sent is logged and does not throw", async () => {
  for (const [eventType, tag] of [
    ["response_event", "[response]"],
    [undefined, "[review]"],
  ]) {
    const t = harness(eventType, () => {
      throw new Error("socket closed");
    });
    await t.h.j({ request_id: "q" }, t.reply, "k", "word");
    await settled();
    assert.deepEqual(t.lines, [`${tag} job_done not delivered: socket closed`]);
  }
});

test("every handler of impl is wrapped and nothing else", () => {
  const h = wrapHandlers({ one: async () => ({}), two: async () => ({}) });
  assert.deepEqual(Object.keys(h), ["one", "two"]);
});
