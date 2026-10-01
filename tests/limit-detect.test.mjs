import { test } from "node:test";
import assert from "node:assert/strict";
import {
  detectUsageLimit,
  turnCompleteEvents,
  emptyTurnStats,
  noteMessage,
} from "../daemon/limit-detect.mjs";

const assistant = (text, extra = {}, model = "claude-sonnet-5-5") => ({
  type: "assistant",
  message: { model, content: [{ type: "text", text }] },
  parent_tool_use_id: null,
  ...extra,
});

// Positive: each must be detected.
const POSITIVE = [
  [
    "observed 2026-09-30 shape (error rate_limit, <synthetic>)",
    assistant(
      "You've hit your session limit · resets 5:30pm (Asia/Seoul)",
      { error: "rate_limit" },
      "<synthetic>",
    ),
  ],
  [
    "billing_error",
    assistant("Credit balance is too low", { error: "billing_error" }, "<synthetic>"),
  ],
  [
    "rate_limit_event rejected",
    {
      type: "rate_limit_event",
      rate_limit_info: { status: "rejected", rateLimitType: "five_hour", resetsAt: 1790757000 },
    },
  ],
  [
    "result is_error 429",
    { type: "result", subtype: "success", is_error: true, api_error_status: 429, result: "429" },
  ],
  [
    "result is_error with limit text, no status",
    {
      type: "result",
      subtype: "success",
      is_error: true,
      api_error_status: null,
      result: "Claude usage limit reached. Try later.",
    },
  ],
  [
    "synthetic text without error field (older producer)",
    assistant("You’ve hit your weekly limit", {}, "<synthetic>"),
  ],
];

// Negative: none may be detected. Several carry limit words on purpose.
const NEGATIVE = [
  [
    "model quotes the limit sentence in a normal answer",
    assistant('The banner said "You\'ve hit your session limit" yesterday.'),
  ],
  [
    "rate_limit_event allowed_warning",
    { type: "rate_limit_event", rate_limit_info: { status: "allowed_warning" } },
  ],
  [
    "successful result mentioning a usage limit",
    {
      type: "result",
      subtype: "success",
      is_error: false,
      result: "The usage limit reached in the rate-limiting step was 3 mM.",
    },
  ],
  [
    "is_error 500 server error",
    {
      type: "result",
      subtype: "success",
      is_error: true,
      api_error_status: 500,
      result: "Internal server error",
    },
  ],
  ["assistant error overloaded", assistant("Overloaded", { error: "overloaded" }, "<synthetic>")],
  [
    "stream_event delta with limit words",
    { type: "stream_event", event: { delta: { type: "text_delta", text: "usage limit reached" } } },
  ],
  [
    "synthetic API error that is not a limit",
    assistant("API Error: 500 upstream", {}, "<synthetic>"),
  ],
  ["null message", null],
];

for (const [name, msg] of POSITIVE) {
  test(`limit detected: ${name}`, () => {
    const r = detectUsageLimit(msg);
    assert.equal(r.limited, true);
    assert.equal(r.kind, "usage_limit");
    assert.ok(r.reason);
  });
}

for (const [name, msg] of NEGATIVE) {
  test(`not a limit: ${name}`, () => {
    assert.equal(detectUsageLimit(msg).limited, false);
  });
}

test("fixture counts meet the 5+5 floor", () => {
  assert.ok(POSITIVE.length >= 5 && NEGATIVE.length >= 5);
});

test("turnCompleteEvents: limit replaces success with usage_limit + error", () => {
  const lim = detectUsageLimit(POSITIVE[0][1]);
  const ev = turnCompleteEvents({ type: "result", subtype: "success", is_error: false }, lim);
  assert.deepEqual(ev[0], { event: "turn_complete", subtype: "usage_limit" });
  assert.equal(ev[1].event, "error");
  assert.match(ev[1].error, /usage limit reached\. You've hit your session limit/);
});

test("turnCompleteEvents: non-limit API error still becomes an error", () => {
  const ev = turnCompleteEvents(
    { type: "result", subtype: "success", is_error: true, api_error_status: 500, result: "boom" },
    null,
  );
  assert.deepEqual(
    ev.map((e) => e.event),
    ["turn_complete", "error"],
  );
  assert.equal(ev[0].subtype, "api_error");
});

test("turnCompleteEvents: a normal result is unchanged", () => {
  assert.deepEqual(
    turnCompleteEvents({ type: "result", subtype: "success", is_error: false }, null),
    [{ event: "turn_complete", subtype: "success" }],
  );
  assert.deepEqual(
    turnCompleteEvents({ type: "result", subtype: "error_max_turns", is_error: true }, null),
    [{ event: "turn_complete", subtype: "error_max_turns" }],
  );
});

// L8. Refusals (structured fields only, SDK 0.3.285 sdk.d.ts 5355, 5392,
// 5618, 5713) and turns that end with nothing.
const REFUSAL_POSITIVE = [
  [
    "model_refusal_no_fallback with a category",
    {
      type: "system",
      subtype: "model_refusal_no_fallback",
      original_model: "claude-sonnet-5-5",
      api_refusal_category: "bio",
      content: "declined",
    },
  ],
  [
    "model_refusal_no_fallback without a category",
    { type: "system", subtype: "model_refusal_no_fallback" },
  ],
  [
    "result success with stop_reason refusal",
    { type: "result", subtype: "success", is_error: false, stop_reason: "refusal", result: "" },
  ],
  [
    "error result with stop_reason refusal",
    { type: "result", subtype: "error_during_execution", is_error: true, stop_reason: "refusal" },
  ],
  [
    "assistant error frame with stop_reason refusal",
    {
      type: "assistant",
      message: { model: "claude-sonnet-5-5", stop_reason: "refusal", content: [] },
    },
  ],
];
const REFUSAL_NEGATIVE = [
  [
    "model_refusal_fallback retries on another model (the turn goes on)",
    { type: "system", subtype: "model_refusal_fallback", trigger: "refusal", direction: "retry" },
  ],
  [
    "answer that talks about a refusal",
    assistant("The reviewer refusal to accept the method is noted; I refuse nothing."),
  ],
  [
    "result mentioning refusal with stop_reason end_turn",
    {
      type: "result",
      subtype: "success",
      is_error: false,
      stop_reason: "end_turn",
      result: "no refusal here",
    },
  ],
  [
    "assistant stop_reason end_turn",
    { type: "assistant", message: { stop_reason: "end_turn", content: [] } },
  ],
  [
    "result stop_reason max_tokens",
    {
      type: "result",
      subtype: "success",
      is_error: false,
      stop_reason: "max_tokens",
      result: "cut",
    },
  ],
  ["system init", { type: "system", subtype: "init" }],
];

test("L8 refusal positive corpus: kind refusal", () => {
  for (const [name, msg] of REFUSAL_POSITIVE) {
    const r = detectUsageLimit(msg);
    assert.deepEqual([r.limited, r.kind], [true, "refusal"], name);
  }
});

test("L8 refusal negative corpus", () => {
  for (const [name, msg] of REFUSAL_NEGATIVE)
    assert.equal(detectUsageLimit(msg).limited, false, name);
  console.log(
    `refusal corpus: ${REFUSAL_POSITIVE.length} positive, ${REFUSAL_NEGATIVE.length} negative`,
  );
});

test("L8 turnCompleteEvents: a refusal ends the turn as refusal + error", () => {
  const hit = detectUsageLimit(REFUSAL_POSITIVE[0][1]);
  const ev = turnCompleteEvents(
    { type: "result", subtype: "success", is_error: false, result: "x" },
    hit,
  );
  assert.deepEqual(ev[0], { event: "turn_complete", subtype: "refusal" });
  assert.match(ev[1].error, /declined this request \(category bio\)/);
});

const ok = (result) => ({ type: "result", subtype: "success", is_error: false, result });
const statsOf = (...msgs) => msgs.reduce((st, m) => noteMessage(st, m), emptyTurnStats());
const toolUse = {
  type: "assistant",
  message: { content: [{ type: "tool_use", name: "office_find", input: {} }] },
};

test("L8 empty turn: no text, no tool call, empty result text", () => {
  for (const result of ["", "   ", undefined]) {
    assert.deepEqual(
      turnCompleteEvents(ok(result), null, statsOf({ type: "user" }, assistant("  "))).map(
        (e) => e.subtype ?? e.event,
      ),
      ["empty", "error"],
      String(result),
    );
  }
});

test("L8 not empty: tool calls, text, no stats, an answer, an error subtype", () => {
  const cases = [
    [ok(""), statsOf(toolUse)],
    [ok(""), statsOf(assistant("Done."))],
    [ok(""), null],
    [ok("Done."), statsOf()],
    [{ type: "result", subtype: "error_max_turns", is_error: true, result: "" }, statsOf()],
  ];
  for (const [result, stats] of cases) {
    assert.notEqual(turnCompleteEvents(result, null, stats)[0].subtype, "empty");
  }
});

// The outcome a review turn_done carries (review-service.mjs turnEnded): the
// pane's "Propose all" reads it there, not from the order of the events.
test("turnOutcome: ok for a plain turn_complete; the error text of a limit, refusal, API error or empty turn", async () => {
  const { turnOutcome } = await import("../daemon/limit-detect.mjs");
  assert.deepEqual(turnOutcome(turnCompleteEvents({ subtype: "success", result: "done" }, null)), {
    ok: true,
  });
  const limit = turnOutcome(
    turnCompleteEvents(
      { subtype: "success" },
      { limited: true, kind: "limit", text: "Resets 5pm." },
    ),
  );
  assert.deepEqual(limit, {
    ok: false,
    error: "Claude usage limit reached. Resets 5pm.",
    subtype: "usage_limit",
  });
  const refusal = turnOutcome(
    turnCompleteEvents({ subtype: "success" }, { limited: true, kind: "refusal", text: "" }),
  );
  assert.deepEqual([refusal.ok, refusal.subtype], [false, "refusal"]);
  const empty = turnOutcome(
    turnCompleteEvents({ subtype: "success", result: "" }, null, { text: 0, tool_use: 0 }),
  );
  assert.deepEqual([empty.ok, empty.subtype], [false, "empty"]);
  assert.deepEqual(turnOutcome([]), { ok: true });
});
