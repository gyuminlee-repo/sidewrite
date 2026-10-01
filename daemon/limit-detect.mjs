// Recognize a Claude usage-limit stop from structured SDK messages, so the
// daemon reports it as an error instead of "turn complete (success)".
//
// Observed 2026-09-30 (session f2748a6f .jsonl, CLI 2.1.285): the limit
// arrives as an assistant message with error "rate_limit", apiErrorStatus
// 429, model "<synthetic>" and text "You've hit your session limit ...".
// SDK 0.3.285 types: SDKAssistantMessage.error (SDKAssistantMessageError),
// SDKResultSuccess.is_error / api_error_status, SDKRateLimitEvent with
// rate_limit_info.status "rejected".
//
// Text is consulted only on messages that are already errors (synthetic
// assistant messages, is_error result messages), never on a normal model answer,
// so a manuscript sentence about a "rate limit" cannot trip it.
//
// L8 (a common failure mode of LLM-assisted revisions): a refusal and a turn that
// ends with no answer at all are errors too. Refusal fields, SDK 0.3.285
// sdk.d.ts: SDKModelRefusalNoFallbackMessage (system/model_refusal_no_fallback,
// 5392), result stop_reason (5618, 5713). model_refusal_fallback (5355) retries
// on another model, so the turn goes on and it is not an error. An empty turn
// needs per-turn counts (noteMessage), because a review turn may end right
// after a tool call with an empty result text.

export const LIMIT_ERRORS = new Set(["rate_limit", "billing_error"]);

export const LIMIT_TEXT_RE =
  /(?:you['’]ve hit your (?:session|usage|weekly|daily|monthly) limit|\busage limit (?:reached|exceeded)\b|\b(?:session|weekly|daily) limit reached\b|\bout of (?:credits|extra usage)\b)/i;

const NO = Object.freeze({ limited: false, kind: null, reason: null, text: "" });

function refusal(reason, category) {
  return { limited: true, kind: "refusal", reason, text: category ? `category ${category}` : "" };
}

function assistantText(msg) {
  const c = msg?.message?.content;
  if (typeof c === "string") return c;
  if (!Array.isArray(c)) return "";
  return c
    .filter((b) => b?.type === "text" && typeof b.text === "string")
    .map((b) => b.text)
    .join("");
}

export function detectUsageLimit(msg) {
  if (!msg || typeof msg !== "object") return NO;
  if (msg.type === "assistant") {
    const text = assistantText(msg);
    if (LIMIT_ERRORS.has(msg.error)) {
      return { limited: true, kind: "usage_limit", reason: `assistant.error=${msg.error}`, text };
    }
    if (msg.message?.model === "<synthetic>" && LIMIT_TEXT_RE.test(text)) {
      return { limited: true, kind: "usage_limit", reason: "synthetic assistant text", text };
    }
    if (msg.message?.stop_reason === "refusal") return refusal("assistant.stop_reason=refusal");
    return NO;
  }
  if (msg.type === "system" && msg.subtype === "model_refusal_no_fallback") {
    return refusal("model_refusal_no_fallback", msg.api_refusal_category);
  }
  if (msg.type === "rate_limit_event") {
    const info = msg.rate_limit_info ?? {};
    if (info.status !== "rejected") return NO;
    const when = Number.isFinite(info.resetsAt)
      ? `resets at ${new Date(info.resetsAt * 1000).toISOString()}`
      : "";
    return {
      limited: true,
      kind: "usage_limit",
      reason: `rate_limit_event ${info.rateLimitType ?? "unknown"}`,
      text: when,
    };
  }
  if (msg.type === "result" && msg.is_error === true) {
    const text = String(msg.result ?? (msg.errors ?? []).join("; "));
    if (msg.api_error_status === 429 || LIMIT_TEXT_RE.test(text)) {
      return {
        limited: true,
        kind: "usage_limit",
        reason: `result.is_error api_error_status=${msg.api_error_status ?? "null"}`,
        text,
      };
    }
  }
  if (msg.type === "result" && msg.stop_reason === "refusal") {
    return refusal("result.stop_reason=refusal");
  }
  return NO;
}

// Per-turn counts of what the model produced, for the empty-turn check.
export function emptyTurnStats() {
  return { text: 0, tool_use: 0 };
}

export function noteMessage(stats, msg) {
  if (!stats || msg?.type !== "assistant" || !Array.isArray(msg.message?.content)) return stats;
  for (const b of msg.message.content) {
    if (b?.type === "text" && String(b.text ?? "").trim()) stats.text += 1;
    else if (b?.type === "tool_use") stats.tool_use += 1;
  }
  return stats;
}

// Events the daemon sends for a `result` message. `usageLimit` is the first
// detectUsageLimit() hit seen during the turn (or null); `stats` the
// noteMessage() counts of the turn (or null: no empty-turn check). Order
// matters: the task pane treats a later "error" as the final status of the turn.
export function turnCompleteEvents(result, usageLimit, stats = null) {
  if (usageLimit?.limited && usageLimit.kind === "refusal") {
    const detail = usageLimit.text ? ` (${usageLimit.text})` : "";
    return [
      { event: "turn_complete", subtype: "refusal" },
      {
        event: "error",
        error: `The model declined this request${detail}. Nothing was changed by this answer; rephrase the request or split the comment.`,
      },
    ];
  }
  if (usageLimit?.limited) {
    const detail = usageLimit.text ? ` ${usageLimit.text}` : "";
    return [
      { event: "turn_complete", subtype: "usage_limit" },
      { event: "error", error: `Claude usage limit reached.${detail}` },
    ];
  }
  // Only a "success" result can hide an API error (is_error true); the error
  // subtypes already say why the turn stopped and keep their subtype.
  if (result?.subtype === "success" && result.is_error === true) {
    const text = String(result.result ?? (result.errors ?? []).join("; ")).slice(0, 500);
    return [
      { event: "turn_complete", subtype: "api_error" },
      { event: "error", error: `The turn ended on an API error: ${text || "(no detail)"}` },
    ];
  }
  if (
    result?.subtype === "success" &&
    !String(result.result ?? "").trim() &&
    stats &&
    stats.text === 0 &&
    stats.tool_use === 0
  ) {
    return [
      { event: "turn_complete", subtype: "empty" },
      { event: "error", error: "The turn ended without an answer or a tool call. Send it again." },
    ];
  }
  return [{ event: "turn_complete", subtype: result?.subtype }];
}

// How a turn ended, from the events turnCompleteEvents made for it:
// { ok: true }, or { ok: false, error, subtype }. Failed: the events hold an
// error (limit, refusal, API error, empty turn), or the SDK result has an
// error subtype (SDKResultError: error_max_turns, error_during_execution,
// error_max_budget_usd, error_max_structured_output_retries, and any
// error_* added later), which comes with no error event. `result` adds the
// SDK's own error texts to the reason.
// The review turn_done carries it (review-service.mjs turnEnded), so the
// task pane does not have to pair turn_done with an error event by timing.
const SUBTYPE_REASONS = {
  error_max_turns: "The turn stopped at the maximum number of turns before it finished.",
  error_max_budget_usd: "The turn stopped at the spending limit.",
  error_max_structured_output_retries:
    "The turn stopped: the model did not give the structured output it was asked for.",
};

export function turnOutcome(events, result = null) {
  const list = events ?? [];
  const done = list.find((e) => e.event === "turn_complete");
  const err = list.find((e) => e.event === "error");
  if (err) return { ok: false, error: err.error, subtype: done?.subtype ?? null };
  const subtype = done?.subtype;
  if (typeof subtype !== "string" || !subtype.startsWith("error")) return { ok: true };
  const detail = (Array.isArray(result?.errors) ? result.errors : [])
    .map((e) => String(e).trim())
    .filter(Boolean)
    .join("; ")
    .slice(0, 500);
  const error =
    subtype === "error_during_execution"
      ? `The turn ended on an error while it ran${detail ? `: ${detail}` : "."}`
      : (SUBTYPE_REASONS[subtype] ?? `The turn ended with ${subtype}.`);
  return { ok: false, error, subtype };
}
