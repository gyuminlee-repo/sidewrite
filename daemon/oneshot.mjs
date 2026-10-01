// Single-shot model requests that must not touch the drafting session:
// comment splitting, LLM ledger checks, and (stage 4, P3) translation and
// questions. No resume, no tools, no MCP, no settings, not persisted.
//
// Returns { ok, text, json?, error?, limited? } and never throws, so callers
// can show a plain message in the task pane.

import { query as sdkQuery } from "@anthropic-ai/claude-agent-sdk";
import { tmpdir } from "node:os";
import { detectUsageLimit } from "./limit-detect.mjs";
import { claudeExecutablePath } from "./query-options.mjs";

export const ONESHOT_MODEL = "haiku";

// The executable is the one the drafting session runs (CC_OFFICE_CLAUDE_PATH,
// query-options.mjs claudeExecutablePath; unset: the SDK's bundled CLI). The
// isolation stays: no tools, no MCP, no settings, nothing persisted.
export function buildOneshotOptions({
  system,
  model = ONESHOT_MODEL,
  cwd = tmpdir(),
  abortController,
  env = process.env,
}) {
  const executable = claudeExecutablePath(env);
  return {
    cwd,
    model,
    systemPrompt: system,
    tools: [],
    mcpServers: {},
    settingSources: [],
    strictMcpConfig: true,
    persistSession: false,
    maxTurns: 1,
    ...(executable ? { pathToClaudeCodeExecutable: executable } : {}),
    ...(abortController ? { abortController } : {}),
  };
}

// Accepts bare JSON, a fenced ```json block, or JSON with prose around it.
export function parseJsonReply(text) {
  const s = String(text ?? "").trim();
  const fence = /```(?:json)?\s*([\s\S]*?)```/i.exec(s);
  const body = fence ? fence[1].trim() : s;
  try {
    return JSON.parse(body);
  } catch {
    /* fall through */
  }
  const a = body.indexOf("{");
  const b = body.lastIndexOf("}");
  if (a !== -1 && b > a) {
    try {
      return JSON.parse(body.slice(a, b + 1));
    } catch {
      /* fall through */
    }
  }
  return undefined;
}

export async function runOneshot(
  { system, prompt, model = ONESHOT_MODEL, expectJson = true, timeoutMs = 120_000 },
  { queryFn = sdkQuery, env = process.env } = {},
) {
  const abortController = new AbortController();
  const timer = setTimeout(() => abortController.abort(), timeoutMs);
  let limited = null;
  let result = null;
  try {
    for await (const msg of queryFn({
      prompt,
      options: buildOneshotOptions({ system, model, abortController, env }),
    })) {
      const lim = detectUsageLimit(msg);
      if (lim.limited && !limited) limited = lim;
      if (msg?.type === "result") result = msg;
    }
  } catch (e) {
    const aborted = abortController.signal.aborted || e?.name === "AbortError";
    return {
      ok: false,
      error: aborted ? `timed out after ${timeoutMs} ms` : String(e?.message ?? e),
      limited,
    };
  } finally {
    clearTimeout(timer);
  }
  if (limited) {
    const what =
      limited.kind === "refusal"
        ? "The model declined this request."
        : "Claude usage limit reached.";
    return { ok: false, error: `${what} ${limited.text}`.trim(), limited };
  }
  if (!result) return { ok: false, error: "no result message" };
  if (result.subtype !== "success" || result.is_error) {
    const detail = String(result.result ?? (result.errors ?? []).join("; ")).slice(0, 300);
    return {
      ok: false,
      error: `one-shot ended with ${result.subtype}${result.is_error ? " (is_error)" : ""}: ${detail}`,
    };
  }
  const text = String(result.result ?? "");
  if (!expectJson) return { ok: true, text };
  const json = parseJsonReply(text);
  if (json === undefined) return { ok: false, error: "reply was not JSON", text };
  return { ok: true, text, json };
}

export function makeOneshot(deps = {}) {
  return (req) => runOneshot(req, deps);
}
