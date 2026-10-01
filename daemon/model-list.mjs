// Claude models for the composer dropdown (P13). The SDK only exposes
// supportedModels() on a live Query, so a throwaway isolated query is opened
// with a prompt stream that never yields, asked once, and closed. No tools,
// no MCP, no settings, not persisted (the isolation of oneshot.mjs); no
// prompt is sent, so it costs no tokens.
//
// fetchClaudeModels never throws: { source: "sdk", models } or
// { source: "fallback", models: FALLBACK_MODELS, error }.

import { query as sdkQuery } from "@anthropic-ai/claude-agent-sdk";
import { tmpdir } from "node:os";
import { claudeExecutablePath } from "./query-options.mjs";

export const ALIASES = ["haiku", "sonnet", "opus"];
export const FALLBACK_MODELS = [
  { value: "haiku", displayName: "Haiku", description: "cheapest", resolvedModel: null },
  { value: "sonnet", displayName: "Sonnet", description: "balanced", resolvedModel: null },
  { value: "opus", displayName: "Opus", description: "most capable", resolvedModel: null },
];
const ID_RE = /^claude-[a-z0-9][a-z0-9.-]*(\[1m\])?$/i;

// The "default" row is left out: the daemon always pins an explicit model.
export function normalizeModels(list) {
  const seen = new Set();
  const out = [];
  for (const m of Array.isArray(list) ? list : []) {
    const value = String(m?.value ?? "").trim();
    if (!value || value === "default" || seen.has(value)) continue;
    seen.add(value);
    out.push({
      value,
      displayName: String(m.displayName ?? value),
      description: String(m.description ?? ""),
      resolvedModel: m.resolvedModel ?? null,
    });
  }
  return out;
}

export async function fetchClaudeModels({
  queryFn = sdkQuery,
  timeoutMs = Number(process.env.CC_OFFICE_MODELS_TIMEOUT_MS) || 20_000,
  env = process.env,
} = {}) {
  let release;
  const idle = new Promise((r) => (release = r));
  async function* never() {
    await idle;
  }
  const abortController = new AbortController();
  const exe = claudeExecutablePath(env);
  let q = null;
  let timer;
  try {
    q = queryFn({
      prompt: never(),
      options: {
        cwd: tmpdir(),
        tools: [],
        mcpServers: {},
        settingSources: [],
        strictMcpConfig: true,
        persistSession: false,
        abortController,
        ...(exe ? { pathToClaudeCodeExecutable: exe } : {}),
      },
    });
    // Read the stream in the background so control replies arrive. Its end
    // or failure carries no result: the list or the timeout below decides.
    const stream = q;
    (async () => {
      for await (const _ of stream) void _;
    })().then(undefined, () => undefined);
    const list = await Promise.race([
      q.supportedModels(),
      new Promise((_, rej) => {
        timer = setTimeout(() => rej(new Error(`timed out after ${timeoutMs} ms`)), timeoutMs);
      }),
    ]);
    const models = normalizeModels(list);
    if (!models.length) throw new Error("empty model list");
    return { source: "sdk", models };
  } catch (e) {
    return { source: "fallback", models: FALLBACK_MODELS, error: String(e?.message ?? e) };
  } finally {
    clearTimeout(timer);
    release();
    q?.close?.();
    abortController.abort();
  }
}

// The model the daemon passes to query(): the three aliases always, anything
// the account lists (by value or by resolved id), and, while the list is
// unavailable, any well-formed claude-* id. Everything else is "sonnet".
export function resolveModelChoice(requested, state) {
  const r = String(requested ?? "").trim();
  if (ALIASES.includes(r)) return r;
  const models = state?.models ?? [];
  if (models.some((m) => m.value === r || m.resolvedModel === r)) return r;
  if (state?.source !== "sdk" && ID_RE.test(r)) return r;
  return "sonnet";
}

// What set_model does with a choice: the model to use and whether the
// session of that pane must be restarted to run on it. after: the account's
// list, once it is known. running: the model the pane's session was started
// on (index.mjs session.model); null when there is no session, or one that
// has not read its model yet (it will read the one decided here).
// The session's own model is the measure, not what the list allowed when
// the request came: a session started while the list was awaited can be on
// the replacement already, and restarting it would be for nothing.
export function modelChange({ requested, after, running = null }) {
  const model = resolveModelChoice(requested, after);
  return { model, restart: running != null && running !== model };
}

export function createModelCache({
  fetch = fetchClaudeModels,
  ttlMs = 10 * 60_000,
  // A failed fetch (a slow CLI start can time the probe out) is tried again
  // soon instead of leaving the aliases for the whole ttl.
  fallbackTtlMs = 60_000,
  now = Date.now,
} = {}) {
  let value = null;
  let at = -Infinity;
  let inflight = null;
  return {
    current: () => value ?? { source: "fallback", models: FALLBACK_MODELS },
    async get({ refresh = false } = {}) {
      const ttl = value?.source === "sdk" ? ttlMs : fallbackTtlMs;
      if (!refresh && value && now() - at < ttl) return value;
      inflight ??= fetch().then((v) => {
        value = v;
        at = now();
        inflight = null;
        return v;
      });
      return inflight;
    },
  };
}
