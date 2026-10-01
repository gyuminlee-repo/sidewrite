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

// The best model for work whose quality matters (edits of the manuscript,
// the response letter): the first row of the Opus family in the account's
// list, by its name or by what it resolves to, else "sonnet". While the list
// is unavailable its stand-in holds the alias "opus", which query() takes.
export function bestClaudeModel(state) {
  const opus = (state?.models ?? []).find(
    (m) => /opus/i.test(m.value) || /opus/i.test(String(m.resolvedModel ?? "")),
  );
  return opus?.value ?? "sonnet";
}

// The model the daemon passes to query(): the three aliases always, anything
// the account lists (by value or by resolved id), and, while the list is
// unavailable, any well-formed claude-* id. No model asked for (a pane that
// stored none) is the best one; a model that was asked for and cannot be
// used is "sonnet".
export function resolveModelChoice(requested, state) {
  const r = String(requested ?? "").trim();
  if (!r) return bestClaudeModel(state);
  if (ALIASES.includes(r)) return r;
  const models = state?.models ?? [];
  if (models.some((m) => m.value === r || m.resolvedModel === r)) return r;
  if (state?.source !== "sdk" && ID_RE.test(r)) return r;
  return "sonnet";
}

// The CLI's notice that a request ended with a refusal and was run again on
// another model (SDK message system/model_refusal_fallback). Unless its
// scope is "local" (a subagent or a side question), the session stays on
// that model from then on: the daemon did not ask for it and names no
// fallback model itself. Seen in a real run of 20 review turns: asked opus,
// claude-opus-5-5 until the notice in the third turn, claude-opus-5 after.
// { session, log, event }: ids and the category only, none of its prose.
export function refusalFallback(msg) {
  const scope = msg.scope ?? "session";
  const category = msg.api_refusal_category ?? null;
  return {
    session: scope !== "local",
    log: `[agent] model refusal fallback original=${msg.original_model} fallback=${msg.fallback_model} scope=${scope} category=${category ?? "-"}`,
    event: {
      event: "model_fallback",
      original: msg.original_model,
      fallback: msg.fallback_model,
      scope,
      category,
    },
  };
}

// A model that declines a manuscript declines it on every turn (a real
// revision: 18 refusals of the same category in 18 turns), so taking the
// session back to the chosen model before each turn spends a refused request
// a turn. The record of a pane's refusals: the first one is followed by one
// more try on the chosen model; a second in a row of the same category, for
// the same model asked, and the document stays on the other model.
// prev: the record so far (or none). f: { asked, original, fallback, category }.
export function declineStep(prev, { asked, original, fallback, category = null }) {
  const again = prev && prev.asked === asked && prev.category === category;
  const count = again ? prev.count + 1 : 1;
  return { asked, original, fallback, category, count, stay: count >= 2 };
}

// The record, when it says to stay and is about the model chosen now.
export function staysOnFallback(record, chosen) {
  return record?.stay && record.asked === chosen ? record : null;
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
