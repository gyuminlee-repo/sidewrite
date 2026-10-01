// Which providers can be used right now, and why not (P14). One row per
// preset for the task pane:
//   { id, tier, label, roles, policy, default_enabled,
//     available, turn_ready, reason, models, has_key, key_source, base_url? }
// available: one-shot judgments can go to it. turn_ready: review turns can.
// reason: why not, or a note for an available one; null when there is
// nothing to say. models: [{ value, displayName }].
//
// turn_note: why review turns are off for an available provider, or null.
//
// A probe only reads: `codex login status`, `agy models`, GET requests, the
// Codex model cache, and whether a key is stored. It never
// writes a CLI's configuration and never asks a model anything. Every command
// and request has a timeout and every provider a deadline, so a row always
// comes back; a failure is a reason in the row, never a throw. A key is used
// for the Authorization header only and is in no row.

import { execFile } from "node:child_process";
import { readFile } from "node:fs/promises";
import { homedir } from "node:os";
import { join } from "node:path";
import { PRESETS, canRun } from "./registry.mjs";
import { resolveCliBin } from "./cli-exec.mjs";

const CMD_TIMEOUT_MS = 10_000;
const AGY_MODELS_TIMEOUT_MS = 30_000;
const OLLAMA_TIMEOUT_MS = 2_000;
const HTTP_TIMEOUT_MS = 5_000;
const DEADLINE_MS = 35_000;
const MAX_MODELS = 200;

// Why a CLI provider is not offered for review turns. Checked with the real
// CLIs on 2026-10-01 against the daemon's MCP endpoint, with the read-only
// sandbox and no approval flag (tests/fixtures/codex_exec_turn_approval.jsonl,
// agy_turn_denied.jsonl): each refused the office tool call itself. This
// add-in never passes a flag that skips an approval, so both stay providers
// for judgments. The texts quote what the CLI said.
export const TURN_BLOCKED = Object.freeze({
  codex:
    'Judgments only. Codex CLI will not call the office tools without an approval: "MCP tool call requires approval, but approval policy is never".',
  antigravity:
    "Judgments only. agy print mode denies the office tools (permission mode request-review, denied_actions: CallMcpTool) unless permission prompts are skipped, which this add-in never does.",
});
const NO_TOOL_LOOP = "Judgments only: no tool loop for this provider in this stage.";

// CC_OFFICE_CLI_TURNS: a comma list of CLI providers to offer for review
// turns all the same, for a test after a CLI changed its approval rules. It
// adds no flag to any command; the CLI still decides on each tool call.
function cliTurnsTried(env) {
  return new Set(
    String(env.CC_OFFICE_CLI_TURNS ?? "")
      .split(",")
      .map((s) => s.trim())
      .filter((s) => s in TURN_BLOCKED),
  );
}

// { code, stdout, stderr, timedOut }. Rejects only when the command cannot
// be started (ENOENT: not installed).
export function execFileDefault(cmd, args, { timeout }) {
  return new Promise((resolve, reject) => {
    const child = execFile(
      cmd,
      args,
      { timeout, windowsHide: true, maxBuffer: 4 * 1024 * 1024 },
      (err, stdout, stderr) => {
        if (!err) return resolve({ code: 0, stdout, stderr });
        if (err.code === "ENOENT") return reject(err);
        resolve({
          code: typeof err.code === "number" ? err.code : -1,
          stdout: stdout ?? "",
          stderr: stderr ?? "",
          timedOut: err.killed === true,
        });
      },
    );
    child.stdin?.end(); // a CLI that reads stdin must not wait for it
  });
}

// `agy models`: a status line, then "<slug>\t<display name>" lines.
export function parseAgyModels(stdout) {
  const out = [];
  for (const line of String(stdout ?? "").split(/\r?\n/)) {
    const [value, name] = line.split("\t");
    if (!name || !/^[\w.:-]+$/.test(value.trim())) continue;
    out.push({ value: value.trim(), displayName: name.trim() });
  }
  return out;
}

// ~/.codex/models_cache.json (a private cache of the CLI; its form is not
// promised): models[] with slug, display_name, visibility "list" | "hide".
export function codexModelsFromCache(text) {
  let data;
  try {
    data = JSON.parse(String(text ?? ""));
  } catch {
    return [];
  }
  return (Array.isArray(data?.models) ? data.models : [])
    .filter((m) => m?.visibility === "list" && typeof m.slug === "string")
    .map((m) => ({ value: m.slug, displayName: String(m.display_name ?? m.slug) }));
}

// The top-level `model = "..."` of ~/.codex/config.toml (before any table).
function codexModelFromConfig(text) {
  const top = String(text ?? "").split(/^\s*\[/m)[0];
  const m = /^\s*model\s*=\s*"([^"]+)"/m.exec(top);
  return m ? [{ value: m[1], displayName: m[1] }] : [];
}

// GET /api/tags of Ollama. A model whose capabilities lack "completion"
// (an embedding model) cannot judge.
export function ollamaChatModels(tags) {
  return (Array.isArray(tags?.models) ? tags.models : [])
    .filter((m) => typeof m?.name === "string")
    .filter((m) => !Array.isArray(m.capabilities) || m.capabilities.includes("completion"))
    .map((m) => ({ value: m.name, displayName: m.name }));
}

const netCode = (e) =>
  e?.name === "TimeoutError" ? "timeout" : (e?.cause?.code ?? e?.code ?? e?.name ?? "error");

// bin: what resolveCliBin returned ({ cmd, prefixArgs } or { cmd: null, error }).
async function run(execFileFn, bin, args, timeout) {
  if (!bin.cmd) return { code: -1, stdout: "", stderr: bin.error, unstartable: bin.error };
  try {
    return await execFileFn(bin.cmd, [...bin.prefixArgs, ...args], { timeout });
  } catch (e) {
    if (e?.code === "ENOENT") return { missing: true };
    return { code: -1, stdout: "", stderr: String(e?.message ?? e) };
  }
}

async function probeClaude({ claudeModels }) {
  const r = await claudeModels();
  return {
    available: true,
    turn_ready: true,
    reason: r.source === "sdk" ? null : "Model list unavailable; aliases only",
    models: (r.models ?? []).map((m) => ({
      value: m.value,
      displayName: m.displayName ?? m.value,
    })),
  };
}

async function probeCodex({ execFileFn, readFileFn, env }) {
  const bin = resolveCliBin("codex", { env, envVar: "CC_OFFICE_CODEX_BIN" });
  const r = await run(execFileFn, bin, ["login", "status"], CMD_TIMEOUT_MS);
  if (r.missing) return { reason: "Codex CLI is not installed" };
  if (r.unstartable) return { reason: r.unstartable };
  if (r.timedOut) return { reason: "codex login status timed out" };
  if (r.code !== 0 || !/Logged in/i.test(`${r.stdout}${r.stderr ?? ""}`)) {
    return { reason: "Not signed in (run: codex login)" };
  }
  const dir = join(homedir(), ".codex");
  let models = await readFileFn(join(dir, "models_cache.json"), "utf8").then(
    codexModelsFromCache,
    () => [],
  );
  if (!models.length) {
    models = await readFileFn(join(dir, "config.toml"), "utf8").then(
      codexModelFromConfig,
      () => [],
    );
  }
  // No list: the CLI runs with the model of its own configuration.
  return { available: true, turn_ready: true, models };
}

async function probeAgy({ execFileFn, env }) {
  const bin = resolveCliBin("agy", { env, envVar: "CC_OFFICE_AGY_BIN" });
  const r = await run(execFileFn, bin, ["models"], AGY_MODELS_TIMEOUT_MS);
  if (r.missing) return { reason: "Antigravity CLI (agy) is not installed" };
  if (r.unstartable) return { reason: r.unstartable };
  if (r.timedOut) return { reason: "agy models timed out" };
  const models = parseAgyModels(r.stdout);
  if (r.code !== 0 || !models.length)
    return { reason: "Not signed in, or no model listed (run: agy)" };
  return { available: true, turn_ready: true, models };
}

async function probeOllama({ fetchFn }, preset) {
  const origin = preset.baseUrl.replace(/\/v1$/, "");
  let res;
  try {
    res = await fetchFn(`${origin}/api/tags`, { signal: AbortSignal.timeout(OLLAMA_TIMEOUT_MS) });
  } catch {
    return { reason: `Ollama is not running (no answer from ${origin})` };
  }
  if (!res.ok) return { reason: `Ollama is not running (${origin} answered HTTP ${res.status})` };
  // The signal of the request also ends a body that stalls after the headers.
  let tags;
  try {
    tags = await res.json();
  } catch (e) {
    return { reason: `Ollama did not finish its answer (${netCode(e)})` };
  }
  const models = ollamaChatModels(tags);
  if (!models.length) return { reason: "Ollama has no chat model (run: ollama pull <model>)" };
  return { available: true, models };
}

async function getJson(fetchFn, url, key) {
  try {
    const res = await fetchFn(url, {
      headers: { authorization: `Bearer ${key}` },
      signal: AbortSignal.timeout(HTTP_TIMEOUT_MS),
    });
    if (res.status === 401 || res.status === 403) {
      return { reason: `Key rejected (HTTP ${res.status})` };
    }
    if (!res.ok) return { reason: `Server answered HTTP ${res.status}` };
    return { json: await res.json() };
  } catch (e) {
    return { reason: `Connection failed (${netCode(e)})` };
  }
}

async function probeKey({ fetchFn }, preset, stored) {
  if (!stored.key) return { reason: "No API key" };
  const base = preset.baseUrl ?? stored.base_url;
  if (!base) return { reason: "No server URL" };
  // OpenRouter lists its models to anyone; only /key says the key is good.
  if (preset.keyCheck) {
    const check = await getJson(fetchFn, `${base}${preset.keyCheck}`, stored.key);
    if (check.reason) return { reason: check.reason };
  }
  const list = await getJson(fetchFn, `${base}/models`, stored.key);
  if (list.reason) return { reason: list.reason };
  const rows = Array.isArray(list.json?.data) ? list.json.data : [];
  const models = rows
    .filter((m) => typeof m?.id === "string")
    .slice(0, MAX_MODELS)
    .map((m) => ({ value: m.id, displayName: m.id }));
  return { available: true, models };
}

function probeOne(deps, preset, stored) {
  if (preset.id === "claude") return probeClaude(deps);
  if (preset.id === "codex") return probeCodex(deps);
  if (preset.id === "antigravity") return probeAgy(deps);
  if (preset.tier === "local") return probeOllama(deps, preset);
  return probeKey(deps, preset, stored);
}

export async function probeAll({
  claudeModels,
  keys,
  execFileFn = execFileDefault,
  readFileFn = readFile,
  fetchFn = fetch,
  env = process.env,
  deadlineMs = DEADLINE_MS,
} = {}) {
  const deps = { claudeModels, execFileFn, readFileFn, fetchFn, env };
  const tried = cliTurnsTried(env);
  return Promise.all(
    PRESETS.map(async (preset) => {
      let stored = { key: null, base_url: null, source: null };
      let timer;
      const found = await Promise.race([
        (async () => {
          if (preset.tier === "key") stored = { ...stored, ...(await keys.get(preset.id)) };
          return probeOne(deps, preset, stored);
        })(),
        new Promise((resolve) => {
          timer = setTimeout(
            () => resolve({ reason: `Probe timed out after ${Math.round(deadlineMs / 1000)} s` }),
            deadlineMs,
          );
        }),
      ]).then(
        (r) => r,
        (e) => ({ reason: `Probe failed: ${String(e?.message ?? e).slice(0, 200)}` }),
      );
      clearTimeout(timer);
      const available = found.available === true;
      const base_url = preset.baseUrl ?? stored.base_url ?? null;
      const blocked = preset.id in TURN_BLOCKED && !tried.has(preset.id);
      const turn_note = !canRun(preset, "turn")
        ? NO_TOOL_LOOP
        : blocked
          ? TURN_BLOCKED[preset.id]
          : tried.has(preset.id)
            ? "Offered for review turns because CC_OFFICE_CLI_TURNS lists it."
            : null;
      return {
        id: preset.id,
        tier: preset.tier,
        label: preset.label,
        roles: [...preset.roles],
        policy: preset.policy,
        default_enabled: preset.defaultEnabled,
        available,
        turn_ready: available && canRun(preset, "turn") && !blocked && found.turn_ready === true,
        turn_note,
        reason: found.reason ?? null,
        models: found.models ?? [],
        has_key: !!stored.key,
        key_source: stored.source ?? null,
        ...(base_url ? { base_url } : {}),
      };
    }),
  );
}

// The rows of the last probe, kept for a minute (the pane asks on every
// connect; `agy models` alone takes seconds).
export function createProbeCache({ probe, ttlMs = 60_000, now = Date.now }) {
  let value = null;
  let at = -Infinity;
  let inflight = null;
  return {
    current: () => value ?? [],
    async get({ refresh = false } = {}) {
      if (!refresh && value && now() - at < ttlMs) return value;
      inflight ??= probe().then(
        (v) => {
          value = v;
          at = now();
          inflight = null;
          return v;
        },
        (e) => {
          inflight = null;
          throw e;
        },
      );
      return inflight;
    },
  };
}
