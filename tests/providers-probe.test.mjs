import { test } from "node:test";
import assert from "node:assert/strict";
import {
  probeAll,
  parseAgyModels,
  codexModelsFromCache,
  ollamaChatModels,
  createProbeCache,
  AGY_NOT_OFFERED,
} from "../daemon/providers/probe.mjs";

const AGY_OUT =
  "Fetching available models...\ngemini-3.8-flash-high\tGemini 3.8 Flash (High)\nclaude-sonnet-4-6\tClaude Sonnet 4.6 (Thinking)\n";
const CODEX_CACHE = JSON.stringify({
  models: [
    { slug: "gpt-6.1-sol", display_name: "GPT-6.1-Sol", visibility: "list" },
    { slug: "gpt-reserve", display_name: "x", visibility: "hide" },
  ],
});
const TAGS = {
  models: [
    { name: "qwen3.5:9b", capabilities: ["completion", "tools"] },
    { name: "bge-m3:latest", capabilities: ["embedding"] },
  ],
};

test("parsers", () => {
  assert.deepEqual(
    parseAgyModels(AGY_OUT).map((m) => m.value),
    ["gemini-3.8-flash-high", "claude-sonnet-4-6"],
  );
  assert.equal(parseAgyModels(AGY_OUT)[0].displayName, "Gemini 3.8 Flash (High)");
  assert.deepEqual(
    codexModelsFromCache(CODEX_CACHE).map((m) => m.value),
    ["gpt-6.1-sol"],
  );
  assert.deepEqual(codexModelsFromCache("{ broken"), []);
  assert.deepEqual(
    ollamaChatModels(TAGS).map((m) => m.value),
    ["qwen3.5:9b"],
  );
  assert.deepEqual(ollamaChatModels(null), []);
});

function deps(over = {}) {
  return {
    claudeModels: async () => ({
      source: "sdk",
      models: [{ value: "sonnet", displayName: "Sonnet" }],
    }),
    execFileFn: async (cmd, args) => {
      if (cmd === "codex" && args.join(" ") === "login status")
        return { code: 0, stdout: "Logged in using ChatGPT\n" };
      if (cmd === "agy" && args[0] === "models") return { code: 0, stdout: AGY_OUT };
      if (cmd === "agy" && args.join(" ") === "mcp list")
        return { code: 0, stdout: "No MCP servers configured.\n" };
      throw Object.assign(new Error("not found"), { code: "ENOENT" });
    },
    readFileFn: async () => CODEX_CACHE,
    fetchFn: async (url) =>
      url.endsWith("/api/tags")
        ? { ok: true, status: 200, json: async () => TAGS }
        : { ok: false, status: 401, json: async () => ({}) },
    keys: {
      get: async () => ({ key: null }),
      publicView: async () => ({ has_key: false, source: null }),
    },
    env: {},
    ...over,
  };
}
const byId = (rows) => Object.fromEntries(rows.map((r) => [r.id, r]));

test("probeAll: available, turn_ready and reasons", async () => {
  const rows = await probeAll(deps());
  const by = byId(rows);
  assert.equal(rows.length, 10);
  assert.equal(by.claude.available, true);
  assert.equal(by.claude.turn_ready, true);
  assert.equal(by.codex.available, true);
  assert.equal(by.codex.reason, null);
  assert.deepEqual(
    by.codex.models.map((m) => m.value),
    ["gpt-6.1-sol"],
  );
  // agy is installed and signed in, and still offered for nothing: its
  // tools cannot be switched off (measured 2026-10-01, agy.mjs).
  assert.equal(by.antigravity.available, false);
  assert.equal(by.antigravity.reason, AGY_NOT_OFFERED);
  assert.equal(
    AGY_NOT_OFFERED,
    "Not offered: the built-in tools of Antigravity CLI (agy) cannot be switched off, so manuscript text would sit next to them.",
  );
  assert.ok(by.antigravity.models.length > 0, "its models are still listed");
  assert.equal(by.ollama.available, true);
  assert.equal(by.ollama.turn_ready, false, "no review turns for a local model in this stage");
  assert.equal(by.openrouter.available, false);
  assert.match(by.openrouter.reason, /No API key/);
  assert.match(by["openai-compatible"].reason, /No API key/);
  // Every row carries what the pane shows: the tier, the data policy, off by default for keys.
  for (const r of rows) {
    assert.ok(r.policy.length > 10, r.id);
    assert.equal(r.default_enabled, r.tier !== "key", r.id);
    assert.equal(typeof r.label, "string");
    assert.equal(r.available || !!r.reason, true, `${r.id}: unavailable needs a reason`);
  }
});

test("probeAll: missing CLI, logged out, network error", async () => {
  const rows = await probeAll(
    deps({
      execFileFn: async (cmd) =>
        cmd === "codex"
          ? { code: 1, stdout: "Not logged in\n" }
          : Promise.reject(Object.assign(new Error("x"), { code: "ENOENT" })),
      fetchFn: async () => {
        throw Object.assign(new TypeError("fetch failed"), { cause: { code: "ECONNREFUSED" } });
      },
      keys: {
        get: async (id) => ({ key: id === "deepseek" ? "k" : null }),
        publicView: async () => ({ has_key: true }),
      },
    }),
  );
  const by = byId(rows);
  assert.match(by.codex.reason, /Not signed in/);
  assert.match(by.antigravity.reason, /not installed/);
  assert.match(by.ollama.reason, /not running/);
  assert.match(by.deepseek.reason, /Connection failed \(ECONNREFUSED\)/);
  for (const id of ["codex", "antigravity", "ollama", "deepseek"]) {
    assert.deepEqual([by[id].available, by[id].turn_ready], [false, false], id);
  }
});

// DeepSeek direct is blocked by TLS inspection on the development machine:
// the probe must say so in the row, not throw.
test("a TLS failure is an unavailable row with the reason", async () => {
  const rows = await probeAll(
    deps({
      fetchFn: async (url) => {
        if (url.startsWith("https://api.deepseek.com")) {
          throw Object.assign(new TypeError("fetch failed"), {
            cause: { code: "SELF_SIGNED_CERT_IN_CHAIN" },
          });
        }
        return { ok: true, status: 200, json: async () => ({ data: [{ id: "m-1" }] }) };
      },
      keys: {
        get: async (id) => ({ key: ["deepseek", "groq"].includes(id) ? "k" : null }),
        publicView: async () => ({}),
      },
    }),
  );
  const by = byId(rows);
  assert.equal(by.deepseek.available, false);
  assert.equal(by.deepseek.reason, "Connection failed (SELF_SIGNED_CERT_IN_CHAIN)");
  assert.equal(by.groq.available, true);
  assert.deepEqual(by.groq.models, [{ value: "m-1", displayName: "m-1" }]);
});

test("key providers: bearer header, rejected key, OpenRouter key check, custom URL, no key in any row", async () => {
  const SECRET = "sk-secret-0123456789";
  const seen = [];
  const rows = await probeAll(
    deps({
      fetchFn: async (url, init) => {
        seen.push([url, init?.headers?.authorization ?? null]);
        if (url.endsWith("/api/tags")) return { ok: true, status: 200, json: async () => TAGS };
        if (url === "https://openrouter.ai/api/v1/key")
          return { ok: false, status: 401, json: async () => ({}) };
        if (url.startsWith("https://api.mistral.ai"))
          return { ok: false, status: 503, json: async () => ({}) };
        return { ok: true, status: 200, json: async () => ({ data: [{ id: "a" }, { id: "b" }] }) };
      },
      keys: {
        get: async (id) =>
          id === "openai-compatible"
            ? { key: SECRET, base_url: "http://127.0.0.1:10100/v1" }
            : { key: ["openrouter", "mistral", "gemini"].includes(id) ? SECRET : null },
        publicView: async () => ({}),
      },
    }),
  );
  const by = byId(rows);
  assert.equal(by.openrouter.reason, "Key rejected (HTTP 401)");
  assert.equal(by.mistral.reason, "Server answered HTTP 503");
  assert.equal(by.gemini.available, true);
  assert.equal(by["openai-compatible"].available, true);
  assert.equal(by["openai-compatible"].base_url, "http://127.0.0.1:10100/v1");
  assert.ok(seen.some(([u]) => u === "http://127.0.0.1:10100/v1/models"));
  for (const [url, auth] of seen) {
    if (!url.endsWith("/api/tags")) assert.equal(auth, `Bearer ${SECRET}`, url);
  }
  assert.equal(JSON.stringify(rows).includes(SECRET), false, "no key in the rows");
  assert.equal(JSON.stringify(rows).includes(SECRET.slice(0, 6)), false);
  assert.deepEqual([by.gemini.has_key, by.groq.has_key], [true, false]);

  const noUrl = byId(
    await probeAll(
      deps({
        keys: { get: async () => ({ key: SECRET, base_url: null }), publicView: async () => ({}) },
      }),
    ),
  );
  assert.match(noUrl["openai-compatible"].reason, /No server URL/);
});

test("probes are read-only and bounded: only status commands run; a hang or a throw becomes a reason", async () => {
  const ran = [];
  await probeAll(
    deps({
      execFileFn: async (cmd, args, opts) => {
        ran.push(`${cmd} ${args.join(" ")}`);
        assert.ok(opts.timeout > 0 && opts.timeout <= 30_000, "every command has a timeout");
        return { code: 0, stdout: cmd === "agy" ? AGY_OUT : "Logged in using ChatGPT\n" };
      },
    }),
  );
  assert.deepEqual(ran.sort(), ["agy mcp list", "agy models", "codex login status"]);

  const t0 = Date.now();
  const rows = await probeAll(
    deps({
      execFileFn: () => new Promise(() => {}), // never answers
      fetchFn: async () => {
        throw Object.assign(new Error("The operation was aborted"), { name: "TimeoutError" });
      },
      claudeModels: async () => {
        throw new Error("boom");
      },
      deadlineMs: 40,
    }),
  );
  assert.ok(Date.now() - t0 < 2000, "did not hang");
  const by = byId(rows);
  assert.match(by.codex.reason, /timed out/);
  assert.match(by.antigravity.reason, /timed out/);
  assert.match(by.ollama.reason, /not running/);
  assert.match(by.claude.reason, /boom/);
  assert.equal(by.codex.available, false);

  const slow = byId(
    await probeAll(
      deps({
        execFileFn: async (cmd) =>
          cmd === "codex" ? { code: -1, stdout: "", timedOut: true } : { code: 0, stdout: AGY_OUT },
      }),
    ),
  );
  assert.match(slow.codex.reason, /timed out/);
});

test("bin overrides come from the environment", async () => {
  const ran = [];
  await probeAll(
    deps({
      env: { CC_OFFICE_CODEX_BIN: "/opt/codex", CC_OFFICE_AGY_BIN: "/opt/agy" },
      execFileFn: async (cmd) => (ran.push(cmd), { code: 1, stdout: "" }),
    }),
  );
  assert.deepEqual([...new Set(ran)].sort(), ["/opt/agy", "/opt/codex"]);
});

test("a codex model falls back to the top-level model of config.toml", async () => {
  const by = byId(
    await probeAll(
      deps({
        readFileFn: async (p) => {
          if (p.endsWith("models_cache.json")) throw new Error("ENOENT");
          return 'x = 1\nmodel = "gpt-6-astra"\n[other]\nmodel = "no"\n';
        },
      }),
    ),
  );
  assert.deepEqual(
    by.codex.models.map((m) => m.value),
    ["gpt-6-astra"],
  );
});

// Review turns. Claude calls the office tools itself. Codex is given no
// tools: it proposes and the daemon applies (turn_mode "proposal"). agy is
// offered for nothing. A local or key provider has no review turns in this
// stage, and its row says so.
test("review turns: Claude as an agent, Codex by proposals, nobody else", async () => {
  const by = byId(await probeAll(deps()));
  assert.deepEqual(
    [by.claude.turn_ready, by.claude.turn_mode, by.claude.turn_note],
    [true, "agent", null],
  );
  assert.deepEqual(
    [by.codex.available, by.codex.turn_ready, by.codex.turn_mode, by.codex.turn_note],
    [true, true, "proposal", null],
  );
  assert.deepEqual(
    [by.antigravity.available, by.antigravity.turn_ready, by.antigravity.turn_mode],
    [false, false, null],
  );
  assert.equal(by.antigravity.reason, AGY_NOT_OFFERED);
  const note = "Judgments only: review turns are not offered for this provider in this stage.";
  for (const id of ["ollama", "groq", "openrouter", "openai-compatible"]) {
    assert.deepEqual([by[id].turn_ready, by[id].turn_mode, by[id].turn_note], [false, null, note]);
  }
  // Codex that is not signed in runs no review turn either.
  const off = byId(await probeAll(deps({ execFileFn: async () => ({ code: 1, stdout: "" }) })));
  assert.deepEqual([off.codex.available, off.codex.turn_ready], [false, false]);
});

// The switch that once offered a CLI for review turns through the office
// tools is gone: nothing in the environment changes a row.
test("no environment variable changes which provider runs review turns", async () => {
  const plain = await probeAll(deps());
  const name = ["CC_OFFICE", "CLI", "TURNS"].join("_");
  const env = { [name]: "codex, antigravity, ollama" };
  assert.deepEqual(await probeAll(deps({ env })), plain);
});

test("probe cache: one probe for concurrent callers, again after the ttl or a refresh", async () => {
  let n = 0;
  let t = 0;
  const cache = createProbeCache({
    probe: async () => [{ id: `p${++n}` }],
    ttlMs: 100,
    now: () => t,
  });
  assert.deepEqual(cache.current(), []);
  const [a, b] = await Promise.all([cache.get(), cache.get()]);
  assert.equal(a, b);
  t = 50;
  await cache.get();
  assert.equal(n, 1);
  t = 150;
  await cache.get();
  assert.equal(n, 2);
  await cache.get({ refresh: true });
  assert.equal(n, 3);
  assert.equal(cache.current()[0].id, "p3");
});

// The probe's requests carry AbortSignal.timeout, which also ends a body
// read: a server that sends headers and stalls is a reason, not a hang.
// (Ollama's limit is 2 s; the key providers use the same signal for /models.)
test("a server that stalls the body of its model list is a row with the reason", async () => {
  const r = await stalledBody("probe");
  assert.ok(r.ms < 4_000, `ended at Ollama's limit, not at the body (${r.ms} ms)`);
  assert.equal(r.ollama.available, false);
  assert.equal(r.ollama.reason, "Ollama did not finish its answer (timeout)");
});

// The real requests run in a process of their own (tests/helpers/
// stalled-body.mjs says why) and print what they got.
async function stalledBody(scenario) {
  const { execFile } = await import("node:child_process");
  const { fileURLToPath } = await import("node:url");
  const helper = fileURLToPath(new URL("./helpers/stalled-body.mjs", import.meta.url));
  const stdout = await new Promise((resolve, reject) =>
    execFile(process.execPath, [helper, scenario], { timeout: 30_000 }, (err, out, errText) =>
      err ? reject(new Error(`${err.message}\n${errText}`)) : resolve(out),
    ),
  );
  return JSON.parse(stdout);
}

// agy is not offered whatever its state; the row still says what is known.
test("Antigravity: not offered for anything; MCP servers it has are named", async () => {
  const tried = byId(await probeAll(deps()));
  assert.deepEqual([tried.antigravity.available, tried.antigravity.turn_ready], [false, false]);
  const ran = [];
  const withServers = byId(
    await probeAll(
      deps({
        execFileFn: async (cmd, args) => {
          ran.push(`${cmd} ${args.join(" ")}`);
          if (cmd === "agy" && args[0] === "models") return { code: 0, stdout: AGY_OUT };
          if (cmd === "agy" && args.join(" ") === "mcp list")
            return { code: 0, stdout: "notes  http://127.0.0.1:9000/sse  enabled\n" };
          return { code: 0, stdout: "Logged in using ChatGPT\n" };
        },
      }),
    ),
  );
  assert.ok(ran.includes("agy mcp list"), "read-only: list, never add or remove");
  assert.equal(
    ran.some((c) => /mcp (add|remove|enable|disable)/.test(c)),
    false,
  );
  assert.equal(
    withServers.antigravity.reason,
    `${AGY_NOT_OFFERED} agy also has MCP servers registered (notes http://127.0.0.1:9000/sse enabled).`,
  );
  // Not installed is still said as that.
  const none = byId(
    await probeAll(
      deps({
        execFileFn: async () => {
          throw Object.assign(new Error("not found"), { code: "ENOENT" });
        },
      }),
    ),
  );
  assert.match(none.antigravity.reason, /not installed/);
});
