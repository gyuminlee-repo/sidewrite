import { test } from "node:test";
import assert from "node:assert/strict";
import {
  FALLBACK_MODELS,
  normalizeModels,
  fetchClaudeModels,
  resolveModelChoice,
  createModelCache,
} from "../daemon/model-list.mjs";

function fakeQuery(behavior, seen = {}) {
  return ({ prompt, options }) => {
    seen.prompt = prompt;
    seen.options = options;
    return {
      supportedModels: () =>
        behavior.reject
          ? Promise.reject(new Error(behavior.reject))
          : behavior.hang
            ? new Promise(() => {})
            : Promise.resolve(behavior.models),
      close: () => (seen.closed = true),
      async *[Symbol.asyncIterator]() {},
    };
  };
}
const SDK = [
  {
    value: "default",
    displayName: "Default (recommended)",
    description: "",
    resolvedModel: "claude-sonnet-5-5",
  },
  {
    value: "sonnet",
    displayName: "Sonnet",
    description: "balanced",
    resolvedModel: "claude-sonnet-5-5",
  },
  { value: "fable", displayName: "Fable", description: "new", resolvedModel: "claude-fable-5" },
  { value: "sonnet", displayName: "dup", description: "" },
];

test("normalize drops default and duplicates, keeps resolvedModel", () => {
  const m = normalizeModels(SDK);
  assert.deepEqual(
    m.map((x) => x.value),
    ["sonnet", "fable"],
  );
  assert.equal(m[1].resolvedModel, "claude-fable-5");
});

test("fetch: isolated probe query, closed after the list", async () => {
  const seen = {};
  const r = await fetchClaudeModels({ queryFn: fakeQuery({ models: SDK }, seen) });
  assert.equal(r.source, "sdk");
  assert.deepEqual(
    r.models.map((x) => x.value),
    ["sonnet", "fable"],
  );
  assert.equal(seen.closed, true);
  assert.deepEqual(seen.options.settingSources, []);
  assert.deepEqual(seen.options.mcpServers, {});
  assert.deepEqual(seen.options.tools, []);
  assert.equal(seen.options.persistSession, false);
  assert.equal(typeof seen.prompt[Symbol.asyncIterator], "function");
});

test("fetch: rejection and timeout fall back to aliases", async () => {
  const a = await fetchClaudeModels({ queryFn: fakeQuery({ reject: "boom" }) });
  assert.deepEqual([a.source, a.error], ["fallback", "boom"]);
  assert.deepEqual(a.models, FALLBACK_MODELS);
  const b = await fetchClaudeModels({ queryFn: fakeQuery({ hang: true }), timeoutMs: 20 });
  assert.equal(b.source, "fallback");
  assert.match(b.error, /timed out/);
});

test("resolveModelChoice: listed, alias, explicit id on fallback, else sonnet", () => {
  const sdk = { source: "sdk", models: normalizeModels(SDK) };
  assert.equal(resolveModelChoice("fable", sdk), "fable");
  assert.equal(resolveModelChoice("claude-fable-5", sdk), "claude-fable-5"); // resolvedModel match
  assert.equal(resolveModelChoice("opus", sdk), "opus"); // alias always allowed
  assert.equal(resolveModelChoice("gpt-5", sdk), "sonnet");
  const fb = { source: "fallback", models: FALLBACK_MODELS };
  assert.equal(resolveModelChoice("claude-fable-5", fb), "claude-fable-5");
  assert.equal(resolveModelChoice("rm -rf", fb), "sonnet");
});

test("cache: one fetch for concurrent callers, refetch after ttl or refresh", async () => {
  let n = 0;
  let t = 0;
  const cache = createModelCache({
    fetch: async () => ({ source: "sdk", models: [{ value: `m${++n}` }] }),
    ttlMs: 100,
    now: () => t,
  });
  const [a, b] = await Promise.all([cache.get(), cache.get()]);
  assert.equal(a, b);
  assert.equal(n, 1);
  t = 50;
  await cache.get();
  assert.equal(n, 1);
  t = 200;
  await cache.get();
  assert.equal(n, 2);
  await cache.get({ refresh: true });
  assert.equal(n, 3);
  assert.equal(cache.current().models[0].value, "m3");
});

// A slow CLI start (seen here: 18 s under load) times the probe out once; the
// alias list must not then stay for the whole ttl.
test("cache: a fallback result is kept only for the short fallback ttl", async () => {
  let n = 0;
  let t = 0;
  const cache = createModelCache({
    fetch: async () =>
      ++n === 1 ? { source: "fallback", models: FALLBACK_MODELS, error: "x" } : { source: "sdk" },
    ttlMs: 1000,
    fallbackTtlMs: 100,
    now: () => t,
  });
  assert.equal((await cache.get()).source, "fallback");
  t = 50;
  assert.equal((await cache.get()).source, "fallback");
  t = 150;
  assert.equal((await cache.get()).source, "sdk");
  assert.equal(n, 2);
});
