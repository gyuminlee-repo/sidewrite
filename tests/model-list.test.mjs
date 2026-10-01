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

// Edits and Writing default to the best model the account can use: a row of
// the Opus family when the list has one, else sonnet.
test("bestClaudeModel: the Opus row of the list, else sonnet; nothing asked for resolves to it", async () => {
  const { bestClaudeModel, modelChange } = await import("../daemon/model-list.mjs");
  const row = (value, resolvedModel = null) => ({ value, displayName: value, resolvedModel });
  const withOpus = {
    source: "sdk",
    models: [row("sonnet", "claude-sonnet-5-5"), row("opus", "claude-opus-5-5"), row("haiku")],
  };
  assert.equal(bestClaudeModel(withOpus), "opus");
  // Named by its id only, or only in what the row resolves to.
  assert.equal(
    bestClaudeModel({ source: "sdk", models: [row("sonnet"), row("claude-opus-5-5")] }),
    "claude-opus-5-5",
  );
  assert.equal(
    bestClaudeModel({ source: "sdk", models: [row("big", "claude-opus-5-5"), row("sonnet")] }),
    "big",
  );
  const sdk = { source: "sdk", models: normalizeModels(SDK) };
  assert.equal(bestClaudeModel(sdk), "sonnet", "an account without an Opus model");
  // The list is not known: the alias, which query() always takes.
  assert.equal(bestClaudeModel({ source: "fallback", models: FALLBACK_MODELS }), "opus");
  assert.equal(bestClaudeModel(null), "sonnet");
  // No model asked for (a pane that stored none) is the best one.
  assert.equal(resolveModelChoice("", withOpus), "opus");
  assert.equal(resolveModelChoice(null, sdk), "sonnet");
  assert.deepEqual(modelChange({ requested: "", after: withOpus, running: "sonnet" }), {
    model: "opus",
    restart: true,
  });
  // A model that was asked for and is not listed is still replaced by sonnet.
  assert.equal(resolveModelChoice("gpt-5", withOpus), "sonnet");
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

// set_model (Codex review of stage 2.5): on the first connection the stored
// model is taken before the account's list is in, so a message sent
// meanwhile starts a session on it. When the list then replaces it, the
// session must follow: the pane shows the replacement. Whether it must is
// read from the model the session really runs on, not guessed from the list
// as it was: a session that started after the list came in is on the
// replacement already, and restarting it dropped the first message.
test("modelChange: a session restarts only when the model it runs on is not the one decided", async () => {
  const { modelChange } = await import("../daemon/model-list.mjs");
  const sdk = { source: "sdk", models: normalizeModels(SDK) };
  // A well-formed id the account does not list runs until the list is in.
  assert.equal(resolveModelChoice("claude-old-3", null), "claude-old-3");
  // The session started on it: it must follow the replacement.
  assert.deepEqual(
    modelChange({ requested: "claude-old-3", after: sdk, running: "claude-old-3" }),
    {
      model: "sonnet",
      restart: true,
    },
  );
  // The session started when the list was in: it runs on the replacement.
  assert.deepEqual(modelChange({ requested: "claude-old-3", after: sdk, running: "sonnet" }), {
    model: "sonnet",
    restart: false,
  });
  // No session, or one that has not read its model yet (it will read the
  // one decided here): nothing to restart.
  for (const running of [null, undefined]) {
    assert.equal(modelChange({ requested: "claude-old-3", after: sdk, running }).restart, false);
    assert.equal(modelChange({ requested: "opus", after: sdk, running }).restart, false);
  }
  // Nothing replaced: no restart.
  for (const requested of ["opus", "claude-fable-5"]) {
    const c = modelChange({ requested, after: sdk, running: requested });
    assert.deepEqual(c, { model: requested, restart: false });
  }
  // A later choice: a restart when the model changes.
  assert.deepEqual(modelChange({ requested: "opus", after: sdk, running: "sonnet" }), {
    model: "opus",
    restart: true,
  });
  // What the list as it was would have allowed plays no part.
  assert.equal(
    modelChange({
      prev: undefined,
      before: null,
      requested: "claude-old-3",
      after: sdk,
      running: "sonnet",
    }).restart,
    false,
  );
});

// The CLI itself changes the model of a session: when a request ends with a
// refusal it runs it again on another model and, unless the notice says the
// scope is local, keeps the session on that one (seen in a real run: asked
// opus, claude-opus-5-5 until the notice, claude-opus-5 for every turn after).
// The daemon says so in its log and to the pane, with ids only.
test("refusalFallback: the log line and the pane event of a model_refusal_fallback notice", async () => {
  const { refusalFallback } = await import("../daemon/model-list.mjs");
  const msg = {
    type: "system",
    subtype: "model_refusal_fallback",
    trigger: "refusal",
    direction: "retry",
    original_model: "claude-opus-5-5",
    fallback_model: "claude-opus-5",
    api_refusal_category: "bio",
    api_refusal_explanation: "prose that is not for a log",
    content: "prose",
  };
  assert.deepEqual(refusalFallback(msg), {
    session: true,
    log: "[agent] model refusal fallback original=claude-opus-5-5 fallback=claude-opus-5 scope=session category=bio",
    event: {
      event: "model_fallback",
      original: "claude-opus-5-5",
      fallback: "claude-opus-5",
      scope: "session",
      category: "bio",
    },
  });
  const local = refusalFallback({ ...msg, scope: "local", api_refusal_category: null });
  assert.equal(local.session, false);
  assert.match(local.log, / scope=local category=-$/);
  assert.equal(JSON.stringify(refusalFallback(msg)).includes("prose"), false);
});
