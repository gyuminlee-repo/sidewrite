import { test } from "node:test";
import assert from "node:assert/strict";
import { createOneshotRouter } from "../daemon/providers/oneshot-router.mjs";

test("routes by the pane's choice and labels the model for receipts", async () => {
  const calls = [];
  const mk = (id) => async (req) => (calls.push([id, req.model]), { ok: true, text: id });
  const router = createOneshotRouter({
    claude: mk("claude"),
    codex: mk("codex"),
    antigravity: mk("agy"),
    http: async (o) => (
      calls.push(["http", o.baseUrl, o.apiKey, o.model]),
      { ok: true, text: "h" }
    ),
    keys: { get: async (id) => ({ key: id === "groq" ? "gk" : null }) },
    getChoice: (key) =>
      ({ a: { id: "groq", model: "llama-x" }, b: { id: "codex", model: "gpt-6.1-sol" } })[key] ??
      null,
  });
  await router.forKey("a")({ system: "s", prompt: "p" });
  await router.forKey("b")({ system: "s", prompt: "p" });
  await router.forKey("z")({ system: "s", prompt: "p" });
  assert.deepEqual(calls, [
    ["http", "https://api.groq.com/openai/v1", "gk", "llama-x"],
    ["codex", "gpt-6.1-sol"],
    ["claude", "haiku"],
  ]);
  assert.equal((await router.pin("a")).model, "groq:llama-x");
  assert.equal((await router.pin("b")).model, "codex:gpt-6.1-sol");
  assert.equal((await router.pin("z")).model, "haiku");
});

function make(choices, over = {}) {
  const calls = [];
  const router = createOneshotRouter({
    claude: async (req) => (calls.push(["claude", req]), { ok: true, text: "c" }),
    codex: async (req) => (calls.push(["codex", req]), { ok: true, text: "x" }),
    antigravity: async (req) => (calls.push(["agy", req]), { ok: true, text: "g" }),
    http: async (o) => (calls.push(["http", o]), { ok: true, text: "h" }),
    keys: { get: async () => ({ key: null, base_url: null }) },
    getChoice: (key) => choices[key] ?? null,
    ...over,
  });
  return { router, calls };
}

test("the request is passed on whole; Claude keeps the chosen alias, the others their model", async () => {
  const { router, calls } = make({
    c: { id: "claude", model: "sonnet" },
    g: { id: "antigravity", model: "gemini-3.8-flash-high" },
    o: { id: "ollama", model: "qwen3.5:9b" },
  });
  const req = { system: "S", prompt: "P", expectJson: false, timeoutMs: 5 };
  await router.forKey("c")(req);
  await router.forKey("g")(req);
  await router.forKey("o")(req);
  assert.deepEqual(calls[0], ["claude", { ...req, model: "sonnet" }]);
  assert.deepEqual(calls[1], ["agy", { ...req, model: "gemini-3.8-flash-high" }]);
  assert.deepEqual(calls[2], [
    "http",
    { ...req, baseUrl: "http://localhost:11434/v1", apiKey: null, model: "qwen3.5:9b" },
  ]);
  assert.equal((await router.pin("c")).model, "sonnet");
  assert.equal((await router.pin("g")).model, "antigravity:gemini-3.8-flash-high");
  assert.equal((await router.pin("o")).model, "ollama:qwen3.5:9b");
});

test("a key provider without a key, a custom server without a URL and an unknown provider are error results", async () => {
  const { router, calls } = make({
    k: { id: "mistral", model: "m" },
    u: { id: "openai-compatible", model: "m" },
    x: { id: "nope", model: "m" },
  });
  const req = { system: "", prompt: "" };
  assert.deepEqual(await router.forKey("k")(req), {
    ok: false,
    error: "No API key for Mistral. Enter one in Setup, Providers.",
  });
  assert.match((await router.forKey("u")(req)).error, /No API key/);
  assert.deepEqual(await router.forKey("x")(req), { ok: false, error: "Unknown provider nope" });
  assert.equal(calls.length, 0, "nothing was sent anywhere, and nothing fell back to Claude");
  assert.equal((await router.pin("x")).model, "nope:m");

  const custom = make(
    { u: { id: "openai-compatible", model: "m" } },
    { keys: { get: async () => ({ key: "kk", base_url: "http://127.0.0.1:10100/v1" }) } },
  );
  await custom.router.forKey("u")(req);
  assert.deepEqual(
    [custom.calls[0][1].baseUrl, custom.calls[0][1].apiKey],
    ["http://127.0.0.1:10100/v1", "kk"],
  );
  const noUrl = make(
    { u: { id: "openai-compatible", model: "m" } },
    { keys: { get: async () => ({ key: "kk", base_url: null }) } },
  );
  assert.match((await noUrl.router.forKey("u")(req)).error, /No server URL/);
});

test("the choice is read when the request runs, a failing key store is an error result, no key in a label", async () => {
  const choices = {};
  const SECRET = "sk-secret-0123456789";
  const { router, calls } = make(choices, {
    keys: {
      get: async (id) => {
        if (id === "groq") throw new Error("provider-keys file is not valid JSON: /x");
        return { key: SECRET, base_url: null };
      },
    },
  });
  const shot = router.forKey("p");
  await shot({ system: "", prompt: "" });
  choices.p = { id: "codex", model: "gpt-6.1-sol" };
  await shot({ system: "", prompt: "" });
  assert.deepEqual(
    calls.map((c) => c[0]),
    ["claude", "codex"],
  );
  choices.p = { id: "groq", model: "m" };
  const bad = await shot({ system: "", prompt: "" });
  assert.deepEqual([bad.ok, /not valid JSON/.test(bad.error)], [false, true]);
  choices.p = { id: "gemini", model: "gemini-x" };
  await shot({ system: "", prompt: "" });
  assert.equal(calls.at(-1)[1].apiKey, SECRET);
  assert.equal((await router.pin("p")).model.includes(SECRET), false);
  assert.equal((await router.pin("p")).model, "gemini:gemini-x");
});

// Codex review of stage 2.5: a job with several requests (a ledger check in
// batches, a split with its requirement extraction) must not change provider
// half way, and its receipts must name the provider that ran it.
test("pin: one job keeps its provider and model while the pane's choice changes", async () => {
  const choices = { p: { id: "ollama", model: "qwen3.5:9b" } };
  const lines = [];
  const { router, calls } = make(choices, { log: (l) => lines.push(l) });
  const job = await router.pin("p");
  assert.deepEqual([job.provider, job.model], ["ollama", "ollama:qwen3.5:9b"]);
  await job.run({ system: "", prompt: "batch 1" });
  choices.p = { id: "codex", model: "gpt-6.1-sol" }; // changed in the pane mid-job
  await job.run({ system: "", prompt: "batch 2" });
  choices.p = null; // back to the default
  await job.run({ system: "", prompt: "batch 3" });
  assert.deepEqual(
    calls.map((c) => [c[0], c[1].model, c[1].prompt]),
    [
      ["http", "qwen3.5:9b", "batch 1"],
      ["http", "qwen3.5:9b", "batch 2"],
      ["http", "qwen3.5:9b", "batch 3"],
    ],
  );
  assert.equal(job.model, "ollama:qwen3.5:9b");
  // The next job takes the choice of its own start.
  const next = await router.pin("p");
  assert.deepEqual([next.provider, next.model], ["claude", "haiku"]);
  // One log line per request, naming the provider and no text of the request.
  assert.deepEqual(
    lines,
    Array(3).fill("[providers] judgment provider=ollama model=ollama:qwen3.5:9b"),
  );
});

test("pin: a choice that cannot be resolved sends nothing anywhere and every request is that error", async () => {
  const lines = [];
  const { router, calls } = make(
    {},
    {
      getChoice: async () => {
        throw new Error("The task pane has not confirmed its provider for judgments yet.");
      },
      log: (l) => lines.push(l),
    },
  );
  const job = await router.pin("p");
  assert.deepEqual([job.provider, job.model], [null, "none"]);
  const r = await job.run({ system: "", prompt: "secret manuscript text" });
  assert.deepEqual(r, {
    ok: false,
    error: "The task pane has not confirmed its provider for judgments yet.",
  });
  assert.deepEqual(await router.forKey("p")({ system: "", prompt: "x" }), r);
  assert.equal(calls.length, 0, "no provider was called, Claude included");
  assert.equal(lines.join("\n").includes("secret manuscript"), false);
  assert.ok(lines.every((l) => l.startsWith("[providers] judgment refused: ")));
});
