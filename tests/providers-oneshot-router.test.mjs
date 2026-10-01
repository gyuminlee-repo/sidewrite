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
    {
      ...req,
      baseUrl: "http://localhost:11434/v1",
      apiKey: null,
      model: "qwen3.5:9b",
      // Ollama: no thinking for a judgment, and a JSON object when asked.
      body: { reasoning_effort: "none" },
      jsonBody: { response_format: { type: "json_object" } },
    },
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

// Writing (a response draft, a translation) is a job of its own role: the
// router asks for the pane's choice of that role, and where the choice names
// no Claude model the default of the role applies, not the judgment model.
test("a writing job uses the pane's writing choice and the writing default of Claude", async () => {
  const calls = [];
  const asked = [];
  const lines = [];
  const choices = {
    oneshot: { id: "ollama", model: "qwen3.5:9b" },
    writing: null,
  };
  const router = createOneshotRouter({
    claude: async (req) => (calls.push(["claude", req.model]), { ok: true, text: "c" }),
    codex: async (req) => (calls.push(["codex", req.model]), { ok: true, text: "x" }),
    antigravity: async () => ({ ok: true, text: "g" }),
    http: async (o) => (calls.push(["http", o.model]), { ok: true, text: "h" }),
    keys: { get: async () => ({ key: null }) },
    getChoice: (key, role) => (asked.push([key, role]), choices[role]),
    claudeDefault: (role) => (role === "writing" ? "opus" : "haiku"),
    log: (l) => lines.push(l),
  });
  const req = { system: "S", prompt: "P" };
  const w = await router.pin("k", { role: "writing" });
  assert.deepEqual([w.role, w.provider, w.model], ["writing", "claude", "opus"]);
  await w.run(req);
  const j = await router.pin("k");
  assert.deepEqual([j.role, j.provider, j.model], ["oneshot", "ollama", "ollama:qwen3.5:9b"]);
  await j.run(req);
  assert.deepEqual(asked, [
    ["k", "writing"],
    ["k", "oneshot"],
  ]);
  assert.deepEqual(calls, [
    ["claude", "opus"],
    ["http", "qwen3.5:9b"],
  ]);
  // The log says which kind of job it was.
  assert.deepEqual(lines, [
    "[providers] writing provider=claude model=opus",
    "[providers] judgment provider=ollama model=ollama:qwen3.5:9b",
  ]);
  // A writing choice with a model, or another provider, is used as it is.
  choices.writing = { id: "claude", model: "sonnet" };
  assert.equal((await router.pin("k", { role: "writing" })).model, "sonnet");
  choices.writing = { id: "codex", model: "gpt-6.1-sol" };
  assert.equal((await router.pin("k", { role: "writing" })).model, "codex:gpt-6.1-sol");
  // Without claudeDefault a judgment is haiku, as before.
  const plain = createOneshotRouter({ claude: async () => ({ ok: true }), getChoice: () => null });
  assert.equal((await plain.pin("k")).model, "haiku");
});

// A job may name the Claude model used when the pane's choice names none
// (CC_OFFICE_RESPONSE_MODEL for writing). That is a default only.
test("a job may name the Claude model used when the pane names none; a choice is never replaced", async () => {
  const { router, calls } = make({
    c: { id: "claude", model: "opus" },
    n: { id: "claude", model: null },
    o: { id: "ollama", model: "qwen3.5:9b" },
  });
  const req = { system: "S", prompt: "P", model: "from-the-request" };
  const labels = [];
  for (const key of ["z", "n", "c", "o"]) {
    const job = await router.pin(key, { claudeModel: "sonnet" });
    labels.push(job.model);
    await job.run(req);
  }
  assert.deepEqual(labels, ["sonnet", "sonnet", "opus", "ollama:qwen3.5:9b"]);
  assert.deepEqual(
    calls.map(([id, o]) => [id, o.model]),
    [
      ["claude", "sonnet"],
      ["claude", "sonnet"],
      ["claude", "opus"],
      ["http", "qwen3.5:9b"],
    ],
  );
  // Without the option nothing changes, and a model in the request is not used.
  const plain = await router.pin("z");
  await plain.run(req);
  assert.deepEqual([plain.model, calls.at(-1)[1].model], ["haiku", "haiku"]);
});

// What a request really ran on goes to the pane: onUsed gets one notice per
// request that was sent, with what was asked for and what the provider said.
test("every request that is sent is reported: role, provider, the model asked for and the one that ran", async () => {
  const used = [];
  const choices = { writing: null, oneshot: { id: "ollama", model: "qwen3.5:9b" } };
  const router = createOneshotRouter({
    claude: async (req) => ({ ok: true, text: "c", model_ran: `claude-${req.model}-5-5` }),
    codex: async () => ({ ok: true, text: "x" }),
    antigravity: async () => ({ ok: true, text: "g" }),
    http: async () => ({ ok: false, error: "HTTP 500" }),
    keys: { get: async () => ({ key: null }) },
    getChoice: (key, role) => {
      if (key === "none") throw new Error("no confirmed choice");
      return choices[role];
    },
    claudeDefault: (role) => (role === "writing" ? "opus" : "haiku"),
    onUsed: (key, u) => used.push([key, u]),
  });
  const w = await router.pin("k", { role: "writing" });
  assert.equal(w.used(), null, "nothing ran yet");
  await w.run({ system: "s", prompt: "p" });
  assert.deepEqual(w.used(), {
    role: "writing",
    provider: "claude",
    asked: "opus",
    ran: "claude-opus-5-5",
    ok: true,
  });
  const j = await router.pin("k");
  await j.run({ system: "s", prompt: "p" });
  // A provider that does not say what ran: ran is null. A failed request is reported too.
  assert.deepEqual(j.used(), {
    role: "oneshot",
    provider: "ollama",
    asked: "qwen3.5:9b",
    ran: null,
    ok: false,
  });
  choices.writing = { id: "codex", model: null };
  const c = await router.pin("k", { role: "writing" });
  await c.run({ system: "s", prompt: "p" });
  assert.deepEqual(used, [
    ["k", w.used()],
    ["k", j.used()],
    ["k", { role: "writing", provider: "codex", asked: null, ran: null, ok: true }],
  ]);
  // A job that was refused sends nothing and reports nothing.
  const refused = await router.pin("none");
  await refused.run({ system: "s", prompt: "p" });
  assert.equal(used.length, 3);
});
