// The provider a pane stored must be the one the daemon uses after a
// restart, from the first request on (Codex review of stage 2.5: a Split
// clicked right after a daemon restart went to the default Claude while the
// stored Ollama choice was still on its way). The real selection and router,
// with fake providers that record every call.
import { test } from "node:test";
import assert from "node:assert/strict";
import { createProviderSelection } from "../daemon/providers/selection.mjs";
import { createOneshotRouter } from "../daemon/providers/oneshot-router.mjs";

const ROWS = [
  {
    id: "claude",
    tier: "subscription",
    label: "Claude (Agent SDK)",
    available: true,
    turn_ready: true,
    models: [{ value: "sonnet" }],
  },
  {
    id: "ollama",
    tier: "local",
    label: "Ollama (local)",
    available: true,
    turn_ready: false,
    turn_note: "Judgments only: review turns are not offered for this provider.",
    models: [{ value: "qwen3.5:9b" }],
  },
  {
    id: "deepseek",
    tier: "key",
    label: "DeepSeek",
    available: false,
    turn_ready: false,
    reason: "Connection failed (ECONNRESET)",
    models: [],
  },
];
const OLLAMA = { id: "ollama", model: "qwen3.5:9b" };
const REQ = { system: "s", prompt: "the text of a manuscript" };
const tick = (ms = 5) => new Promise((r) => setTimeout(r, ms));

// A daemon: its selection and its router. calls: every provider call made.
function daemon({ waitMs = 60 } = {}) {
  const calls = [];
  const sel = createProviderSelection({ waitMs });
  const router = createOneshotRouter({
    claude: async (req) => (calls.push(["claude", req.model]), { ok: true, text: "c" }),
    codex: async (req) => (calls.push(["codex", req.model]), { ok: true, text: "x" }),
    antigravity: async (req) => (calls.push(["agy", req.model]), { ok: true, text: "g" }),
    http: async (o) => (calls.push(["http", o.baseUrl, o.model]), { ok: true, text: "h" }),
    keys: { get: async () => ({ key: "k", base_url: null }) },
    getChoice: (key) => sel.ready(key, "oneshot"),
  });
  return { sel, router, calls };
}

test("a judgment asked before the pane's stored choice arrived waits for it and uses it", async () => {
  const d = daemon({ waitMs: 2_000 });
  d.sel.begin("k"); // the pane said hello
  const early = d.router.forKey("k")(REQ); // Split, clicked at once
  await tick();
  assert.deepEqual(d.calls, [], "nothing goes anywhere while the choice is not confirmed");
  assert.deepEqual(d.sel.get("k").state, {
    turn: "restoring",
    writing: "restoring",
    oneshot: "restoring",
  });
  // The pane's stored choice arrives; the probe it is checked against is slow.
  let probed;
  const rows = new Promise((r) => (probed = r));
  const restored = d.sel.restore("k", { turn: null, oneshot: OLLAMA }, () => rows);
  await tick();
  assert.deepEqual(d.calls, [], "still waiting: the stored provider is not checked yet");
  // No stored choice for review turns: the default applies at once.
  assert.equal(d.sel.get("k").state.turn, "ready");
  assert.deepEqual(await d.sel.ready("k", "turn"), { id: "claude", model: null });
  probed(ROWS);
  assert.deepEqual((await restored).state, { turn: "ready", writing: "ready", oneshot: "ready" });
  assert.deepEqual(await early, { ok: true, text: "h" });
  assert.deepEqual(d.calls, [["http", "http://localhost:11434/v1", "qwen3.5:9b"]]);
});

test("a stored provider that did not pass: the request fails with the reason and nothing is sent anywhere", async () => {
  const d = daemon();
  d.sel.begin("k");
  const early = d.router.forKey("k")(REQ);
  const out = await d.sel.restore(
    "k",
    { turn: null, oneshot: { id: "deepseek", model: "x" } },
    async () => ROWS,
  );
  assert.equal(out.state.oneshot, "refused");
  assert.equal(out.error.oneshot, "DeepSeek is not available: Connection failed (ECONNRESET)");
  const r = await early;
  assert.deepEqual(r, {
    ok: false,
    error:
      "The stored provider for judgments cannot be used: DeepSeek is not available: Connection failed (ECONNRESET). Nothing was sent to any provider. Choose one in Setup, Providers.",
  });
  // Later requests too: no fallback to Claude until the user picks.
  assert.equal((await d.router.forKey("k")(REQ)).ok, false);
  assert.equal((await d.router.pin("k")).provider, null);
  assert.deepEqual(d.calls, []);
  // The user picks one: judgments run again, on that provider.
  d.sel.set("k", "oneshot", OLLAMA, ROWS);
  assert.equal((await d.router.forKey("k")(REQ)).ok, true);
  assert.deepEqual(
    d.calls.map((c) => c[0]),
    ["http"],
  );
});

test("a pane that never sends its choice: the request waits its limit, then fails; the daemon does not guess", async () => {
  const d = daemon({ waitMs: 40 });
  // No hello bookkeeping either (an old pane build, a second pane): the same.
  const t0 = Date.now();
  const r = await d.router.forKey("other")(REQ);
  assert.ok(Date.now() - t0 >= 35, "it waited");
  assert.equal(r.ok, false);
  assert.equal(
    r.error,
    "The task pane has not confirmed its provider for judgments (waited 0.04 s). Nothing was sent to any provider. Choose one in Setup, Providers.",
  );
  await assert.rejects(
    d.sel.ready("other", "turn"),
    /has not confirmed its provider for review turns/,
  );
  assert.deepEqual(d.calls, []);
});

test("no stored choice at all: the documented default applies, Claude with its judgment model", async () => {
  const d = daemon();
  d.sel.begin("k");
  const out = await d.sel.restore("k", { turn: null, oneshot: null }, () => {
    throw new Error("no probe is needed for the defaults");
  });
  assert.deepEqual(out.state, { turn: "ready", writing: "ready", oneshot: "ready" });
  assert.deepEqual(await d.router.forKey("k")(REQ), { ok: true, text: "c" });
  assert.deepEqual(d.calls, [["claude", "haiku"]]);
});

test("a new connection starts over: the choice of the last one is not used until it is confirmed again", async () => {
  const d = daemon({ waitMs: 40 });
  d.sel.begin("k");
  await d.sel.restore("k", { turn: null, oneshot: OLLAMA }, async () => ROWS);
  await d.router.forKey("k")(REQ);
  d.sel.begin("k"); // the socket dropped and the pane said hello again
  assert.equal(d.sel.get("k").state.oneshot, "restoring");
  assert.equal((await d.router.forKey("k")(REQ)).ok, false);
  assert.equal(d.calls.length, 1, "only the request of the first connection was sent");
});

test("the choice survives a daemon restart through the pane: the new daemon uses what the pane stored", async () => {
  const pane = { turnProvider: null, oneshotProvider: OLLAMA }; // the pane's settings
  for (const run of ["first daemon", "daemon after the restart"]) {
    const d = daemon({ waitMs: 2_000 });
    d.sel.begin("k");
    const early = d.router.forKey("k")(REQ);
    await d.sel.restore(
      "k",
      { turn: pane.turnProvider, oneshot: pane.oneshotProvider },
      async () => ROWS,
    );
    assert.equal((await early).ok, true, run);
    assert.deepEqual(d.calls, [["http", "http://localhost:11434/v1", "qwen3.5:9b"]], run);
  }
});

test("review turns: a stored provider that cannot run them is refused; the probe failing is a refusal too", async () => {
  const d = daemon();
  d.sel.begin("k");
  await d.sel.restore("k", { turn: OLLAMA, oneshot: null }, async () => ROWS);
  await assert.rejects(
    d.sel.ready("k", "turn"),
    /^Error: The stored provider for review turns cannot be used: Ollama \(local\) cannot run review turns\. Judgments only: review turns are not offered.*Nothing was sent to any provider\. Choose one in Setup, Providers\.$/,
  );
  assert.deepEqual(await d.sel.ready("k", "oneshot"), null);

  const e = daemon();
  e.sel.begin("k");
  const out = await e.sel.restore("k", { turn: null, oneshot: OLLAMA }, async () => {
    throw new Error("probe crashed");
  });
  assert.deepEqual(
    [out.state.oneshot, out.error.oneshot],
    ["refused", "The providers could not be checked: probe crashed"],
  );
  assert.equal((await e.router.forKey("k")(REQ)).ok, false);
  assert.deepEqual(e.calls, []);
});

test("a choice made by hand while the stored one is still being checked wins", async () => {
  const d = daemon({ waitMs: 2_000 });
  d.sel.begin("k");
  let probed;
  const restored = d.sel.restore(
    "k",
    { turn: null, oneshot: { id: "deepseek", model: "x" } },
    () => new Promise((r) => (probed = r)),
  );
  d.sel.set("k", "oneshot", OLLAMA, ROWS);
  probed(ROWS);
  const out = await restored;
  assert.deepEqual([out.state.oneshot, out.oneshot], ["ready", OLLAMA]);
  await d.router.forKey("k")(REQ);
  assert.deepEqual(
    d.calls.map((c) => c[0]),
    ["http"],
  );
});
