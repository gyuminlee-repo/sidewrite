import { test } from "node:test";
import assert from "node:assert/strict";
import { createProviderSelection, providersListReply } from "../daemon/providers/selection.mjs";

const ROWS = [
  {
    id: "claude",
    tier: "subscription",
    label: "Claude (Agent SDK)",
    available: true,
    turn_ready: true,
    turn_note: null,
    reason: null,
    models: [{ value: "sonnet" }],
  },
  {
    id: "codex",
    tier: "subscription",
    label: "Codex CLI",
    available: true,
    turn_ready: false,
    turn_note: "Judgments only: review turns are not offered for this provider.",
    reason: null,
    models: [{ value: "gpt-6.1-sol" }],
  },
  {
    id: "ollama",
    tier: "local",
    label: "Ollama (local)",
    available: true,
    turn_ready: false,
    turn_note: "Judgments only: review turns are not offered for this provider.",
    reason: null,
    models: [{ value: "qwen3.5:9b" }],
  },
  {
    id: "deepseek",
    tier: "key",
    label: "DeepSeek",
    available: false,
    turn_ready: false,
    turn_note: "Judgments only: review turns are not offered for this provider.",
    reason: "Connection failed (ECONNRESET)",
    models: [],
    has_key: true,
  },
];

test("a pane starts with no confirmed choice; a choice is per pane and per role", () => {
  const sel = createProviderSelection();
  // Nothing is assumed for a pane that has not said what it uses.
  assert.deepEqual(sel.get("a"), {
    turn: null,
    oneshot: null,
    writing: null,
    state: { turn: "restoring", writing: "restoring", oneshot: "restoring" },
    error: { turn: null, writing: null, oneshot: null },
  });
  sel.set("a", "oneshot", { id: "ollama", model: "qwen3.5:9b" }, ROWS);
  assert.deepEqual(sel.get("a").oneshot, { id: "ollama", model: "qwen3.5:9b" });
  assert.deepEqual(sel.get("a").state, {
    turn: "restoring",
    writing: "restoring",
    oneshot: "ready",
  });
  assert.equal(sel.get("b").state.oneshot, "restoring");
  sel.set("a", "turn", { id: "claude" }, ROWS);
  assert.deepEqual(sel.choiceFor("a", "turn"), { id: "claude", model: null });
  sel.set("a", "oneshot", { id: "claude" }, ROWS);
  assert.deepEqual(sel.choiceFor("a", "oneshot"), { id: "claude", model: null });
  sel.drop("a");
  assert.equal(sel.get("a").state.oneshot, "restoring");
});

test("only a provider that passed its check can be chosen, and the refusal says why", () => {
  const sel = createProviderSelection();
  assert.throws(
    () => sel.set("a", "turn", { id: "codex", model: "gpt-6.1-sol" }, ROWS),
    /Codex CLI cannot run review turns\. Judgments only: review turns are not offered/,
  );
  assert.throws(
    () => sel.set("a", "turn", { id: "ollama", model: "qwen3.5:9b" }, ROWS),
    /Ollama \(local\) cannot run review turns\. Judgments only: review turns are not offered/,
  );
  assert.throws(
    () => sel.set("a", "oneshot", { id: "deepseek", model: "x" }, ROWS),
    /DeepSeek is not available: Connection failed \(ECONNRESET\)/,
  );
  assert.throws(() => sel.set("a", "oneshot", { id: "nope" }, ROWS), /Unknown provider nope/);
  assert.throws(() => sel.set("a", "chat", { id: "claude" }, ROWS), /Unknown role chat/);
  // Nothing was stored by a refused choice.
  assert.deepEqual(
    [sel.get("a").turn, sel.get("a").oneshot, sel.get("a").state],
    [null, null, { turn: "restoring", writing: "restoring", oneshot: "restoring" }],
  );
});

test("the model must be one the provider lists; a provider without a list takes any name or none", () => {
  const sel = createProviderSelection();
  assert.throws(
    () => sel.set("a", "oneshot", { id: "ollama", model: "llama-not-here" }, ROWS),
    /Ollama \(local\) does not list the model llama-not-here/,
  );
  assert.throws(
    () => sel.set("a", "oneshot", { id: "ollama", model: "a b; rm" }, ROWS),
    /does not list the model/,
  );
  const noList = [{ ...ROWS[1], models: [] }];
  assert.deepEqual(sel.set("a", "oneshot", { id: "codex" }, noList).oneshot, {
    id: "codex",
    model: null,
  });
  assert.throws(
    () => sel.set("a", "oneshot", { id: "codex", model: "bad name!" }, noList),
    /not a model name/,
  );
  // A model is needed where the request cannot go without one.
  assert.throws(
    () => sel.set("a", "oneshot", { id: "ollama" }, ROWS),
    /Choose a model for Ollama \(local\)/,
  );
});

test("the reply for the pane: rows and the pane's selection, and no key in it", () => {
  const sel = createProviderSelection();
  sel.set("a", "oneshot", { id: "ollama", model: "qwen3.5:9b" }, ROWS);
  const r = providersListReply(ROWS, sel.get("a"));
  assert.equal(r.providers.length, ROWS.length);
  assert.deepEqual(r.selection.oneshot, { id: "ollama", model: "qwen3.5:9b" });
  const leaky = [{ ...ROWS[3], key: "sk-secret-1", token: "t", apiKey: "sk-secret-2" }];
  const out = JSON.stringify(providersListReply(leaky, sel.get("a")));
  assert.equal(out.includes("sk-secret"), false, "fields that could hold a key are not passed on");
  assert.equal(JSON.parse(out).providers[0].has_key, true);
});

// Writing (response drafts, translations, questions) is a role of its own:
// its provider is chosen apart from the one of the judgments.
test("writing is a role of its own: any available provider, a model where one is needed", async () => {
  const sel = createProviderSelection({ waitMs: 20 });
  sel.set("a", "oneshot", { id: "ollama", model: "qwen3.5:9b" }, ROWS);
  assert.equal(sel.get("a").writing, null);
  assert.equal(sel.get("a").state.writing, "restoring");
  sel.set("a", "writing", { id: "claude", model: "sonnet" }, ROWS);
  assert.deepEqual(await sel.ready("a", "writing"), { id: "claude", model: "sonnet" });
  assert.deepEqual(await sel.ready("a", "oneshot"), { id: "ollama", model: "qwen3.5:9b" });
  // Codex writes, though it runs no review turn.
  sel.set("a", "writing", { id: "codex", model: "gpt-6.1-sol" }, ROWS);
  assert.deepEqual(sel.choiceFor("a", "writing"), { id: "codex", model: "gpt-6.1-sol" });
  assert.throws(() => sel.set("a", "writing", { id: "ollama" }, ROWS), /Choose a model for Ollama/);
  assert.throws(
    () => sel.set("a", "writing", { id: "deepseek", model: "x" }, ROWS),
    /DeepSeek is not available/,
  );
  // Restored with the others; nothing stored is the default (null: Claude
  // with the best model, oneshot-router.mjs), and a pane of an older build
  // that sends no writing choice gets it too.
  const b = createProviderSelection({ waitMs: 20 });
  b.begin("k");
  const out = await b.restore(
    "k",
    { turn: null, writing: { id: "ollama", model: "qwen3.5:9b" }, oneshot: null },
    async () => ROWS,
  );
  assert.deepEqual(out.writing, { id: "ollama", model: "qwen3.5:9b" });
  b.begin("old");
  const old = await b.restore("old", { turn: null, oneshot: null }, async () => ROWS);
  assert.deepEqual([old.state.writing, old.writing], ["ready", null]);
  assert.equal(await b.ready("old", "writing"), null);
  // Refused as any stored choice: the request of that role sends nothing.
  b.begin("bad");
  await b.restore("bad", { writing: { id: "deepseek", model: "x" } }, async () => ROWS);
  await assert.rejects(
    b.ready("bad", "writing"),
    /The stored provider for writing cannot be used: DeepSeek is not available/,
  );
});
