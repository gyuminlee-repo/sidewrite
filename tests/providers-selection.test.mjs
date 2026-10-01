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
    turn_note: "Judgments only. Codex CLI will not call the office tools without an approval.",
    reason: null,
    models: [{ value: "gpt-6.1-sol" }],
  },
  {
    id: "ollama",
    tier: "local",
    label: "Ollama (local)",
    available: true,
    turn_ready: false,
    turn_note: "Judgments only: no tool loop for this provider in this stage.",
    reason: null,
    models: [{ value: "qwen3.5:9b" }],
  },
  {
    id: "deepseek",
    tier: "key",
    label: "DeepSeek",
    available: false,
    turn_ready: false,
    turn_note: "Judgments only: no tool loop for this provider in this stage.",
    reason: "Connection failed (ECONNRESET)",
    models: [],
    has_key: true,
  },
];

test("defaults: Claude for both roles; a choice is per pane and per role", () => {
  const sel = createProviderSelection();
  assert.deepEqual(sel.get("a"), { turn: { id: "claude", model: null }, oneshot: null });
  assert.equal(sel.choiceFor("a", "oneshot"), null);
  sel.set("a", "oneshot", { id: "ollama", model: "qwen3.5:9b" }, ROWS);
  assert.deepEqual(sel.choiceFor("a", "oneshot"), { id: "ollama", model: "qwen3.5:9b" });
  assert.equal(sel.choiceFor("b", "oneshot"), null);
  assert.deepEqual(sel.choiceFor("a", "turn"), { id: "claude", model: null });
  sel.set("a", "oneshot", { id: "claude" }, ROWS);
  assert.deepEqual(sel.choiceFor("a", "oneshot"), { id: "claude", model: null });
  sel.drop("a");
  assert.equal(sel.choiceFor("a", "oneshot"), null);
});

test("only a provider that passed its check can be chosen, and the refusal says why", () => {
  const sel = createProviderSelection();
  assert.throws(
    () => sel.set("a", "turn", { id: "codex", model: "gpt-6.1-sol" }, ROWS),
    /Codex CLI cannot run review turns\. Judgments only\. Codex CLI will not call/,
  );
  assert.throws(
    () => sel.set("a", "turn", { id: "ollama", model: "qwen3.5:9b" }, ROWS),
    /Ollama \(local\) cannot run review turns\. Judgments only: no tool loop/,
  );
  assert.throws(
    () => sel.set("a", "oneshot", { id: "deepseek", model: "x" }, ROWS),
    /DeepSeek is not available: Connection failed \(ECONNRESET\)/,
  );
  assert.throws(() => sel.set("a", "oneshot", { id: "nope" }, ROWS), /Unknown provider nope/);
  assert.throws(() => sel.set("a", "chat", { id: "claude" }, ROWS), /Unknown role chat/);
  // Nothing was stored by a refused choice.
  assert.deepEqual(sel.get("a"), { turn: { id: "claude", model: null }, oneshot: null });
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
