import { test } from "node:test";
import assert from "node:assert/strict";
import {
  providerRows,
  choiceAfterProbe,
  providerStatus,
  modelRows,
} from "../taskpane/shared/providers-view.js";

const LIST = [
  {
    id: "claude",
    tier: "subscription",
    label: "Claude",
    roles: ["turn", "oneshot"],
    available: true,
    turn_ready: true,
    turn_note: null,
    reason: null,
    policy: "p",
    models: [{ value: "sonnet", displayName: "Sonnet 5.5" }],
  },
  {
    id: "antigravity",
    tier: "subscription",
    label: "agy",
    roles: ["turn", "oneshot"],
    available: true,
    turn_ready: false,
    turn_note: "Judgments only. agy print mode denies the office tools.",
    reason: null,
    policy: "p",
    models: [{ value: "g", displayName: "G" }],
  },
  {
    id: "ollama",
    tier: "local",
    label: "Ollama (local)",
    roles: ["oneshot"],
    available: true,
    turn_ready: false,
    turn_note: "Judgments only: review turns are not offered for this provider.",
    reason: null,
    policy: "never leaves this machine",
    models: [{ value: "qwen3.5:9b", displayName: "qwen3.5:9b" }],
  },
  {
    id: "deepseek",
    tier: "key",
    label: "DeepSeek",
    roles: ["oneshot"],
    available: false,
    turn_ready: false,
    turn_note: "Judgments only: review turns are not offered for this provider.",
    reason: "Connection failed (ECONNRESET)",
    policy: "China",
    has_key: true,
    key_source: "file",
    default_enabled: false,
    models: [],
  },
  {
    id: "groq",
    tier: "key",
    label: "Groq",
    roles: ["oneshot"],
    available: false,
    turn_ready: false,
    turn_note: "Judgments only: review turns are not offered for this provider.",
    reason: "No API key",
    policy: "30 days",
    has_key: false,
    key_source: null,
    default_enabled: false,
    models: [],
  },
];

test("turn menu: the CLI and Claude providers, only a turn-ready one enabled; judgment menu: available only", () => {
  const t = providerRows(LIST, "turn");
  assert.deepEqual(
    t.map((r) => r.id),
    ["claude", "antigravity"],
  );
  assert.deepEqual(
    t.filter((r) => r.enabled).map((r) => r.id),
    ["claude"],
  );
  assert.match(t.find((r) => r.id === "antigravity").note, /denies the office tools/);
  assert.equal(
    t.some((r) => r.id === "deepseek"),
    false,
  );
  const o = providerRows(LIST, "oneshot");
  assert.deepEqual(
    o.filter((r) => r.enabled).map((r) => r.id),
    ["claude", "antigravity", "ollama"],
  );
  assert.match(o.find((r) => r.id === "deepseek").note, /Connection failed/);
  assert.equal(o.find((r) => r.id === "deepseek").policy, "China");
  assert.equal(o.find((r) => r.id === "antigravity").note, "", "fine for a judgment");
  // What the option shows.
  assert.equal(
    o.find((r) => r.id === "deepseek").text,
    "DeepSeek (Connection failed (ECONNRESET))",
  );
  assert.equal(o.find((r) => r.id === "ollama").text, "Ollama (local)");
  assert.equal(t.find((r) => r.id === "antigravity").text, "agy (judgments only)");
});

test("a stored choice that stopped passing falls back to Claude", () => {
  assert.deepEqual(choiceAfterProbe({ id: "deepseek", model: "x" }, LIST, "oneshot"), {
    id: "claude",
    model: null,
    dropped: "deepseek",
  });
  assert.deepEqual(choiceAfterProbe({ id: "antigravity", model: "g" }, LIST, "oneshot"), {
    id: "antigravity",
    model: "g",
  });
  assert.deepEqual(choiceAfterProbe({ id: "antigravity", model: "g" }, LIST, "turn"), {
    id: "claude",
    model: null,
    dropped: "antigravity",
  });
  assert.deepEqual(choiceAfterProbe(null, LIST, "oneshot"), { id: "claude", model: null });
  assert.deepEqual(choiceAfterProbe({ id: "gone" }, LIST, "oneshot"), {
    id: "claude",
    model: null,
    dropped: "gone",
  });
  // A stored model the provider no longer lists is dropped, the provider kept.
  assert.deepEqual(choiceAfterProbe({ id: "ollama", model: "old:1b" }, LIST, "oneshot"), {
    id: "ollama",
    model: "qwen3.5:9b",
  });
});

test("the status line of a provider row", () => {
  const by = Object.fromEntries(LIST.map((r) => [r.id, providerStatus(r)]));
  assert.deepEqual(by.claude, {
    state: "ok",
    text: "Available for review turns and judgments.",
    keyNote: null,
  });
  assert.equal(by.antigravity.state, "ok");
  assert.equal(
    by.antigravity.text,
    "Available for judgments. Judgments only. agy print mode denies the office tools.",
  );
  assert.equal(by.ollama.text, "Available for judgments.");
  assert.deepEqual(by.deepseek, {
    state: "off",
    text: "Not available: Connection failed (ECONNRESET)",
    keyNote: "A key is stored in the key file.",
  });
  assert.equal(by.groq.keyNote, "Off until you enter a key.");
  assert.equal(
    providerStatus({ ...LIST[4], has_key: true, key_source: "env" }).keyNote,
    "A key comes from the environment.",
  );
});

test("model rows: the provider's list, the current one selected, a first row for none", () => {
  assert.deepEqual(modelRows(LIST[2], "qwen3.5:9b"), [
    { value: "qwen3.5:9b", label: "qwen3.5:9b", selected: true },
  ]);
  assert.deepEqual(modelRows(LIST[0], null), [
    { value: "", label: "Default model", selected: true },
    { value: "sonnet", label: "Sonnet 5.5", selected: false },
  ]);
  assert.deepEqual(modelRows(LIST[3], null), [
    { value: "", label: "No model listed", selected: true },
  ]);
  assert.deepEqual(modelRows(undefined, null), []);
});
