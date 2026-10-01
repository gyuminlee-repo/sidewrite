import { test } from "node:test";
import assert from "node:assert/strict";
import { PRESETS, presetById, canRun } from "../daemon/providers/registry.mjs";

test("three tiers, ten presets, ids unique", () => {
  assert.equal(new Set(PRESETS.map((p) => p.id)).size, PRESETS.length);
  assert.equal(PRESETS.length, 10);
  assert.deepEqual(
    PRESETS.filter((p) => p.tier === "subscription").map((p) => p.id),
    ["claude", "codex", "antigravity"],
  );
  assert.deepEqual(
    PRESETS.filter((p) => p.tier === "local").map((p) => p.id),
    ["ollama"],
  );
  assert.deepEqual(
    PRESETS.filter((p) => p.tier === "key").map((p) => p.id),
    ["openrouter", "deepseek", "groq", "mistral", "gemini", "openai-compatible"],
  );
});

test("only subscription CLIs run review turns; every preset can judge", () => {
  for (const p of PRESETS) {
    assert.equal(canRun(p, "turn"), p.tier === "subscription", p.id);
    assert.equal(canRun(p, "oneshot"), true, p.id);
  }
  assert.equal(canRun(null, "oneshot"), false);
});

test("key presets are off by default and carry a data policy", () => {
  for (const p of PRESETS.filter((x) => x.tier === "key")) {
    assert.equal(p.defaultEnabled, false);
    assert.ok(p.policy.length > 10, p.id);
  }
  for (const p of PRESETS) assert.ok(p.policy.length > 10, p.id);
  assert.match(presetById("ollama").policy, /never leaves this machine/);
  assert.equal(presetById("nope"), null);
});

test("base URLs: https for hosted key presets, none for the custom one, loopback for Ollama", () => {
  for (const p of PRESETS.filter((x) => x.tier === "key" && x.id !== "openai-compatible")) {
    assert.match(p.baseUrl, /^https:\/\//, p.id);
  }
  assert.equal(presetById("openai-compatible").baseUrl, null);
  assert.match(presetById("ollama").baseUrl, /^http:\/\/localhost:11434\//);
  assert.equal(Object.isFrozen(PRESETS), true);
});

// Request fields of a preset (openai-compat.mjs body, jsonBody). Only Ollama
// has them: other servers refuse fields they do not know.
test("only the Ollama preset carries request fields of its own", () => {
  const ollama = presetById("ollama");
  assert.deepEqual(ollama.body, { reasoning_effort: "none" });
  assert.deepEqual(ollama.jsonBody, { response_format: { type: "json_object" } });
  for (const p of PRESETS) {
    if (p.id !== "ollama") assert.equal("body" in p || "jsonBody" in p, false, p.id);
  }
});
