import { test } from "node:test";
import assert from "node:assert/strict";
import { modelLabel, ranOther, roleText, modelLineText } from "../taskpane/shared/model-line.js";

const MODELS = [
  { value: "opus", displayName: "Opus 5.5", resolvedModel: "claude-opus-5-5" },
  { value: "sonnet", displayName: "Sonnet 5.5", resolvedModel: "claude-sonnet-5-5" },
  { value: "haiku", displayName: "Haiku 4.5", resolvedModel: "claude-haiku-4-5" },
];

test("modelLabel: provider and model as a person reads them", () => {
  assert.equal(modelLabel({ provider: "claude", model: "opus" }, MODELS), "Claude · Opus 5.5");
  assert.equal(
    modelLabel({ provider: "claude", model: "claude-sonnet-5-5" }, MODELS),
    "Claude · Sonnet 5.5",
  );
  assert.equal(
    modelLabel({ provider: "ollama", model: "qwen3.5:9b" }, MODELS),
    "Ollama · qwen3.5:9b",
  );
  assert.equal(modelLabel({ provider: "codex", model: null }, MODELS), "Codex · its default model");
  assert.equal(modelLabel({ provider: "newone", model: "m" }, MODELS), "newone · m");
});

test("ranOther: only a difference that can be told is one", () => {
  assert.equal(ranOther("opus", "claude-opus-5-5", MODELS), false);
  assert.equal(ranOther("opus", "claude-opus-5-5[1m]", MODELS), false);
  assert.equal(ranOther("opus", "claude-sonnet-5-5", MODELS), true);
  assert.equal(ranOther("claude-opus-5-5", "claude-opus-5-5", MODELS), false);
  // No list: an alias is in the id of its family, a longer id holds the shorter.
  assert.equal(ranOther("opus", "claude-opus-5-5", null), false);
  assert.equal(ranOther("opus", "claude-sonnet-5-5", null), true);
  assert.equal(ranOther("qwen3.5:9b", "qwen3.5:9b-q4", null), false);
  // A date at the end of an id is the same model (the list may name it with
  // or without); a longer version is another one.
  assert.equal(ranOther("haiku", "claude-haiku-4-5-20251001", MODELS), false);
  assert.equal(ranOther("claude-haiku-4-5-20251001", "claude-haiku-4-5", null), false);
  assert.equal(ranOther("claude-opus-5", "claude-opus-5-5", null), true);
  assert.equal(
    ranOther("claude-opus-5", "claude-opus-5-5", [
      { value: "claude-opus-5", displayName: "Opus 5", resolvedModel: "claude-opus-5" },
    ]),
    true,
  );
  // The provider did not say, or nothing was asked for by name.
  assert.equal(ranOther("opus", null, MODELS), false);
  assert.equal(ranOther(null, "gpt-6", MODELS), false);
});

test("roleText: what last ran; the choice while nothing ran; a difference is said", () => {
  const choice = { provider: "claude", model: "opus" };
  assert.equal(
    roleText("turn", { used: null, choice }, MODELS),
    "Edits: Claude · Opus 5.5 (not used yet)",
  );
  assert.equal(
    roleText(
      "turn",
      { used: { provider: "claude", asked: "opus", ran: "claude-opus-5-5" }, choice },
      MODELS,
    ),
    "Edits: Claude · Opus 5.5",
  );
  assert.equal(
    roleText(
      "turn",
      { used: { provider: "claude", asked: "opus", ran: "claude-sonnet-5-5" }, choice },
      MODELS,
    ),
    "Edits: asked for Claude · Opus 5.5, ran on Claude · Sonnet 5.5",
  );
  // The provider does not say what ran: what was asked for, said as that.
  assert.equal(
    roleText(
      "oneshot",
      {
        used: { provider: "ollama", asked: "qwen3.5:9b", ran: null },
        choice: { provider: "ollama", model: "qwen3.5:9b" },
      },
      MODELS,
    ),
    "Judgments: Ollama · qwen3.5:9b",
  );
  // The choice was changed since: the next request will use another one.
  assert.equal(
    roleText(
      "writing",
      {
        used: { provider: "claude", asked: "sonnet", ran: "claude-sonnet-5-5" },
        choice: { provider: "claude", model: "opus" },
      },
      MODELS,
    ),
    "Writing: Claude · Sonnet 5.5 (next: Claude · Opus 5.5)",
  );
  // A request that failed is said as that.
  assert.equal(
    roleText(
      "writing",
      { used: { provider: "claude", asked: "opus", ran: null, ok: false }, choice },
      MODELS,
    ),
    "Writing: Claude · Opus 5.5 (last request failed)",
  );
});

test("modelLineText: the three kinds of work in one line", () => {
  const state = {
    turn: { used: null, choice: { provider: "claude", model: "opus" } },
    writing: { used: null, choice: { provider: "claude", model: "opus" } },
    oneshot: { used: null, choice: { provider: "claude", model: "haiku" } },
  };
  assert.equal(
    modelLineText(state, MODELS),
    "Edits: Claude · Opus 5.5 (not used yet)   Writing: Claude · Opus 5.5 (not used yet)   Judgments: Claude · Haiku 4.5 (not used yet)",
  );
});

test("fallbackNotice: what the chat says when the CLI moved the session to another model", async () => {
  const { fallbackNotice } = await import("../taskpane/shared/model-line.js");
  const models = [
    ...MODELS,
    { value: "claude-opus-5", displayName: "Opus 5", resolvedModel: "claude-opus-5" },
  ];
  assert.equal(
    fallbackNotice(
      { original: "claude-opus-5-5", fallback: "claude-opus-5", scope: "session" },
      models,
    ),
    "Opus 5.5 declined a request of this turn, and Claude Code ran it again on Opus 5. The next turn asks Opus 5.5 once more; if it declines again, this document stays on Opus 5.",
  );
  assert.equal(
    fallbackNotice(
      { original: "claude-opus-5-5", fallback: "claude-opus-5", scope: "local" },
      models,
    ),
    "Opus 5.5 declined a request of this turn, and Claude Code ran it again on Opus 5. Only that answer came from Opus 5; the session keeps its model.",
  );
});

// The chosen model declined two turns in a row: the document stays on the
// other model, and that is a state, not a warning repeated every turn.
test("a model that declines the manuscript: the notice once, then a state in the line and a note for Setup", async () => {
  const { fallbackNotice, roleText, declinedNote } = await import("../taskpane/shared/model-line.js");
  const models = [
    ...MODELS,
    { value: "claude-opus-5", displayName: "Opus 5", resolvedModel: "claude-opus-5" },
  ];
  const f = { original: "claude-opus-5-5", fallback: "claude-opus-5", scope: "session", category: "bio" };
  assert.equal(
    fallbackNotice({ ...f, stay: true }, models),
    "Opus 5.5 declined this manuscript again (bio). Edits of this document now stay on Opus 5: asking Opus 5.5 before every turn would spend a refused request each time. To use another model, choose it in the chat box or in Setup.",
  );
  const choice = { provider: "claude", model: "opus" };
  const once = { provider: "claude", asked: "opus", ran: "claude-opus-5", declined: { ...f, stay: false } };
  assert.equal(roleText("turn", { used: once, choice }, models), "Edits: asked for Claude · Opus 5.5, ran on Claude · Opus 5");
  const stays = { ...once, declined: { ...f, stay: true } };
  assert.equal(
    roleText("turn", { used: stays, choice }, models),
    "Edits: Claude · Opus 5 (Opus 5.5 declines this manuscript: bio)",
  );
  assert.equal(
    roleText("turn", { used: { ...stays, declined: { ...f, category: null, stay: true } }, choice }, models),
    "Edits: Claude · Opus 5 (Opus 5.5 declines this manuscript)",
  );
  assert.equal(
    declinedNote(stays, models),
    "Opus 5.5 declines this manuscript (bio), so edits run on Opus 5. You can choose another model here.",
  );
  assert.equal(declinedNote(once, models), "");
  assert.equal(declinedNote(null, models), "");
});

