import { test } from "node:test";
import assert from "node:assert/strict";
import {
  modelMenuOptions,
  isStoredModelValid,
  effectiveModelNote,
  modelDisplayName,
} from "../taskpane/shared/model-menu.js";

const M = [
  { value: "sonnet", displayName: "Sonnet", resolvedModel: "claude-sonnet-5-5", description: "b" },
  { value: "fable", displayName: "Fable", resolvedModel: "claude-fable-5", description: "n" },
];

test("labels show the alias and the real id", () => {
  const o = modelMenuOptions(M, "fable");
  assert.deepEqual(
    o.map((x) => x.value),
    ["sonnet", "fable"],
  );
  assert.equal(o[1].label, "Fable · claude-fable-5");
});

test("a stored choice missing from the list is kept and marked", () => {
  const o = modelMenuOptions(M, "claude-opus-5");
  assert.equal(o.at(-1).value, "claude-opus-5");
  assert.match(o.at(-1).label, /not in list/);
});

test("stored value check replaces the old three-alias migration", () => {
  assert.equal(isStoredModelValid("fable"), true);
  assert.equal(isStoredModelValid("claude-sonnet-5-5[1m]"), true);
  assert.equal(isStoredModelValid("default"), false);
  assert.equal(isStoredModelValid(""), false);
  assert.equal(isStoredModelValid("a b"), false);
});

// F2: the daemon answers set_model with the model it will use.
test("set_model answer: no note when the choice is kept, one line when it was replaced", () => {
  assert.equal(effectiveModelNote("fable", "fable"), null);
  assert.equal(
    effectiveModelNote("claude-opus-9", "sonnet"),
    'Model "claude-opus-9" is not available for this account. Using sonnet.',
  );
  assert.equal(effectiveModelNote("", "sonnet"), null); // nothing was asked for
  assert.equal(effectiveModelNote("fable", undefined), null); // an answer without a model
});

// F3: the name next to the connection status. No family list in the code.
test("display name: the SDK name when listed, else the family in title case", () => {
  const L = [
    { value: "opus", displayName: "Opus 5.5", resolvedModel: "claude-opus-5-5" },
    { value: "claude-fable-5-1", displayName: "Fable 5.1", resolvedModel: "claude-fable-5-1" },
  ];
  assert.equal(modelDisplayName("opus", L), "Opus 5.5");
  assert.equal(modelDisplayName("claude-opus-5-5", L), "Opus 5.5"); // by resolved id
  assert.equal(modelDisplayName("claude-opus-5-5[1m]", L), "Opus 5.5"); // context suffix
  assert.equal(modelDisplayName("claude-fable-5-1", L), "Fable 5.1");
  assert.equal(modelDisplayName("claude-fable-5-1", []), "Fable");
  assert.equal(modelDisplayName("claude-haiku-4-5-20251001", L), "Haiku");
  assert.equal(modelDisplayName("sonnet", null), "Sonnet");
  assert.equal(modelDisplayName("claude-zephyr-7", L), "Zephyr"); // a family not known today
  assert.equal(modelDisplayName("", L), "");
});
