import { test } from "node:test";
import assert from "node:assert/strict";
import {
  modelMenuOptions,
  isStoredModelValid,
  effectiveModelNote,
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
