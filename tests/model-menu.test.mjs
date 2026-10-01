import { test } from "node:test";
import assert from "node:assert/strict";
import { modelMenuOptions, isStoredModelValid } from "../taskpane/shared/model-menu.js";

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
