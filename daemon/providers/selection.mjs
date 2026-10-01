// Which provider a pane uses for each role (P14):
//   turn     review turns. Default Claude; its model is the composer's.
//   oneshot  judgments (split, requirements, ledger checks and suggestions).
//            Default null: Claude with its judgment model (oneshot-router.mjs).
// A choice is accepted only for a provider whose probe row passed
// (probe.mjs): available for a judgment, turn_ready for a review turn. A
// refusal carries the row's own reason, which the task pane shows.

const ROLES = ["turn", "oneshot"];
const MODEL_NAME = /^[\w.:/@[\]-]{1,120}$/;
// Providers whose request cannot be made without a model name.
const needsModel = (row) => row.tier === "local" || row.tier === "key";

// Row fields the task pane gets. A row never holds a key (probe.mjs); the
// list is closed all the same, so a field added later is not sent unseen.
const ROW_FIELDS = [
  "id",
  "tier",
  "label",
  "roles",
  "policy",
  "default_enabled",
  "available",
  "turn_ready",
  "turn_note",
  "reason",
  "models",
  "has_key",
  "key_source",
  "base_url",
];

export function providersListReply(rows, selection) {
  return {
    providers: (rows ?? []).map((r) =>
      Object.fromEntries(ROW_FIELDS.filter((k) => k in r).map((k) => [k, r[k]])),
    ),
    selection,
  };
}

export function createProviderSelection() {
  const byKey = new Map(); // pane key -> { turn?, oneshot? }
  const get = (key) => ({
    turn: byKey.get(key)?.turn ?? { id: "claude", model: null },
    oneshot: byKey.get(key)?.oneshot ?? null,
  });
  return {
    get,
    choiceFor: (key, role) => get(key)[role],
    drop: (key) => byKey.delete(key),
    // rows: the last probe (probe.mjs probeAll). Returns the pane's selection.
    set(key, role, { id, model = null } = {}, rows) {
      if (!ROLES.includes(role)) throw new Error(`Unknown role ${role}`);
      const row = (rows ?? []).find((r) => r.id === id);
      if (!row) throw new Error(`Unknown provider ${id}`);
      if (!row.available) {
        throw new Error(
          `${row.label} is not available: ${row.reason ?? "it did not pass its check"}`,
        );
      }
      if (role === "turn" && !row.turn_ready) {
        throw new Error(
          `${row.label} cannot run review turns. ${row.turn_note ?? row.reason ?? ""}`.trim(),
        );
      }
      const name = model == null || model === "" ? null : String(model);
      const listed = (row.models ?? []).map((m) => m.value);
      if (name && listed.length && !listed.includes(name)) {
        throw new Error(`${row.label} does not list the model ${name}`);
      }
      if (name && !MODEL_NAME.test(name)) throw new Error(`${name} is not a model name`);
      if (!name && role === "oneshot" && needsModel(row)) {
        throw new Error(`Choose a model for ${row.label}`);
      }
      byKey.set(key, { ...(byKey.get(key) ?? {}), [role]: { id, model: name } });
      return get(key);
    },
  };
}
