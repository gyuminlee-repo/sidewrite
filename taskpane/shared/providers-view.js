// Providers section of the Setup tab (P14): pure helpers over the daemon's
// providers_list rows ({ id, tier, label, roles, available, turn_ready,
// turn_note, reason, policy, models, has_key, key_source, default_enabled }).
// DOM glue is in providers-pane.js.

// A local or key provider is called over HTTP and needs a model name; a CLI
// or Claude may run with its own default.
const needsModel = (row) => row?.tier === "local" || row?.tier === "key";

// The rows of one menu: every provider. role "turn" (edits of the
// manuscript): enabled when turn_ready. role "writing" and "oneshot"
// (judgments): enabled when available. A row that cannot be chosen says why:
// shortly in its text, in full in its note.
export function providerRows(list, role) {
  return (list ?? []).map((r) => {
    const enabled = role === "turn" ? !!r.turn_ready : !!r.available;
    const note = enabled
      ? ""
      : !r.available
        ? (r.reason ?? "not available")
        : (r.turn_note ?? r.reason ?? "not ready for review turns");
    const short = enabled ? "" : !r.available ? note : "not for edits";
    return {
      id: r.id,
      label: r.label,
      text: short ? `${r.label} (${short})` : r.label,
      enabled,
      note,
      policy: r.policy ?? "",
    };
  });
}

// A stored choice ({ id, model } or null) against a fresh list: kept when the
// provider can still be chosen for the role, else Claude with `dropped` set
// to the id that was given up. A model the provider no longer lists is
// replaced by its first one (or by none where a default exists).
export function choiceAfterProbe(stored, list, role) {
  if (!stored?.id) return { id: "claude", model: null };
  const row = providerRows(list, role).find((r) => r.id === stored.id);
  if (!row?.enabled) return { id: "claude", model: null, dropped: stored.id };
  const full = (list ?? []).find((r) => r.id === stored.id);
  const listed = (full?.models ?? []).map((m) => m.value);
  let model = stored.model ?? null;
  if (listed.length && !listed.includes(model)) model = needsModel(full) ? listed[0] : null;
  return { id: stored.id, model };
}

// The status of one provider in the list under the menus.
export function providerStatus(r) {
  const keyNote =
    r.tier !== "key"
      ? null
      : r.has_key
        ? r.key_source === "env"
          ? "A key comes from the environment."
          : "A key is stored in the key file."
        : "Off until you enter a key.";
  if (!r.available) {
    return {
      state: "off",
      text: `Not available: ${r.reason ?? "it did not pass its check"}`,
      keyNote,
    };
  }
  const text = r.turn_ready
    ? "Available for review turns and judgments."
    : (r.roles ?? []).includes("turn") && r.turn_note
      ? `Available for judgments. ${r.turn_note}`
      : "Available for judgments.";
  return { state: "ok", text, keyNote };
}

// The rows of a model menu for the chosen provider. defaultLabel: the text
// of the first row of a provider that runs without a model being named.
export function modelRows(row, current, defaultLabel = "Default model") {
  if (!row) return [];
  const models = (row.models ?? []).map((m) => ({
    value: m.value,
    label: m.displayName ?? m.value,
    selected: m.value === current,
  }));
  if (needsModel(row)) {
    return models.length ? models : [{ value: "", label: "No model listed", selected: true }];
  }
  return [{ value: "", label: defaultLabel, selected: !current }, ...models];
}
