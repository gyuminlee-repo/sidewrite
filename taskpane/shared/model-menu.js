// Composer model dropdown rows from the daemon's models_list (P13).

// models: [{ value, displayName, resolvedModel, description }]. current: the
// stored choice; kept as a marked last row when the list does not hold it.
export function modelMenuOptions(models, current) {
  const rows = (models ?? []).map((m) => ({
    value: m.value,
    label:
      m.resolvedModel && m.resolvedModel !== m.value
        ? `${m.displayName} · ${m.resolvedModel}`
        : m.displayName,
    title: m.description ?? "",
  }));
  if (current && !rows.some((r) => r.value === current)) {
    rows.push({ value: current, label: `${current} (not in list)`, title: "" });
  }
  return rows;
}

// The rows of the model menus before the account's list is known.
export const ALIAS_MODELS = [
  { value: "haiku", displayName: "Haiku", description: "cheapest", resolvedModel: null },
  { value: "sonnet", displayName: "Sonnet", description: "balanced", resolvedModel: null },
  { value: "opus", displayName: "Opus", description: "most capable", resolvedModel: null },
];

// The model of edits (review turns and the chat), one value for the chat box
// and for Setup: the stored choice; else the default the daemon answered
// set_model with (the best model the account lists); else, before it
// answered, the alias it takes for that.
export function editsModelValue(stored, daemonDefault) {
  return stored || daemonDefault || "opus";
}

// A stored model choice the pane may send to the daemon, which checks it
// against the account's list. "default" was retired: a model is always named.
export function isStoredModelValid(v) {
  return typeof v === "string" && v !== "default" && /^[A-Za-z0-9][\w.:\-[\]]{0,80}$/.test(v);
}

// The model of edits a pane stored, or null when it stored no choice. Until
// the default became the best model, every saved setting carried model:
// "sonnet" (the default of the time), chosen or not. So a stored "sonnet"
// counts as a choice only with the mark the menus set when the user picks a
// model (modelChosen); without it, it is that old default. Any other stored
// model can only have come from the menu and is kept.
const OLD_DEFAULT = "sonnet";
export function storedEditsModel(settings) {
  const model = settings?.model;
  if (!isStoredModelValid(model)) return null;
  return settings.modelChosen === true || model !== OLD_DEFAULT ? model : null;
}

// The daemon answers set_model with the model it will use (it checks the
// choice against the account's list). A line for the chat when that is not
// the model the pane asked for, else null.
export function effectiveModelNote(requested, effective) {
  if (!requested || !effective || requested === effective) return null;
  return `Model "${requested}" is not available for this account. Using ${effective}.`;
}

// The short name shown next to the connection status. models: the daemon's
// models_list, or null before it answered. The SDK's display name when the
// alias or id is listed (a "[1m]" context suffix is ignored); else the
// family of a claude-<family>-<version> id, or the alias, in title case.
// No list of families is kept here: a new family needs no change.
export function modelDisplayName(idOrAlias, models) {
  const s = String(idOrAlias || "").trim();
  if (!s) return "";
  const id = s.replace(/\[[^\]]*\]$/, "");
  const hit = (models ?? []).find((m) => m.value === id || m.resolvedModel === id);
  if (hit?.displayName) return hit.displayName;
  const family = id.replace(/^claude-/i, "").split("-")[0] || id;
  return family.charAt(0).toUpperCase() + family.slice(1).toLowerCase();
}
