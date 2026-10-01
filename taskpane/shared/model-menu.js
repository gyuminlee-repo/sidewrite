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

// A stored model choice the pane may send to the daemon, which checks it
// against the account's list. "default" was retired: a model is always named.
export function isStoredModelValid(v) {
  return typeof v === "string" && v !== "default" && /^[A-Za-z0-9][\w.:\-[\]]{0,80}$/.test(v);
}

// The daemon answers set_model with the model it will use (it checks the
// choice against the account's list). A line for the chat when that is not
// the model the pane asked for, else null.
export function effectiveModelNote(requested, effective) {
  if (!requested || !effective || requested === effective) return null;
  return `Model "${requested}" is not available for this account. Using ${effective}.`;
}
