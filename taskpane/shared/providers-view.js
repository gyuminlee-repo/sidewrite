// Providers section of the Setup tab (P14): pure helpers over the daemon's
// providers_list rows ({ id, tier, label, roles, available, turn_ready,
// turn_mode, turn_note, reason, policy, models, has_key, key_source,
// default_enabled }). turn_mode "proposal": the provider is given no tools
// in a review turn; it returns proposed changes and the daemon applies them.
// DOM glue is in providers-pane.js.

// A local or key provider is called over HTTP and needs a model name; a CLI
// or Claude may run with its own default.
const needsModel = (row) => row?.tier === "local" || row?.tier === "key";

// The tooltip of a proposing provider in the review turn menu.
export const PROPOSAL_TITLE =
  "Returns a list of proposed changes. This add-in checks each one and applies it as a tracked change with a rationale comment. The CLI does not touch Word.";

// The rows of one menu. role "turn": the providers that can run review turns
// at all, enabled when turn_ready. role "oneshot": every provider, enabled
// when available. A row that cannot be chosen says why in its text; a
// proposing provider says "proposals only" (proposal: true).
export function providerRows(list, role) {
  return (list ?? [])
    .filter((r) => role !== "turn" || (r.roles ?? []).includes("turn"))
    .map((r) => {
      const enabled = role === "turn" ? !!r.turn_ready : !!r.available;
      const proposal = role === "turn" && enabled && r.turn_mode === "proposal";
      const note = enabled
        ? proposal
          ? PROPOSAL_TITLE
          : ""
        : !r.available
          ? (r.reason ?? "not available")
          : (r.turn_note ?? r.reason ?? "not ready for review turns");
      const short = enabled ? (proposal ? "proposals only" : "") : note;
      return {
        id: r.id,
        label: r.label,
        text: short ? `${r.label} (${short})` : r.label,
        enabled,
        note,
        policy: r.policy ?? "",
        proposal,
      };
    });
}

// What the author is told before the first review turn of a proposing
// provider: what leaves the computer, what the CLI adds, and who edits.
// { text, use: the label of the button that agrees }.
export function proposalNotice(row) {
  return {
    use: `Use ${row.label}`,
    text:
      `Review turns with ${row.label} send text off this computer, to the CLI and the service behind it: ` +
      "the reviewer comment as written, the manuscript paragraphs shown for it (your selection, or the whole manuscript), " +
      "your author instructions and the forbidden list. The CLI adds its own global instructions (for Codex, your AGENTS.md) to every request. " +
      (row.policy ? `${row.policy} ` : "") +
      "It does not edit the document: it returns proposed changes, and this add-in checks each one and applies it as a tracked change with a rationale comment. " +
      "Preview prompt in the Review tab shows what is sent.",
  };
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
    ? r.turn_mode === "proposal"
      ? "Available for review turns (proposals only) and judgments."
      : "Available for review turns and judgments."
    : (r.roles ?? []).includes("turn") && r.turn_note
      ? `Available for judgments. ${r.turn_note}`
      : "Available for judgments.";
  return { state: "ok", text, keyNote };
}

// The rows of a model menu for the chosen provider.
export function modelRows(row, current) {
  if (!row) return [];
  const models = (row.models ?? []).map((m) => ({
    value: m.value,
    label: m.displayName ?? m.value,
    selected: m.value === current,
  }));
  if (needsModel(row)) {
    return models.length ? models : [{ value: "", label: "No model listed", selected: true }];
  }
  return [{ value: "", label: "Default model", selected: !current }, ...models];
}
