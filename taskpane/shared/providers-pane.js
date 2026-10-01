// Providers section of the Setup tab (P14): which provider and model do
// each kind of work, out of the ones the daemon's probe passed. Edits (role
// "turn": review turns and the chat), Writing (response drafts,
// translations, questions) and Judgments (role "oneshot"). Pure helpers are
// in providers-view.js.
//
// The model of Edits is not stored here: it is the model of the chat box
// (taskpane.js, settings.model). editsModel = { get(), rows(), set(value) }
// reads and changes that one value, so the two menus cannot differ.
//
// Text goes in through textContent only. An API key is read from its box,
// sent once in provider_set_key and the box emptied before the answer comes;
// it is never put in the settings, a message or any other request.

import { providerRows, choiceAfterProbe, providerStatus, modelRows } from "./providers-view.js";

const $ = (id) => document.getElementById(id);
const ROLES = ["turn", "writing", "oneshot"];
const SETTING = { turn: "turnProvider", writing: "writingProvider", oneshot: "oneshotProvider" };
const WORDS = { turn: "review turns", writing: "writing", oneshot: "judgments" };
const STOPPED = { turn: "Review turns", writing: "Drafts and translations", oneshot: "Judgments" };
// The first row of a Claude model menu: what runs when no model is named.
const DEFAULT_MODEL_LABEL = { writing: "Best model (default)", oneshot: "Default model" };
// The probe starts CLIs and calls servers; a choice waits for it.
const TIMEOUT = { timeoutMs: 60_000 };
// The one key provider without a fixed server: it takes a URL with its key.
const CUSTOM_ID = "openai-compatible";

function el(tag, { text, cls, ...props } = {}, kids = []) {
  const e = document.createElement(tag);
  if (text != null) e.textContent = text;
  if (cls) e.className = cls;
  Object.assign(e, props);
  e.append(...kids.filter(Boolean)); // a part left out is null, not the text "null"
  return e;
}

export function initProvidersPane({
  sendRequest,
  settings,
  saveSettings,
  editsModel = null,
  // Called after every draw of the menus.
  onRender = () => {},
}) {
  let list = [];
  // What the menus show: the choice the daemon confirmed, or the stored one
  // it refused (the notes then say so).
  const sel = Object.fromEntries(ROLES.map((r) => [r, { id: "claude", model: null }]));
  // Why a stored choice is not in use, per role, and a daemon that gave no
  // answer. Shown until the user picks a provider that passes.
  const notes = { turn: null, writing: null, oneshot: null, daemon: null };
  let waiting = null; // the restore_providers request under way

  const ask = (type, payload) =>
    sendRequest(type, payload, TIMEOUT).then(
      (r) => r,
      (e) => ({ ok: false, error: e?.message ?? String(e) }),
    );
  const labelOf = (id) => list.find((r) => r.id === id)?.label ?? id;

  function message(text) {
    const m = $("providers-message");
    m.textContent = text ?? "";
    m.hidden = !text;
  }

  const showNotes = () =>
    message([notes.daemon, notes.turn, notes.writing, notes.oneshot].filter(Boolean).join(" "));

  function store(role, choice) {
    // Claude without a model is the daemon's default: nothing to remember.
    settings[SETTING[role]] = choice.id === "claude" && !choice.model ? null : choice;
    saveSettings(settings);
  }

  function renderMenus() {
    for (const role of ROLES) {
      const menu = $(`provider-${role}`);
      menu.replaceChildren(
        ...providerRows(list, role).map((r) =>
          el("option", { value: r.id, text: r.text, disabled: !r.enabled, title: r.note }),
        ),
      );
      menu.value = sel[role].id;
      const models = $(`provider-${role}-model`);
      // Claude's edits run on the model of the chat box: the same value,
      // shown here and changed through editsModel.
      const ofChat = role === "turn" && sel.turn.id === "claude";
      const chat = ofChat && editsModel ? editsModel.get() : null;
      const opts = chat
        ? editsModel
            .rows()
            .map((r) => ({ value: r.value, label: r.label, selected: r.value === chat }))
        : ofChat
          ? [{ value: "", label: "Model of the chat box", selected: true }]
          : modelRows(
              list.find((r) => r.id === sel[role].id),
              sel[role].model,
              DEFAULT_MODEL_LABEL[role],
            );
      models.replaceChildren(
        ...opts.map((o) => el("option", { value: o.value, text: o.label, selected: o.selected })),
      );
      models.value = chat ?? (ofChat ? "" : (sel[role].model ?? ""));
      models.disabled = chat ? false : ofChat || opts.length < 2;
    }
  }

  function keyControls(r) {
    const box = el("input", { type: "password", placeholder: "API key", autocomplete: "off" });
    const url =
      r.id === CUSTOM_ID
        ? el("input", { type: "text", placeholder: "https://host/v1", value: r.base_url ?? "" })
        : null;
    const save = el("button", { type: "button", text: "Save key", cls: "btn-secondary btn-small" });
    save.addEventListener("click", () => {
      const key = box.value.trim();
      box.value = "";
      if (!key) return undefined;
      return setKey({ id: r.id, key, ...(url ? { base_url: url.value.trim() } : {}) });
    });
    const kids = [box, url, save];
    if (r.key_source === "file") {
      const remove = el("button", {
        type: "button",
        text: "Remove key",
        cls: "btn-secondary btn-small",
      });
      remove.addEventListener("click", () => setKey({ id: r.id, key: "" }));
      kids.push(remove);
    }
    return el("div", { cls: "path-row provider-key" }, kids);
  }

  function renderList() {
    $("providers-list").replaceChildren(
      ...list.map((r) => {
        const st = providerStatus(r);
        const row = el("div", { cls: "provider-row" }, [
          el("div", { cls: "provider-name", text: r.label }),
          el("div", { cls: `provider-status provider-${st.state}`, text: st.text }),
          r.policy ? el("small", { cls: "provider-policy", text: r.policy }) : null,
          st.keyNote ? el("small", { cls: "provider-key-note", text: st.keyNote }) : null,
          r.tier === "key" ? keyControls(r) : null,
        ]);
        row.setAttribute("data-provider", r.id);
        row.setAttribute("data-state", st.state);
        return row;
      }),
    );
  }

  function render() {
    renderMenus();
    renderList();
    // The choices the daemon accepted, readable by the acceptance driver.
    $("providers-root").setAttribute("data-selection", JSON.stringify(sel));
    onRender();
  }

  // Send one choice. A refusal is shown and nothing changes.
  async function choose(role, id, model) {
    const r = await ask("set_provider", { role, id, model });
    if (!r.ok) {
      message(r.error ?? "The provider could not be set.");
      render();
      return false;
    }
    sel[role] = { id, model };
    store(role, sel[role]);
    notes[role] = null;
    showNotes();
    render();
    return true;
  }

  // Right after hello, before the provider list and whatever else: tell the
  // daemon which providers this pane stored. Until it answers the daemon
  // sends nothing of this pane anywhere (daemon/providers/selection.mjs), and
  // the Review tab holds its requests (whenRestored). A stored choice the
  // daemon refuses is shown and kept: the menus stay on it and nothing is
  // switched to Claude for the user.
  function restoreStored() {
    const stored = Object.fromEntries(ROLES.map((r) => [r, settings[SETTING[r]] ?? null]));
    for (const role of ROLES) sel[role] = stored[role] ?? { id: "claude", model: null };
    for (const role of ROLES) notes[role] = null;
    const mine = ask("restore_providers", stored).then((r) => {
      if (waiting !== mine) return; // a newer connection asked again
      waiting = null;
      if (!r.ok || !r.selection?.state) {
        const why = r.error ?? "no answer";
        notes.daemon = `The daemon did not confirm the stored providers (${why}). Judgments and review turns are stopped until it does: reopen the pane, or choose a provider here.`;
      } else {
        notes.daemon = null;
        for (const role of ROLES) {
          if (r.selection.state[role] !== "refused") continue;
          const why = String(r.selection.error?.[role] ?? "it was refused").replace(/\.$/, "");
          notes[role] =
            `The stored provider for ${WORDS[role]} cannot be used: ${why}. ` +
            `${STOPPED[role]} are stopped until you choose a provider here.`;
        }
      }
      showNotes();
      render();
    });
    waiting = mine;
    message(
      "Waiting for the daemon to confirm the stored providers. Judgments and review turns start after that.",
    );
    return mine;
  }

  async function reload({ refresh = false } = {}) {
    const r = await ask("providers_list", refresh ? { refresh: true } : {});
    if (!r.ok || !Array.isArray(r.providers)) {
      message(`Provider list unavailable (${r.error ?? "no list in the answer"}).`);
      return;
    }
    list = r.providers;
    render();
  }

  async function setKey(payload) {
    const r = await ask("provider_set_key", payload);
    if (!r.ok) {
      message(r.error ?? "The key could not be stored.");
      return;
    }
    if (Array.isArray(r.providers)) list = r.providers;
    showNotes();
    render();
  }

  for (const role of ROLES) {
    $(`provider-${role}`).addEventListener("change", (e) => {
      // The provider's first model where it needs one, else its default.
      const c = choiceAfterProbe({ id: e.target.value, model: null }, list, role);
      return choose(role, c.id, c.model);
    });
    $(`provider-${role}-model`).addEventListener("change", (e) =>
      role === "turn" && sel.turn.id === "claude" && editsModel
        ? editsModel.set(e.target.value)
        : choose(role, sel[role].id, e.target.value || null),
    );
  }
  $("providers-recheck").addEventListener("click", () => reload({ refresh: true }));

  return {
    reload,
    restoreStored,
    // Draw again: the model of the chat box changed.
    render,
    // What the menus hold, per role: { id, model }.
    selection: () => structuredClone(sel),
    restoring: () => waiting !== null,
    // The request the Review tab waits for before it asks the daemon for a
    // judgment or a review turn; null when there is nothing to wait for.
    whenRestored: () => waiting,
    // For the status line of a review turn; null when Claude runs it.
    turnLabel() {
      if (sel.turn.id === "claude") return null;
      return `Proposing with ${labelOf(sel.turn.id)} · ${sel.turn.model ?? "default model"}`;
    },
  };
}
