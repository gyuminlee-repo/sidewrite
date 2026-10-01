// Providers section of the Setup tab (P14): which provider runs review turns
// and which one makes the judgments, out of the ones the daemon's probe
// passed. Pure helpers are in providers-view.js.
//
// Text goes in through textContent only. An API key is read from its box,
// sent once in provider_set_key and the box emptied before the answer comes;
// it is never put in the settings, a message or any other request.

import { providerRows, choiceAfterProbe, providerStatus, modelRows } from "./providers-view.js";

const $ = (id) => document.getElementById(id);
const ROLES = ["turn", "oneshot"];
const SETTING = { turn: "turnProvider", oneshot: "oneshotProvider" };
const DROPPED = {
  turn: (label) => `${label} can no longer run review turns; review turns use Claude.`,
  oneshot: (label) => `${label} is no longer available for judgments; judgments use Claude.`,
};
// The probe starts CLIs and calls servers; a choice waits for it.
const TIMEOUT = { timeoutMs: 60_000 };
// The one key provider without a fixed server: it takes a URL with its key.
const CUSTOM_ID = "openai-compatible";

function el(tag, { text, cls, ...props } = {}, kids = []) {
  const e = document.createElement(tag);
  if (text != null) e.textContent = text;
  if (cls) e.className = cls;
  Object.assign(e, props);
  e.append(...kids);
  return e;
}

export function initProvidersPane({ sendRequest, settings, saveSettings }) {
  let list = [];
  // What the daemon holds for this pane, as far as the pane told it.
  const sel = { turn: { id: "claude", model: null }, oneshot: { id: "claude", model: null } };

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
      // Claude's review turns run on the model chosen in the chat box.
      const ofChat = role === "turn" && sel.turn.id === "claude";
      const opts = ofChat
        ? [{ value: "", label: "Model of the chat box", selected: true }]
        : modelRows(
            list.find((r) => r.id === sel[role].id),
            sel[role].model,
          );
      models.replaceChildren(
        ...opts.map((o) => el("option", { value: o.value, text: o.label, selected: o.selected })),
      );
      models.value = ofChat ? "" : (sel[role].model ?? "");
      models.disabled = ofChat || opts.length < 2;
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
    message("");
    render();
    return true;
  }

  // After a (re)connect the daemon has its defaults for this pane: send the
  // stored choices again, each checked against the fresh list first.
  async function restore() {
    const notes = [];
    for (const role of ROLES) {
      const stored = settings[SETTING[role]] ?? null;
      const c = choiceAfterProbe(stored, list, role);
      sel[role] = { id: c.id, model: c.model };
      if (c.dropped) {
        notes.push(DROPPED[role](labelOf(c.dropped)));
        store(role, sel[role]);
      }
      if (!stored) continue;
      const r = await ask("set_provider", { role, ...sel[role] });
      if (!r.ok) {
        notes.push(r.error ?? `The stored provider for ${role} was refused.`);
        sel[role] = { id: "claude", model: null };
      }
    }
    message(notes.join(" "));
  }

  async function reload({ refresh = false } = {}) {
    const r = await ask("providers_list", refresh ? { refresh: true } : {});
    if (!r.ok || !Array.isArray(r.providers)) {
      message(`Provider list unavailable (${r.error ?? "no list in the answer"}).`);
      return;
    }
    list = r.providers;
    await restore();
    render();
  }

  async function setKey(payload) {
    const r = await ask("provider_set_key", payload);
    if (!r.ok) {
      message(r.error ?? "The key could not be stored.");
      return;
    }
    if (Array.isArray(r.providers)) list = r.providers;
    message("");
    render();
  }

  for (const role of ROLES) {
    $(`provider-${role}`).addEventListener("change", (e) => {
      // The provider's first model where it needs one, else its default.
      const c = choiceAfterProbe({ id: e.target.value, model: null }, list, role);
      return choose(role, c.id, c.model);
    });
    $(`provider-${role}-model`).addEventListener("change", (e) =>
      choose(role, sel[role].id, e.target.value || null),
    );
  }
  $("providers-recheck").addEventListener("click", () => reload({ refresh: true }));

  return {
    reload,
    // For the status line of a review turn; null when Claude runs it.
    turnLabel() {
      if (sel.turn.id === "claude") return null;
      return `Proposing with ${labelOf(sel.turn.id)} · ${sel.turn.model ?? "default model"}`;
    },
  };
}
