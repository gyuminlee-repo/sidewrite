// Behavior of the Providers section glue (providers-pane.js) on a minimal
// fake DOM. `sendRequest` is a stub that plays the daemon.
import { test } from "node:test";
import assert from "node:assert/strict";

class FakeEl {
  constructor(tag, id = null) {
    this.tagName = tag.toUpperCase();
    this.id = id;
    this.children = [];
    this.listeners = {};
    this.attributes = {};
    this.textContent = "";
    this.hidden = false;
    this.value = "";
    this.disabled = false;
    this.selected = false;
    this.className = "";
  }
  addEventListener(ev, fn) {
    (this.listeners[ev] ??= []).push(fn);
  }
  append(...kids) {
    this.children.push(...kids.filter((k) => k && typeof k === "object"));
  }
  replaceChildren(...kids) {
    this.children = [];
    this.append(...kids);
  }
  setAttribute(k, v) {
    this.attributes[k] = v;
  }
  async fire(ev) {
    for (const fn of this.listeners[ev] ?? []) await fn({ target: this });
  }
}

function fakeDocument() {
  const byId = new Map();
  return {
    getElementById(id) {
      if (!byId.has(id)) byId.set(id, new FakeEl("div", id));
      return byId.get(id);
    },
    createElement: (tag) => new FakeEl(tag),
  };
}

function findAll(node, pred, out = []) {
  if (pred(node)) out.push(node);
  for (const k of node.children) findAll(k, pred, out);
  return out;
}

const NO_LOOP = "Judgments only: no tool loop for this provider in this stage.";
const rows = () => [
  {
    id: "claude",
    tier: "subscription",
    label: "Claude (Agent SDK)",
    roles: ["turn", "oneshot"],
    available: true,
    turn_ready: true,
    turn_note: null,
    reason: null,
    policy: "Anthropic terms.",
    models: [{ value: "sonnet", displayName: "Sonnet" }],
  },
  {
    id: "codex",
    tier: "subscription",
    label: "Codex CLI",
    roles: ["turn", "oneshot"],
    available: true,
    turn_ready: false,
    turn_note: "Judgments only. Codex CLI will not call the office tools without an approval.",
    reason: null,
    policy: "OpenAI terms.",
    models: [{ value: "gpt-6.1-sol", displayName: "GPT 6.1 Sol" }],
  },
  {
    id: "ollama",
    tier: "local",
    label: "Ollama (local)",
    roles: ["oneshot"],
    available: true,
    turn_ready: false,
    turn_note: NO_LOOP,
    reason: null,
    policy: "Never leaves this machine.",
    models: [
      { value: "qwen3.5:9b", displayName: "qwen3.5:9b" },
      { value: "gemma4:12b", displayName: "gemma4:12b" },
    ],
  },
  {
    id: "groq",
    tier: "key",
    label: "Groq",
    roles: ["oneshot"],
    available: false,
    turn_ready: false,
    turn_note: NO_LOOP,
    reason: "No API key",
    policy: "Kept 30 days.",
    has_key: false,
    key_source: null,
    default_enabled: false,
    models: [],
  },
];

async function setup({ settings = {}, respond = {}, list = rows() } = {}) {
  globalThis.document = fakeDocument();
  const { initProvidersPane } = await import("../taskpane/shared/providers-pane.js");
  const requests = [];
  const saved = [];
  const sendRequest = async (type, payload, opts) => {
    requests.push({ type, payload, opts });
    if (respond[type]) return respond[type](payload);
    if (type === "providers_list") return { ok: true, providers: list, selection: {} };
    return { ok: true };
  };
  const pane = initProvidersPane({
    sendRequest,
    settings,
    saveSettings: (s) => saved.push(JSON.stringify(s)),
  });
  await pane.reload();
  const $ = (id) => document.getElementById(id);
  const sent = (type) => requests.filter((r) => r.type === type).map((r) => r.payload);
  return { pane, requests, saved, settings, $, sent };
}

const optionTexts = (sel) => sel.children.map((o) => o.textContent);
const rowOf = (t, id) =>
  findAll(t.$("providers-list"), (n) => n.attributes["data-provider"] === id)[0];
const textsOf = (node) => findAll(node, (n) => !!n.textContent).map((n) => n.textContent);

test("the menus list what passed; a provider that cannot be chosen is disabled and says why", async () => {
  const t = await setup();
  assert.deepEqual(optionTexts(t.$("provider-turn")), [
    "Claude (Agent SDK)",
    "Codex CLI (judgments only)",
  ]);
  assert.deepEqual(
    t.$("provider-turn").children.map((o) => o.disabled),
    [false, true],
  );
  assert.deepEqual(optionTexts(t.$("provider-oneshot")), [
    "Claude (Agent SDK)",
    "Codex CLI",
    "Ollama (local)",
    "Groq (No API key)",
  ]);
  assert.equal(t.$("provider-oneshot").children[3].disabled, true);
  assert.equal(t.$("provider-turn").value, "claude");
  assert.equal(t.$("provider-oneshot").value, "claude");
  // Claude's review turns use the model of the chat box.
  assert.equal(t.$("provider-turn-model").disabled, true);
  // The list under the menus: status, the reason, the data policy.
  const codex = textsOf(rowOf(t, "codex"));
  assert.ok(codex.includes("Codex CLI"));
  assert.ok(codex.some((x) => /^Available for judgments\. Judgments only\. Codex CLI/.test(x)));
  assert.ok(codex.includes("OpenAI terms."));
  const groq = textsOf(rowOf(t, "groq"));
  assert.ok(groq.includes("Not available: No API key"));
  assert.ok(groq.includes("Off until you enter a key."));
  assert.ok(groq.includes("Kept 30 days."));
  // Only a key provider has a key box, and it is a password box.
  const boxes = findAll(
    t.$("providers-list"),
    (n) => n.tagName === "INPUT" && n.type === "password",
  );
  assert.equal(boxes.length, 1);
  assert.equal(findAll(rowOf(t, "ollama"), (n) => n.tagName === "INPUT").length, 0);
  // Nothing was sent but the list request: the defaults are the daemon's own.
  assert.deepEqual(
    t.requests.map((r) => r.type),
    ["providers_list"],
  );
});

test("choosing a judgment provider sends it with its first model and stores the choice", async () => {
  const t = await setup();
  const sel = t.$("provider-oneshot");
  sel.value = "ollama";
  await sel.fire("change");
  assert.deepEqual(t.sent("set_provider"), [
    { role: "oneshot", id: "ollama", model: "qwen3.5:9b" },
  ]);
  assert.deepEqual(t.settings.oneshotProvider, { id: "ollama", model: "qwen3.5:9b" });
  assert.equal(t.saved.length, 1);
  assert.deepEqual(optionTexts(t.$("provider-oneshot-model")), ["qwen3.5:9b", "gemma4:12b"]);

  const m = t.$("provider-oneshot-model");
  m.value = "gemma4:12b";
  await m.fire("change");
  assert.deepEqual(t.sent("set_provider")[1], {
    role: "oneshot",
    id: "ollama",
    model: "gemma4:12b",
  });
  assert.deepEqual(t.settings.oneshotProvider, { id: "ollama", model: "gemma4:12b" });
  assert.equal(t.pane.turnLabel(), null, "review turns are still Claude's");
});

test("a refused choice is shown and not stored; the menu goes back", async () => {
  const t = await setup({
    respond: {
      set_provider: () => ({ ok: false, error: "Codex CLI is not available: not logged in" }),
    },
  });
  const sel = t.$("provider-oneshot");
  sel.value = "codex";
  await sel.fire("change");
  assert.equal(t.$("providers-message").textContent, "Codex CLI is not available: not logged in");
  assert.equal(t.$("providers-message").hidden, false);
  assert.equal(sel.value, "claude");
  assert.equal(t.settings.oneshotProvider, undefined);
  assert.equal(t.saved.length, 0);
});

test("a stored choice is sent again after a reconnect; one that stopped passing goes back to Claude", async () => {
  const kept = await setup({
    settings: { oneshotProvider: { id: "ollama", model: "gemma4:12b" } },
  });
  assert.deepEqual(kept.sent("set_provider"), [
    { role: "oneshot", id: "ollama", model: "gemma4:12b" },
  ]);
  assert.equal(kept.$("provider-oneshot").value, "ollama");
  assert.equal(kept.$("provider-oneshot-model").value, "gemma4:12b");

  const gone = await setup({
    settings: {
      oneshotProvider: { id: "groq", model: "x" },
      turnProvider: { id: "codex", model: "gpt-6.1-sol" },
    },
  });
  assert.deepEqual(gone.sent("set_provider"), [
    { role: "turn", id: "claude", model: null },
    { role: "oneshot", id: "claude", model: null },
  ]);
  assert.equal(gone.settings.oneshotProvider, null);
  assert.equal(gone.settings.turnProvider, null);
  assert.match(gone.$("providers-message").textContent, /Codex CLI can no longer run review turns/);
  assert.match(
    gone.$("providers-message").textContent,
    /Groq is no longer available for judgments/,
  );
  assert.equal(gone.$("provider-oneshot").value, "claude");
});

test("a key goes to the daemon once, the box is emptied, and no setting or other request holds it", async () => {
  const SECRET = "gsk_live_0123456789abcdef";
  const after = rows().map((r) =>
    r.id === "groq"
      ? {
          ...r,
          available: true,
          reason: null,
          has_key: true,
          key_source: "file",
          models: [{ value: "llama-4", displayName: "llama-4" }],
        }
      : r,
  );
  let box;
  const t = await setup({
    respond: {
      provider_set_key: () => {
        // The box is already empty while the request is on its way.
        assert.equal(box.value, "");
        return { ok: true, providers: after, selection: {} };
      },
    },
  });
  const row = rowOf(t, "groq");
  box = findAll(row, (n) => n.tagName === "INPUT" && n.type === "password")[0];
  const save = findAll(row, (n) => n.tagName === "BUTTON" && n.textContent === "Save key")[0];
  box.value = SECRET;
  await save.fire("click");
  assert.deepEqual(t.sent("provider_set_key"), [{ id: "groq", key: SECRET }]);
  // Rendered from the reply: Groq can be chosen now, its key box is new and empty.
  assert.equal(t.$("provider-oneshot").children[3].disabled, false);
  assert.ok(textsOf(rowOf(t, "groq")).includes("A key is stored in the key file."));
  assert.equal(
    findAll(document.getElementById("providers-list"), (n) => n.value === SECRET).length,
    0,
  );
  const elsewhere = t.requests.filter((r) => r.type !== "provider_set_key");
  assert.equal(JSON.stringify(elsewhere).includes(SECRET), false);
  assert.equal(JSON.stringify(t.settings).includes(SECRET), false);
  assert.equal(t.saved.join("").includes(SECRET), false);
  assert.equal(t.$("providers-message").textContent.includes(SECRET), false);

  // A stored key can be removed: an empty key is sent.
  const remove = findAll(
    rowOf(t, "groq"),
    (n) => n.tagName === "BUTTON" && n.textContent === "Remove key",
  )[0];
  await remove.fire("click");
  assert.deepEqual(t.sent("provider_set_key")[1], { id: "groq", key: "" });
});

test("an empty key box sends nothing; Check again asks for a fresh probe", async () => {
  const t = await setup();
  const row = rowOf(t, "groq");
  const save = findAll(row, (n) => n.tagName === "BUTTON" && n.textContent === "Save key")[0];
  await save.fire("click");
  assert.deepEqual(t.sent("provider_set_key"), []);
  await t.$("providers-recheck").fire("click");
  assert.deepEqual(t.sent("providers_list"), [{}, { refresh: true }]);
});

test("the list request failing leaves a message and the menus as they were", async () => {
  const t = await setup({
    respond: {
      providers_list: () => {
        throw new Error('Request "providers_list" timed out');
      },
    },
  });
  assert.match(t.$("providers-message").textContent, /Provider list unavailable \(Request/);
  assert.equal(t.$("provider-oneshot").children.length, 0);
});
