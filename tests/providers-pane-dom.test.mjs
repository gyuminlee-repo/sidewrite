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
  // As in a browser: anything that is not a node becomes a text node, so an
  // appended null shows as the text "null".
  append(...kids) {
    for (const k of kids) {
      if (k instanceof FakeEl) this.children.push(k);
      else this.children.push(Object.assign(new FakeEl("#text"), { textContent: String(k) }));
    }
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

const NO_LOOP = "Judgments only: review turns are not offered for this provider.";
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
    roles: ["oneshot"], // judgments only, as the daemon's rows have it (registry.mjs)
    available: true,
    turn_ready: false,
    turn_note: NO_LOOP,
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

const confirmed = (p, over = {}) => ({
  ok: true,
  selection: {
    turn: p.turn ?? { id: "claude", model: null },
    writing: p.writing ?? null,
    oneshot: p.oneshot ?? null,
    state: { turn: "ready", writing: "ready", oneshot: "ready" },
    error: { turn: null, writing: null, oneshot: null },
    ...over,
  },
});

async function setup({
  settings = {},
  respond = {},
  list = rows(),
  connect = true,
  editsModel = undefined,
} = {}) {
  globalThis.document = fakeDocument();
  const { initProvidersPane } = await import("../taskpane/shared/providers-pane.js");
  const requests = [];
  const saved = [];
  const sendRequest = async (type, payload, opts) => {
    requests.push({ type, payload, opts });
    if (respond[type]) return respond[type](payload);
    if (type === "providers_list") return { ok: true, providers: list, selection: {} };
    // The daemon confirms what the pane stored (selection.mjs restore).
    if (type === "restore_providers") return confirmed(payload);
    return { ok: true };
  };
  const pane = initProvidersPane({
    sendRequest,
    settings,
    saveSettings: (s) => saved.push(JSON.stringify(s)),
    ...(editsModel ? { editsModel } : {}),
  });
  // As taskpane.js: the stored choices right after hello, the list on welcome.
  if (connect) {
    await pane.restoreStored();
    await pane.reload();
  }
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
  // Edits: Claude only; the others are listed, disabled, each with its reason.
  assert.deepEqual(optionTexts(t.$("provider-turn")), [
    "Claude (Agent SDK)",
    "Codex CLI (not for edits)",
    "Ollama (local) (not for edits)",
    "Groq (No API key)",
  ]);
  assert.deepEqual(
    t.$("provider-turn").children.map((o) => [o.disabled, o.title]),
    [
      [false, ""],
      [true, NO_LOOP],
      [true, NO_LOOP],
      [true, "No API key"],
    ],
  );
  // Writing: what is available, as for judgments.
  assert.deepEqual(optionTexts(t.$("provider-writing")), optionTexts(t.$("provider-oneshot")));
  assert.equal(t.$("provider-writing").value, "claude");
  assert.deepEqual(optionTexts(t.$("provider-writing-model")), ["Best model (default)", "Sonnet"]);
  assert.equal(optionTexts(t.$("provider-oneshot-model"))[0], "Default model");
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
  assert.ok(codex.includes("Available for judgments."));
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
  // A row part that is left out (no key note, no key box) leaves no text.
  assert.deepEqual(
    findAll(t.$("providers-list"), (n) => n.tagName === "#TEXT").map((n) => n.textContent),
    [],
  );
  // The stored choices first (none here: the daemon's defaults), then the list.
  assert.deepEqual(
    t.requests.map((r) => [r.type, r.payload]),
    [
      ["restore_providers", { turn: null, writing: null, oneshot: null }],
      ["providers_list", {}],
    ],
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
  // What the daemon accepted, for the acceptance driver (scripts/e2e).
  assert.deepEqual(JSON.parse(t.$("providers-root").attributes["data-selection"]), {
    turn: { id: "claude", model: null },
    writing: { id: "claude", model: null },
    oneshot: { id: "ollama", model: "gemma4:12b" },
  });
});

// Writing (response drafts, translations, questions) has a provider and a
// model of its own, stored apart from the one of the judgments.
test("choosing a writing provider and model sends them under their own role and stores them", async () => {
  const t = await setup({ settings: { oneshotProvider: { id: "ollama", model: "gemma4:12b" } } });
  assert.equal(t.$("provider-writing").value, "claude");
  const m = t.$("provider-writing-model");
  m.value = "sonnet";
  await m.fire("change");
  assert.deepEqual(t.sent("set_provider"), [{ role: "writing", id: "claude", model: "sonnet" }]);
  assert.deepEqual(t.settings.writingProvider, { id: "claude", model: "sonnet" });
  const sel = t.$("provider-writing");
  sel.value = "codex";
  await sel.fire("change");
  assert.deepEqual(t.sent("set_provider")[1], { role: "writing", id: "codex", model: null });
  assert.deepEqual(t.settings.writingProvider, { id: "codex", model: null });
  // The judgments keep theirs.
  assert.deepEqual(t.settings.oneshotProvider, { id: "ollama", model: "gemma4:12b" });
  assert.equal(t.$("provider-oneshot").value, "ollama");
  // Sent with the others after a reconnect.
  const again = await setup({ settings: t.settings });
  assert.deepEqual(again.sent("restore_providers"), [
    {
      turn: null,
      writing: { id: "codex", model: null },
      oneshot: { id: "ollama", model: "gemma4:12b" },
    },
  ]);
  assert.equal(again.$("provider-writing").value, "codex");
});

// The model of edits is one value: the one the chat box holds. Setup shows
// it and changes it through the same function, so the two cannot differ.
test("the model of edits in Setup is the chat box's: shown from it, changed through it", async () => {
  let model = "opus";
  const set = [];
  const editsModel = {
    get: () => model,
    rows: () => [
      { value: "sonnet", label: "Sonnet 5.5", title: "" },
      { value: "opus", label: "Opus 5.5", title: "" },
    ],
    set: (v) => (set.push(v), (model = v)),
  };
  const t = await setup({ editsModel });
  const m = t.$("provider-turn-model");
  assert.equal(m.disabled, false);
  assert.deepEqual(optionTexts(m), ["Sonnet 5.5", "Opus 5.5"]);
  assert.equal(m.value, "opus");
  m.value = "sonnet";
  await m.fire("change");
  assert.deepEqual(set, ["sonnet"]);
  // Not a provider choice: nothing goes to set_provider and nothing is stored here.
  assert.deepEqual(t.sent("set_provider"), []);
  assert.equal(t.saved.length, 0);
  // A change made in the chat box shows here after render().
  model = "opus";
  t.pane.render();
  assert.equal(t.$("provider-turn-model").value, "opus");
  // What the pane chose, for the line that says what runs.
  assert.deepEqual(t.pane.selection(), {
    turn: { id: "claude", model: null },
    writing: { id: "claude", model: null },
    oneshot: { id: "claude", model: null },
  });
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

// Codex review of stage 2.5: the stored choice must reach the daemon before
// anything can be asked of it, and a choice that cannot be used is shown,
// never replaced by Claude behind the user's back.
test("the stored choices go to the daemon right after hello, before and without the provider list", async () => {
  let answer;
  const t = await setup({
    connect: false,
    settings: { oneshotProvider: { id: "ollama", model: "gemma4:12b" } },
    respond: { restore_providers: (p) => new Promise((r) => (answer = () => r(confirmed(p)))) },
  });
  const restoring = t.pane.restoreStored();
  assert.deepEqual(
    t.requests.map((r) => [r.type, r.payload]),
    [
      [
        "restore_providers",
        { turn: null, writing: null, oneshot: { id: "ollama", model: "gemma4:12b" } },
      ],
    ],
  );
  // Until the daemon answers, the Review tab holds its requests (whenRestored).
  assert.equal(t.pane.restoring(), true);
  assert.equal(
    t.$("providers-message").textContent,
    "Waiting for the daemon to confirm the stored providers. Judgments and review turns start after that.",
  );
  let released = false;
  t.pane.whenRestored().then(() => (released = true));
  await new Promise((r) => setTimeout(r, 5));
  assert.equal(released, false);
  answer();
  await restoring;
  await new Promise((r) => setTimeout(r, 0));
  assert.deepEqual([t.pane.restoring(), released], [false, true]);
  assert.equal(t.pane.whenRestored(), null, "nothing to wait for any more");
  assert.equal(t.$("providers-message").textContent, "");
  // The list comes later; the menus then show the confirmed choice.
  await t.pane.reload();
  assert.equal(t.$("provider-oneshot").value, "ollama");
  assert.equal(t.$("provider-oneshot-model").value, "gemma4:12b");
  assert.deepEqual(t.sent("set_provider"), [], "restore is the only message that sets it");
  assert.deepEqual(t.settings.oneshotProvider, { id: "ollama", model: "gemma4:12b" });
});

test("a stored choice the daemon refuses is shown and kept: no fallback to Claude, the menu stays on it", async () => {
  const stored = {
    oneshotProvider: { id: "groq", model: "x" },
    turnProvider: { id: "codex", model: "gpt-6.1-sol" },
  };
  const t = await setup({
    settings: structuredClone(stored),
    respond: {
      restore_providers: (p) =>
        confirmed(p, {
          turn: null,
          oneshot: null,
          state: { turn: "refused", oneshot: "refused" },
          error: {
            turn: "Codex CLI cannot run review turns. Judgments only: review turns are not offered for this provider.",
            oneshot: "Groq is not available: No API key",
          },
        }),
    },
  });
  assert.equal(
    t.$("providers-message").textContent,
    "The stored provider for review turns cannot be used: Codex CLI cannot run review turns. Judgments only: review turns are not offered for this provider. Review turns are stopped until you choose a provider here. " +
      "The stored provider for judgments cannot be used: Groq is not available: No API key. Judgments are stopped until you choose a provider here.",
  );
  assert.equal(t.$("providers-message").hidden, false);
  // The menus stay on the stored choices, which their rows mark as not usable.
  assert.equal(t.$("provider-oneshot").value, "groq");
  assert.equal(t.$("provider-turn").value, "codex");
  assert.equal(t.$("provider-oneshot").children[3].textContent, "Groq (No API key)");
  // Nothing was changed for the user: no Claude was set, the settings are as stored.
  assert.deepEqual(t.sent("set_provider"), []);
  assert.deepEqual(t.settings, stored);
  assert.equal(t.saved.length, 0);
  // Picking a provider that passes ends it.
  const sel = t.$("provider-oneshot");
  sel.value = "ollama";
  await sel.fire("change");
  assert.deepEqual(t.sent("set_provider"), [
    { role: "oneshot", id: "ollama", model: "qwen3.5:9b" },
  ]);
  assert.match(t.$("providers-message").textContent, /^The stored provider for review turns/);
  assert.doesNotMatch(t.$("providers-message").textContent, /for judgments cannot/);
});

test("no answer to the stored choices: the pane says so and stops waiting; the choice is not dropped", async () => {
  const t = await setup({
    settings: { oneshotProvider: { id: "ollama", model: "gemma4:12b" } },
    respond: {
      restore_providers: () => {
        throw new Error('Request "restore_providers" timed out');
      },
    },
  });
  assert.equal(
    t.$("providers-message").textContent,
    'The daemon did not confirm the stored providers (Request "restore_providers" timed out). Judgments and review turns are stopped until it does: reopen the pane, or choose a provider here.',
  );
  assert.equal(t.pane.restoring(), false);
  assert.deepEqual(t.settings.oneshotProvider, { id: "ollama", model: "gemma4:12b" });
  assert.equal(t.$("provider-oneshot").value, "ollama");
});

test("after a daemon restart the same stored choices are sent again, unchanged", async () => {
  const settings = { oneshotProvider: { id: "ollama", model: "gemma4:12b" }, turnProvider: null };
  const first = await setup({ settings });
  // The daemon restarts; the pane reconnects with the settings it kept.
  const second = await setup({ settings });
  for (const t of [first, second]) {
    assert.deepEqual(t.sent("restore_providers"), [
      { turn: null, writing: null, oneshot: { id: "ollama", model: "gemma4:12b" } },
    ]);
    assert.equal(t.requests[0].type, "restore_providers");
    assert.equal(t.$("provider-oneshot").value, "ollama");
  }
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
