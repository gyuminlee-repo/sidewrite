// Behavior of the Tools tab glue (tools-pane.js) on a minimal fake DOM, as
// tests/response-pane-dom.test.mjs does for the Response tab. No Office.js:
// `word` and `sendRequest` are stubs; a job answers { accepted, job_id } and
// its result comes as a tools_event.
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
    this.checked = false;
    this.disabled = false;
    this.files = [];
    this.classList = { toggle() {} };
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
    for (const fn of this.listeners[ev] ?? []) await fn();
  }
  click() {
    return this.fire("click");
  }
}

function fakeDocument(missing = []) {
  const byId = new Map();
  return {
    getElementById(id) {
      if (missing.includes(id)) return null;
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
const texts = (node) => findAll(node, (n) => n.textContent).map((n) => n.textContent);
const tick = () => new Promise((r) => setTimeout(r, 5));

const JOBS = new Set([
  "selection_translate",
  "selection_ask",
  "profile_rules_extract",
  "marked_build",
  "anonymize",
]);
const STATE = {
  ok: true,
  doc_name: "paper.docx",
  profile_name: "generic",
  policy_packs: { review_scope: true },
};

// respond: { <type>: (payload) => answer }. hold: job types whose job_done is
// kept back until deliver(type) is called.
async function setup({ respond = {}, word = {}, hold = [], missing = [], settings = {} } = {}) {
  globalThis.document = fakeDocument(missing);
  const { initToolsPane } = await import("../taskpane/shared/tools-pane.js");
  const requests = [];
  const held = new Map();
  let n = 0;
  let pane = null;
  const sendRequest = async (type, payload) => {
    requests.push({ type, payload });
    const answer = () =>
      respond[type] ? respond[type](payload) : type === "tools_state" ? STATE : { ok: true };
    if (!JOBS.has(type)) return answer();
    const job_id = `j${++n}`;
    const deliver = () =>
      pane.onEvent({ type: "tools_event", event: "job_done", job_id, result: answer() });
    if (hold.includes(type)) held.set(type, deliver);
    else setTimeout(deliver, 0);
    return { ok: true, accepted: true, job_id };
  };
  const saved = [];
  pane = initToolsPane({
    sendRequest,
    settings,
    saveSettings: (s) => saved.push({ ...s }),
    word: {
      selectionText: async () => "The cells were grown at 30 °C.",
      ...word,
    },
  });
  await pane.reload();
  const $ = (id) => document.getElementById(id);
  return {
    pane,
    requests,
    saved,
    $,
    of: (type) => requests.filter((r) => r.type === type),
    deliver: (type) => held.get(type)(),
    message: () => $("tools-message").textContent,
  };
}

test("the tab says which saved file the tools work on", async () => {
  const t = await setup();
  assert.match(t.$("tools-doc").textContent, /^Saved file: paper\.docx\. Profile: generic\./);
  assert.equal(t.message(), "");
});

test("a document that is not saved: the tab says so and shows the daemon's reason", async () => {
  const t = await setup({
    respond: {
      tools_state: () => ({ ok: false, error: "Save the document to a local folder first." }),
    },
  });
  assert.match(t.$("tools-doc").textContent, /^Save the document to a local folder/);
  assert.equal(t.message(), "Save the document to a local folder first.");
});

// ---- selection: translate and ask --------------------------------------------

test("translate: the selection and the language are sent, the answer is shown, Word is only read", async () => {
  const wordCalls = [];
  const t = await setup({
    settings: { toolsLang: "Korean" },
    word: { selectionText: async () => (wordCalls.push("selectionText"), " Grown at 30 °C. ") },
    respond: {
      selection_translate: (p) => ({
        ok: true,
        text: "30 °C 에서 배양했다.",
        lang: p.lang,
        model: "sonnet",
      }),
    },
  });
  assert.equal(t.$("tools-lang").value, "Korean");
  await t.$("tools-translate").click();
  await tick();
  assert.deepEqual(t.of("selection_translate")[0].payload, {
    text: "Grown at 30 °C.",
    lang: "Korean",
  });
  assert.deepEqual(texts(t.$("tools-selection")), [
    'Translation into Korean (sonnet) for: "Grown at 30 °C."',
    "30 °C 에서 배양했다.",
  ]);
  assert.deepEqual(wordCalls, ["selectionText"]);
  assert.equal(t.saved.at(-1).toolsLang, "Korean");
  assert.equal(t.message(), "");
});

test("ask: the question goes with the selection and the answer is shown", async () => {
  const t = await setup({
    respond: { selection_ask: () => ({ ok: true, text: "At 30 °C.", model: "haiku" }) },
  });
  t.$("tools-question").value = "At what temperature?";
  await t.$("tools-ask").click();
  await tick();
  assert.deepEqual(t.of("selection_ask")[0].payload, {
    text: "The cells were grown at 30 °C.",
    question: "At what temperature?",
  });
  assert.deepEqual(texts(t.$("tools-selection")), [
    'Answer (haiku) for: "The cells were grown at 30 °C."',
    "At 30 °C.",
  ]);
});

test("nothing selected, no language or no question: said in the tab, nothing is sent", async () => {
  let t = await setup({ word: { selectionText: async () => "  " } });
  t.$("tools-lang").value = "Korean";
  await t.$("tools-translate").click();
  assert.equal(t.message(), "Select some text in the document first.");
  t = await setup();
  await t.$("tools-translate").click();
  assert.equal(t.message(), "Name the language to translate into.");
  await t.$("tools-ask").click();
  assert.equal(t.message(), "Type a question about the selected text.");
  assert.equal(t.of("selection_translate").length + t.of("selection_ask").length, 0);
});

test("a failed job shows the daemon's reason and leaves the buttons usable", async () => {
  const t = await setup({
    respond: { selection_ask: () => ({ ok: false, error: "the answer failed: usage limit" }) },
  });
  t.$("tools-question").value = "Why?";
  await t.$("tools-ask").click();
  await tick();
  assert.equal(t.message(), "the answer failed: usage limit");
  assert.equal(t.$("tools-ask").disabled, false);
});

test("a new connection ends a job that was waiting, with an error, and the buttons come back", async () => {
  const t = await setup({ hold: ["selection_ask"] });
  t.$("tools-question").value = "Why?";
  const pressed = t.$("tools-ask").click();
  await tick();
  assert.equal(t.$("tools-ask").disabled, true);
  await t.pane.reload(); // welcome of the new connection
  await pressed;
  await tick();
  assert.match(t.message(), /connection to the daemon was lost/);
  assert.equal(t.$("tools-ask").disabled, false);
  assert.equal(t.$("tools-translate").disabled, false);
});

test("a page without the Tools tab (Excel) gets a pane that does nothing", async () => {
  const t = await setup({ missing: ["tools-root"] });
  assert.equal(t.requests.length, 0);
  t.pane.onEvent({ event: "job_done", job_id: "x", result: { ok: true } });
});
