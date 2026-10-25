// Behavior of the Tools tab glue (tools-pane.js) on a minimal fake DOM, as
// tests/response-pane-dom.test.mjs does for the Response tab. No Office.js:
// `word` and `sendRequest` are stubs; a job answers { accepted, job_id } and
// its result comes as a tools_event.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

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
async function setup({
  respond = {},
  word = {},
  hold = [],
  missing = [],
  settings = {},
  usedLabel,
} = {}) {
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
  const closers = [];
  pane = initToolsPane({
    sendRequest,
    onDisconnect: (fn) => closers.push(fn),
    usedLabel,
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
    // The socket closed (taskpane.js ws.onclose).
    close: () => closers.forEach((fn) => fn()),
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

test("a closed socket ends a waiting job at once, with an error, and the buttons come back", async () => {
  const t = await setup({ hold: ["selection_ask"] });
  t.$("tools-question").value = "Why?";
  const pressed = t.$("tools-ask").click();
  await tick();
  assert.equal(t.$("tools-ask").disabled, true);
  t.close(); // no welcome follows: the daemon does not come back
  await pressed;
  await tick();
  assert.match(t.message(), /connection to the daemon was lost/);
  assert.equal(t.$("tools-ask").disabled, false);
  assert.equal(t.$("tools-translate").disabled, false);
});

test("a job whose job_done comes while the socket is open still ends with its answer", async () => {
  const t = await setup({
    hold: ["selection_ask"],
    respond: { selection_ask: () => ({ ok: true, answer: "Because of the heat." }) },
  });
  t.$("tools-question").value = "Why?";
  const pressed = t.$("tools-ask").click();
  await tick();
  t.deliver("selection_ask");
  await pressed;
  await tick();
  assert.doesNotMatch(t.message(), /connection to the daemon was lost/);
  assert.equal(t.$("tools-ask").disabled, false);
});

// ---- journal guideline -------------------------------------------------------

const CANDIDATES = [
  {
    kind: "instruction",
    text: "Provide 4 to 6 keywords.",
    scope: "manuscript",
    check: { type: "count", target: "keywords", min: 4, max: 6 },
    source: { kind: "guideline", sentence: "Provide 4 to 6 keywords." },
  },
  {
    kind: "instruction",
    text: "Avoid jargon.",
    scope: "manuscript",
    check: { type: "llm" },
    source: { kind: "guideline", sentence: "Avoid jargon." },
  },
];

test("guideline: candidates are listed with their sentence; only the ticked ones are added", async () => {
  const t = await setup({
    respond: {
      profile_rules_extract: () => ({ ok: true, candidates: CANDIDATES, warnings: ["s3: kept"] }),
      profile_rules_apply: () => ({ ok: true, added: 1, skipped: 0, profile_keys: [] }),
    },
  });
  assert.equal(t.$("tools-rules-add").hidden, true);
  t.$("tools-guideline").value = " Provide 4 to 6 keywords. Avoid jargon. ";
  await t.$("tools-rules-extract").click();
  await tick();
  assert.deepEqual(t.of("profile_rules_extract")[0].payload, {
    text: "Provide 4 to 6 keywords. Avoid jargon.",
  });
  assert.deepEqual(texts(t.$("tools-rules")), [
    "Count rule: at least 4 and at most 6 keywords",
    'Guideline: "Provide 4 to 6 keywords."',
    "Rule the model checks",
    'Guideline: "Avoid jargon."',
    "s3: kept",
  ]);
  assert.equal(t.$("tools-rules-add").hidden, false);
  const boxes = findAll(t.$("tools-rules"), (n) => n.tagName === "INPUT");
  assert.deepEqual(
    boxes.map((b) => b.checked),
    [true, true],
  );
  boxes[1].checked = false;
  // The text box changed after the extraction: the text that was read is sent.
  t.$("tools-guideline").value = "something else";
  await t.$("tools-rules-add").click();
  assert.deepEqual(t.of("profile_rules_apply")[0].payload, {
    text: "Provide 4 to 6 keywords. Avoid jargon.",
    candidates: [CANDIDATES[0]],
  });
  assert.equal(t.message(), "1 rule added to the ledger.");
  assert.deepEqual(texts(t.$("tools-rules")), []);
  assert.equal(t.$("tools-rules-add").hidden, true);
});

test("guideline: nothing pasted or nothing ticked is said in the tab", async () => {
  const t = await setup({
    respond: { profile_rules_extract: () => ({ ok: true, candidates: CANDIDATES, warnings: [] }) },
  });
  await t.$("tools-rules-extract").click();
  assert.equal(t.message(), "Paste the guideline text first.");
  t.$("tools-guideline").value = "Avoid jargon.";
  await t.$("tools-rules-extract").click();
  await tick();
  for (const b of findAll(t.$("tools-rules"), (n) => n.tagName === "INPUT")) b.checked = false;
  await t.$("tools-rules-add").click();
  assert.equal(t.message(), "Pick at least one rule.");
  assert.equal(t.of("profile_rules_apply").length, 0);
});

test("the policy pack box shows the profile and saves an override of this paper", async () => {
  const t = await setup({
    respond: {
      profile_set: (p) => ({ ok: true, profile: { policy_packs: p.overrides.policy_packs } }),
    },
  });
  const box = t.$("tools-policy-scope");
  assert.equal(box.checked, true);
  box.checked = false;
  await box.fire("change");
  assert.deepEqual(t.of("profile_set")[0].payload, {
    overrides: { policy_packs: { review_scope: false } },
  });
  assert.equal(box.checked, false);
  assert.equal(t.message(), "Review turns no longer carry rules 2 to 4 for this paper.");
});

// ---- checks ------------------------------------------------------------------

test("checks: the pane reads Word, lists the findings, and a finding selects its paragraph", async () => {
  const wordCalls = [];
  const SNAP = { paragraphs: [{ index: 0, text: "Abstract" }], comments: [] };
  const t = await setup({
    word: {
      collectReviewSnapshot: async () => (wordCalls.push("snapshot"), SNAP),
      collectHeadings: async () => (
        wordCalls.push("headings"),
        [{ index: 0, level: 1, text: "Abstract" }]
      ),
      readFormatSnapshot: async () => (
        wordCalls.push("format"),
        { paragraphs: [{ index: 0 }], tables_read: true }
      ),
      selectParagraphByIndex: async (i) => wordCalls.push(`select ${i}`),
    },
    respond: {
      format_check: () => ({
        ok: true,
        format: {
          rules: ["statistics symbols in italics"],
          error: null,
          findings: [{ paragraph_index: 9, rule: "stat", find: "p", context: "(p < 0.05)" }],
        },
        counts: [{ id: "L1", text: "3 to 6 keywords.", status: "ok", count: 3, violations: [] }],
      }),
    },
  });
  await t.$("tools-check").click();
  assert.deepEqual(wordCalls, ["snapshot", "headings", "format"]);
  assert.deepEqual(t.of("format_check")[0].payload, {
    snapshot: SNAP,
    headings: [{ index: 0, level: 1, text: "Abstract" }],
    format: { paragraphs: [{ index: 0 }], tables_read: true },
  });
  assert.deepEqual(texts(t.$("tools-checks")), [
    "Format rules on: statistics symbols in italics, empty tables.",
    'Paragraph 9: "p" should be italic in "(p < 0.05)"',
    "L1 ok (3): 3 to 6 keywords.",
  ]);
  assert.equal(t.message(), "1 findings. Nothing was changed.");
  const go = findAll(t.$("tools-checks"), (n) => n.tagName === "BUTTON");
  assert.equal(go.length, 1);
  await go[0].click();
  assert.deepEqual(wordCalls.slice(3), ["select 9"]);
  // The check sent no request that changes anything.
  assert.deepEqual(
    t.requests.map((r) => r.type),
    ["tools_state", "format_check"],
  );
});

// ---- marked copy -------------------------------------------------------------

const docx = (name = "v1.docx", size = 3) => ({
  name,
  size,
  arrayBuffer: async () => Uint8Array.of(80, 75, 3).buffer,
});

test("marked copy: the earlier version and the mode are sent, the result is listed", async () => {
  const t = await setup({
    respond: {
      marked_build: (p) => ({
        ok: true,
        name: "paper_marked.docx",
        mode: p.mode,
        report: { paragraphs: { n_eq: 18, n_mod: 6, n_ins: 0, n_del: 0 }, warnings: [] },
        notice: "Read from disk.",
      }),
    },
  });
  t.$("tools-baseline").files = [docx()];
  t.$("tools-marked-mode").value = "color";
  await t.$("tools-marked-build").click();
  await tick();
  assert.deepEqual(t.of("marked_build")[0].payload, { mode: "color", baseline_base64: "UEsD" });
  assert.deepEqual(texts(t.$("tools-marked")), [
    "Written: paper_marked.docx (colours).",
    "Paragraphs: 6 changed, 0 added, 0 removed, 18 unchanged.",
    "Read from disk.",
  ]);
  assert.equal(t.message(), "");
});

test("marked copy: a refusal is shown as the tool wrote it, with its paragraph number", async () => {
  const refusal = 'paragraph 12 ("The cells were"): a text direction mark (w:dir) was added';
  const t = await setup({ respond: { marked_build: () => ({ ok: false, error: refusal }) } });
  t.$("tools-baseline").files = [docx()];
  await t.$("tools-marked-build").click();
  await tick();
  assert.equal(t.of("marked_build")[0].payload.mode, "track_changes");
  assert.deepEqual(texts(t.$("tools-marked")), ["No marked copy was written.", refusal]);
});

test("marked copy: no file, a file that is not a docx or one too large is said and not sent", async () => {
  const t = await setup();
  await t.$("tools-marked-build").click();
  assert.equal(t.message(), "Choose the earlier version (.docx) first.");
  t.$("tools-baseline").files = [docx("v1.doc")];
  await t.$("tools-marked-build").click();
  assert.equal(t.message(), "v1.doc is not a .docx file.");
  t.$("tools-baseline").files = [docx("v1.docx", 80 * 1024 * 1024)];
  await t.$("tools-marked-build").click();
  assert.equal(t.message(), "v1.docx is too large to send to the daemon.");
  assert.equal(t.of("marked_build").length, 0);
});

// ---- anonymized copy ---------------------------------------------------------

test("anonymized copy: the verifier's result is listed under the button", async () => {
  const t = await setup({
    respond: {
      anonymize: () => ({
        ok: true,
        name: "paper_anonymized.docx",
        report: {
          total: 0,
          body_equal: true,
          markup_equal: true,
          structure_ok: true,
          categories: {},
        },
        notice: "The text itself is not changed.",
      }),
    },
  });
  await t.$("tools-anonymize").click();
  await tick();
  assert.deepEqual(t.of("anonymize")[0].payload, {});
  assert.deepEqual(texts(t.$("tools-anonymized")), [
    "Written: paper_anonymized.docx.",
    "Verifier: total 0 left; text unchanged: yes; tracked changes unchanged: yes; structure intact: yes.",
    "The text itself is not changed.",
  ]);
  assert.equal(t.message(), "");
});

test("anonymized copy: a copy that failed its check is shown with its report; a tool error is a message", async () => {
  let t = await setup({
    respond: {
      anonymize: () => ({
        ok: false,
        name: "paper_anonymized.INCOMPLETE.docx",
        error: "The copy is incomplete: the text changed.",
        report: {
          total: 0,
          body_equal: false,
          markup_equal: true,
          structure_ok: true,
          categories: {},
        },
      }),
    },
  });
  await t.$("tools-anonymize").click();
  await tick();
  assert.match(
    texts(t.$("tools-anonymized"))[0],
    /set aside as paper_anonymized\.INCOMPLETE\.docx/,
  );
  assert.match(texts(t.$("tools-anonymized"))[2], /text unchanged: NO/);
  t = await setup({ respond: { anonymize: () => ({ ok: false, error: "python3 not found" }) } });
  await t.$("tools-anonymize").click();
  await tick();
  assert.equal(t.message(), "python3 not found");
  assert.deepEqual(texts(t.$("tools-anonymized")), []);
});

// ---- the page ----------------------------------------------------------------

test("the Word page has the Tools tab with every element the pane uses, and nothing of the parts that are not connected", async () => {
  const html = await readFile(new URL("../taskpane/word/index.html", import.meta.url), "utf8");
  const pane = await readFile(new URL("../taskpane/shared/tools-pane.js", import.meta.url), "utf8");
  assert.match(html, /<button class="tab" data-tab="tools">Tools<\/button>/);
  assert.match(html, /<section class="tab-content" data-tab="tools" hidden>/);
  const ids = new Set([...pane.matchAll(/(?:\$|onClick)\("(tools-[a-z-]+)"/g)].map((m) => m[1]));
  assert.ok(ids.size >= 15, `${ids.size} ids`);
  for (const id of [...ids, "tools-root"]) assert.ok(html.includes(`id="${id}"`), id);
  // Modules that exist and are not connected to the pane have no control here.
  const tab = html.slice(
    html.indexOf('data-tab="tools" hidden'),
    html.indexOf("<!-- Setup tab content -->"),
  );
  for (const word of [/deepl/i, /languagetool/i, /crossref/i, /\bcsl\b/i, /consent/i, /unit/i])
    assert.doesNotMatch(tab, word);
});

test("a page without the Tools tab (Excel) gets a pane that does nothing", async () => {
  const t = await setup({ missing: ["tools-root"] });
  assert.equal(t.requests.length, 0);
  t.pane.onEvent({ event: "job_done", job_id: "x", result: { ok: true } });
});

test("the head of an answer names what really ran, when the pane knows it", async () => {
  const t = await setup({
    settings: { toolsLang: "Korean" },
    word: { selectionText: async () => "Grown at 30 °C." },
    usedLabel: (role) => (role === "writing" ? "Claude · Opus 5.5" : null),
    respond: {
      selection_translate: (p) => ({ ok: true, text: "번역", lang: p.lang, model: "opus" }),
    },
  });
  await t.$("tools-translate").click();
  await tick();
  assert.equal(
    texts(t.$("tools-selection"))[0],
    'Translation into Korean (Claude · Opus 5.5) for: "Grown at 30 °C."',
  );
});

// The Tools tab said "Save the document to a local folder" before the daemon
// had answered which file the pane is on (an e2e run read that line of a
// saved document). Until the answer the line is empty; the hint is for a
// document the daemon says is not saved; another failure is an error, not
// that hint.
test("the file line: empty until the daemon answers, the save hint only for an unsaved document", async () => {
  globalThis.document = fakeDocument([]);
  const { initToolsPane } = await import("../taskpane/shared/tools-pane.js");
  let answer = null;
  const pane = initToolsPane({
    sendRequest: async () => answer,
    word: {},
  });
  const line = () => document.getElementById("tools-doc").textContent;
  assert.equal(line(), "", "before the first answer");
  answer = { ok: false, error: "Save the document to a local folder first." };
  await pane.reload();
  assert.match(line(), /^Save the document to a local folder to use the tools\.$/);
  answer = { ok: false, error: "timed out" };
  await pane.reload();
  assert.equal(line(), "", "an error says nothing about the file");
  assert.equal(document.getElementById("tools-message").textContent, "timed out");
  answer = { ok: true, doc_name: "paper.docx", profile_name: "generic", policy_packs: {} };
  await pane.reload();
  assert.match(line(), /^Saved file: paper\.docx\. Profile: generic\./);
});
