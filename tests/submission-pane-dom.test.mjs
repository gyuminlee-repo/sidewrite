// The submission checklist in the Review tab (review-pane.js, go-to.js) on a
// minimal fake DOM, and the markup of its section (word/index.html). No
// Office.js: `word`, `sendRequest` and the job of the Tools tab are stubs.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { createGoTo } from "../taskpane/shared/go-to.js";
import {
  submissionCount,
  submissionGoLabel,
  SUBMISSION_STATE_LABELS,
} from "../taskpane/shared/review-view.js";

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
    this.title = "";
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
  async click() {
    for (const fn of this.listeners.click ?? []) await fn();
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
const findAll = (node, pred, out = []) => {
  if (pred(node)) out.push(node);
  for (const k of node.children) findAll(k, pred, out);
  return out;
};

const comment = (ref, text, status) => ({
  ref,
  reviewer: 1,
  text,
  section: "",
  requirements: [],
  status,
  changes: [],
  blocked: [],
  questions: [],
});
const REVISION = {
  comments: [
    comment("R1.1", "Please state the unit.", "done"),
    comment("R1.2", "Name the antifoam.", "in_progress"),
  ],
  ledger: [],
  external: [],
  snapshot: { paragraphs: [], ids: [] },
};
const ROWS = [
  {
    id: "answers",
    file: "paper.docx",
    state: "ready",
    text: "Every comment has an answer.",
    go: null,
  },
  {
    id: "requests",
    file: "paper.docx",
    state: "not_ready",
    text: "1 request has no accepted change: R1.2.",
    go: { tab: "review", target: "review-comments" },
  },
  {
    id: "changes_decided",
    file: "paper.docx",
    state: "not_ready",
    text: "1 paragraph still holds tracked changes. Accept or reject them in Word.",
    go: { paragraph: 4 },
  },
  {
    id: "format_counts",
    file: "paper.docx",
    state: "not_applicable",
    text: "No format rule and no count rule is set for this paper.",
    go: null,
  },
  {
    id: "letter",
    file: "paper.docx",
    state: "not_ready",
    text: "The response letter has not been exported yet.",
    go: { tab: "response", target: "response-export" },
  },
];

async function setup({
  result = { ok: true, rows: ROWS, doc_name: "paper.docx" },
  word = {},
} = {}) {
  globalThis.document = fakeDocument();
  const { initReviewPane } = await import("../taskpane/shared/review-pane.js");
  const jobs = [];
  const went = [];
  const selected = [];
  const pane = initReviewPane({
    sendRequest: async () => ({ ok: true, revision: REVISION, path: "/m/paper.revision.json" }),
    word: {
      collectReviewSnapshot: async () => ({ paragraphs: [{ index: 0, text: "T" }], comments: [] }),
      collectHeadings: async () => [{ index: 0, level: 1, text: "T" }],
      readFormatSnapshot: async () => ({ paragraphs: [{ index: 0, runs: [] }] }),
      selectParagraphByIndex: async (i) => selected.push(i),
      ...word,
    },
    onTurnStarted() {},
    isTurnInFlight: () => false,
    goTo: (tab, target) => went.push([tab, target]),
    toolsJob: async (type, payload) => {
      jobs.push({ type, payload });
      return typeof result === "function" ? result() : result;
    },
  });
  await pane.reload();
  const $ = (id) => document.getElementById(id);
  const rows = () =>
    findAll($("review-submit-rows"), (n) => n.attributes["data-row"]).map((n) => ({
      id: n.attributes["data-row"],
      cls: n.attributes.class,
      texts: n.children.map((k) => k.textContent),
      button: n.children.find((k) => k.tagName === "BUTTON") ?? null,
    }));
  return { pane, jobs, went, selected, $, rows };
}

test("before a check the header counts the comments that are not done and no row is shown", async () => {
  const t = await setup();
  assert.equal(t.$("review-submit-count").textContent, "(1 not done)");
  assert.deepEqual(t.rows(), []);
});

test("Check now sends the document as read and renders one row per answer, with state and sentence", async () => {
  const t = await setup();
  await t.$("review-submit-check").click();
  assert.equal(t.jobs.length, 1);
  assert.equal(t.jobs[0].type, "submission_check");
  assert.deepEqual(Object.keys(t.jobs[0].payload).sort(), ["format", "headings", "snapshot"]);
  assert.equal(t.jobs[0].payload.headings.length, 1);
  const rows = t.rows();
  assert.deepEqual(
    rows.map((r) => r.id),
    ROWS.map((r) => r.id),
  );
  assert.deepEqual(rows[0].texts, ["Ready", "Every comment has an answer."]);
  assert.deepEqual(rows[1].texts.slice(0, 2), [
    "Not ready",
    "1 request has no accepted change: R1.2.",
  ]);
  assert.equal(rows[3].texts[0], "Not needed");
  assert.match(rows[1].cls, /review-submit-not_ready/);
  // A row that is ready or not needed has no button.
  assert.equal(rows[0].button, null);
  assert.equal(rows[3].button, null);
  assert.equal(t.$("review-message").textContent, "");
});

test("after a check the header counts the rows that are not ready, and keeps doing so across renders", async () => {
  const t = await setup();
  await t.$("review-submit-check").click();
  assert.equal(t.$("review-submit-count").textContent, "(3 not ready)");
  // Another action of the tab renders again: the count stays the checklist's.
  t.pane.onEvent({ event: "figures_changed" });
  await new Promise((r) => setTimeout(r, 0));
  assert.equal(t.$("review-submit-count").textContent, "(3 not ready)");
  assert.equal(t.rows().length, ROWS.length);
  // A new connection drops the result: it was of the document as it was then.
  await t.pane.reload();
  assert.equal(t.$("review-submit-count").textContent, "(1 not done)");
  assert.deepEqual(t.rows(), []);
});

test("the button of a row goes to where it is fixed: a tab and an element, or a paragraph in Word", async () => {
  const t = await setup();
  await t.$("review-submit-check").click();
  const [, requests, changes, , letter] = t.rows();
  assert.equal(requests.button.textContent, "Go there");
  await requests.button.click();
  assert.equal(letter.button.textContent, "Go to the Response tab");
  await letter.button.click();
  assert.deepEqual(t.went, [
    ["review", "review-comments"],
    ["response", "response-export"],
  ]);
  assert.equal(changes.button.textContent, "Show in Word");
  await changes.button.click();
  assert.deepEqual(t.selected, [4]);
});

test("a check that fails says why and leaves the rows of the check before it", async () => {
  let n = 0;
  const t = await setup({
    result: () =>
      n++
        ? { ok: false, error: "Save the document to a local folder first." }
        : { ok: true, rows: ROWS },
  });
  await t.$("review-submit-check").click();
  await t.$("review-submit-check").click();
  assert.equal(t.$("review-message").textContent, "Save the document to a local folder first.");
  assert.equal(t.rows().length, ROWS.length);
});

test("formatting Word cannot give is sent as null: the other rows are still asked for", async () => {
  const t = await setup({
    word: {
      readFormatSnapshot: async () => {
        throw new Error("GeneralException");
      },
    },
  });
  await t.$("review-submit-check").click();
  assert.equal(t.jobs[0].payload.format, null);
  assert.equal(t.rows().length, ROWS.length);
});

test("view helpers: the header count, the state labels and the button labels", () => {
  assert.equal(submissionCount(ROWS), "(3 not ready)");
  assert.equal(submissionCount([]), "(0 not ready)");
  assert.deepEqual(Object.keys(SUBMISSION_STATE_LABELS), ["ready", "not_ready", "not_applicable"]);
  assert.equal(submissionGoLabel(null), "");
  assert.equal(submissionGoLabel({ paragraph: 0 }), "Show in Word");
  assert.equal(submissionGoLabel({ tab: "tools", target: "tools-check" }), "Go to the Tools tab");
  assert.equal(submissionGoLabel({ tab: "review", target: "review-clean-tool" }), "Go there");
});

// ---- go-to.js: the tab is switched and the target gets the focus -----------

function page() {
  const calls = [];
  const make = (id, { tabIndex = 0, parent = null, details = false } = {}) => ({
    id,
    tabIndex,
    open: false,
    parentElement: parent,
    attrs: {},
    isDetails: details,
    setAttribute(k, v) {
      this.attrs[k] = v;
    },
    closest(sel) {
      assert.equal(sel, "details");
      for (let e = this; e; e = e.parentElement) if (e.isDetails) return e;
      return null;
    },
    scrollIntoView: () => calls.push(["scroll", id]),
    focus: () => calls.push(["focus", id]),
  });
  const outer = make("review-more", { details: true });
  const fold = make("review-submit-box", { details: true, parent: outer });
  const els = {
    "review-clean-tool": make("review-clean-tool", { parent: fold }),
    "review-comments": make("review-comments", { tabIndex: -1 }),
  };
  return { calls, outer, fold, els, doc: { getElementById: (id) => els[id] ?? null } };
}

test("goTo switches the tab, opens the folds around the target, scrolls to it and focuses it", () => {
  const p = page();
  const tabs = [];
  const goTo = createGoTo((t) => tabs.push(t), p.doc);
  assert.equal(goTo("review", "review-clean-tool"), true);
  assert.deepEqual(tabs, ["review"]);
  assert.equal(p.fold.open, true);
  assert.equal(p.outer.open, true);
  assert.deepEqual(p.calls, [
    ["scroll", "review-clean-tool"],
    ["focus", "review-clean-tool"],
  ]);
  assert.deepEqual(p.els["review-clean-tool"].attrs, {}, "a button takes the focus as it is");
});

test("goTo makes a target that takes no focus take it, and still shows the tab of an unknown id", () => {
  const p = page();
  const tabs = [];
  const goTo = createGoTo((t) => tabs.push(t), p.doc);
  goTo("review", "review-comments");
  assert.deepEqual(p.els["review-comments"].attrs, { tabindex: "-1" });
  assert.deepEqual(p.calls.at(-1), ["focus", "review-comments"]);
  assert.equal(goTo("tools", "no-such-id"), false);
  assert.deepEqual(tabs, ["review", "tools"]);
});

// ---- the markup: the targets exist and the cleanup controls stay reachable --

const html = await readFile(new URL("../taskpane/word/index.html", import.meta.url), "utf8");
const BOX = html.slice(
  html.indexOf('id="review-submit-box"'),
  html.indexOf("</details>", html.indexOf('id="review-submit-box"')),
);

test("Check now stands above the cleanup buttons in the Before submission section", () => {
  const at = (s) => BOX.indexOf(s);
  assert.ok(at('id="review-submit-check"') > 0);
  assert.ok(at('id="review-submit-check"') < at('id="review-submit-rows"'));
  assert.ok(at('id="review-submit-rows"') < at('id="review-clean-tool"'));
  assert.ok(at('id="review-clean-tool"') < at('id="review-delete-all"'));
  assert.match(BOX, />\s*Check now\s*</);
});

test("the cleanup controls keep their ids and are reached as the acceptance harness reaches them", async () => {
  // pane-driver.cjs opens the fold by pressing its summary, then presses the
  // button by its id: the section is one <details> with a <summary>, the
  // buttons are in it directly (no fold inside the fold) and not hidden.
  const driver = await readFile(new URL("../scripts/e2e/pane-driver.cjs", import.meta.url), "utf8");
  assert.match(driver, /openFold\(page, "review-submit-box"\)/);
  assert.match(driver, /page\.click\("#review-clean-tool"\)/);
  assert.match(driver, /const btn = "#review-delete-all"/);
  assert.match(html, /<details class="review-fold" id="review-submit-box" hidden>\s*<summary>/);
  assert.equal(BOX.includes("<details"), false, "no fold inside the section");
  for (const id of ["review-clean-tool", "review-delete-all", "review-submit-count"]) {
    const tag = BOX.slice(
      BOX.lastIndexOf("<", BOX.indexOf(`id="${id}"`)),
      BOX.indexOf(">", BOX.indexOf(`id="${id}"`)),
    );
    assert.ok(tag.includes(`id="${id}"`), id);
    assert.equal(/\bhidden\b/.test(tag), false, `${id} is not hidden`);
  }
});

test("every place a row leads to is an element of the pane", async () => {
  const source = await readFile(new URL("../daemon/submission-check.mjs", import.meta.url), "utf8");
  const targets = [...source.matchAll(/target: "([a-z-]+)"/g)].map((m) => m[1]);
  assert.ok(targets.length >= 5);
  for (const id of targets) assert.ok(html.includes(`id="${id}"`), `${id} is not in the pane`);
  const tabs = [...source.matchAll(/tab: "([a-z]+)"/g)].map((m) => m[1]);
  for (const tab of tabs)
    assert.ok(html.includes(`class="tab-content" data-tab="${tab}"`), `${tab} is not a tab`);
});
