// A supplementary file in the task pane, step A1, on a minimal fake DOM (as
// tests/review-pane-dom.test.mjs): the Review tab before and after joining
// the revision of a manuscript, the Response tab of a supplementary file, and
// the line of linked files in the manuscript's pane. No Office.js.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import {
  joinLabel,
  linkedFilesText,
  partOfText,
  PART_NO_TURNS,
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
    this.checked = false;
    this.disabled = false;
    this.title = "";
    this.classList = { toggle() {} };
  }
  addEventListener(ev, fn) {
    (this.listeners[ev] ??= []).push(fn);
  }
  // As the DOM does: what is no node becomes text, null included.
  append(...kids) {
    for (const k of kids) {
      if (k && typeof k === "object") this.children.push(k);
      else if (k !== "")
        this.children.push(Object.assign(new FakeEl("#text"), { textContent: String(k) }));
    }
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

function findAll(node, pred, out = []) {
  if (pred(node)) out.push(node);
  for (const k of node.children) findAll(k, pred, out);
  return out;
}
const buttons = (node, text) =>
  findAll(node, (n) => n.tagName === "BUTTON" && n.textContent === text);
const texts = (node) => findAll(node, (n) => n.textContent).map((n) => n.textContent);

const MAIN = { doc: "/m/paper.docx", name: "paper.docx", label: "SI" };
const comment = (ref, text) => ({
  ref,
  reviewer: 1,
  text,
  status: "pending",
  requirements: [{ id: "q2", text: "give the unit" }],
  changes: [],
  blocked: [],
  questions: [],
  requirements_outside: [],
});
const REVISION = {
  comments: [
    comment("R1.1", "State the incubation temperature."),
    comment("R1.2", "Add the control strain to Table S3."),
    comment("R2.1", "Do not call the method novel."),
  ],
  ledger: [
    {
      id: "L1",
      kind: "forbid",
      text: "do not write novel",
      scope: "manuscript",
      check: { type: "regex", pattern: "novel", flags: "i", expect: "absent" },
      status: "unchecked",
      violation_count: 0,
      violations: [],
      unsure: [],
    },
  ],
  external: [],
  snapshot: null,
};
const EMPTY = { comments: [], ledger: [], external: [], snapshot: null };
const FREE = { part_of: null, parts: [], note: null };
const CANDIDATES = [
  { doc: "/m/paper.docx", name: "paper.docx", comments: 23, updated_at: "t" },
  { doc: "/m/other.docx", name: "other.docx", comments: 1, updated_at: "t" },
];

// world: what the daemon holds for the document of this pane; the join and
// leave requests change it as the daemon would.
async function setup(world) {
  globalThis.document = fakeDocument();
  const { initReviewPane } = await import("../taskpane/shared/review-pane.js");
  const requests = [];
  const sendRequest = async (type, payload) => {
    requests.push({ type, payload });
    if (type === "review_link_candidates") return { ok: true, candidates: world.candidates ?? [] };
    if (type === "review_link_join") {
      if (world.refuse) return { ok: false, error: world.refuse };
      world.revision = REVISION;
      world.linked = { part_of: MAIN, parts: [], note: null };
    }
    if (type === "review_link_leave") {
      world.revision = EMPTY;
      world.linked = FREE;
    }
    return {
      ok: true,
      revision: world.revision,
      path: "/m/SI.revision.json",
      ...(world.linked === undefined ? {} : { linked: world.linked }),
    };
  };
  const pane = initReviewPane({
    sendRequest,
    word: {},
    onTurnStarted() {},
    isTurnInFlight: () => false,
  });
  await pane.reload();
  const $ = (id) => document.getElementById(id);
  return { pane, requests, $, sent: (type) => requests.filter((r) => r.type === type) };
}

test("the words of the link", () => {
  assert.equal(
    joinLabel({ name: "paper.docx", comments: 23 }),
    "Join the revision of paper.docx, 23 comments",
  );
  assert.equal(
    joinLabel({ name: "paper.docx", comments: 1 }),
    "Join the revision of paper.docx, 1 comment",
  );
  assert.equal(
    linkedFilesText([{ name: "SI.docx" }, { name: "S2.docx" }]),
    "Linked files: SI.docx, S2.docx",
  );
  assert.match(partOfText(MAIN), /edited in paper\.docx/);
  for (const s of [joinLabel(CANDIDATES[0]), partOfText(MAIN), PART_NO_TURNS]) {
    assert.equal(s.includes(String.fromCharCode(0x2014)), false, "no em dash");
  }
});

test("a file with no review: the manuscripts of the folder as one-click candidates and a path field; nothing joins by itself", async () => {
  const t = await setup({ revision: EMPTY, linked: FREE, candidates: CANDIDATES });
  const box = t.$("review-link-box");
  assert.deepEqual([box.hidden, box.open], [false, true]);
  assert.equal(box.attributes.class, "review-fold review-link");
  assert.deepEqual(
    t.$("review-link-candidates").children.map((b) => [b.tagName, b.textContent]),
    [
      ["BUTTON", "Join the revision of paper.docx, 23 comments"],
      ["BUTTON", "Join the revision of other.docx, 1 comment"],
    ],
  );
  assert.equal(t.$("review-linked").hidden, true);
  assert.equal(t.sent("review_link_join").length, 0, "a load joins nothing");
  // No id or path of a candidate in the view.
  assert.equal(
    texts(t.$("review-link-candidates")).some((s) => s.includes("/m/")),
    false,
  );
  // The typed path.
  t.$("review-link-path").value = "..\\paper\\paper.docx";
  await t.$("review-link-join").click();
  assert.deepEqual(t.sent("review_link_join")[0].payload, { main: "..\\paper\\paper.docx" });
  assert.equal(t.$("review-link-path").value, "");
});

test("no candidate in the folder: the box is there, closed, with the path field", async () => {
  const t = await setup({ revision: EMPTY, linked: FREE, candidates: [] });
  assert.deepEqual([t.$("review-link-box").hidden, t.$("review-link-box").open], [false, false]);
  assert.equal(t.$("review-link-candidates").children.length, 0);
  assert.equal(t.$("review-link-box").attributes.class, "review-fold review-link review-link-last");
});

test("a refused join says why and joins nothing", async () => {
  const refuse = "paper.docx has no saved reviewer comments.";
  const t = await setup({ revision: EMPTY, linked: FREE, candidates: CANDIDATES, refuse });
  await t.$("review-link-candidates").children[0].click();
  assert.equal(t.$("review-message").textContent, refuse);
  assert.equal(t.$("review-link-box").hidden, false);
  assert.equal(t.$("review-linked").hidden, true);
});

test("after the join: the rows of the manuscript, read-only, and no enabled Propose changes", async () => {
  const t = await setup({ revision: EMPTY, linked: FREE, candidates: CANDIDATES });
  await t.$("review-link-candidates").children[0].click();
  assert.deepEqual(t.sent("review_link_join")[0].payload, { main: "/m/paper.docx" });
  assert.equal(t.$("review-link-box").hidden, true);
  // The line that says where the list comes from, with the way out.
  const line = t.$("review-linked");
  assert.equal(line.hidden, false);
  assert.deepEqual(texts(line), [partOfText(MAIN), "Leave this revision"]);
  assert.equal(buttons(line, "Leave this revision").length, 1);
  // The same rows as the manuscript.
  const list = t.$("review-comments");
  assert.deepEqual(
    list.children.map((c) => c.attributes["data-ref"]),
    ["R1.1", "R1.2", "R2.1"],
  );
  assert.equal(findAll(list, (n) => n.attributes["data-req-id"] === "q2").length, 3);
  // Propose changes: in no card, the sentence in its place. A turn of the
  // pane that ends brings no button back.
  t.pane.setTurnBusy(false);
  assert.equal(buttons(list, "Propose changes").length, 0);
  assert.equal(texts(list).filter((s) => s === `${PART_NO_TURNS}.`).length, 3);
  assert.equal(t.$("review-propose-all").hidden, true);
  assert.equal(t.$("review-more").hidden, true);
  // What the manuscript owns cannot be changed here.
  assert.equal(buttons(list, "Edit").length, 0);
  assert.equal(buttons(list, "Save requirements").length, 0);
  assert.equal(buttons(list, "Preview prompt").length, 0);
  assert.equal(findAll(list, (n) => n.tagName === "TEXTAREA").length, 0);
  assert.equal(texts(list).filter((s) => s === "(edited in paper.docx)").length, 3);
  assert.equal(t.$("review-paste-section").hidden, true);
  assert.equal(t.$("review-start").hidden, true);
  // The rules: shown and checked here, edited in the manuscript.
  const ledger = t.$("review-ledger");
  assert.equal(ledger.children.length, 1);
  assert.deepEqual(
    [
      buttons(ledger, "Check").length,
      buttons(ledger, "Edit").length,
      buttons(ledger, "Remove").length,
    ],
    [1, 0, 0],
  );
  assert.deepEqual(
    [t.$("review-ledger-note").hidden, t.$("review-ledger-note").textContent],
    [false, "These rules are edited in paper.docx."],
  );
  for (const id of [
    "review-ledger-form",
    "review-ledger-advanced",
    "ledger-import",
    "ledger-suggest-review",
  ]) {
    assert.equal(t.$(id).hidden, true, id);
  }
  // The submission check is the manuscript's.
  assert.equal(t.$("review-submit-check").hidden, true);
  assert.equal(t.$("review-submit-help").textContent, "The submission check runs in paper.docx.");
});

test("Leave this revision: asked of the daemon, and the join box is back", async () => {
  const t = await setup({
    revision: REVISION,
    linked: { part_of: MAIN, parts: [], note: null },
    candidates: CANDIDATES,
  });
  assert.equal(t.sent("review_link_candidates").length, 0, "a linked file asks for no candidates");
  await buttons(t.$("review-linked"), "Leave this revision")[0].click();
  assert.equal(t.sent("review_link_leave").length, 1);
  assert.equal(t.$("review-linked").hidden, true);
  assert.equal(t.$("review-link-box").hidden, false);
  assert.equal(t.$("review-comments").children.length, 0);
  assert.equal(t.$("review-paste-section").hidden, false);
});

test("a note of the daemon about the manuscript is shown in the line", async () => {
  const note = "The review file of paper.docx was not found, so this list may be out of date.";
  const t = await setup({ revision: REVISION, linked: { part_of: MAIN, parts: [], note } });
  assert.ok(texts(t.$("review-linked")).includes(note));
});

test("the manuscript's pane: one line names the linked files, and everything else is as before", async () => {
  const parts = [{ doc: "/m/SI.docx", name: "SI.docx", label: "SI" }];
  const t = await setup({ revision: REVISION, linked: { part_of: null, parts, note: null } });
  assert.equal(t.$("review-linked").hidden, false);
  assert.deepEqual(texts(t.$("review-linked")), ["Linked files: SI.docx"]);
  assert.equal(t.$("review-link-box").hidden, true);
  const list = t.$("review-comments");
  assert.deepEqual(
    buttons(list, "Propose changes").map((b) => b.disabled),
    [false, false, false],
  );
  assert.equal(buttons(list, "Edit").length, 3);
  assert.equal(texts(list).includes(`${PART_NO_TURNS}.`), false);
  assert.equal(buttons(t.$("review-ledger"), "Remove").length, 1);
  assert.equal(t.$("review-ledger-note").hidden, true);
  assert.equal(t.$("review-ledger-form").hidden, false);
  assert.equal(t.$("review-paste-section").hidden, false);
  assert.equal(t.$("review-submit-check").hidden, false);
});

test("a manuscript with a review and no link shows neither the line nor the box", async () => {
  const t = await setup({ revision: REVISION, linked: FREE, candidates: CANDIDATES });
  assert.deepEqual([t.$("review-linked").hidden, t.$("review-link-box").hidden], [true, true]);
  assert.equal(t.sent("review_link_candidates").length, 0);
});

test("a daemon from before links: no box, and no request it does not know", async () => {
  const t = await setup({ revision: EMPTY, linked: undefined });
  assert.equal(t.$("review-link-box").hidden, true);
  assert.equal(t.$("review-linked").hidden, true);
  assert.equal(t.sent("review_link_candidates").length, 0);
});

test("linked_changed: the pane loads its file again", async () => {
  const world = { revision: REVISION, linked: FREE };
  const t = await setup(world);
  const loads = t.sent("review_load").length;
  world.linked = {
    part_of: null,
    parts: [{ doc: "/m/SI.docx", name: "SI.docx", label: "SI" }],
    note: null,
  };
  t.pane.onEvent({ type: "review_event", event: "linked_changed", doc: "SI.docx" });
  await new Promise((r) => setTimeout(r, 5));
  assert.equal(t.sent("review_load").length, loads + 1);
  assert.deepEqual(texts(t.$("review-linked")), ["Linked files: SI.docx"]);
});

// ---- the Response tab ---------------------------------------------------------

async function responseSetup(linked) {
  globalThis.document = fakeDocument();
  const { initResponsePane } = await import("../taskpane/shared/response-pane.js");
  const requests = [];
  const world = { linked };
  const sendRequest = async (type) => {
    requests.push(type);
    if (type === "profile_get")
      return {
        ok: true,
        name: "generic",
        path: null,
        profile: null,
        overrides: {},
        available: [],
        problems: [],
      };
    return { ok: true, revision: REVISION, linked: world.linked };
  };
  const pane = initResponsePane({ sendRequest, word: {} });
  await pane.reload();
  return { pane, requests, world, $: (id) => document.getElementById(id) };
}

test("the Response tab of a supplementary file is one sentence", async () => {
  const t = await responseSetup({ part_of: MAIN, parts: [], note: null });
  assert.deepEqual(
    [t.$("response-part-note").hidden, t.$("response-part-note").textContent],
    [false, "The letter is written from paper.docx."],
  );
  assert.equal(t.$("response-main").hidden, true);
  // The link goes: the tab is back after the event.
  t.world.linked = FREE;
  t.pane.onEvent({ type: "review_event", event: "linked_changed", doc: "SI.docx" });
  await new Promise((r) => setTimeout(r, 5));
  assert.deepEqual([t.$("response-part-note").hidden, t.$("response-main").hidden], [true, false]);
});

test("the Response tab of a manuscript, linked or not, is the whole tab", async () => {
  for (const linked of [
    FREE,
    { part_of: null, parts: [{ name: "SI.docx" }], note: null },
    undefined,
  ]) {
    const t = await responseSetup(linked);
    assert.deepEqual(
      [t.$("response-part-note").hidden, t.$("response-main").hidden],
      [true, false],
    );
    assert.equal(t.$("response-part-note").textContent, "");
  }
});

// ---- the page -----------------------------------------------------------------

test("the page holds the elements of the link, and the Response tab is inside its wrapper", async () => {
  const html = await readFile(new URL("../taskpane/word/index.html", import.meta.url), "utf8");
  for (const id of [
    "review-linked",
    "review-link-box",
    "review-link-candidates",
    "review-link-path",
    "review-link-join",
    "review-ledger-note",
    "review-ledger-form",
    "review-submit-help",
    "response-part-note",
    "response-main",
  ]) {
    assert.equal(html.split(`id="${id}"`).length, 2, `${id} once in the page`);
  }
  assert.match(html, /<summary>This file belongs to another manuscript's revision<\/summary>/);
  const at = (s) => html.indexOf(s);
  const end = at('id="tools-root"');
  assert.ok(at('id="response-part-note"') < at('id="response-main"'));
  assert.ok(at('id="review-link-box"') < at('id="review-start"'), "the join box comes first");
  for (const id of [
    "response-next",
    "response-comments",
    "response-export",
    "response-profile-details",
  ]) {
    assert.ok(at('id="response-main"') < at(`id="${id}"`) && at(`id="${id}"`) < end, id);
  }
  // The dispatcher passes the event to both tabs.
  const pane = await readFile(new URL("../taskpane/shared/taskpane.js", import.meta.url), "utf8");
  assert.match(pane, /if \(msg\.event === "linked_changed"\) responsePane\?\.onEvent\(msg\);/);
});
