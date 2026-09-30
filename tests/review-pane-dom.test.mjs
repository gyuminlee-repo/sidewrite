// Behavior of the Review tab glue (review-pane.js) on a minimal fake DOM:
// just enough of document/element for initReviewPane to render and for
// click listeners to run. No Office.js: `word` and `sendRequest` are stubs.
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

function findAll(node, pred, out = []) {
  if (pred(node)) out.push(node);
  for (const k of node.children) findAll(k, pred, out);
  return out;
}

async function setup({ revision, word = {}, respond = {} }) {
  globalThis.document = fakeDocument();
  const { initReviewPane } = await import("../taskpane/shared/review-pane.js");
  const requests = [];
  const sendRequest = async (type, payload) => {
    requests.push({ type, payload });
    if (respond[type]) return respond[type](payload);
    return { ok: true, revision, path: "/m/paper.revision.json" };
  };
  const pane = initReviewPane({
    sendRequest,
    word,
    onTurnStarted() {},
    isTurnInFlight: () => false,
  });
  await pane.reload();
  return { pane, requests, $: (id) => document.getElementById(id) };
}

const comment = (id, content, extra = {}) => ({
  id,
  author: "A",
  content,
  resolved: false,
  anchor_text: "t",
  paragraph_index: 0,
  ...extra,
});

test("Delete ALL: a comment added after the first press is not deleted unseen", async () => {
  let doc = [comment("1", "Check units"), comment("2", "[R1.1] a / b / c")];
  const cleared = [];
  const backups = [];
  const word = {
    readAllComments: async () => structuredClone(doc),
    clearAllComments: async () => {
      cleared.push(doc.length);
      const n = doc.length;
      doc = [];
      return { cleared_comments: n };
    },
    listComments: async () => ({ comment_count: doc.length }),
  };
  const respond = {
    comments_backup: ({ backup }) => {
      backups.push(backup);
      return {
        ok: true,
        count: backup.comments.length,
        reply_count: backup.reply_count,
        path: "/m/paper.comments_backup.json",
      };
    },
  };
  const t = await setup({ revision: { comments: [], ledger: [] }, word, respond });
  const btn = t.$("review-delete-all");

  await btn.click();
  assert.match(btn.textContent, /Back up and delete 2 comments/);

  // Someone adds a comment with a reply between the two presses.
  doc.push(comment("3", "New remark", { replies: [{ author: "B", content: "agreed" }] }));
  await btn.click();
  assert.deepEqual(cleared, [], "nothing may be deleted when the list changed");
  assert.deepEqual(backups, []);
  assert.match(btn.textContent, /Back up and delete 3 comments/);
  assert.match(t.$("review-message").textContent, /1 added.*Nothing deleted/);

  await btn.click();
  assert.deepEqual(cleared, [3]);
  assert.equal(backups.length, 1);
  assert.deepEqual(
    backups[0].comments.map((c) => c.id),
    ["1", "2", "3"],
  );
  assert.equal(backups[0].reply_count, 1);
  assert.match(t.$("review-message").textContent, /Backed up 3, deleted 3 of 3, remaining 0/);
});

test("Delete ALL: a failed backup deletes nothing", async () => {
  const doc = [comment("1", "Check units")];
  const cleared = [];
  const word = {
    readAllComments: async () => structuredClone(doc),
    clearAllComments: async () => (cleared.push(1), { cleared_comments: 1 }),
    listComments: async () => ({ comment_count: 0 }),
  };
  const respond = {
    comments_backup: () => ({ ok: false, error: "disk full" }),
  };
  const t = await setup({ revision: { comments: [], ledger: [] }, word, respond });
  const btn = t.$("review-delete-all");
  await btn.click();
  await btn.click();
  assert.deepEqual(cleared, []);
  assert.match(t.$("review-message").textContent, /disk full/);
});
