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
  const sendRequest = async (type, payload, opts) => {
    requests.push({ type, payload, opts });
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
  const s = fakeCommentStore([comment("1", "Check units"), comment("2", "[R1.1] a / b / c")]);
  const backups = [];
  const respond = backupOk((backup) => backups.push(backup));
  const t = await setup({ revision: { comments: [], ledger: [] }, word: s.word, respond });
  const btn = t.$("review-delete-all");

  await btn.click();
  assert.match(btn.textContent, /Back up and delete 2 comments/);

  // Someone adds a comment with a reply between the two presses.
  s.doc.push(comment("3", "New remark", { replies: [{ author: "B", content: "agreed" }] }));
  await btn.click();
  assert.deepEqual(s.deleted, [], "nothing may be deleted when the list changed");
  assert.deepEqual(backups, []);
  assert.match(btn.textContent, /Back up and delete 3 comments/);
  assert.match(t.$("review-message").textContent, /1 added.*Nothing deleted/);

  await btn.click();
  assert.deepEqual(s.deleted, ["1", "2", "3"]);
  assert.equal(backups.length, 1);
  assert.deepEqual(
    backups[0].comments.map((c) => c.id),
    ["1", "2", "3"],
  );
  assert.deepEqual(s.passed, ["1", "2", "3"]);
  assert.equal(backups[0].reply_count, 1);
  assert.match(t.$("review-message").textContent, /Backed up 3, deleted 3 of 3/);
  assert.doesNotMatch(t.$("review-message").textContent, /MISMATCH/);
});

// A fake Word comment store. clearAllComments deletes everything (the old
// path); deleteBackedUpComments deletes only unchanged backed-up threads,
// with the same decision helper the Office function uses.
function fakeCommentStore(initial) {
  const store = { doc: initial, deleted: [], passed: null };
  store.word = {
    readAllComments: async () => structuredClone(store.doc),
    clearAllComments: async () => {
      store.deleted.push(...store.doc.map((c) => c.id));
      const n = store.doc.length;
      store.doc = [];
      return { cleared_comments: n };
    },
    deleteBackedUpComments: async (rows) => {
      const { planBackupDelete } = await import("../taskpane/shared/review-view.js");
      store.passed = rows.map((r) => r.id);
      const plan = planBackupDelete(rows, store.doc);
      store.deleted.push(...plan.delete_ids);
      store.doc = store.doc.filter((c) => !plan.delete_ids.includes(c.id));
      return plan;
    },
    listComments: async () => ({
      comment_count: store.doc.length,
      comments: store.doc.map((c) => ({ id: c.id, text: c.content })),
    }),
  };
  return store;
}

const backupOk = (onBackup) => ({
  comments_backup: ({ backup }) => {
    onBackup?.(backup);
    return {
      ok: true,
      count: backup.comments.length,
      reply_count: backup.reply_count,
      path: "/m/paper.comments_backup.json",
    };
  },
});

test("Delete ALL: a comment added while the backup is written is not deleted", async () => {
  const s = fakeCommentStore([comment("1", "Check units")]);
  const respond = backupOk(() => s.doc.push(comment("2", "Added during the backup")));
  const t = await setup({ revision: { comments: [], ledger: [] }, word: s.word, respond });
  const btn = t.$("review-delete-all");
  await btn.click();
  await btn.click();
  assert.deepEqual(s.deleted, ["1"]);
  assert.deepEqual(
    s.doc.map((c) => c.id),
    ["2"],
  );
  const msg = t.$("review-message").textContent;
  assert.match(msg, /Backed up 1, deleted 1 of 1/);
  assert.match(msg, /1 comment\(s\) added after the backup were not deleted/);
  assert.doesNotMatch(msg, /MISMATCH/);
});

test("Delete ALL: a thread changed after the backup is kept and asks for a new press", async () => {
  const s = fakeCommentStore([comment("1", "Check units"), comment("2", "Cite the method")]);
  const respond = backupOk(() => {
    s.doc[1] = { ...s.doc[1], replies: [{ author: "B", content: "Will do" }] };
  });
  const t = await setup({ revision: { comments: [], ledger: [] }, word: s.word, respond });
  const btn = t.$("review-delete-all");
  await btn.click();
  await btn.click();
  assert.deepEqual(s.deleted, ["1"]);
  assert.deepEqual(
    s.doc.map((c) => c.id),
    ["2"],
  );
  assert.match(
    t.$("review-message").textContent,
    /1 changed after the backup and not deleted.*press again/,
  );
  // The next press shows the count again instead of deleting.
  await btn.click();
  assert.deepEqual(s.deleted, ["1"]);
  assert.match(btn.textContent, /Back up and delete 1 comments/);
});

test("Delete ALL: a failed backup deletes nothing", async () => {
  const doc = [comment("1", "Check units")];
  const cleared = [];
  const word = {
    readAllComments: async () => structuredClone(doc),
    clearAllComments: async () => (cleared.push(1), { cleared_comments: 1 }),
    deleteBackedUpComments: async () => (cleared.push(1), { delete_ids: ["1"] }),
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

test("Allow exception: the second press sends one review_override", async () => {
  const blocked = {
    id: "R1.1-b1",
    args: { find: "held at 30 degrees", replace: "as suggested by the reviewer, held at 30 °C" },
    blocked_by: ["L2"],
    created_at: "2026-01-01T00:00:00Z",
  };
  const revision = {
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "State the unit.",
        requirements: [],
        status: "in_progress",
        changes: [],
        blocked: [blocked],
      },
    ],
    ledger: [],
  };
  const respond = { review_override: () => ({ ok: true, applied: true, revision }) };
  const t = await setup({ revision, respond });
  const allow = () =>
    findAll(t.$("review-comments"), (n) => /Allow exception|Apply despite/.test(n.textContent));

  assert.equal(allow().length, 1);
  await allow()[0].click();
  assert.equal(t.requests.filter((r) => r.type === "review_override").length, 0);
  // render() rebuilt the row; the rebuilt button must still be armed.
  assert.match(allow()[0].textContent, /Apply despite L2\?/);
  await allow()[0].click();
  const overrides = t.requests.filter((r) => r.type === "review_override");
  assert.equal(overrides.length, 1);
  assert.deepEqual(overrides[0].payload, { ref: "R1.1", blocked_id: "R1.1-b1" });
});

test("outcome unknown: Allow exception is disabled; re-apply needs two presses and sends force", async () => {
  const blocked = {
    id: "R1.1-b1",
    args: { find: "held at 30 degrees", replace: "as suggested by the reviewer, held at 30 °C" },
    blocked_by: ["L2"],
    created_at: "2026-01-01T00:00:00Z",
    state: "outcome_unknown",
    attempt: null,
  };
  const revision = {
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "State the unit.",
        requirements: [],
        status: "in_progress",
        changes: [],
        blocked: [blocked],
      },
    ],
    ledger: [],
  };
  const respond = { review_override: () => ({ ok: true, applied: true, revision }) };
  const t = await setup({ revision, respond });
  const find = (re) => findAll(t.$("review-comments"), (n) => re.test(n.textContent));
  const overrides = () => t.requests.filter((r) => r.type === "review_override");

  const allow = find(/^Allow exception$/);
  assert.equal(allow.length, 1);
  assert.equal(allow[0].disabled, true);
  assert.equal(find(/Refresh status/).length, 1);

  await find(/Re-apply anyway/)[0].click();
  assert.equal(overrides().length, 0);
  await find(/Re-apply anyway/)[0].click();
  assert.equal(overrides().length, 1);
  assert.deepEqual(overrides()[0].payload, { ref: "R1.1", blocked_id: "R1.1-b1", force: true });
});

test("override: waits past the tool timeout and keeps the row disabled while it runs", async () => {
  const blocked = {
    id: "R1.1-b1",
    args: { find: "held at 30 degrees", replace: "as suggested by the reviewer, held at 30 °C" },
    blocked_by: ["L2"],
    created_at: "2026-01-01T00:00:00Z",
  };
  const revision = {
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "State the unit.",
        requirements: [],
        status: "in_progress",
        changes: [],
        blocked: [blocked],
      },
    ],
    ledger: [],
  };
  let release;
  const gate = new Promise((r) => (release = r));
  const respond = {
    review_override: async () => {
      await gate;
      return { ok: true, applied: true, revision: { ...revision, comments: [] } };
    },
  };
  const t = await setup({ revision, respond });
  const find = (re) => findAll(t.$("review-comments"), (n) => re.test(n.textContent));
  await find(/Allow exception/)[0].click();
  const running = find(/Apply despite/)[0].click();
  await new Promise((r) => setTimeout(r, 0));
  const busy = find(/^Applying\.\.\.$/);
  assert.equal(busy.length, 1);
  assert.equal(busy[0].disabled, true);
  assert.equal(find(/Allow exception|Apply despite/).length, 0);
  release();
  await running;
  const sent = t.requests.filter((r) => r.type === "review_override");
  assert.equal(sent.length, 1);
  assert.ok(sent[0].opts.timeoutMs >= 60_000);
  assert.equal(find(/^Applying/).length, 0);
});

test("a blocked edit the daemon reports applying shows no enabled button", async () => {
  const revision = {
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "State the unit.",
        requirements: [],
        status: "in_progress",
        changes: [],
        blocked: [
          {
            id: "R1.1-b1",
            args: { find: "x", replace: "y" },
            blocked_by: ["L2"],
            created_at: "2026-01-01T00:00:00Z",
            state: "applying",
            attempt: null,
          },
        ],
      },
    ],
    ledger: [],
  };
  const t = await setup({ revision });
  const buttons = findAll(t.$("review-comments"), (n) =>
    /^(Allow exception|Re-apply anyway|Applying\.\.\.)$/.test(n.textContent),
  );
  assert.equal(buttons.length, 1);
  assert.equal(buttons[0].disabled, true);
});

// A chat turn (not a Review turn) that re-renders the Review tab disables
// every "Propose changes" button; the end of that turn must enable them
// again, whether it ended normally or on an error. No review turn_done comes.
test("Propose changes follows the running turn: chat re-render disables, chat end enables", async () => {
  const { createTurnState } = await import("../taskpane/shared/turn-state.js");
  globalThis.document = fakeDocument();
  const { initReviewPane } = await import("../taskpane/shared/review-pane.js");
  const revision = {
    comments: ["R1.1", "R1.2"].map((ref) => ({
      ref,
      reviewer: 1,
      text: "State the unit.",
      requirements: [],
      status: "pending",
      changes: [],
      blocked: [],
    })),
    ledger: [],
  };
  let pane = null;
  const turn = createTurnState({ getReviewPane: () => pane });
  pane = initReviewPane({
    sendRequest: async () => ({ ok: true, revision, prompt: "p" }),
    word: {},
    onTurnStarted() {},
    isTurnInFlight: () => turn.inFlight,
  });
  await pane.reload();
  const propose = () =>
    findAll(document.getElementById("review-comments"), (n) => n.textContent === "Propose changes");
  const states = () => propose().map((b) => b.disabled);
  assert.deepEqual(states(), [false, false]);

  for (const end of ["turn_complete", "error", "auth_error"]) {
    turn.begin(); // a chat message was sent
    const preview = findAll(
      document.getElementById("review-comments"),
      (n) => n.textContent === "Preview prompt",
    );
    await preview[0].click(); // re-renders the Review tab mid-turn
    assert.deepEqual(states(), [true, true], `${end}: disabled while the chat runs`);
    turn.onAssistantEvent({ event: end });
    assert.deepEqual(states(), [false, false], `${end}: enabled once the turn ends`);
  }
  turn.begin();
  assert.deepEqual(states(), [true, true], "a new turn disables without a re-render");
  turn.end(); // connection dropped (ws.onclose)
  assert.deepEqual(states(), [false, false]);
});
