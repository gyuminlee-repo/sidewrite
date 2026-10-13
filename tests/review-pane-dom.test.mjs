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
    listeners: {},
    addEventListener(ev, fn) {
      (this.listeners[ev] ??= []).push(fn);
    },
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

async function setup({ revision, word = {}, respond = {}, providersReady, options = {} }) {
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
    providersReady,
    ...options,
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
  // The daemon is told which comments are gone (rationale comments among them).
  assert.deepEqual(
    t.requests.filter((r) => r.type === "review_comments_cleaned").map((r) => r.payload.ids),
    [["1", "2", "3"]],
  );
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

test("Save shows the daemon's warning about comments it kept", async () => {
  const revision = { comments: [], ledger: [] };
  const warning = "Not removed because work on it is in progress: R1.1";
  let t = null;
  t = await setup({
    revision,
    respond: {
      review_split: () => {
        t.pane.onEvent({
          type: "review_event",
          event: "job_done",
          job_id: "j1",
          result: {
            ok: true,
            method: "regex",
            warnings: [],
            comments: [{ ref: "R1.3", reviewer: 1, text: "Rename it." }],
          },
        });
        return { ok: true, accepted: true, job_id: "j1" };
      },
      review_save_comments: () => ({ ok: true, revision, warnings: [warning] }),
    },
  });
  await t.$("review-split").click();
  await t.$("review-save").click();
  assert.equal(t.$("review-message").textContent, warning);
  assert.equal(t.$("review-message").hidden, false);
});

test("Save sends the section label the split gave a comment, and an empty one where it gave none", async () => {
  const revision = { comments: [], ledger: [] };
  let t = null;
  t = await setup({
    revision,
    respond: {
      review_split: () => {
        t.pane.onEvent({
          type: "review_event",
          event: "job_done",
          job_id: "j1",
          result: {
            ok: true,
            method: "regex",
            warnings: [],
            comments: [
              { ref: "R1.1", reviewer: 1, text: "Add a control.", section: "Major comments" },
              { ref: "R1.2", reviewer: 1, text: "Add units." },
            ],
          },
        });
        return { ok: true, accepted: true, job_id: "j1" };
      },
      review_save_comments: () => ({ ok: true, revision, warnings: [] }),
    },
  });
  await t.$("review-split").click();
  await t.$("review-save").click();
  const sent = t.requests.find((r) => r.type === "review_save_comments").payload.comments;
  assert.deepEqual(
    sent.map((c) => [c.ref, c.section, "section" in c]),
    [
      ["R1.1", "Major comments", true],
      ["R1.2", "", true],
    ],
  );
});

test("a comment card shows its section as a label; a comment without one shows none", async () => {
  const revision = {
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "Shorten this section.",
        section: "Abstract",
        requirements: [],
        status: "pending",
        changes: [],
        blocked: [],
      },
      {
        ref: "R1.2",
        reviewer: 1,
        text: "Add units.",
        requirements: [],
        status: "pending",
        changes: [],
        blocked: [],
      },
    ],
    ledger: [],
  };
  const t = await setup({ revision });
  const labels = findAll(
    t.$("review-comments"),
    (n) => n.attributes.class === "review-comment-section",
  );
  assert.deepEqual(
    labels.map((n) => n.textContent),
    ["Abstract"],
  );
  // The cards stay in ref order, the label sits in the card of its comment.
  const cards = findAll(t.$("review-comments"), (n) => n.attributes.class === "review-comment");
  assert.equal(cards.length, 2);
  assert.equal(findAll(cards[0], (n) => n.attributes.class === "review-comment-section").length, 1);
  assert.equal(findAll(cards[1], (n) => n.attributes.class === "review-comment-section").length, 0);
});

test("the draft row has a section field: Save sends an edited section and clears an emptied one", async () => {
  const revision = { comments: [], ledger: [] };
  let t = null;
  t = await setup({
    revision,
    respond: {
      review_split: () => {
        t.pane.onEvent({
          type: "review_event",
          event: "job_done",
          job_id: "j1",
          result: {
            ok: true,
            method: "regex",
            warnings: [],
            comments: [
              { ref: "R1.1", reviewer: 1, text: "Shorten this section.", section: "Abstract" },
              { ref: "R1.2", reviewer: 1, text: "Expand this section.", section: "Methods" },
              { ref: "R1.3", reviewer: 1, text: "Add units." },
              { ref: "R1.4", reviewer: 1, text: "State n." },
            ],
          },
        });
        return { ok: true, accepted: true, job_id: "j1" };
      },
      review_save_comments: () => ({ ok: true, revision, warnings: [] }),
    },
  });
  await t.$("review-split").click();
  const fields = findAll(t.$("review-draft"), (n) => n.attributes.class === "review-section");
  assert.deepEqual(
    fields.map((n) => n.value),
    ["Abstract", "Methods", "", ""],
  );
  const type = (node, value) => {
    node.value = value;
    for (const fn of node.listeners.input ?? []) fn();
  };
  type(fields[0], "Abstract and title ");
  type(fields[1], "  ");
  type(fields[2], "Minor comments");
  await t.$("review-save").click();
  const sent = t.requests.find((r) => r.type === "review_save_comments").payload.comments;
  assert.deepEqual(
    sent.map((c) => [c.ref, c.section, "section" in c]),
    [
      ["R1.1", "Abstract and title", true],
      ["R1.2", "", true],
      ["R1.3", "Minor comments", true],
      ["R1.4", "", true],
    ],
  );
});

// A saved comment has a section; the review is split again from text whose
// section headings were taken out. The empty field must clear the stored
// label: a save without the key would keep it (replaceComments).
test("a re-split without section headings sends an empty section for every row", async () => {
  const revision = {
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "Shorten this section.",
        section: "Abstract",
        requirements: [],
        status: "pending",
        changes: [],
        blocked: [],
      },
    ],
    ledger: [],
  };
  let t = null;
  t = await setup({
    revision,
    respond: {
      review_split: () => {
        t.pane.onEvent({
          type: "review_event",
          event: "job_done",
          job_id: "j1",
          result: {
            ok: true,
            method: "regex",
            warnings: [],
            comments: [
              { ref: "R1.1", reviewer: 1, text: "Shorten the abstract." },
              { ref: "R1.2", reviewer: 1, text: "Expand the methods." },
            ],
          },
        });
        return { ok: true, accepted: true, job_id: "j1" };
      },
      review_save_comments: () => ({ ok: true, revision, warnings: [] }),
    },
  });
  await t.$("review-split").click();
  const fields = findAll(t.$("review-draft"), (n) => n.attributes.class === "review-section");
  assert.deepEqual(
    fields.map((n) => n.value),
    ["", ""],
  );
  await t.$("review-save").click();
  const sent = t.requests.find((r) => r.type === "review_save_comments").payload.comments;
  assert.deepEqual(
    sent.map((c) => [c.ref, c.section]),
    [
      ["R1.1", ""],
      ["R1.2", ""],
    ],
  );
});

// Codex review 7: review_propose answers after the agent message is queued
// and the revision file is written, so turn_complete and turn_done can come
// first. The turn is taken before the request and a turn that already ended
// is not started again when the answer arrives.
async function proposeRace(order) {
  const { createTurnState } = await import("../taskpane/shared/turn-state.js");
  globalThis.document = fakeDocument();
  const { initReviewPane } = await import("../taskpane/shared/review-pane.js");
  const revision = {
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "State the unit.",
        requirements: [],
        status: "pending",
        changes: [],
        blocked: [],
      },
    ],
    ledger: [],
  };
  let answer;
  const answered = new Promise((r) => (answer = r));
  const counts = { accepted: 0, rejected: 0, user_edited: 0, pending: 0, unlocated: 0 };
  let composerDisabled = false;
  let pane = null;
  const turn = createTurnState({
    setComposerDisabled: (v) => (composerDisabled = v),
    getReviewPane: () => pane,
  });
  const sendRequest = async (type) => {
    if (type === "review_refresh") return { ok: true, rows: [], counts, external: [], revision };
    if (type === "review_propose") return answered;
    return { ok: true, revision };
  };
  pane = initReviewPane({
    sendRequest,
    word: { collectReviewSnapshot: async () => ({ paragraphs: [], comments: [] }) },
    // taskpane.js: take the turn, return its release.
    onTurnStarted() {
      const id = turn.begin();
      return () => turn.endIf(id);
    },
    isTurnInFlight: () => turn.inFlight,
  });
  await pane.reload();
  const propose = () =>
    findAll(document.getElementById("review-comments"), (n) => n.textContent === "Propose changes");
  const ends = () => {
    turn.onAssistantEvent({ event: "turn_complete" });
    pane.onEvent({ event: "turn_done", ref: "R1.1" });
  };
  try {
    const click = propose()[0].click();
    for (let i = 0; i < 5; i++) await null; // the handler reaches review_propose
    if (order === "answer-first") {
      answer({
        ok: true,
        started: true,
        ref: "R1.1",
        started_at: "2026-09-30T00:00:00Z",
        revision,
      });
      await click;
      assert.equal(turn.inFlight, true, "running after the answer");
      assert.deepEqual(
        propose().map((b) => b.disabled),
        [true],
      );
      ends();
    } else if (order === "end-first") {
      ends();
      answer({
        ok: true,
        started: true,
        ref: "R1.1",
        started_at: "2026-09-30T00:00:00Z",
        revision,
      });
      await click;
    } else {
      answer({ ok: false, error: "No such comment" });
      await click;
      assert.match(document.getElementById("review-message").textContent, /No such comment/);
    }
    for (let i = 0; i < 20; i++) await null; // turn_done reload settles
    assert.equal(turn.inFlight, false, `${order}: turn ended`);
    assert.equal(composerDisabled, false, `${order}: chat input enabled`);
    assert.deepEqual(
      propose().map((b) => b.disabled),
      [false],
      `${order}: Propose changes enabled`,
    );
    assert.equal(document.getElementById("review-progress").hidden, true, `${order}: no progress`);
  } finally {
    await pane.reload(); // stops a progress timer a failed check left running
  }
}

test("propose: answer then turn end leaves the buttons and the chat input enabled", () =>
  proposeRace("answer-first"));
test("propose: turn end before the answer does not start the turn again", () =>
  proposeRace("end-first"));
test("propose: a failed request releases the turn it took", () => proposeRace("failed"));

test("Refresh: Word refusing to resolve a comment still shows the counts", async () => {
  const change = { id: "R1.1-1", word_comment_id: "70", state: "accepted", rationale: {} };
  const revision = { comments: [{ ref: "R1.1", changes: [change], blocked: [] }], ledger: [] };
  const counts = { accepted: 1, rejected: 0, user_edited: 0, pending: 0, unlocated: 0 };
  const tried = [];
  const t = await setup({
    revision,
    word: {
      collectReviewSnapshot: async () => ({
        paragraphs: [],
        comments: [comment("70", "[R1.1] a / b / c")],
      }),
      resolveComment: async (id) => {
        tried.push(id);
        throw new Error("GeneralException");
      },
    },
    respond: {
      review_refresh: () => ({
        ok: true,
        rows: [{ change_id: "R1.1-1", state: "accepted" }],
        counts,
        external: [{ id: "X1" }],
        revision,
      }),
    },
  });
  await t.$("review-refresh").click();
  assert.deepEqual(tried, ["70"]);
  const m = t.$("review-message").textContent;
  assert.match(m, /^Every change is decided in Word\./);
  assert.match(m, /1 change was made outside the tool/);
  assert.match(m, /Word would not mark 1 rationale comment as resolved/);
  // The numbers are under Details.
  const d = t.$("review-message-details").textContent;
  assert.match(d, /^Accepted 1, rejected 0/);
  assert.match(d, /Outside changes: 1 /);
  assert.match(d, /1 rationale comment\(s\) not marked resolved \(Word: GeneralException;/);
  assert.equal(t.$("review-message-more").hidden, false);
});

test("Clean up tool comments tells the daemon which comment ids are gone", async () => {
  const revision = { comments: [], ledger: [] };
  const t = await setup({
    revision,
    word: {
      deleteToolComments: async () => ({
        before: { tool: 2, user: 1 },
        after: { tool: 0, user: 1 },
        deleted_ids: ["60", "61"],
      }),
    },
  });
  await t.$("review-clean-tool").click();
  const sent = t.requests.filter((r) => r.type === "review_comments_cleaned");
  assert.deepEqual(
    sent.map((r) => r.payload.ids),
    [["60", "61"]],
  );
  assert.match(t.$("review-message").textContent, /Tool comments 2 -> 0/);
});

// Codex review 8: two quick clicks on Propose changes both used to wait for
// refresh() before taking the turn. The second request was refused by the
// daemon and its release ended the first turn, so the chat input came back
// while the first turn ran. The first click now takes the turn before any
// wait and a second click while it runs does nothing.
test("propose: a double click sends one review_propose and keeps the chat input disabled", async () => {
  const { createTurnState } = await import("../taskpane/shared/turn-state.js");
  globalThis.document = fakeDocument();
  const { initReviewPane } = await import("../taskpane/shared/review-pane.js");
  const revision = {
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "State the unit.",
        requirements: [],
        status: "pending",
        changes: [],
        blocked: [],
      },
    ],
    ledger: [],
  };
  const counts = { accepted: 0, rejected: 0, user_edited: 0, pending: 0, unlocated: 0 };
  let composerDisabled = false;
  let pane = null;
  const turn = createTurnState({
    setComposerDisabled: (v) => (composerDisabled = v),
    getReviewPane: () => pane,
  });
  let answer;
  const answered = new Promise((r) => (answer = r));
  const proposals = [];
  const sendRequest = async (type) => {
    if (type === "review_refresh") return { ok: true, rows: [], counts, external: [], revision };
    if (type === "review_propose") {
      proposals.push(type);
      // The daemon runs one turn per pane and refuses a second one.
      if (proposals.length > 1) return { ok: false, error: "The agent is still working" };
      return answered;
    }
    return { ok: true, revision };
  };
  pane = initReviewPane({
    sendRequest,
    word: { collectReviewSnapshot: async () => ({ paragraphs: [], comments: [] }) },
    onTurnStarted() {
      const id = turn.begin();
      return () => turn.endIf(id);
    },
    isTurnInFlight: () => turn.inFlight,
  });
  await pane.reload();
  const button = findAll(
    document.getElementById("review-comments"),
    (n) => n.textContent === "Propose changes",
  )[0];
  try {
    const first = button.click();
    const second = button.click();
    for (let i = 0; i < 20; i++) await null;
    await second;
    assert.equal(proposals.length, 1, "one review_propose");
    assert.equal(turn.inFlight, true, "the first turn still runs");
    assert.equal(composerDisabled, true, "chat input stays disabled");
    answer({ ok: true, started: true, ref: "R1.1", started_at: "2026-09-30T00:00:00Z", revision });
    await first;
    assert.equal(composerDisabled, true, "still disabled until the turn ends");
    turn.onAssistantEvent({ event: "turn_complete" });
    pane.onEvent({ event: "turn_done", ref: "R1.1" });
    for (let i = 0; i < 20; i++) await null;
    assert.equal(composerDisabled, false, "enabled after the turn ends");
  } finally {
    await pane.reload();
  }
});

// Codex review 8: review_split and ledger_check answer at once with a job id
// and send the result later in a job_done review_event, so a model that
// takes longer than any request timeout still shows its result. The event
// may also come before the answer.
const SPLIT_RESULT = {
  ok: true,
  method: "llm",
  comments: [{ ref: "R1.1", reviewer: 1, text: "State the unit.", requirements: [] }],
  warnings: [],
};
for (const order of ["answer-first", "event-first"]) {
  test(`split and ledger check: the result comes in job_done (${order})`, async () => {
    let t = null;
    const later = [];
    const job = (id, result) => () => {
      const ev = { type: "review_event", event: "job_done", job_id: id, result };
      if (order === "event-first") t.pane.onEvent(ev);
      else later.push(ev);
      return { ok: true, accepted: true, job_id: id };
    };
    const revision = {
      comments: [],
      ledger: [{ id: "L1", kind: "forbid", text: "t", check: { type: "llm" } }],
    };
    t = await setup({
      revision,
      word: { collectReviewSnapshot: async () => ({ paragraphs: [], comments: [] }) },
      respond: {
        review_split: job("j1", SPLIT_RESULT),
        ledger_check: job("j2", {
          ok: true,
          outcomes: [
            { id: "L1", status: "violations", violations: [{ paragraph_index: 3, match: "x" }] },
          ],
          revision,
        }),
      },
    });
    const split = t.$("review-split").click();
    const check = t.$("review-check-all").click();
    for (let i = 0; i < 20; i++) await null;
    if (order === "answer-first") {
      assert.equal(t.$("review-draft-actions").hidden, true, "nothing before job_done");
      for (const ev of later.reverse()) t.pane.onEvent(ev);
    }
    await split;
    await check;
    assert.equal(t.$("review-draft-actions").hidden, false, `${order}: split shown`);
    assert.match(t.$("review-message").textContent, /^Split into 1 comment \(read by Claude\)/);
    // The rule is named by its words; the paragraph by what it starts with (the
    // snapshot of this test holds none, so "a paragraph").
    assert.deepEqual(
      t.$("review-violations").children.map((b) => b.textContent),
      ["t: found in 1 place", "a paragraph has \u201cx\u201d"],
    );
    for (const r of t.requests.filter((q) => ["review_split", "ledger_check"].includes(q.type))) {
      assert.equal(r.opts, undefined, `${r.type}: default request timeout`);
    }
  });
}

test("a job still waiting when the daemon connection restarts ends with an error", async () => {
  const t = await setup({
    revision: { comments: [], ledger: [] },
    respond: { review_split: () => ({ ok: true, accepted: true, job_id: "j9" }) },
  });
  const split = t.$("review-split").click();
  for (let i = 0; i < 20; i++) await null;
  await t.pane.reload();
  await split;
  assert.match(t.$("review-message").textContent, /connection/i);
});

// Codex review 9: a render (an event, another button) rebuilds the comment
// cards; what the author typed and did not save must come back.
function typeInto(node, value) {
  node.value = value;
  for (const fn of node.listeners.input ?? []) fn();
}
async function settle() {
  for (let i = 0; i < 20; i++) await null;
}

test("unsaved requirement text survives a re-render and is cleared only by a save", async () => {
  const revision = {
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "State the unit.",
        requirements: [{ id: "q1", text: "State the unit" }],
        status: "in_progress",
        changes: [],
        blocked: [],
      },
    ],
    ledger: [],
  };
  let failSave = true;
  const t = await setup({
    revision,
    respond: {
      review_set_requirements: () =>
        failSave ? { ok: false, error: "disk full" } : { ok: true, revision },
    },
  });
  const reqs = () => findAll(t.$("review-comments"), (n) => n.attributes.class === "review-reqs");
  const save = () =>
    findAll(t.$("review-comments"), (n) => n.textContent === "Save requirements")[0];

  typeInto(reqs()[0], "State the unit\nGive the n");
  t.pane.onEvent({ type: "review_event", event: "change_recorded", ref: "R1.1" });
  await settle();
  assert.equal(reqs()[0].value, "State the unit\nGive the n", "draft kept after an event");

  await save().click(); // fails: the draft stays
  assert.equal(reqs()[0].value, "State the unit\nGive the n", "draft kept after a failed save");

  failSave = false;
  await save().click();
  assert.deepEqual(
    t.requests.filter((r) => r.type === "review_set_requirements").at(-1).payload.requirements,
    ["State the unit", "Give the n"],
  );
  assert.equal(reqs()[0].value, "State the unit", "saved: the stored list shows again");
});

test("split draft rows, the paste box and the ledger form survive a re-render", async () => {
  const t = await setup({
    revision: { comments: [], ledger: [] },
    respond: {
      review_split: () => {
        queueMicrotask(() =>
          t.pane.onEvent({
            type: "review_event",
            event: "job_done",
            job_id: "j1",
            result: SPLIT_RESULT,
          }),
        );
        return { ok: true, accepted: true, job_id: "j1" };
      },
    },
  });
  t.$("review-input").value = "Reviewer 1: pasted review";
  await t.$("review-split").click();
  const row = () => findAll(t.$("review-draft"), (n) => n.attributes.class === "review-reqs")[0];
  typeInto(row(), "State the unit in degrees C");
  t.$("ledger-text").value = "Never write data shows";
  t.$("ledger-pattern").value = "data shows";
  t.pane.onEvent({ type: "review_event", event: "change_recorded", ref: "R1.1" });
  await settle();
  assert.equal(row().value, "State the unit in degrees C");
  assert.equal(t.$("review-input").value, "Reviewer 1: pasted review");
  assert.equal(t.$("ledger-text").value, "Never write data shows");
  assert.equal(t.$("ledger-pattern").value, "data shows");
});

// P15: listeners of a non-click event on a fake element.
async function fire(node, ev) {
  for (const fn of node.listeners[ev] ?? []) await fn();
}

test("Load review from file fills the paste box; a wrong type only shows an error", async () => {
  const t = await setup({ revision: { comments: [], ledger: [] } });
  const input = t.$("review-file-input");
  let opened = 0;
  input.addEventListener("click", () => (opened += 1));
  await t.$("review-load-file").click();
  assert.equal(opened, 1, "the button opens the file picker");

  input.files = [{ name: "review.txt", size: 20, text: async () => "\uFEFFReviewer 1\r\nOne." }];
  await fire(input, "change");
  assert.equal(t.$("review-input").value, "Reviewer 1\nOne.");
  assert.match(t.$("review-message").textContent, /Loaded review\.txt/);

  input.files = [{ name: "review.docx", size: 20, text: async () => "PK" }];
  await fire(input, "change");
  assert.equal(t.$("review-input").value, "Reviewer 1\nOne.", "the box is left as it was");
  assert.match(t.$("review-message").textContent, /\.txt or \.md/);
  assert.equal(t.requests.filter((r) => r.type !== "review_load").length, 0, "no daemon request");
});

test("ledger: Export shows the file path; Import sends the items and reports added and skipped", async () => {
  const revision = { comments: [], ledger: [] };
  const item = { kind: "forbid", text: "no novel", scope: "manuscript", check: { type: "llm" } };
  const t = await setup({
    revision,
    respond: {
      ledger_export: () => ({ ok: true, path: "/m/paper.ledger.json", count: 2 }),
      ledger_import: () => ({ ok: true, added: 1, skipped: 0, revision }),
    },
  });
  await t.$("ledger-export").click();
  assert.match(t.$("review-message").textContent, /2 ledger rule\(s\).*\/m\/paper\.ledger\.json/);

  const input = t.$("ledger-file-input");
  input.files = [{ name: "l.json", size: 50, text: async () => JSON.stringify({ items: [item] }) }];
  await fire(input, "change");
  assert.deepEqual(t.requests.find((r) => r.type === "ledger_import").payload, {
    doc: { items: [item] },
    mode: "append",
  });
  assert.match(t.$("review-message").textContent, /Added 1 ledger rule\(s\), skipped 0/);

  input.files = [{ name: "l.json", size: 5, text: async () => "{ nope" }];
  await fire(input, "change");
  assert.match(t.$("review-message").textContent, /not JSON/);
  assert.equal(t.requests.filter((r) => r.type === "ledger_import").length, 1);
});

test("ledger suggestions: a checklist from the review text; Add selected imports the ticked rules", async () => {
  const revision = { comments: [], ledger: [] };
  const cue = {
    kind: "forbid",
    text: 'Do not use "novel".',
    scope: "manuscript",
    check: { type: "regex", pattern: "novel", flags: "i", expect: "absent" },
    why: "forbid cue",
    sentence_index: 1,
  };
  const model = {
    kind: "instruction",
    text: "Use SI units throughout.",
    scope: "manuscript",
    check: { type: "llm" },
    why: "model label",
    sentence_index: 2,
  };
  let t = null;
  t = await setup({
    revision,
    respond: {
      ledger_suggest: () => {
        queueMicrotask(() =>
          t.pane.onEvent({
            type: "review_event",
            event: "job_done",
            job_id: "j5",
            result: { ok: true, method: "regex+llm", candidates: [cue, model], warnings: [] },
          }),
        );
        return { ok: true, accepted: true, job_id: "j5" };
      },
      ledger_import: () => ({ ok: true, added: 2, skipped: 0, revision }),
    },
  });
  t.$("review-input").value = "Reviewer 1 ...";
  await t.$("ledger-suggest-review").click();
  assert.equal(t.requests.find((r) => r.type === "ledger_suggest").payload.text, "Reviewer 1 ...");
  const boxes = () => findAll(t.$("ledger-candidates"), (n) => n.attributes.type === "checkbox");
  // A cue-based candidate starts ticked, a model label does not.
  assert.deepEqual(
    boxes().map((b) => b.checked),
    [true, false],
  );
  assert.equal(t.$("ledger-candidates-actions").hidden, false);
  assert.match(t.$("review-message").textContent, /2 candidate rule\(s\)/);

  boxes()[1].checked = true;
  await fire(boxes()[1], "change");
  await t.$("ledger-candidates-add").click();
  const strip = ({ kind, text, scope, check }) => ({ kind, text, scope, check });
  assert.deepEqual(t.requests.find((r) => r.type === "ledger_import").payload, {
    doc: { items: [strip(cue), strip(model)] },
    mode: "append",
  });
  assert.equal(boxes().length, 0, "the list is cleared once added");
  assert.equal(t.$("ledger-candidates-actions").hidden, true);
});

// P15 "Propose all": one review turn per open comment, the next one only
// after the turn_done of the one before.
// The text of an element and of everything under it (the fake keeps text per element).
const textOf = (n) => [n.textContent, ...n.children.map(textOf)].join("");
// What the card of comment `ref` says about its last turn.
const resultOf = (t, ref) => {
  const card = findAll(t.$("review-comments"), (n) => n.attributes?.["data-ref"] === ref)[0];
  const r = card
    ? findAll(card, (n) => String(n.attributes?.class ?? "").startsWith("review-card-result"))[0]
    : null;
  return r ? textOf(r) : "";
};

async function proposeAllSetup({ providersReady, usedLabel } = {}) {
  const { createTurnState } = await import("../taskpane/shared/turn-state.js");
  globalThis.document = fakeDocument();
  const { initReviewPane } = await import("../taskpane/shared/review-pane.js");
  const c = (ref, status) => ({
    ref,
    reviewer: 1,
    text: "State the unit.",
    requirements: [],
    status,
    changes: [],
    blocked: [],
  });
  const revision = {
    comments: [
      c("R1.1", "pending"),
      c("R1.2", "done"),
      c("R1.3", "pending"),
      c("R2.1", "in_progress"),
    ],
    ledger: [],
  };
  const counts = { accepted: 0, rejected: 0, user_edited: 0, pending: 0, unlocated: 0 };
  let pane = null;
  const turn = createTurnState({ getReviewPane: () => pane });
  const proposals = [];
  // fail[type] = "message": the next request of that type is refused, once.
  const fail = {};
  // doc.comments: the comments Word reports in the next snapshot.
  const doc = { comments: [] };
  const sent = []; // [type, payload, opts] of every request
  const sendRequest = async (type, payload, opts) => {
    sent.push([type, payload, opts]);
    if (fail[type]) {
      const error = fail[type];
      delete fail[type];
      return { ok: false, error };
    }
    if (type === "review_refresh") return { ok: true, rows: [], counts, external: [], revision };
    if (type === "review_propose") {
      proposals.push([payload.ref, payload.paragraphs.length]);
      return { ok: true, started: true, ref: payload.ref, started_at: Date.now(), revision };
    }
    return { ok: true, revision, path: "/m/paper.revision.json" };
  };
  pane = initReviewPane({
    sendRequest,
    word: {
      collectReviewSnapshot: async () => ({ paragraphs: [], comments: doc.comments }),
      selectionParagraphs: async () => [{ index: 1, text: "selected" }],
    },
    onTurnStarted() {
      const id = turn.begin();
      return () => turn.endIf(id);
    },
    isTurnInFlight: () => turn.inFlight,
    providersReady,
    usedLabel,
  });
  await pane.reload();
  const $ = (id) => document.getElementById(id);
  const wait = () => new Promise((r) => setTimeout(r, 5));
  // What the daemon sends when the turn of `ref` ends: turn_done says how.
  const endTurn = (ref, error = null) => {
    turn.onAssistantEvent({ event: "turn_complete" });
    pane.onEvent({
      type: "review_event",
      event: "turn_done",
      ref,
      ok: !error,
      ...(error ? { error } : {}),
    });
  };
  return { pane, turn, proposals, sent, fail, doc, revision, $, wait, endTurn };
}

test("Propose all sends the next review_propose only after turn_done; an error stops the queue", async () => {
  const t = await proposeAllSetup();
  try {
    t.$("review-use-selection").checked = true; // not used by Propose all
    const all = t.$("review-propose-all").click();
    await t.wait();
    assert.deepEqual(t.proposals, [["R1.1", 0]]);
    assert.equal(t.$("review-propose-stop").hidden, false);
    assert.equal(t.$("review-propose-all").disabled, true);
    assert.match(t.$("review-progress").textContent, /^Propose all: 1\/3 R1\.1\. R1\.1 running/);
    const cards = () =>
      findAll(t.$("review-comments"), (n) => n.textContent === "Propose changes").map(
        (b) => b.disabled,
      );
    assert.deepEqual(cards(), [true, true, true, true]);

    await t.wait();
    assert.equal(t.proposals.length, 1, "no second turn while the first one runs");
    await t.$("review-propose-all").click(); // a second press while it runs
    assert.equal(t.proposals.length, 1);

    t.endTurn("R1.1");
    await t.wait();
    assert.deepEqual(
      t.proposals.map((p) => p[0]),
      ["R1.1", "R1.3"],
      "the done comment R1.2 is skipped",
    );
    assert.equal(t.turn.inFlight, true);

    t.endTurn("R1.3", "Claude usage limit reached.");
    await all;
    assert.equal(t.proposals.length, 2, "R2.1 is not started after the failure");
    assert.match(
      t.$("review-message").textContent,
      /Propose all stopped at R1\.3: Claude usage limit reached\. 1 of 3 proposed\./,
    );
    assert.equal(t.$("review-propose-stop").hidden, true);
    assert.equal(t.$("review-propose-all").disabled, false);
    assert.equal(t.turn.inFlight, false);
    assert.deepEqual(cards(), [false, false, false, false]);
    assert.equal(t.$("review-progress").hidden, true);
  } finally {
    await t.pane.reload();
  }
});

test("Propose all: Stop after this one lets the turn end and starts no other; a reconnect ends the wait", async () => {
  const t = await proposeAllSetup();
  try {
    const all = t.$("review-propose-all").click();
    await t.wait();
    await t.$("review-propose-stop").click();
    await t.wait();
    assert.equal(t.proposals.length, 1);
    t.endTurn("R1.1");
    await all;
    assert.equal(t.proposals.length, 1);
    assert.match(t.$("review-message").textContent, /Propose all stopped before R1\.3\. 1 of 3/);

    // A daemon restart never sends turn_done: welcome (reload) ends the wait.
    const again = t.$("review-propose-all").click();
    await t.wait();
    assert.equal(t.proposals.length, 2);
    t.turn.end(); // ws.onclose
    await t.pane.reload();
    await again;
    assert.equal(t.proposals.length, 2);
    assert.match(t.$("review-message").textContent, /stopped at R1\.1: The connection/);
  } finally {
    await t.pane.reload();
  }
});

// The usage-limit path of the daemon sends turn_complete, turn_done, then the
// error event. turn_done says the turn failed, so the queue stops whenever
// the error event comes: here after the refresh of turn_done has finished.
test("Propose all: a failed turn_done stops the queue even when the error event comes late", async () => {
  const t = await proposeAllSetup();
  try {
    const all = t.$("review-propose-all").click();
    await t.wait();
    t.endTurn("R1.1", "Claude usage limit reached.");
    await all; // the refresh of turn_done is over, the run has ended
    t.turn.onAssistantEvent({ event: "error", error: "Claude usage limit reached." });
    await t.wait();
    assert.equal(t.proposals.length, 1, "R1.3 is not started");
    assert.match(
      t.$("review-message").textContent,
      /^Propose all stopped at R1\.1: Claude usage limit reached\. 0 of 3 proposed\./,
    );
    assert.equal(t.pane.onTurnError, undefined, "no timing-dependent error hook");
  } finally {
    await t.pane.reload();
  }
});

// An error event alone decides nothing: the turn_done of that turn does.
test("Propose all: an error event without a failed turn_done does not end the wait", async () => {
  const t = await proposeAllSetup();
  try {
    const all = t.$("review-propose-all").click();
    await t.wait();
    t.turn.onAssistantEvent({ event: "error", error: "x" });
    await t.wait();
    assert.equal(t.proposals.length, 1);
    assert.equal(t.$("review-propose-stop").hidden, false, "still waiting for turn_done");
    t.endTurn("R1.1", "x");
    await all;
    assert.match(t.$("review-message").textContent, /stopped at R1\.1: x/);
  } finally {
    await t.pane.reload();
  }
});

// The turn ran, but the reload or the refresh that follows its turn_done
// failed: the pane does not know the state of that comment, so the queue stops
// there and does not count it. What the turn recorded is saved by the daemon;
// the message says so and names the way out.
const unchecked = (ref, error) =>
  `The turn for ${ref} ran and its record is saved, but the status check after it failed (${error}); press Refresh status to check again.`;

test("Propose all: every turn and its status check succeed", async () => {
  const t = await proposeAllSetup();
  try {
    const all = t.$("review-propose-all").click();
    for (const ref of ["R1.1", "R1.3", "R2.1"]) {
      await t.wait();
      t.endTurn(ref);
    }
    await all;
    assert.deepEqual(
      t.proposals.map((p) => p[0]),
      ["R1.1", "R1.3", "R2.1"],
    );
    assert.equal(t.$("review-message").textContent, "Propose all: 3 of 3 proposed.");
  } finally {
    await t.pane.reload();
  }
});

test("Propose all: a failed reload after the last turn is not counted as proposed", async () => {
  const t = await proposeAllSetup();
  try {
    const all = t.$("review-propose-all").click();
    for (const ref of ["R1.1", "R1.3"]) {
      await t.wait();
      t.endTurn(ref);
    }
    await t.wait();
    t.fail.review_load = "revision file is locked";
    t.endTurn("R2.1");
    await all;
    assert.equal(t.proposals.length, 3);
    assert.equal(
      t.$("review-message").textContent,
      `Propose all stopped at R2.1: ${unchecked("R2.1", "revision file is locked")} R2.1 is not counted: 2 of 3 proposed.`,
    );
    assert.doesNotMatch(t.$("review-message").textContent, /lost|undone|discarded/i);
    assert.equal(t.$("review-propose-all").disabled, false);
    assert.equal(t.turn.inFlight, false);
  } finally {
    delete t.fail.review_load;
    await t.pane.reload();
  }
});

test("Propose all: a failed refresh after a turn in the middle starts no other turn", async () => {
  const t = await proposeAllSetup();
  try {
    const all = t.$("review-propose-all").click();
    await t.wait();
    t.endTurn("R1.1");
    await t.wait();
    assert.equal(t.proposals.length, 2);
    // Once: the refresh before the turn of R2.1 would succeed.
    t.fail.review_refresh = "Word is busy";
    t.endTurn("R1.3");
    // Not `await all`: a queue that went on would wait for the turn of R2.1.
    await t.wait();
    await t.wait();
    assert.deepEqual(
      t.proposals.map((p) => p[0]),
      ["R1.1", "R1.3"],
      "R2.1 is not started",
    );
    assert.equal(
      t.$("review-message").textContent,
      `Propose all stopped at R1.3: ${unchecked("R1.3", "Word is busy")} R1.3 is not counted: 1 of 3 proposed.`,
    );
    assert.equal(t.$("review-propose-stop").hidden, true);
    assert.equal(t.turn.inFlight, false);
    await all;
  } finally {
    delete t.fail.review_refresh;
    await t.pane.reload();
  }
});

// A turn that failed stops the queue with its own error, whether the status
// check after it worked or not.
test("Propose all: the error of a failed turn is the one shown when its status check fails too", async () => {
  const t = await proposeAllSetup();
  try {
    const all = t.$("review-propose-all").click();
    await t.wait();
    t.fail.review_load = "revision file is locked";
    t.endTurn("R1.1", "Claude usage limit reached.");
    await all;
    assert.equal(t.proposals.length, 1);
    assert.equal(
      t.$("review-message").textContent,
      "Propose all stopped at R1.1: Claude usage limit reached. 0 of 3 proposed.",
    );
  } finally {
    delete t.fail.review_load;
    await t.pane.reload();
  }
});

test("one Propose changes: a failed status check after the turn says what to do", async () => {
  const t = await proposeAllSetup();
  try {
    const [first] = findAll(t.$("review-comments"), (n) => n.textContent === "Propose changes");
    await first.click();
    assert.deepEqual(t.proposals, [["R1.1", 0]]);
    t.fail.review_refresh = "Word is busy";
    t.endTurn("R1.1");
    await t.wait();
    assert.equal(t.$("review-message").textContent, unchecked("R1.1", "Word is busy"));
  } finally {
    delete t.fail.review_refresh;
    await t.pane.reload();
  }
});

// The status check worked and found that the changes on record and the tool
// comments in Word do not match (an edit that was not recorded, or a
// rationale comment that is gone). The next turn would start on that state:
// the queue stops, as for a check that failed.
const mismatched = (ref, missing, unrecorded) =>
  `The turn for ${ref} ran and its record is saved, but the changes on record and the tool comments in Word do not match (missing ${missing}, unrecorded ${unrecorded}); check those in Word, then press Refresh status.`;

test("Propose all: a mismatch found by the status check after a turn starts no other turn", async () => {
  const t = await proposeAllSetup();
  try {
    const all = t.$("review-propose-all").click();
    await t.wait();
    // A tool comment in Word that no recorded change has.
    t.doc.comments = [comment("70", "[R1.1] a / b / c")];
    t.endTurn("R1.1");
    // Not `await all`: a queue that went on would wait for the turn of R1.3.
    await t.wait();
    await t.wait();
    assert.deepEqual(
      t.proposals.map((p) => p[0]),
      ["R1.1"],
      "R1.3 is not started",
    );
    assert.equal(
      t.$("review-message").textContent,
      `Propose all stopped at R1.1: ${mismatched("R1.1", "-", "70")} R1.1 is not counted: 0 of 3 proposed.`,
    );
    assert.equal(t.turn.inFlight, false);
    await all;
  } finally {
    await t.pane.reload();
  }
});

test("one Propose changes: a mismatch after the turn says what to do; Refresh status keeps its counts", async () => {
  const t = await proposeAllSetup();
  try {
    const [first] = findAll(t.$("review-comments"), (n) => n.textContent === "Propose changes");
    await first.click();
    t.doc.comments = [comment("70", "[R1.1] a / b / c")];
    t.endTurn("R1.1");
    await t.wait();
    assert.equal(t.$("review-message").textContent, mismatched("R1.1", "-", "70"));
    await t.$("review-refresh").click();
    assert.match(
      t.$("review-message").textContent,
      /^The changes on record and the comments in Word do not match\./,
    );
    assert.match(
      t.$("review-message-details").textContent,
      /^Accepted 0, .* found 1 \(MISMATCH: missing -, unrecorded 70\)$/,
    );
  } finally {
    await t.pane.reload();
  }
});

// P15: the status refreshes by itself when the pane gets the focus back and
// when any turn ends, without touching the message line.
test("auto refresh: focus and a turn end refresh quietly; not while there are no comments", async () => {
  const listeners = {};
  globalThis.window = { addEventListener: (ev, fn) => (listeners[ev] = fn) };
  const counts = { accepted: 0, rejected: 0, user_edited: 0, pending: 0, unlocated: 0 };
  const wait = () => new Promise((r) => setTimeout(r, 5));
  const refreshes = (t) => t.requests.filter((r) => r.type === "review_refresh").length;
  const word = { collectReviewSnapshot: async () => ({ paragraphs: [], comments: [] }) };
  try {
    const revision = {
      comments: [
        { ref: "R1.1", reviewer: 1, text: "x", status: "pending", changes: [], blocked: [] },
      ],
      ledger: [],
    };
    const t = await setup({
      revision,
      word,
      respond: {
        review_refresh: () => ({ ok: true, rows: [], counts, external: [], revision }),
      },
    });
    t.$("review-message").textContent = "KEEP";
    listeners.focus();
    await wait();
    assert.equal(refreshes(t), 1);
    assert.equal(t.$("review-message").textContent, "KEEP", "the message line is not replaced");
    assert.match(t.$("review-auto").textContent, /^Auto-refreshed .*No changes yet\./);
    listeners.focus(); // again at once: too soon
    await wait();
    assert.equal(refreshes(t), 1);
    assert.equal(t.pane.onAnyTurnEnded(), "started"); // a chat turn ended
    await wait();
    assert.equal(refreshes(t), 2);

    const empty = await setup({ revision: { comments: [], ledger: [] }, word });
    listeners.focus();
    assert.equal(empty.pane.onAnyTurnEnded(), "busy");
    await wait();
    assert.equal(refreshes(empty), 0);
  } finally {
    delete globalThis.window;
  }
});

// Codex review 14: a normal edit Word did not answer is kept in the file with
// no forbid rule. It is no exception: no Allow button, only the re-apply.
test("an unanswered normal edit: named as such, no Allow exception, re-apply with force", async () => {
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
            args: { find: "30 degrees", replace: "30 °C" },
            blocked_by: [],
            created_at: "2026-01-01T00:00:00Z",
            state: "outcome_unknown",
            attempt: null,
          },
        ],
      },
    ],
    ledger: [],
  };
  const respond = { review_override: () => ({ ok: true, applied: true, revision }) };
  const t = await setup({ revision, respond });
  const find = (re) => findAll(t.$("review-comments"), (n) => re.test(n.textContent));
  assert.equal(find(/^No answer from Word: "30 °C" $/).length, 1);
  assert.equal(find(/Blocked by/).length, 0);
  assert.equal(find(/^Allow exception$/).length, 0);
  assert.equal(find(/Refresh status/).length, 1);
  assert.equal(
    findAll(t.$("review-decisions"), (n) =>
      /^R1\.1: Word did not answer an edit \(\u201c30 °C\u201d\)\. Open the comment to check it\.$/.test(
        n.textContent,
      ),
    ).length,
    1,
  );
  assert.equal(find(/this edit may already be applied/).length, 1);
  await find(/Re-apply anyway/)[0].click();
  await find(/Re-apply anyway/)[0].click();
  const sent = t.requests.filter((r) => r.type === "review_override");
  assert.deepEqual(sent[0].payload, { ref: "R1.1", blocked_id: "R1.1-b1", force: true });
});

// Codex review 15: Preview prompt sent only the ref, so with "send selected
// paragraphs" on it showed the prompt without them while Propose changes
// sent the selection.
test("Preview prompt and Propose changes send the same paragraphs under the selection option", async () => {
  const revision = {
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "State the unit.",
        requirements: [],
        status: "pending",
        changes: [],
        blocked: [],
      },
    ],
    ledger: [],
  };
  const counts = { accepted: 0, rejected: 0, user_edited: 0, pending: 0, unlocated: 0 };
  const selected = [
    { id: "U4", index: 4, text: "Cells were held at 30 degrees." },
    { id: "U5", index: 5, text: "Growth was followed for 2 h." },
  ];
  let selectionReads = 0;
  const respond = {
    review_refresh: () => ({ ok: true, rows: [], counts, external: [], revision }),
    review_preview: () => ({ ok: true, prompt: "PROMPT" }),
    review_propose: () => ({
      ok: true,
      started: true,
      ref: "R1.1",
      started_at: "2026-09-30T00:00:00Z",
      revision,
    }),
  };
  const t = await setup({
    revision,
    respond,
    word: {
      collectReviewSnapshot: async () => ({ paragraphs: [], comments: [] }),
      selectionParagraphs: async () => {
        selectionReads += 1;
        return selected;
      },
    },
  });
  const button = (label) => findAll(t.$("review-comments"), (n) => n.textContent === label)[0];
  // wait_ms belongs to review_propose alone (how long the pane waits for it).
  const sent = (type) =>
    t.requests.filter((r) => r.type === type).map(({ payload: { wait_ms: _, ...rest } }) => rest);

  // Option on: both carry the selection.
  t.$("review-use-selection").checked = true;
  await button("Preview prompt").click();
  await button("Propose changes").click();
  assert.deepEqual(sent("review_preview"), [{ ref: "R1.1", paragraphs: selected }]);
  assert.deepEqual(sent("review_propose"), [{ ref: "R1.1", paragraphs: selected }]);
  assert.deepEqual(sent("review_preview")[0], sent("review_propose")[0]);
  assert.equal(selectionReads, 2);
  assert.equal(t.$("review-preview").textContent, "PROMPT");

  // Option off: both carry none, and the selection is not read.
  t.$("review-use-selection").checked = false;
  await button("Preview prompt").click();
  await button("Propose changes").click();
  assert.deepEqual(sent("review_preview")[1], { ref: "R1.1", paragraphs: [] });
  assert.deepEqual(sent("review_preview")[1], sent("review_propose")[1]);
  assert.equal(selectionReads, 2);
});

// An edit of a turn was taken back and Word took an earlier pending change
// of the paragraph with it: the pane refreshes, and the card of the comment
// that was proposed says one sentence that stays after the turn ended.
test("undo_took_more refreshes the status and puts one sentence in the card of the comment, which stays when the turn ends", async () => {
  const SENTENCE =
    "While taking back an edit that would have damaged a citation, Word also took back an earlier pending change in the same paragraph. Check the paragraph and ask for that change again if you still want it.";
  const revision = {
    comments: ["R1.1", "R1.2"].map((ref) => ({
      ref,
      reviewer: 1,
      text: "Reword the sentence.",
      requirements: [],
      status: "in_progress",
      changes: [],
      blocked: [],
    })),
    ledger: [],
  };
  const counts = { accepted: 0, rejected: 0, user_edited: 0, pending: 0, unlocated: 0 };
  const now = { paragraphs: [{ index: 0, text: "A material." }], comments: [] };
  const t = await setup({
    revision,
    respond: {
      review_refresh: () => ({ ok: true, rows: [], counts, external: [], revision }),
      review_propose: () => ({ ok: true, started: true, ref: "R1.2", revision }),
    },
    word: { collectReviewSnapshot: async () => now },
  });
  const notices = () =>
    findAll(t.$("review-comments"), (n) => /review-card-notice/.test(n.attributes.class ?? "")).map(
      (n) => n.textContent,
    );
  const refreshes = () => t.requests.filter((r) => r.type === "review_refresh").length;
  assert.deepEqual(notices(), []);
  t.pane.onEvent({
    type: "review_event",
    event: "undo_took_more",
    ref: "R1.2",
    paragraph_index: 0,
  });
  await settle();
  assert.equal(refreshes(), 1, "the review state is read from the document again");
  assert.deepEqual(notices(), [SENTENCE]);
  assert.doesNotMatch(SENTENCE, /[\u2014]|R1\.\d-\d|[0-9a-f]{8}/);
  // The end of the turn writes its own result line; the sentence stays.
  t.pane.onEvent({ type: "review_event", event: "turn_done", ref: "R1.2", ok: true });
  await settle();
  assert.deepEqual(notices(), [SENTENCE]);
  // It is in the card of that comment only.
  const card = findAll(
    t.$("review-comments"),
    (n) => n.attributes["data-ref"] === "R1.2" || n.attributes["data-comment"] === "R1.2",
  );
  if (card.length)
    assert.equal(
      findAll(card[0], (n) => /review-card-notice/.test(n.attributes.class ?? "")).length,
      1,
    );
});

// H2: a tracked edit the task pane refused because its paragraph changed
// (StaleTarget). The daemon has the document only from a status refresh, so
// it asks for one. H1: an attempt that is still with Word is not listed.
test("stale_target refreshes the status with the document as it is now; a running attempt is not shown", async () => {
  const row = (id, state, blocked_by) => ({ id, args: { replace: "x" }, blocked_by, state });
  const revision = {
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "State the unit.",
        requirements: [],
        status: "in_progress",
        changes: [],
        blocked: [row("R1.1-b1", "applying", []), row("R1.1-b2", "outcome_unknown", [])],
      },
    ],
    ledger: [],
  };
  const counts = { accepted: 0, rejected: 0, user_edited: 0, pending: 0, unlocated: 0 };
  const now = { paragraphs: [{ index: 0, text: "A novel material." }], comments: [] };
  const t = await setup({
    revision,
    respond: {
      review_refresh: () => ({ ok: true, rows: [], counts, external: [], revision }),
    },
    word: { collectReviewSnapshot: async () => now },
  });
  const rows = () =>
    findAll(t.$("review-comments"), (n) => n.attributes.class === "review-blocked").length;
  assert.equal(rows(), 1, "the unanswered edit only");
  assert.equal(t.requests.filter((r) => r.type === "review_refresh").length, 0);
  t.pane.onEvent({ type: "review_event", event: "stale_target", ref: "R1.1" });
  await settle();
  assert.deepEqual(
    t.requests.filter((r) => r.type === "review_refresh").map((r) => r.payload),
    // saved: this Word did not say whether the document has unsaved changes.
    [{ snapshot: now, saved: null }],
  );
  // H1: Word applied an edit and the daemon could not save its record.
  const loads = () => t.requests.filter((r) => r.type === "review_load").length;
  const before = loads();
  t.pane.onEvent({
    type: "review_event",
    event: "record_failed",
    ref: "R1.1",
    attempt_id: "R1.1-b2",
  });
  await settle();
  const m = t.$("review-message");
  assert.equal(
    m.textContent,
    "R1.1: the edit was applied in Word but saving its record failed; press Refresh status to match them again.",
  );
  assert.equal(m.hidden, false);
  assert.equal(loads(), before + 1, "the list is loaded again");
  // The load found the attempt gone: the daemon saved the change with it.
  t.pane.onEvent({
    type: "review_event",
    event: "record_failed",
    ref: "R1.1",
    attempt_id: "R1.1-b3",
  });
  await settle();
  assert.equal(m.textContent, "R1.1: the edit was applied in Word; its record is saved now.");
  assert.equal(loads(), before + 2);
});

// Codex review of stage 2.5: right after a reconnect the daemon has not yet
// confirmed the provider the pane stored. A Review action clicked then is
// held, with a line that says why, and sent once the daemon answered.
test("a Review action clicked while the stored provider is being confirmed waits, then runs", async () => {
  let confirm;
  let pending = new Promise((r) => (confirm = r));
  const t = await setup({
    revision: { comments: [], ledger: [] },
    providersReady: () => pending,
    respond: {
      review_split: () => ({ ok: false, error: "split was asked" }),
    },
  });
  t.$("review-input").value = "Reviewer 1\nComment 1: Use SI units.";
  const clicked = t.$("review-split").click();
  await new Promise((r) => setTimeout(r, 5));
  assert.equal(
    t.$("review-message").textContent,
    "Waiting for the daemon to confirm the stored provider choice...",
  );
  assert.equal(t.requests.filter((r) => r.type === "review_split").length, 0, "held, not sent");
  pending = null; // the daemon answered
  confirm();
  await clicked;
  assert.equal(t.requests.filter((r) => r.type === "review_split").length, 1);
  assert.equal(t.$("review-message").textContent, "split was asked");
  // Nothing to wait for: the next click goes straight through.
  await t.$("review-split").click();
  assert.equal(t.requests.filter((r) => r.type === "review_split").length, 2);
});

// ---- Done asks once more while something is left open ---------------------------
const doneComment = (changes, requirements = []) => ({
  ref: "R1.1",
  reviewer: 1,
  text: "State the unit.",
  status: "in_progress",
  requirements,
  changes,
  blocked: [],
});
const doneChange = (id, state) => ({
  id,
  state,
  kind: "replace",
  before: "b",
  after: "a",
  rationale: { what: "w", why: "y", addresses: "a", requirement_ids: [] },
});
async function chooseDone(t) {
  const select = findAll(t.$("review-comments"), (n) => n.tagName === "SELECT")[0];
  select.value = "done";
  for (const fn of select.listeners.change) await fn();
}
const statusSent = (t) => t.requests.filter((r) => r.type === "review_set_status");

test("Done with changes not decided in Word asks once more, then marks it done", async () => {
  const revision = {
    comments: [doneComment([doneChange("R1.1-1", "pending"), doneChange("R1.1-2", "unlocated")])],
    ledger: [],
  };
  const t = await setup({ revision });
  await chooseDone(t);
  assert.equal(statusSent(t).length, 0);
  assert.equal(
    t.$("review-message").textContent,
    "R1.1: 2 changes are not decided in Word (1 pending, 1 not found). Choose Done again to mark it done anyway.",
  );
  await chooseDone(t);
  assert.deepEqual(statusSent(t)[0].payload, { ref: "R1.1", status: "done" });
});

test("Done with a confirmed ledger rule in violation asks once more; with nothing open it is sent at once", async () => {
  const comments = [doneComment([doneChange("R1.1-1", "accepted")])];
  const dirty = await setup({
    revision: {
      comments,
      ledger: [
        {
          id: "L1",
          kind: "forbid",
          text: "no novel",
          scope: "manuscript",
          check: { type: "regex", pattern: "novel", flags: "i" },
          status: "violations",
          violation_count: 2,
        },
      ],
    },
  });
  await chooseDone(dirty);
  assert.equal(statusSent(dirty).length, 0);
  assert.equal(
    dirty.$("review-message").textContent,
    "R1.1: the ledger check has violations (L1: 2). Choose Done again to mark it done anyway.",
  );
  await chooseDone(dirty);
  assert.equal(statusSent(dirty).length, 1);

  const clean = await setup({ revision: { comments, ledger: [] } });
  await chooseDone(clean);
  assert.equal(statusSent(clean).length, 1);
});

// Codex review of stage 2.5: Propose all did not wait for the daemon to
// confirm the stored provider, as every other action does. Its request then
// timed out in the pane while the daemon still waited at its gate, and the
// turn the author saw fail could start later.
test("Propose all waits for the stored provider to be confirmed before its first request", async () => {
  let confirm;
  let pending = new Promise((r) => (confirm = r));
  const t = await proposeAllSetup({ providersReady: () => pending });
  try {
    const before = t.sent.length;
    const all = t.$("review-propose-all").click();
    await t.wait();
    assert.equal(t.sent.length, before, "nothing is sent while the provider is not confirmed");
    assert.equal(
      t.$("review-message").textContent,
      "Waiting for the daemon to confirm the stored provider choice...",
    );
    // A second press while it waits starts no second run.
    const again = t.$("review-propose-all").click();
    pending = null;
    confirm();
    await t.wait();
    await again;
    assert.deepEqual(t.proposals, [["R1.1", 0]]);
    assert.equal(t.$("review-propose-stop").hidden, false);
    for (const ref of ["R1.1", "R1.3", "R2.1"]) {
      t.endTurn(ref);
      await t.wait();
    }
    await all;
    assert.equal(t.$("review-message").textContent, "Propose all: 3 of 3 proposed.");
    assert.equal(t.proposals.length, 3);
  } finally {
    await t.pane.reload();
  }
});

// The pane says how long it waits for the answer of review_propose, and
// waits exactly that long: the daemon starts no turn for a request the pane
// has given up (review-service.mjs).
test("review_propose carries the time the pane waits for its answer", async () => {
  const t = await proposeAllSetup();
  try {
    const [first] = findAll(t.$("review-comments"), (n) => n.textContent === "Propose changes");
    await first.click();
    const [, payload, opts] = t.sent.find(([type]) => type === "review_propose");
    assert.equal(payload.wait_ms, 10_000);
    assert.deepEqual(opts, { timeoutMs: 10_000 });
    t.endTurn("R1.1");
    await t.wait();
  } finally {
    await t.pane.reload();
  }
});

// The s25 acceptance run: a change Word applied without its rationale
// comment ended a Propose all run as "the changes on record and the tool
// comments in Word do not match (missing R1.1-1 ...)". It still ends the
// run (every change needs its rationale comment), and says what happened.
test("turn_done after a change applied without its rationale comment: the run ends and the line says so", async () => {
  const change = {
    id: "R1.1-1",
    kind: "replace",
    state: "pending",
    find: "a",
    text: "b",
    word_comment_id: null,
    comment_missing: "CommentInsertFailed",
    rationale: { what: "w", why: "y", addresses: "a" },
  };
  const revision = {
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "State the unit.",
        requirements: [],
        status: "in_progress",
        changes: [change],
        blocked: [],
      },
    ],
    ledger: [],
  };
  const counts = { accepted: 0, rejected: 0, user_edited: 0, pending: 1, unlocated: 0 };
  const t = await setup({
    revision,
    respond: { review_refresh: () => ({ ok: true, rows: [], counts, external: [], revision }) },
    word: { collectReviewSnapshot: async () => ({ paragraphs: [], comments: [] }) },
  });
  t.pane.onEvent({ type: "review_event", event: "turn_done", ref: "R1.1", ok: true });
  await settle();
  await new Promise((r) => setTimeout(r, 5));
  assert.equal(
    t.$("review-message").textContent,
    "The turn for R1.1 ran and its record is saved, but R1.1-1 was applied in Word and its rationale comment could not be added; press Add rationale comment on that change, then Refresh status.",
  );
});

test("Add rationale comment: the text comes from the daemon, Word gets the comment, the daemon gets its id", async () => {
  const change = {
    id: "R1.1-1",
    kind: "replace",
    state: "pending",
    find: "degrees",
    text: "degrees Celsius",
    paragraph_id: "U9",
    paragraph_index: 9,
    word_comment_id: null,
    comment_missing: "CommentInsertFailed",
    rationale: { what: "w", why: "y", addresses: "a" },
  };
  const withComment = { ...change, word_comment_id: "3" };
  const rev = (changes) => ({
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "State the unit.",
        requirements: [],
        status: "in_progress",
        changes,
        blocked: [],
      },
    ],
    ledger: [],
  });
  const wordCalls = [];
  let wordAnswer = { ok: true, word_comment_id: "77", existed: false };
  const t = await setup({
    revision: rev([change, withComment]),
    respond: {
      review_comment_text: (p) => ({
        ok: true,
        change_id: p.change_id,
        comment_text: "[R1.1] w / y / a",
        // Comment ids the daemon has on other changes: handed on to Word.
        known_comment_ids: ["31"],
        change: {
          kind: "replace",
          find: "degrees",
          text: "degrees Celsius",
          paragraph_id: "U9",
          paragraph_index: 9,
        },
      }),
      review_comment_set: () => ({
        ok: true,
        revision: rev([{ ...change, word_comment_id: "77", comment_missing: null }, withComment]),
      }),
    },
    word: { addChangeComment: async (a) => (wordCalls.push(a), wordAnswer) },
  });
  const buttons = () =>
    findAll(t.$("review-comments"), (n) => n.textContent === "Add rationale comment");
  assert.equal(buttons().length, 1, "only on the change without a comment");
  const sent = (type) => t.requests.filter((r) => r.type === type).map((r) => r.payload);

  // Word has no place for it: said, and nothing is recorded.
  wordAnswer = { ok: false, code: "CommentAnchorMissing", message: "no place for the comment" };
  await buttons()[0].click();
  assert.match(t.$("review-message").textContent, /no place for the comment/);
  assert.deepEqual(sent("review_comment_set"), []);

  wordAnswer = { ok: true, word_comment_id: "77", existed: false };
  await buttons()[0].click();
  assert.deepEqual(sent("review_comment_text").at(-1), { ref: "R1.1", change_id: "R1.1-1" });
  assert.deepEqual(wordCalls.at(-1), {
    kind: "replace",
    find: "degrees",
    text: "degrees Celsius",
    paragraph_id: "U9",
    paragraph_index: 9,
    comment_text: "[R1.1] w / y / a",
    known_comment_ids: ["31"],
  });
  assert.deepEqual(sent("review_comment_set"), [
    { ref: "R1.1", change_id: "R1.1-1", word_comment_id: "77" },
  ]);
  assert.equal(t.$("review-message").textContent, "R1.1-1: rationale comment added.");
  assert.equal(buttons().length, 0, "the change has its comment now");
});

// Run e2e-20261005-151048: with the Word window minimized, Word applied the
// tracked edits and refused every rationale comment with AccessDenied; the
// pane's page stayed visible, so nothing told the author. Such a change is
// recorded as CommentWaitingForWord; the pane says how many wait and adds
// them itself once Word takes them (on a timer, on focus, before a refresh).
// Fixture: an invented manual of a lighthouse keeper's logbook.
function waitingWorld({ ids = ["R1.1-1", "R1.1-2"], code = "CommentWaitingForWord" } = {}) {
  const changes = ids.map((id) => ({
    id,
    kind: "replace",
    state: "pending",
    find: `lamp ${id}`,
    text: `lantern ${id}`,
    paragraph_index: 2,
    word_comment_id: null,
    comment_missing: code,
    comment_error: "step=insert_comment word=AccessDenied at=Range.insertComment msg=AccessDenied",
    rationale: { what: "w", why: "y", addresses: "a" },
  }));
  const revision = {
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "Name the light by its trade word.",
        requirements: [],
        status: "in_progress",
        changes,
        blocked: [],
      },
    ],
    ledger: [],
  };
  const word = { calls: [], denied: true, allow: Infinity, nextId: 70, gate: null };
  const accessDenied = () =>
    Object.assign(new Error("AccessDenied"), {
      code: "AccessDenied",
      debugInfo: { code: "AccessDenied", errorLocation: "Range.insertComment" },
    });
  const respond = {
    review_comment_text: (p) => ({
      ok: true,
      change_id: p.change_id,
      comment_text: `[R1.1] ${p.change_id}`,
      known_comment_ids: [],
      change: { kind: "replace", find: `lamp ${p.change_id}`, paragraph_index: 2 },
    }),
    // As the daemon does (revision-store.mjs setChangeComment).
    review_comment_set: (p) => {
      const ch = changes.find((c) => c.id === p.change_id);
      if (ch.word_comment_id != null) return { ok: false, error: `${ch.id} has its comment` };
      Object.assign(ch, {
        word_comment_id: p.word_comment_id,
        comment_missing: null,
        comment_error: null,
      });
      return { ok: true, revision };
    },
  };
  const addChangeComment = async (a) => {
    word.calls.push(a.comment_text);
    if (word.gate) await word.gate;
    if (word.denied || word.allow <= 0) throw accessDenied();
    word.allow -= 1;
    return { ok: true, word_comment_id: String(word.nextId++), existed: false };
  };
  return { revision, changes, word, respond, addChangeComment };
}
const waitFor = async (cond, ms = 2000) => {
  const end = Date.now() + ms;
  while (!cond()) {
    if (Date.now() > end) throw new Error("timed out");
    await new Promise((r) => setTimeout(r, 5));
  }
};
const showVisible = () => {
  document.hidden = false;
  for (const fn of document.listeners.visibilitychange ?? []) fn();
};
const pauseMs = (ms) => new Promise((r) => setTimeout(r, ms));

test("comments waiting for Word: the notice counts them, a timer adds them once Word takes them, then it goes", async () => {
  const w = waitingWorld();
  const t = await setup({
    revision: w.revision,
    respond: w.respond,
    word: { addChangeComment: w.addChangeComment },
    options: { commentRetryMs: 15 },
  });
  const notice = t.$("review-comments-waiting");
  const sets = () => t.requests.filter((r) => r.type === "review_comment_set").length;
  assert.equal(notice.hidden, false);
  assert.equal(
    notice.textContent,
    "Word is minimized or not ready: 2 rationale comments are waiting. Open the Word window and they will be added.",
  );
  // Word still refuses: the timer tries, one comment per try (it stops at the
  // first AccessDenied), and nothing is recorded.
  await waitFor(() => w.word.calls.length >= 3);
  assert.equal(sets(), 0);
  assert.ok(w.word.calls.every((c) => c === "[R1.1] R1.1-1"));
  assert.match(notice.textContent, /: 2 rationale comments are waiting/);
  // Word takes one comment, then refuses again: the count goes down by one.
  w.word.allow = 1;
  w.word.denied = false;
  await waitFor(() => sets() === 1);
  await waitFor(() => /: 1 rationale comment is waiting/.test(notice.textContent));
  assert.equal(w.changes[0].word_comment_id, "70");
  assert.equal(w.changes[0].comment_error, null);
  // Word takes comments again: the last one goes in and the notice is gone.
  w.word.allow = Infinity;
  await waitFor(() => notice.hidden === true);
  assert.equal(notice.textContent, "");
  assert.equal(sets(), 2);
  assert.deepEqual(
    w.changes.map((c) => [c.word_comment_id, c.comment_missing]),
    [
      ["70", null],
      ["71", null],
    ],
  );
  // Nothing waits: the timer has stopped and Word is asked nothing more.
  const n = w.word.calls.length;
  await pauseMs(60);
  assert.equal(w.word.calls.length, n);
});

test("comments waiting for Word: the pane shown again adds them at once; a try never runs twice at once", async () => {
  const w = waitingWorld({ ids: ["R1.1-1"] });
  const t = await setup({
    revision: w.revision,
    respond: w.respond,
    word: { addChangeComment: w.addChangeComment },
    options: { commentRetryMs: 1e6 },
  });
  const notice = t.$("review-comments-waiting");
  try {
    showVisible();
    await waitFor(() => w.word.calls.length === 1);
    await pauseMs(5);
    assert.equal(notice.hidden, false, "still refused: still waiting");
    // Word takes it; triggers while Word is adding it make one call.
    w.word.denied = false;
    let open;
    w.word.gate = new Promise((r) => (open = r));
    showVisible();
    showVisible();
    await pauseMs(5);
    const [button] = findAll(
      t.$("review-comments"),
      (n) => n.textContent === "Add rationale comment",
    );
    await button.click(); // the button runs under the same key: refused
    assert.equal(w.word.calls.length, 2);
    open();
    w.word.gate = null;
    await waitFor(() => notice.hidden === true);
    assert.equal(t.requests.filter((r) => r.type === "review_comment_set").length, 1);
    assert.equal(w.word.calls.length, 2);
    showVisible();
    await pauseMs(5);
    assert.equal(w.word.calls.length, 2, "nothing waits any more");
  } finally {
    delete document.hidden;
  }
});

test("comments waiting for Word: Refresh status adds them before it counts", async () => {
  const w = waitingWorld({ ids: ["R1.1-1"] });
  const counts = { accepted: 0, rejected: 0, user_edited: 0, pending: 1, unlocated: 0 };
  const order = [];
  const t = await setup({
    revision: w.revision,
    respond: {
      ...w.respond,
      review_refresh: () => (
        order.push("refresh"),
        { ok: true, rows: [], counts, external: [], revision: w.revision }
      ),
    },
    word: {
      addChangeComment: async (a) => (order.push("comment"), w.addChangeComment(a)),
      collectReviewSnapshot: async () => ({
        paragraphs: [],
        comments: [{ id: "70", content: "[R1.1] R1.1-1" }],
      }),
    },
    options: { commentRetryMs: 1e6 },
  });
  w.word.denied = false;
  await t.$("review-refresh").click();
  assert.deepEqual(order, ["comment", "refresh"]);
  assert.equal(t.$("review-comments-waiting").hidden, true);
  assert.doesNotMatch(t.$("review-message-details").textContent, /MISMATCH|waiting/);
});

test("a waiting comment that fails for another reason is left to its button; no notice for it", async () => {
  const w = waitingWorld({ ids: ["R1.1-1"] });
  let calls = 0;
  const t = await setup({
    revision: w.revision,
    respond: w.respond,
    word: {
      addChangeComment: async () => (
        calls++,
        { ok: false, code: "CommentAnchorMissing", message: "no place for the comment" }
      ),
    },
    options: { commentRetryMs: 10 },
  });
  await waitFor(() => calls === 1);
  await waitFor(() => t.$("review-comments-waiting").hidden === true);
  await pauseMs(50);
  assert.equal(calls, 1, "not tried again by the pane");
  assert.equal(
    findAll(t.$("review-comments"), (n) => n.textContent === "Add rationale comment").length,
    1,
  );
});

test("comments not waiting for Word (CommentInsertFailed) get no notice and no automatic try", async () => {
  const w = waitingWorld({ ids: ["R1.1-1"], code: "CommentInsertFailed" });
  const t = await setup({
    revision: w.revision,
    respond: w.respond,
    word: { addChangeComment: w.addChangeComment },
    options: { commentRetryMs: 10 },
  });
  await pauseMs(50);
  assert.equal(t.$("review-comments-waiting").hidden, true);
  assert.equal(w.word.calls.length, 0);
});

// Run e2e-20261005-154136: the first refresh of the session found nothing
// waiting, and that try, ending without an await, left its finished promise
// stored for good: no later trigger tried the five comments Word refused
// while minimized, also after Word was restored.
test("comments waiting for Word: a try with nothing waiting does not stop later tries", async () => {
  const w = waitingWorld({ ids: ["R1.1-1"], code: null });
  const t = await setup({
    revision: w.revision,
    respond: w.respond,
    word: { addChangeComment: w.addChangeComment },
    options: { commentRetryMs: 1e6 },
  });
  try {
    showVisible(); // nothing waits yet
    await pauseMs(5);
    assert.equal(w.word.calls.length, 0);
    // A later turn: Word refused this comment while minimized.
    w.changes[0].comment_missing = "CommentWaitingForWord";
    w.word.denied = false;
    showVisible();
    await waitFor(() => w.word.calls.length === 1);
    await waitFor(() => t.requests.some((r) => r.type === "review_comment_set"));
    assert.equal(w.changes[0].word_comment_id, "70");
  } finally {
    delete document.hidden;
  }
});

// Each try at a waiting comment goes to the daemon log: which change, what
// started it, what came of it.
test("comments waiting for Word: every try is logged with its trigger and outcome", async () => {
  const w = waitingWorld({ ids: ["R1.1-1", "R1.1-2"] });
  const t = await setup({
    revision: w.revision,
    respond: w.respond,
    word: { addChangeComment: w.addChangeComment },
    options: { commentRetryMs: 1e6 },
  });
  const logged = () =>
    t.requests
      .filter((r) => r.type === "review_comment_retry")
      .map((r) => [r.payload.change_id, r.payload.trigger, r.payload.outcome]);
  try {
    // Still refused: one try, stopped at the first AccessDenied.
    showVisible();
    await waitFor(() => logged().length === 1);
    assert.deepEqual(logged(), [["R1.1-1", "visibility", "AccessDenied"]]);
    // The button of the first holds its key: the next try skips it.
    w.word.denied = false;
    let open;
    w.word.gate = new Promise((r) => (open = r));
    const [button] = findAll(
      t.$("review-comments"),
      (n) => n.textContent === "Add rationale comment",
    );
    const pressed = button.click();
    await pauseMs(5);
    w.word.gate = null; // the try below goes through at once
    await t.pane.addWaitingComments(); // as before a letter export
    open();
    await pressed;
    await waitFor(() => logged().length === 4);
    assert.deepEqual(logged().slice(1), [
      ["R1.1-1", "export", "skipped_busy"],
      ["R1.1-2", "export", "added"],
      ["R1.1-1", "button", "added"],
    ]);
    // An idle tick logs nothing.
    const n = logged().length;
    showVisible();
    await pauseMs(5);
    assert.equal(logged().length, n);
  } finally {
    delete document.hidden;
  }
});

test("a waiting comment that fails for another reason is logged with its code and message", async () => {
  const w = waitingWorld({ ids: ["R1.1-1"] });
  const t = await setup({
    revision: w.revision,
    respond: w.respond,
    word: {
      addChangeComment: async () => ({
        ok: false,
        code: "CommentAnchorMissing",
        message: "no place for the comment",
      }),
    },
    options: { commentRetryMs: 10 },
  });
  await waitFor(() => t.requests.some((r) => r.type === "review_comment_retry"));
  const r = t.requests.find((x) => x.type === "review_comment_retry");
  assert.deepEqual(r.payload, {
    change_id: "R1.1-1",
    trigger: "timer",
    outcome: "CommentAnchorMissing: R1.1-1: no place for the comment",
  });
});

test("turn_done with a comment waiting for Word: the card says it waits, the turn is not unchecked", async () => {
  const w = waitingWorld({ ids: ["R1.1-1"] });
  const counts = { accepted: 0, rejected: 0, user_edited: 0, pending: 1, unlocated: 0 };
  const t = await setup({
    revision: w.revision,
    respond: {
      ...w.respond,
      review_refresh: () => ({ ok: true, rows: [], counts, external: [], revision: w.revision }),
      review_propose: () => ({ ok: true, revision: w.revision }),
    },
    word: {
      addChangeComment: w.addChangeComment,
      collectReviewSnapshot: async () => ({ paragraphs: [], comments: [] }),
    },
    options: { commentRetryMs: 1e6 },
  });
  const [propose] = findAll(t.$("review-comments"), (n) => n.textContent === "Propose changes");
  await propose.click();
  t.pane.onEvent({ type: "review_event", event: "turn_done", ref: "R1.1", ok: true });
  await settle();
  await pauseMs(5);
  const texts = findAll(t.$("review-comments"), (n) =>
    /waiting for the Word window to be opened/.test(n.textContent),
  );
  assert.equal(texts.length, 1);
  assert.match(
    texts[0].textContent,
    /The rationale comment of R1\.1-1 is waiting for the Word window to be opened; it will be added then\./,
  );
  assert.doesNotMatch(t.$("review-message").textContent, /could not be added|press Add rationale/);
  assert.doesNotMatch(t.$("review-message-details").textContent, /MISMATCH/);
});

// Codex review (P2): "Add rationale comment" pressed twice quickly ran twice
// at once: Word got two comments and the daemon refused the second id, which
// left a comment no change has. A button that writes to Word runs once at a
// time; a press while it runs does nothing.
const deferred = () => {
  let resolve;
  const promise = new Promise((r) => (resolve = r));
  return { promise, resolve };
};
const tick = () => new Promise((r) => setTimeout(r, 5));

test("Add rationale comment: a second press while it runs adds no second comment", async () => {
  const change = {
    id: "R1.1-1",
    kind: "replace",
    state: "pending",
    find: "degrees",
    text: "degrees Celsius",
    paragraph_id: "U9",
    paragraph_index: 9,
    word_comment_id: null,
    comment_missing: "CommentInsertFailed",
    rationale: { what: "w", why: "y", addresses: "a" },
  };
  const rev = (changes) => ({
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "State the unit.",
        requirements: [],
        status: "in_progress",
        changes,
        blocked: [],
      },
    ],
    ledger: [],
  });
  let gate = deferred();
  let answer = { ok: false, code: "CommentAnchorMissing", message: "no place for the comment" };
  const wordCalls = [];
  const t = await setup({
    revision: rev([change]),
    respond: {
      review_comment_text: (p) => ({
        ok: true,
        change_id: p.change_id,
        comment_text: "[R1.1] w / y / a",
        known_comment_ids: [],
        change: { kind: "replace", find: "degrees", text: "degrees Celsius", paragraph_index: 9 },
      }),
      review_comment_set: () => ({
        ok: true,
        revision: rev([{ ...change, word_comment_id: "77", comment_missing: null }]),
      }),
    },
    word: {
      addChangeComment: async (a) => {
        wordCalls.push(a);
        await gate.promise;
        return answer;
      },
    },
  });
  const buttons = () =>
    findAll(t.$("review-comments"), (n) => n.textContent === "Add rationale comment");
  const sent = (type) => t.requests.filter((r) => r.type === type).length;

  // A run that fails: the button can be pressed again afterwards.
  const failing = buttons()[0].click();
  await tick();
  assert.equal(buttons()[0].disabled, true, "disabled while it runs");
  gate.resolve();
  await failing;
  assert.match(t.$("review-message").textContent, /no place for the comment/);
  assert.equal(buttons()[0].disabled, false, "enabled again after a failed run");
  assert.equal(wordCalls.length, 1);

  // Two presses, the second while Word is still adding the comment.
  gate = deferred();
  answer = { ok: true, word_comment_id: "77", existed: false };
  const first = buttons()[0].click();
  await tick();
  const second = buttons()[0].click();
  await tick();
  assert.equal(wordCalls.length, 2, "one more call to Word, not two");
  gate.resolve();
  await Promise.all([first, second]);
  assert.equal(wordCalls.length, 2);
  assert.equal(sent("review_comment_text"), 2);
  assert.equal(sent("review_comment_set"), 1);
  assert.equal(t.$("review-message").textContent, "R1.1-1: rationale comment added.");
  assert.equal(buttons().length, 0);
});

test("Refresh status, Clean up tool comments and Delete all run once at a time", async () => {
  const counts = { accepted: 0, rejected: 0, user_edited: 0, pending: 0, unlocated: 0 };
  const revision = { comments: [], ledger: [] };
  const gates = { snapshot: deferred(), clean: deferred(), read: deferred() };
  const calls = { snapshot: 0, clean: 0, read: 0 };
  const t = await setup({
    revision,
    respond: {
      review_refresh: () => ({ ok: true, rows: [], counts, external: [], revision }),
    },
    word: {
      collectReviewSnapshot: async () => {
        calls.snapshot += 1;
        await gates.snapshot.promise;
        return { paragraphs: [], comments: [] };
      },
      deleteToolComments: async () => {
        calls.clean += 1;
        await gates.clean.promise;
        return { deleted_ids: [], before: [], after: [] };
      },
      readAllComments: async () => {
        calls.read += 1;
        await gates.read.promise;
        return [];
      },
    },
  });
  for (const [id, name] of [
    ["review-refresh", "snapshot"],
    ["review-clean-tool", "clean"],
    ["review-delete-all", "read"],
  ]) {
    const first = t.$(id).click();
    await tick();
    assert.equal(t.$(id).disabled, true, `${id} is disabled while it runs`);
    const second = t.$(id).click();
    await tick();
    assert.equal(calls[name], 1, `${id}: the second press does nothing`);
    gates[name].resolve();
    await Promise.all([first, second]);
    assert.equal(t.$(id).disabled, false, `${id} is enabled again`);
    assert.equal(calls[name], 1);
  }
  assert.equal(t.requests.filter((r) => r.type === "review_refresh").length, 1);
});

test("Propose all says what the turns ran on, when the pane knows it", async () => {
  const roles = [];
  const t = await proposeAllSetup({
    usedLabel: (role) => (roles.push(role), "Claude · Opus 5.5"),
  });
  try {
    const all = t.$("review-propose-all").click();
    for (const ref of ["R1.1", "R1.3", "R2.1"]) {
      await t.wait();
      t.endTurn(ref);
      await t.wait();
    }
    await all;
    assert.equal(
      t.$("review-message").textContent,
      "Propose all with Claude · Opus 5.5: 3 of 3 proposed.",
    );
    assert.deepEqual([...new Set(roles)], ["turn"]);
  } finally {
    await t.pane.reload();
  }
});

// The author pressed Propose changes on one comment: when its turn is over
// the tab says what came of it and what it ran on, in front of the status.
test("a single Propose changes ends with a line: how many changes, for which comment, with what", async () => {
  const t = await proposeAllSetup({
    usedLabel: (role) => (role === "turn" ? "Claude · Opus 5.5" : null),
  });
  const buttons = () => findAll(t.$("review-comments"), (n) => n.textContent === "Propose changes");
  try {
    // No change came of the turn.
    await buttons()[0].click();
    await t.wait();
    assert.deepEqual(
      t.proposals.map((p) => p[0]),
      ["R1.1"],
    );
    t.endTurn("R1.1");
    await t.wait();
    await t.wait();
    assert.equal(
      resultOf(t, "R1.1"),
      "No change was proposed for R1.1 with Claude · Opus 5.5: the reason is below.",
    );
    assert.match(t.$("review-message-details").textContent, /^Accepted 0, /);
    // Two changes on record after the turn, each with its comment in Word.
    await buttons()[0].click();
    await t.wait();
    const c = t.revision.comments[0];
    c.changes.push(
      {
        id: "R1.1-1",
        word_comment_id: "w1",
        state: "pending",
        kind: "replace",
        before: "b",
        after: "a",
        rationale: { what: "w", why: "y", addresses: "a", requirement_ids: [] },
      },
      {
        id: "R1.1-2",
        word_comment_id: "w2",
        state: "pending",
        kind: "replace",
        before: "b",
        after: "a",
        rationale: { what: "w", why: "y", addresses: "a", requirement_ids: [] },
      },
    );
    t.doc.comments.push(
      { id: "w1", content: "[R1.1] a / b / c" },
      { id: "w2", content: "[R1.1] d / e / f" },
    );
    t.endTurn("R1.1");
    await t.wait();
    await t.wait();
    assert.equal(resultOf(t, "R1.1"), "Proposed 2 changes for R1.1 with Claude · Opus 5.5.");
    assert.match(t.$("review-message-details").textContent, /^Accepted 0, /);
    // An edit the forbid list blocked is said, also when nothing else came.
    await buttons()[0].click();
    await t.wait();
    c.blocked.push({
      id: "R1.1-b1",
      args: { find: "x", replace: "novel x" },
      blocked_by: ["L2"],
      created_at: "2026-01-01T00:00:00Z",
    });
    t.endTurn("R1.1");
    await t.wait();
    await t.wait();
    assert.equal(
      resultOf(t, "R1.1"),
      "No change was proposed for R1.1 with Claude · Opus 5.5: 1 edit was blocked (see the blocked list of the comment).",
    );
    // A turn that failed gets no such line.
    await buttons()[0].click();
    await t.wait();
    t.endTurn("R1.1", "Claude usage limit reached.");
    await t.wait();
    await t.wait();
    assert.equal(/Proposed|No change was proposed/.test(resultOf(t, "R1.1")), false);
    assert.match(resultOf(t, "R1.1"), /usage limit reached/);
  } finally {
    await t.pane.reload();
  }
});

// What the model asked the author instead of guessing (office_ask_author):
// on the card of the comment, each with a box for the answer.
test("a comment card shows the questions for the author; Save answer sends the answer", async () => {
  const card = (questions) => ({
    ref: "R1.1",
    reviewer: 1,
    text: "The run time does not agree.",
    requirements: [{ id: "q1", text: "make the run time consistent" }],
    status: "in_progress",
    changes: [],
    blocked: [],
    questions,
  });
  const revision = {
    comments: [
      card([
        {
          id: "a1",
          requirement_id: "q1",
          question: "Which run time is right?",
          why: "Two values.",
          answer: "",
          answered_at: null,
        },
        {
          id: "a2",
          requirement_id: null,
          question: "How many flasks?",
          why: "Not said.",
          answer: "12",
          answered_at: "2026-01-01T00:00:00Z",
        },
      ]),
      { ...card([]), ref: "R1.2" },
    ],
    ledger: [],
    snapshot: {
      paragraphs: ["Catalogue", "Shelf marks were checked against the ledger by two clerks."],
      ids: ["k0", "k1"],
    },
  };
  // The paragraph the answer of a2 would go into; a1 names none, and an id
  // the manuscript no longer has shows nothing.
  revision.comments[0].questions[1].paragraph_id = "k1";
  revision.comments[0].questions.push({
    id: "a3",
    requirement_id: null,
    question: "Which year?",
    why: "Not said.",
    answer: "2019",
    answered_at: "2026-01-01T00:00:00Z",
    paragraph_id: "gone",
  });
  const t = await setup({ revision });
  const all = () => findAll(t.$("review-comments"), () => true);
  const texts = all().map((n) => n.textContent);
  const where = all().filter(
    (n) => n.attributes.class === "review-reqs-line review-question-place",
  );
  assert.equal(where.length, 1);
  assert.equal(where[0].textContent, "Where the answer goes:");
  assert.deepEqual(
    findAll(where[0], (n) => /review-paragraph-link/.test(n.attributes.class ?? "")).map(
      (n) => n.textContent,
    ),
    ["Catalogue: \u201cShelf marks were checked against the ledger by...\u201d"],
  );
  assert.ok(texts.includes("a1 (q1): Which run time is right?"));
  assert.ok(texts.includes("Why it is asked: Two values."));
  assert.ok(texts.includes("a2: How many flasks?"));
  const answers = () => all().filter((n) => n.attributes.class === "review-answer");
  assert.deepEqual(
    answers().map((n) => n.value),
    ["", "12", "2019"],
  );
  typeInto(answers()[0], "3,020 hours.");
  // The pane is told a question was noted: it loads again, the draft stays.
  t.pane.onEvent({ type: "review_event", event: "comment_noted", ref: "R1.1" });
  await settle();
  assert.equal(t.requests.filter((r) => r.type === "review_load").length, 2);
  assert.equal(answers()[0].value, "3,020 hours.");
  await all()
    .filter((n) => n.textContent === "Save answer")[0]
    .click();
  assert.deepEqual(t.requests.filter((r) => r.type === "review_question_answer").at(-1).payload, {
    ref: "R1.1",
    id: "a1",
    answer: "3,020 hours.",
  });
});

// A turn that handles one request of three and says nothing of the others:
// the line after the turn and the card both name the two left open.
test("after a turn the line names the requests still open; the card shows the state of each", async () => {
  const t = await proposeAllSetup({
    usedLabel: (role) => (role === "turn" ? "Claude · Opus 5.5" : null),
  });
  const c = t.revision.comments[0];
  c.requirements = [
    { id: "q1", text: "state the unit" },
    { id: "q2", text: "give the n" },
    { id: "q3", text: "name the test" },
  ];
  const buttons = () => findAll(t.$("review-comments"), (n) => n.textContent === "Propose changes");
  // What the reviewer asks: a row per request with its state, per card.
  const lines = () =>
    findAll(t.$("review-comments"), (n) =>
      ["review-req", "review-req-none"].includes(n.attributes.class),
    ).map(textOf);
  try {
    await t.pane.reload();
    assert.deepEqual(lines().slice(0, 4), [
      "state the unitNot yet",
      "give the nNot yet",
      "name the testNot yet",
      "No requests listed yet. Use Edit to add them, one per line; until then a turn lists them itself.",
    ]);
    await buttons()[0].click();
    await t.wait();
    c.changes.push({
      id: "R1.1-1",
      word_comment_id: "w1",
      state: "pending",
      kind: "replace",
      before: "b",
      after: "a",
      rationale: { what: "w", why: "y", addresses: "a", requirement_ids: ["q1"] },
    });
    t.doc.comments.push({ id: "w1", content: "[R1.1] a / b / c" });
    t.endTurn("R1.1");
    await t.wait();
    await t.wait();
    assert.equal(
      resultOf(t, "R1.1"),
      "Proposed 1 change for R1.1 with Claude · Opus 5.5. 2 requests of R1.1 have no change yet: \u201cgive the n\u201d; \u201cname the test\u201d.",
    );
    assert.deepEqual(lines().slice(0, 3), [
      "state the unitChanged",
      "give the nNot yet",
      "name the testNot yet",
    ]);
  } finally {
    await t.pane.reload();
  }
});

// The term a change took out is still in other paragraphs of the last
// status refresh: the card of the comment says where.
test("a comment card names the places where a term its changes took out is still present", async () => {
  const revision = {
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "Gene A is not essential: the claim is not supported.",
        requirements: [],
        status: "in_progress",
        changes: [
          {
            id: "R1.1-1",
            word_comment_id: "w1",
            state: "pending",
            kind: "replace",
            find: "is essential for",
            text: "was associated with",
            narrowed_from: null,
            before: "Gene A is essential for growth.",
            after: "Gene A was associated with growth.",
            rationale: {
              what: "w",
              why: "Essentiality is not shown.",
              addresses: "a",
              requirement_ids: [],
            },
          },
        ],
        blocked: [],
      },
    ],
    ledger: [],
    snapshot: {
      taken_at: "2026-01-01T00:00:00Z",
      paragraphs: [
        "Gene A is essential.",
        "Gene A was associated with growth.",
        "An essential gene.",
      ],
      ids: null,
    },
  };
  const selected = [];
  const t = await setup({
    revision,
    word: { selectParagraphByIndex: async (i) => selected.push(i) },
  });
  const residue = findAll(t.$("review-comments"), (n) => n.attributes.class === "review-residue");
  assert.deepEqual(residue.map(textOf), [
    "\u201cessential\u201d was taken out by this comment's changes. It is still in 2 places:" +
      "\u201cGene A is essential.\u201d\u201cAn essential gene.\u201d",
  ]);
  // Each place is a link that selects the paragraph in Word.
  const links = findAll(residue[0], (n) => /review-paragraph-link/.test(n.attributes.class ?? ""));
  assert.equal(links.length, 2);
  await links[1].click();
  assert.deepEqual(selected, [2]);
});

// "Check revised paragraphs": a status refresh first (the check reads the
// document as it is), then the job; its findings stay under the button.
test("Check revised paragraphs refreshes the status, runs the check and lists what it found", async () => {
  const revision = { comments: [], ledger: [] };
  const consistency = {
    checked_at: "2026-01-02T03:04:05.000Z",
    model: "claude-opus-9",
    paragraphs: [{ index: 2, id: "p2", change_ids: ["R1.1-1", "R2.3-1"], many: false }],
    findings: [
      {
        paragraph_index: 2,
        paragraph_id: "p2",
        kind: "contradiction",
        detail: "It says fast and slow.",
        change_ids: ["R1.1-1"],
      },
    ],
    warnings: [],
  };
  const counts = { accepted: 0, rejected: 0, user_edited: 0, pending: 0, unlocated: 0 };
  const selected = [];
  let t = null;
  t = await setup({
    revision,
    word: {
      collectReviewSnapshot: async () => ({ paragraphs: [], comments: [] }),
      selectParagraphByIndex: async (i) => selected.push(i),
    },
    respond: {
      review_refresh: () => ({ ok: true, rows: [], counts, external: [], revision }),
      review_check_paragraphs: () => {
        queueMicrotask(() =>
          t.pane.onEvent({
            type: "review_event",
            event: "job_done",
            job_id: "j7",
            result: { ok: true, consistency, revision: { ...revision, consistency } },
          }),
        );
        return { ok: true, accepted: true, job_id: "j7" };
      },
    },
  });
  assert.deepEqual(t.$("review-consistency").children, []);
  await t.$("review-check-paragraphs").click();
  assert.deepEqual(
    t.requests.map((r) => r.type).filter((x) => x !== "review_load"),
    ["review_refresh", "review_check_paragraphs"],
  );
  assert.equal(t.$("review-message").textContent, "Checked 1 revised paragraph: 1 problem.");
  const rows = t.$("review-consistency").children;
  assert.deepEqual(
    rows.map((n) => n.textContent),
    ["a paragraph: the paragraph contradicts itself. It says fast and slow."],
  );
  await rows[0].click();
  assert.deepEqual(selected, [2]);
});

// "Rewrite this paragraph" under a paragraph that holds three or more
// changes: a review turn of the comment with the most changes there, told
// which paragraph to write anew.
test("Rewrite this paragraph starts a review turn that carries the paragraph", async () => {
  const t = await proposeAllSetup({});
  t.revision.consistency = {
    checked_at: "2026-01-02T03:04:05.000Z",
    model: null,
    paragraphs: [{ index: 5, id: "p5", change_ids: ["R1.3-1", "R1.3-2", "R1.1-4"], many: true }],
    findings: [],
    warnings: [],
  };
  try {
    await t.pane.reload();
    const buttons = findAll(
      t.$("review-consistency"),
      (n) => n.textContent === "Rewrite this paragraph",
    );
    assert.equal(buttons.length, 1);
    await buttons[0].click();
    await t.wait();
    const sent = t.sent.filter(([type]) => type === "review_propose").at(-1);
    assert.deepEqual([sent[1].ref, sent[1].rewrite_paragraph, sent[1].paragraphs], ["R1.3", 5, []]);
    t.endTurn("R1.3");
    await t.wait();
    await t.wait();
  } finally {
    await t.pane.reload();
  }
});

// "Check other places (N)" on the card of a comment that has changes: a
// second turn of that comment for the places its changes did not touch
// (follow-places.js). Once it ran, the card says so and offers it no more.
test("Check other places starts a follow-through turn of the comment and then shows what it did", async () => {
  const t = await proposeAllSetup({});
  t.revision.snapshot = {
    taken_at: "2026-01-01T00:00:00Z",
    paragraphs: [
      "Deck survey",
      "Rust was reduced on every span after recoating (Fig. 2).",
      "Figure 2. Rust coverage per span.",
    ],
    ids: ["k0", "k1", "k2"],
  };
  const card = t.revision.comments.find((c) => c.ref === "R1.3");
  card.changes = [
    {
      id: "R1.3-1",
      word_comment_id: "w1",
      state: "pending",
      kind: "replace",
      find: "was eliminated",
      text: "was reduced",
      narrowed_from: null,
      paragraph_index: 1,
      paragraph_id: "k1",
      before: "Rust was eliminated on every span after recoating (Fig. 2).",
      after: "Rust was reduced on every span after recoating (Fig. 2).",
      rationale: { what: "w", why: "y", addresses: "a", requirement_ids: [] },
    },
  ];
  t.doc.comments.push({ id: "w1", content: "[R1.3] a / b / c" });
  const buttons = () =>
    findAll(t.$("review-comments"), (n) => /^Check other places/.test(n.textContent));
  const lines = () =>
    findAll(t.$("review-comments"), (n) => /^Other places checked/.test(n.textContent));
  try {
    await t.pane.reload();
    assert.deepEqual(
      buttons().map((b) => b.textContent),
      ["Check other places (2)"],
    );
    assert.equal(lines().length, 0);
    await buttons()[0].click();
    await t.wait();
    const sent = t.sent.filter(([type]) => type === "review_propose").at(-1);
    assert.deepEqual([sent[1].ref, sent[1].follow_through, sent[1].paragraphs], ["R1.3", true, []]);
    // Like every button that starts a turn, it waits while one runs.
    assert.equal(buttons()[0].disabled, true);
    card.follow_through = {
      at: "2026-01-02T00:00:00Z",
      candidate_ids: ["k0", "k2"],
      change_ids: [],
      questions: 0,
      seen_change_ids: ["R1.3-1"],
    };
    t.endTurn("R1.3");
    await t.wait();
    await t.wait();
    assert.equal(buttons().length, 0);
    // The line of the card, and what came of the turn: not the "No change
    // was proposed" of a turn of the comment itself.
    // Twice: the line of the card, and what came of the turn.
    assert.deepEqual(
      lines().map((n) => n.textContent),
      Array(2).fill("Other places checked: 2 (no change needed)."),
    );
    assert.equal(resultOf(t, "R1.3"), "Other places checked: 2 (no change needed).");
    assert.equal(
      findAll(t.$("review-comments"), (n) => /No change was proposed/.test(n.textContent)).length,
      0,
    );
  } finally {
    card.changes = [];
    delete card.follow_through;
    t.doc.comments.length = 0;
    delete t.revision.snapshot;
    await t.pane.reload();
  }
});

test("a deletion inside the author's unaccepted insertion gets one line on its card; others none", async () => {
  const { IN_INSERTION_NOTE } = await import("../taskpane/shared/review-view.js");
  const base = {
    kind: "delete",
    state: "pending",
    text: "",
    word_comment_id: "7",
    rationale: { what: "w", why: "y", addresses: "a" },
  };
  const changes = [
    // Part of the insertion stays: flagged.
    {
      ...base,
      id: "R1.1-1",
      find: "thick ",
      in_insertion: true,
      insertion_left: "New bearings with ",
      insertion_right: "steel plates.",
    },
    // In an insertion, but its place there was not known: no remainders.
    { ...base, id: "R1.1-2", find: "north ", in_insertion: true },
    // Plain text.
    { ...base, id: "R1.1-3", find: "daily ", in_insertion: false },
    // The whole of an insertion the tool made: nothing of it is left.
    {
      ...base,
      id: "R1.1-4",
      find: "Spare bolts.",
      in_insertion: true,
      insertion_left: "",
      insertion_right: "",
    },
    { ...base, id: "R1.1-5", find: "old " },
  ];
  const revision = {
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "Shorten.",
        requirements: [],
        status: "in_progress",
        changes,
        blocked: [],
      },
    ],
    ledger: [],
  };
  const t = await setup({ revision });
  const wraps = findAll(t.$("review-comments"), (n) => n.attributes.class === "review-change-wrap");
  assert.equal(wraps.length, changes.length);
  const lines = wraps.map((w) =>
    findAll(w, (n) => n.attributes.class === "review-flag review-flag-in-insertion").map(
      (n) => n.textContent,
    ),
  );
  assert.deepEqual(lines, [[IN_INSERTION_NOTE], [], [], [], []]);
  assert.match(IN_INSERTION_NOTE, /no deletion mark inside an unaccepted insertion/);
  assert.match(IN_INSERTION_NOTE, /reject that insertion or type the words again/);
});
