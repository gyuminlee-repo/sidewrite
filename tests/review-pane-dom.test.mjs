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
  assert.match(m, /^Accepted 1, rejected 0/);
  assert.match(m, /Outside changes: 1 /);
  assert.match(m, /1 rationale comment\(s\) not marked resolved \(Word: GeneralException;/);
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
    assert.match(t.$("review-message").textContent, /Split by llm: 1 comment/);
    assert.deepEqual(
      t.$("review-violations").children.map((b) => b.textContent),
      ['L1: paragraph 3 "x"'],
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
      /^R1\.1 edit with no answer from Word: "30 °C"$/.test(n.textContent),
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
  const sent = (type) => t.requests.filter((r) => r.type === type).map((r) => r.payload);

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
