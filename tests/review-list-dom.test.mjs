// The Review tab as a list: the tab it opens on, the card that explains the
// first steps, one row per comment with its state in words, one open card at a
// time, the work of a comment inside its card. review-pane.js on a minimal
// fake DOM (as review-pane-dom.test.mjs); the page itself is read as text.
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
    this.title = "";
    this.scrolled = 0;
    const set = new Set();
    this.classList = {
      toggle: (c, on) => ((on ?? !set.has(c)) ? set.add(c) : set.delete(c)),
      contains: (c) => set.has(c),
    };
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
  scrollIntoView() {
    this.scrolled += 1;
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
const textOf = (n) => [n.textContent, ...n.children.map(textOf)].join("");
const klass = (n, name) =>
  String(n.attributes?.class ?? "")
    .split(" ")
    .includes(name);

const change = (id, state) => ({
  id,
  word_comment_id: id,
  state,
  kind: "replace",
  before: "b",
  after: "a",
  rationale: { what: "w", why: "y", addresses: "a", requirement_ids: [] },
});
const comment = (ref, extra = {}) => ({
  ref,
  reviewer: 1,
  text: `Comment ${ref}.\nSecond line of the comment.`,
  section: "",
  requirements: [],
  status: "in_progress",
  changes: [],
  blocked: [],
  questions: [],
  ...extra,
});

async function setup({ comments, ledger = [], external = [] }) {
  globalThis.document = fakeDocument();
  const { initReviewPane } = await import("../taskpane/shared/review-pane.js");
  const revision = { comments, ledger, external, snapshot: { paragraphs: [], ids: [] } };
  let inFlight = false;
  const stopped = [];
  const sendRequest = async (type, payload) => {
    if (type === "review_refresh") {
      return {
        ok: true,
        rows: [],
        external,
        revision,
        counts: { accepted: 0, rejected: 0, user_edited: 0, pending: 0, unlocated: 0 },
      };
    }
    if (type === "review_propose") return { ok: true, revision, started_at: Date.now() };
    return { ok: true, revision, path: "/m/paper.revision.json" };
  };
  const pane = initReviewPane({
    sendRequest,
    word: { collectReviewSnapshot: async () => ({ paragraphs: [], comments: [] }) },
    onTurnStarted() {
      inFlight = true;
      return () => (inFlight = false);
    },
    isTurnInFlight: () => inFlight,
    stopTurn: () => stopped.push(1),
    lastReply: () => "The reviewer asks for data that is not in this file.",
  });
  await pane.reload();
  const $ = (id) => document.getElementById(id);
  const wait = () => new Promise((r) => setTimeout(r, 5));
  const cards = () => findAll($("review-comments"), (n) => klass(n, "review-comment"));
  const card = (ref) => cards().find((c) => c.attributes["data-ref"] === ref);
  const head = (ref) => findAll(card(ref), (n) => klass(n, "review-comment-head"))[0];
  const body = (ref) => findAll(card(ref), (n) => klass(n, "review-comment-body"))[0];
  const stateOf = (ref) => textOf(findAll(card(ref), (n) => klass(n, "review-comment-state"))[0]);
  const part = (ref, name) => findAll(card(ref), (n) => klass(n, name))[0];
  const endTurn = async (ref) => {
    inFlight = false;
    pane.onEvent({ event: "turn_done", ref, ok: true });
    await wait();
    await wait();
  };
  return { pane, revision, $, wait, cards, card, head, body, stateOf, part, endTurn, stopped };
}

const read = (p) => readFile(new URL(p, import.meta.url), "utf8");

test("the pane opens on the Review tab; the Chat tab is still there, not first", async () => {
  const html = await read("../taskpane/word/index.html");
  assert.match(html, /<body data-active-tab="review"/);
  const tabs = [...html.matchAll(/<button class="tab" data-tab="([a-z]+)"/g)].map((m) => m[1]);
  assert.equal(tabs[0], "review");
  assert.ok(tabs.includes("chat") && tabs.indexOf("chat") > 0, tabs.join(","));
  const open = (name) => new RegExp(`<section class="tab-content" data-tab="${name}"( hidden)?>`);
  assert.equal(open("review").exec(html)[1], undefined, "the Review section is shown");
  assert.equal(open("chat").exec(html)[1], " hidden", "the Chat section is not");
  assert.match(html, /<button class="tab" data-tab="review" aria-current="page">/);
  assert.doesNotMatch(html, /data-tab="chat" aria-current="page"/);
});

test("Setup puts the Providers block first", async () => {
  const html = await read("../taskpane/word/index.html");
  const setup = html.slice(html.indexOf('data-tab="setup"'));
  const order = ["<h3>Providers</h3>", "<h3>Workspace</h3>", "<h3>Context files</h3>"].map((h) =>
    setup.indexOf(h),
  );
  assert.ok(
    order.every((i) => i > 0),
    String(order),
  );
  assert.deepEqual(
    [...order].sort((a, b) => a - b),
    order,
  );
});

test("no review yet: the three steps show and everything with nothing to show is hidden", async () => {
  const t = await setup({ comments: [] });
  assert.equal(t.$("review-start").hidden, false);
  for (const id of [
    "review-summary",
    "review-waiting-box",
    "review-comments-section",
    "review-external-box",
    "review-submit-box",
  ]) {
    assert.equal(t.$(id).hidden, true, `${id} is hidden`);
  }
  assert.equal(t.$("review-paste-box").hidden, false, "the paste box is there to paste into");
  const html = await read("../taskpane/word/index.html");
  const start = html.slice(
    html.indexOf('id="review-start"'),
    html.indexOf('id="review-waiting-box"'),
  );
  assert.equal((start.match(/<li>/g) ?? []).length, 3, "three steps");
  assert.match(start, /Paste the review/);
  assert.match(start, /ask for changes/);
  assert.match(start, /Response tab/);
});

test("a saved review: the start card goes, the paste box shrinks to one line and opens again", async () => {
  const t = await setup({
    comments: [comment("R1.1"), comment("R1.2"), comment("R2.1"), comment("E.1")],
  });
  assert.equal(t.$("review-start").hidden, true);
  assert.equal(t.$("review-comments-section").hidden, false);
  assert.equal(t.$("review-ledger-box").hidden, false);
  assert.equal(t.$("review-submit-box").hidden, false);
  assert.equal(t.$("review-paste-box").hidden, true);
  assert.equal(t.$("review-paste-summary").hidden, false);
  assert.equal(
    t.$("review-paste-summary-text").textContent,
    "Review: 4 comments from 2 reviewers and the editor",
  );
  await t.$("review-paste-toggle").click();
  assert.equal(t.$("review-paste-box").hidden, false);
  assert.equal(t.$("review-paste-toggle").textContent, "Hide review text");
  await t.$("review-paste-toggle").click();
  assert.equal(t.$("review-paste-box").hidden, true);
});

test("each comment is a row: its ref, the first line, a state in words; the header counts", async () => {
  const t = await setup({
    comments: [
      comment("R1.1"),
      comment("R1.2", { changes: [change("R1.2-1", "pending"), change("R1.2-2", "pending")] }),
      comment("R1.3", {
        questions: [{ id: "a1", requirement_id: null, question: "q?", why: "w", answer: "" }],
      }),
      comment("R1.4", { changes: [change("R1.4-1", "accepted")] }),
      comment("R1.5", { status: "done" }),
      comment("R1.6", {
        blocked: [{ id: "b1", state: "blocked", blocked_by: ["L1"], args: { replace: "x" } }],
      }),
      comment("R1.7", {
        requirements: [{ id: "q1", text: "elsewhere" }],
        requirements_outside: [{ id: "q1", why: "a file of its own" }],
      }),
    ],
  });
  const row = (ref) => textOf(t.head(ref));
  assert.match(row("R1.1"), /^R1\.1Comment R1\.1\.Not started$/);
  assert.equal(t.stateOf("R1.1"), "Not started");
  assert.equal(t.stateOf("R1.2"), "2 changes to review in Word");
  assert.equal(t.stateOf("R1.3"), "Needs your answer");
  assert.equal(t.stateOf("R1.4"), "All changes decided");
  assert.equal(t.stateOf("R1.5"), "Done");
  assert.equal(t.stateOf("R1.6"), "Needs your decision");
  assert.equal(t.stateOf("R1.7"), "Belongs to another file");
  // The first line only: the second line of a comment is not in the row.
  assert.doesNotMatch(row("R1.1"), /Second line/);
  // 7 comments, 1 done; R1.2, R1.3 and R1.6 need the author.
  assert.equal(t.$("review-summary").textContent, "7 comments, 1 done. 3 need you.");
  assert.equal(t.$("review-summary").hidden, false);
});

test("one card is open at a time; a pressed row opens its card and scrolls it into view", async () => {
  const t = await setup({ comments: [comment("R1.1"), comment("R1.2"), comment("R1.3")] });
  const hiddenOf = () => ["R1.1", "R1.2", "R1.3"].map((r) => t.body(r).hidden);
  assert.deepEqual(hiddenOf(), [true, true, true], "all closed to begin with");
  await t.head("R1.1").click();
  assert.deepEqual(hiddenOf(), [false, true, true]);
  assert.equal(t.head("R1.1").attributes["aria-expanded"], "true");
  assert.equal(t.card("R1.1").scrolled, 1, "the opened card is scrolled into view");
  await t.head("R1.3").click();
  assert.deepEqual(hiddenOf(), [true, true, false], "R1.1 closed when R1.3 opened");
  assert.equal(t.head("R1.1").attributes["aria-expanded"], "false");
  await t.head("R1.3").click();
  assert.deepEqual(hiddenOf(), [true, true, true], "pressing the open row closes it");
});

test("an open card stays open when the tab is drawn again", async () => {
  const t = await setup({ comments: [comment("R1.1"), comment("R1.2")] });
  await t.head("R1.2").click();
  await t.$("review-paste-toggle").click(); // any action ends in a new draw of the list
  assert.deepEqual([t.body("R1.1").hidden, t.body("R1.2").hidden], [true, false]);
});

test("Propose changes: the card itself shows the work, a way to stop, then what came of it", async () => {
  const t = await setup({ comments: [comment("R1.1"), comment("R1.2")] });
  await t.head("R1.1").click();
  const propose = findAll(t.body("R1.1"), (n) => n.textContent === "Propose changes")[0];
  await propose.click();
  await t.wait();
  const working = t.part("R1.1", "review-card-progress");
  assert.equal(working.hidden, false, "the card shows that work is going on");
  assert.match(textOf(t.part("R1.1", "review-card-clock")), /^Working on this comment: \d+:\d\d$/);
  assert.match(t.stateOf("R1.1"), /^Working \d+:\d\d$/);
  assert.equal(t.part("R1.2", "review-card-progress").hidden, true, "the other card shows none");
  assert.equal(t.stateOf("R1.2"), "Not started");
  // A way to stop, inside the card.
  await findAll(working, (n) => n.textContent === "Stop")[0].click();
  assert.equal(t.stopped.length, 1);
  // The turn ends with no change: its card says so, and why.
  await t.endTurn("R1.1");
  assert.equal(t.part("R1.1", "review-card-progress").hidden, true);
  const result = textOf(t.part("R1.1", "review-card-result"));
  assert.match(result, /^No change was proposed for R1\.1/);
  assert.match(result, /The reviewer asks for data that is not in this file\./);
  assert.equal(t.stateOf("R1.1"), "No change proposed");
});

test("Propose all: the bar at the top shows the run and each row keeps its own state", async () => {
  const t = await setup({ comments: [comment("R1.1"), comment("R1.2")] });
  const all = t.$("review-propose-all").click();
  await t.wait();
  assert.equal(t.$("review-progress").hidden, false);
  assert.match(t.$("review-progress").textContent, /^Propose all: 1\/2 R1\.1\. R1\.1 running/);
  assert.match(t.stateOf("R1.1"), /^Working /);
  assert.equal(t.stateOf("R1.2"), "Not started");
  assert.equal(t.$("review-propose-stop").hidden, false);
  await t.endTurn("R1.1");
  await t.wait();
  assert.match(t.$("review-progress").textContent, /^Propose all: 2\/2 R1\.2\./);
  assert.equal(t.stateOf("R1.1"), "No change proposed");
  assert.match(t.stateOf("R1.2"), /^Working /);
  await t.endTurn("R1.2");
  await all;
  assert.equal(t.$("review-progress").hidden, true);
});

test("Waiting for you shows only when something waits; a request of a comment nobody worked on does not count", async () => {
  const idle = await setup({
    comments: [comment("R1.1", { requirements: [{ id: "q1", text: "state the unit" }] })],
  });
  assert.equal(idle.$("review-waiting-box").hidden, true, "nothing waits before the first turn");
  const asked = await setup({
    comments: [
      comment("R1.1", {
        requirements: [{ id: "q1", text: "state the unit" }],
        questions: [
          { id: "a1", requirement_id: "q1", question: "Which unit?", why: "w", answer: "" },
        ],
      }),
    ],
  });
  assert.equal(asked.$("review-waiting-box").hidden, false);
  assert.equal(asked.$("review-waiting-count").textContent, "(1)");
});

test("Changes made outside the tool, the ledger and Before submission carry a count; the first is hidden when empty", async () => {
  const none = await setup({ comments: [comment("R1.1")] });
  assert.equal(none.$("review-external-box").hidden, true);
  const t = await setup({
    comments: [comment("R1.1"), comment("R1.2", { status: "done" })],
    ledger: [
      {
        id: "L1",
        kind: "forbid",
        text: "x",
        scope: "manuscript",
        check: { type: "llm" },
        status: "unchecked",
      },
    ],
    external: [
      { id: "X1", kind: "modified", paragraph_index: 1, before: "a", after: "b", lost_facts: [] },
      { id: "X2", kind: "deleted", paragraph_index: 2, before: "a", after: "", lost_facts: [] },
    ],
  });
  assert.equal(t.$("review-external-box").hidden, false);
  assert.equal(t.$("review-external-count").textContent, "(2)");
  assert.equal(t.$("review-ledger-count").textContent, "(1)");
  assert.equal(t.$("review-submit-count").textContent, "(1 not done)");
});

test("Delete ALL comments is the last thing of its section, under a warning", async () => {
  const html = await read("../taskpane/word/index.html");
  const box = html.slice(
    html.indexOf('id="review-submit-box"'),
    html.indexOf('<details class="review-fold" id="review-advanced"'),
  );
  assert.ok(box.indexOf("review-clean-tool") < box.indexOf("review-delete-all"));
  assert.match(box, /Deletes every comment in the document/);
  assert.ok(
    box.indexOf("Deletes every comment") < box.indexOf("review-delete-all"),
    "the warning comes first",
  );
  // Nothing but the closing tags follows the button.
  assert.match(
    box.slice(box.indexOf("review-delete-all")),
    /^[\s\S]*?<\/button>\s*<\/div>\s*<\/details>\s*$/,
  );
});

test("with no daemon answer Refresh status is still there to press; Propose all and More are not", async () => {
  globalThis.document = fakeDocument();
  const { initReviewPane } = await import("../taskpane/shared/review-pane.js");
  const pane = initReviewPane({
    sendRequest: async () => ({ ok: false, error: "no daemon" }),
    word: {},
    onTurnStarted: () => () => {},
    isTurnInFlight: () => false,
  });
  await pane.reload();
  const $ = (id) => document.getElementById(id);
  assert.equal($("review-comments-section").hidden, false);
  assert.equal($("review-propose-all").hidden, true);
  assert.equal($("review-more").hidden, true);
  assert.equal($("review-message").textContent, "no daemon");
  assert.equal($("review-start").hidden, false);
});
