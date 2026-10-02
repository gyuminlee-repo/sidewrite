// The words of the Review tab: no internal id in anything the author reads, a
// paragraph named by its section and first words, the requests as a list with
// an Edit link, what waits for the author opens its comment, what is for
// experts stays under Advanced (closed). review-pane.js on a minimal fake DOM.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { paragraphLabel } from "../taskpane/shared/paragraph-label.js";

class FakeEl {
  constructor(tag, id = null) {
    this.tagName = tag.toUpperCase();
    this.id = id;
    this.children = [];
    this.listeners = {};
    this.attributes = {};
    this.textContent = "";
    this.hidden = false;
    this.open = false;
    this.value = "";
    this.checked = false;
    this.disabled = false;
    this.title = "";
    this.classList = { toggle() {}, contains: () => false };
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
  scrollIntoView() {}
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
// Everything a reader can get from an element: its text, its tooltip, its attributes.
const readable = (n) =>
  [n.textContent, n.title, ...Object.values(n.attributes), ...n.children.map(readable)].join("\n");
const klass = (n, name) =>
  String(n.attributes?.class ?? "")
    .split(" ")
    .includes(name);

const GUID = /[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}/i;
const ID = (n) => `2AD85338-6AF0-45DD-B6BC-359DEDDAAF${String(n).padStart(2, "0")}`;
const PARAS = [
  "Abstract",
  "Cultures of the strain were grown in a defined mineral medium in sealed flasks.",
  "Materials and Methods",
  "The incubation temperature was 30 degrees and the cell-free extract was prepared by centrifugation.",
  "Results",
  "The extract showed a clear increase in activity after the clarified lysate was added.",
];
const snapshot = { paragraphs: PARAS, ids: PARAS.map((_, i) => ID(i)) };

const change = (over = {}) => ({
  id: "R1.1-1",
  word_comment_id: "70",
  state: "pending",
  kind: "replace",
  find: "cell-free extract",
  text: "clarified lysate",
  before: "b",
  after: "a",
  rationale: { what: "w", why: "y", addresses: "a", requirement_ids: ["q1"] },
  lost_facts: [],
  unsourced_terms: [],
  from_review: [],
  narrowed_from: null,
  ...over,
});

function reviewWith(extra = {}) {
  return {
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        section: "Major comments",
        text: 'The Methods call it a "cell-free extract", but cells remain. Please call it a "clarified lysate".',
        status: "in_progress",
        requirements: [
          { id: "q1", text: "call it a clarified lysate" },
          { id: "q2", text: "name the centrifuge" },
          { id: "q3", text: "add the raw data" },
          { id: "q4", text: "give the speed" },
        ],
        requirements_outside: [{ id: "q3", why: "A file of its own." }],
        questions: [
          { id: "a1", requirement_id: "q4", question: "Which speed?", why: "w", answer: "" },
        ],
        changes: [
          change({
            lost_facts: ["12,000 x g"],
            narrowed_from: { find: "the cell-free extract was", text: "the clarified lysate was" },
          }),
        ],
        blocked: [],
        // The model reported two places by id; only the first is in the snapshot.
        same_term_at: [ID(3), "FFFFFFFF-0000-0000-0000-000000000000"],
      },
    ],
    ledger: [],
    snapshot,
    consistency: {
      paragraphs: [{ index: 3, id: ID(3), change_ids: ["R1.1-1", "R1.1-2", "R1.1-3"], many: true }],
      findings: [
        {
          paragraph_index: 3,
          paragraph_id: ID(3),
          kind: "contradiction",
          detail: "It says fast and slow.",
          change_ids: ["R1.1-1"],
        },
      ],
      warnings: [],
    },
    external: [
      {
        id: "X1",
        kind: "modified",
        paragraph_index: 5,
        old_index: 5,
        before: "The yield was 0.82 and 0.91.",
        after: "The yield was high.",
        lost_facts: ["0.82", "0.91"],
      },
      {
        id: "X2",
        kind: "deleted",
        paragraph_index: null,
        old_index: 1,
        before: "Cultures of the strain were grown in a defined mineral medium.",
        after: "",
        lost_facts: [],
      },
    ],
    ...extra,
  };
}

async function setup(revision, word = {}) {
  globalThis.document = fakeDocument();
  const { initReviewPane } = await import("../taskpane/shared/review-pane.js");
  const sendRequest = async (type) => ({
    ok: true,
    revision,
    path: "/m/paper.revision.json",
    ...(type === "review_refresh"
      ? {
          rows: [],
          external: revision.external,
          counts: { accepted: 0, rejected: 0, user_edited: 0, pending: 1, unlocated: 0 },
        }
      : {}),
  });
  const pane = initReviewPane({
    sendRequest,
    word: { collectReviewSnapshot: async () => ({ paragraphs: PARAS, comments: [] }), ...word },
    onTurnStarted: () => () => {},
    isTurnInFlight: () => false,
  });
  await pane.reload();
  const $ = (id) => document.getElementById(id);
  return { pane, $ };
}

test("no paragraph id appears in anything the author can read", async () => {
  const t = await setup(reviewWith());
  // Show the findings of "Check revised paragraphs" too.
  await t.pane.reload();
  const all = [
    "review-comments",
    "review-consistency",
    "review-external",
    "review-decisions",
    "review-ledger",
    "review-violations",
  ]
    .map((id) => readable(t.$(id)))
    .join("\n");
  assert.ok(all.length > 400, "something was drawn");
  assert.doesNotMatch(all, GUID);
  assert.doesNotMatch(all, /FFFFFFFF-0000/);
});

test("a paragraph is named by its section and first words, as a link that selects it in Word", async () => {
  const selected = [];
  const t = await setup(reviewWith(), { selectParagraphByIndex: async (i) => selected.push(i) });
  const links = findAll(t.$("review-comments"), (n) => klass(n, "review-paragraph-link"));
  const labels = links.map((n) => n.textContent);
  // The reviewer names "cell-free extract": it is still in paragraph 3 and 5? only 3 holds it.
  assert.ok(
    labels.includes(
      "Materials and Methods: “The incubation temperature was 30 degrees and the...”",
    ),
    labels.join(" | "),
  );
  // The id the model reported that the snapshot does not hold is left out.
  assert.equal(links.length >= 1, true);
  await links[0].click();
  assert.equal(selected.length, 1);
  // The findings under "Check revised paragraphs" and the outside changes use the same names.
  const rows = t.$("review-consistency").children.filter((n) => klass(n, "review-violation"));
  assert.match(rows[0].textContent, /^Materials and Methods: “The incubation temperature/);
  assert.match(rows[0].textContent, /the paragraph contradicts itself\. It says fast and slow\.$/);
  const outside = findAll(t.$("review-external"), (n) => klass(n, "review-violation")).map(
    (n) => n.textContent,
  );
  assert.equal(
    outside[0],
    "A paragraph was changed outside the tool: Results: “The extract showed a clear increase in activity...”. The old text had 0.82; 0.91, and the new text does not. Check that this is meant.",
  );
  assert.equal(
    outside[1],
    "A paragraph was deleted outside the tool: “Cultures of the strain were grown in a...”.",
  );
});

test("paragraphLabel: the section above, the first words, a quiet fallback", () => {
  assert.equal(
    paragraphLabel(snapshot, 1).label,
    "Abstract: “Cultures of the strain were grown in a...”",
  );
  // A heading is named by itself, with no section.
  assert.equal(paragraphLabel(snapshot, 2).label, "“Materials and Methods”");
  assert.equal(paragraphLabel(snapshot, 99).label, "a paragraph");
  assert.equal(paragraphLabel(null, 0).label, "a paragraph");
  // A very long heading is cut.
  const long = {
    paragraphs: ["A very long heading that goes on and on and on", "Body text here."],
  };
  assert.match(paragraphLabel(long, 1).label, /^A very long heading that goes\.\.\.: /);
});

test("the requests are a list with a state in words; the box to change them opens from Edit", async () => {
  const t = await setup(reviewWith());
  const rows = findAll(t.$("review-comments"), (n) => klass(n, "review-req")).map(textOf);
  assert.deepEqual(rows, [
    "call it a clarified lysateChanged",
    "name the centrifugeNot yet",
    "add the raw dataIn another file: A file of its own.",
    "give the speedNeeds your answer",
  ]);
  const area = findAll(
    t.$("review-comments"),
    (n) => n.tagName === "TEXTAREA" && klass(n, "review-reqs"),
  )[0];
  const save = findAll(t.$("review-comments"), (n) => n.textContent === "Save requirements")[0];
  assert.deepEqual([area.hidden, save.hidden], [true, true], "closed until Edit");
  await findAll(t.$("review-comments"), (n) => n.textContent === "Edit")[0].click();
  const open = findAll(
    t.$("review-comments"),
    (n) => n.tagName === "TEXTAREA" && klass(n, "review-reqs"),
  )[0];
  assert.equal(open.hidden, false);
  assert.equal(
    open.value,
    "call it a clarified lysate\nname the centrifuge\nadd the raw data\ngive the speed",
  );
  // No code of a request ("q1") in what is shown.
  assert.doesNotMatch(rows.join(" "), /\bq\d\b/);
});

test("Narrowed and Lost become sentences that say what happened and whether to act", async () => {
  const t = await setup(reviewWith());
  const flags = findAll(t.$("review-comments"), (n) => klass(n, "review-flag")).map(
    (n) => n.textContent,
  );
  assert.deepEqual(flags, [
    "The old wording had 12,000 x g. The new wording does not. Reject the change in Word if that is not meant.",
    "Only the words that differ are marked in Word, not the whole phrase Claude gave. Nothing to do.",
  ]);
});

test("Waiting for you: one sentence each, and pressing it opens its comment", async () => {
  const t = await setup(reviewWith());
  assert.equal(t.$("review-waiting-box").hidden, false);
  const rows = t.$("review-decisions").children.filter((n) => klass(n, "review-decision"));
  const texts = rows.map((n) => n.textContent);
  assert.ok(
    texts.includes("R1.1 (Major comments): nothing changes yet for “name the centrifuge”."),
    texts.join(" | "),
  );
  assert.ok(
    texts.includes("R1.1 (Major comments): Claude waits for your answer about “give the speed”."),
  );
  assert.doesNotMatch(texts.join(" "), /\bq\d\b/);
  const head = findAll(t.$("review-comments"), (n) => klass(n, "review-comment-head"))[0];
  const body = findAll(t.$("review-comments"), (n) => klass(n, "review-comment-body"))[0];
  assert.equal(body.hidden, true);
  await rows[0].click();
  assert.equal(body.hidden, false, "its card is open");
  assert.equal(head.attributes["aria-expanded"], "true");
  // Pressing again does not close it: it opens what it is about.
  await rows[0].click();
  assert.equal(body.hidden, false);
});

test("what is for experts is under Advanced, closed to begin with", async () => {
  const html = await readFile(new URL("../taskpane/word/index.html", import.meta.url), "utf8");
  const section = (id) =>
    html.slice(html.indexOf(`id="${id}"`) - 40, html.indexOf(`id="${id}"`) + 40);
  for (const id of ["review-advanced", "review-ledger-advanced"]) {
    assert.match(section(id), /<details[^>]*id="[^"]+"[^>]*>/);
    assert.doesNotMatch(section(id).match(/<details[^>]*>/)[0], /\bopen\b/, `${id} is closed`);
  }
  // The file path, the regex inputs and the check type sit inside those sections.
  const adv = html.slice(html.indexOf('id="review-advanced"'));
  assert.match(adv.slice(0, adv.indexOf("</details>")), /id="review-file"/);
  const ledger = html.slice(
    html.indexOf('id="review-ledger-advanced"'),
    html.indexOf('id="review-ledger"'),
  );
  for (const id of [
    "ledger-pattern",
    "ledger-flags",
    "ledger-check-type",
    "ledger-expect",
    "ledger-suggest",
  ]) {
    assert.ok(ledger.includes(`id="${id}"`), `${id} is under Advanced`);
  }
  // The primary form is the kind, the words and Add.
  const form = html.slice(
    html.indexOf('class="review-ledger-form"'),
    html.indexOf('id="review-ledger-advanced"'),
  );
  for (const id of ["ledger-kind", "ledger-text", "review-ledger-add"])
    assert.ok(form.includes(`id="${id}"`));
  assert.ok(!form.includes("ledger-pattern"));
  // In a card: Preview prompt is under Advanced, and that is closed.
  const t = await setup(reviewWith());
  const adv2 = findAll(
    t.$("review-comments"),
    (n) => n.tagName === "DETAILS" && klass(n, "review-card-advanced"),
  )[0];
  assert.equal(adv2.open, false);
  assert.ok(findAll(adv2, (n) => n.textContent === "Preview prompt").length === 1);
  const stillOut = findAll(t.$("review-comments"), (n) => n.textContent === "Preview prompt");
  assert.equal(stillOut.length, 1, "no other Preview prompt in the card");
});

test("Add drafts the check from the words typed", async () => {
  const added = [];
  globalThis.document = fakeDocument();
  const { initReviewPane } = await import("../taskpane/shared/review-pane.js");
  const revision = { comments: [], ledger: [], external: [], snapshot };
  const pane = initReviewPane({
    sendRequest: async (type, payload) => {
      if (type === "ledger_add") added.push(payload.item);
      return { ok: true, revision, path: "/m/p.json" };
    },
    word: {},
    onTurnStarted: () => () => {},
    isTurnInFlight: () => false,
  });
  await pane.reload();
  const $ = (id) => document.getElementById(id);
  const add = async (kind, text) => {
    $("ledger-kind").value = kind;
    $("ledger-text").value = text;
    $("ledger-check-type").value = "regex";
    $("ledger-pattern").value = "";
    await $("review-ledger-add").click();
  };
  await add("forbid", 'Do not write "novel"');
  await add("instruction", "Keep the tone modest");
  assert.deepEqual(
    added.map((i) => [i.kind, i.check.type, i.check.pattern ?? null]),
    [
      ["forbid", "regex", "novel"],
      ["instruction", "llm", null],
    ],
  );
  assert.equal(
    $("review-message").textContent,
    "Rule added. Claude reads each paragraph to check it.",
  );
});
