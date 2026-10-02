// Figure lane, step 3: from a comment card of the Word pane to the figures
// page and back. The address of the page, the link on a card, the page that
// opens at the manuscript and the comment of the link, and the message that
// tells the pane of a decision taken on the page.
import { test } from "node:test";
import assert from "node:assert/strict";
import { figuresUrl, hasFigureWork } from "../taskpane/shared/review-view.js";
import { fakeRunPy, figureWorld } from "./helpers/figure-world.mjs";
import { figuresPage as openPage } from "./helpers/figures-page.mjs";

const tick = (ms = 5) => new Promise((r) => setTimeout(r, ms));

// ---- the address -------------------------------------------------------------

test("figuresUrl: the daemon's own address, with the document and the comment encoded", () => {
  const doc = "C:\\Users\\me\\My paper #2 & notes\\paper v2.docx";
  const url = figuresUrl("http://127.0.0.1:50123", doc, "R1.1");
  assert.ok(url.startsWith("http://127.0.0.1:50123/figures/?"), url);
  assert.equal(/[ #\\]/.test(url), false, "no raw space, # or backslash");
  // The page reads back exactly what the pane named.
  const read = new URL(url);
  assert.equal(read.pathname, "/figures/");
  assert.equal(read.searchParams.get("doc"), doc);
  assert.equal(read.searchParams.get("ref"), "R1.1");
  assert.equal(read.hash, "");
  // Another origin gives another address: none is written into the helper.
  assert.equal(
    figuresUrl("http://127.0.0.1:47834/", "/m/paper.docx"),
    "http://127.0.0.1:47834/figures/?doc=%2Fm%2Fpaper.docx",
  );
  assert.equal(figuresUrl("http://127.0.0.1:1"), "http://127.0.0.1:1/figures/");
});

test("hasFigureWork: a listed request in another file, or a figure change in any state", () => {
  const requirements = [{ id: "q1", text: "add the control curve" }];
  assert.equal(hasFigureWork({ requirements }), false);
  assert.equal(hasFigureWork({}), false);
  assert.equal(
    hasFigureWork({ requirements, requirements_outside: [{ id: "q1", why: "w" }] }),
    true,
  );
  // A note about a request that is no longer listed does not count.
  assert.equal(
    hasFigureWork({ requirements, requirements_outside: [{ id: "q9", why: "w" }] }),
    false,
  );
  for (const state of ["proposed", "adopted", "rejected"])
    assert.equal(hasFigureWork({ requirements, figure_changes: [{ id: "f", state }] }), true);
});

// ---- the Review tab on a minimal fake DOM (as review-list-dom.test.mjs) --------

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
  scrollIntoView() {}
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
const klass = (n, name) =>
  String(n.attributes?.class ?? "")
    .split(" ")
    .includes(name);
const textOf = (n) => [n.textContent, ...n.children.map(textOf)].join("");

const DOC = "/m/My paper/paper.docx";
const ORIGIN = "http://127.0.0.1:50123";
const comment = (ref, extra = {}) => ({
  ref,
  reviewer: 1,
  text: `Comment ${ref}.`,
  section: "",
  requirements: [{ id: "q1", text: "add the control curve to Figure 2" }],
  status: "in_progress",
  changes: [],
  blocked: [],
  questions: [],
  ...extra,
});
const OUTSIDE = { requirements_outside: [{ id: "q1", why: "A picture file.", at: "t" }] };
const figure = (state) => ({ id: "R1.3-f1", state, requirement_ids: ["q1"] });

async function reviewPane(revision) {
  globalThis.document = fakeDocument();
  const { initReviewPane, openInBrowser } = await import("../taskpane/shared/review-pane.js");
  const sent = [];
  const pane = initReviewPane({
    sendRequest: async (type) => {
      sent.push(type);
      return { ok: true, revision: revision.now, path: "/m/paper.revision.json" };
    },
    word: { collectReviewSnapshot: async () => ({ paragraphs: [], comments: [] }) },
    onTurnStarted: () => () => {},
    isTurnInFlight: () => false,
    docUrl: () => DOC,
    daemonOrigin: () => ORIGIN,
  });
  await pane.reload();
  const $ = (id) => document.getElementById(id);
  const card = (ref) =>
    findAll($("review-comments"), (n) => klass(n, "review-comment")).find(
      (c) => c.attributes["data-ref"] === ref,
    );
  const link = (root) => findAll(root, (n) => klass(n, "review-figures-open"))[0];
  const address = (root) => findAll(root, (n) => klass(n, "review-figures-address"))[0];
  const stateOf = (ref) => textOf(findAll(card(ref), (n) => klass(n, "review-comment-state"))[0]);
  return { pane, sent, $, card, link, address, stateOf, openInBrowser };
}

test("a card has the link to the figures page only with a request in another file or a figure change", async () => {
  const revision = {
    now: {
      comments: [
        comment("R1.1"),
        comment("R1.2", OUTSIDE),
        comment("R1.3", { figure_changes: [figure("rejected")] }),
      ],
      ledger: [],
      external: [],
      snapshot: { paragraphs: [], ids: [] },
    },
  };
  const t = await reviewPane(revision);
  assert.equal(t.link(t.card("R1.1")), undefined, "a comment about the text has no link");
  for (const ref of ["R1.2", "R1.3"]) {
    const a = t.link(t.card(ref));
    assert.equal(a.tagName, "A");
    assert.equal(a.textContent, "Figures and other files");
    // The address comes from the origin and the document the pane was given.
    assert.equal(a.attributes.href, figuresUrl(ORIGIN, DOC, ref));
    assert.equal(a.attributes.target, "_blank");
    // The same address as text, to copy when nothing opens.
    assert.equal(t.address(t.card(ref)).textContent, a.attributes.href);
  }
  assert.ok(t.link(t.card("R1.2")).attributes.href.startsWith(`${ORIGIN}/figures/?doc=%2Fm%2F`));
  // One link for the whole manuscript, without a comment.
  const all = t.link(t.$("review-figures-all"));
  assert.equal(all.attributes.href, figuresUrl(ORIGIN, DOC));
  assert.equal(new URL(all.attributes.href).searchParams.has("ref"), false);
});

test("the link opens through Office where Office has the browser window API, and is a plain link elsewhere", async () => {
  const revision = {
    now: { comments: [comment("R1.2", OUTSIDE)], ledger: [], external: [], snapshot: {} },
  };
  const t = await reviewPane(revision);
  const a = t.link(t.card("R1.2"));
  const press = () => {
    let prevented = 0;
    for (const fn of a.listeners.click) fn({ preventDefault: () => (prevented += 1) });
    return prevented;
  };
  const before = globalThis.Office;
  try {
    // No Office (a browser): the link is followed as it is.
    delete globalThis.Office;
    assert.equal(press(), 0);
    // Office without the requirement set (Word on the web): the same.
    const asked = [];
    const opened = [];
    const office = (supported) => ({
      context: {
        requirements: {
          isSetSupported: (name, version) => {
            asked.push([name, version]);
            return supported;
          },
        },
        ui: { openBrowserWindow: (url) => opened.push(url) },
      },
    });
    globalThis.Office = office(false);
    assert.equal(press(), 0);
    assert.deepEqual(opened, []);
    // Office with it: Office opens the address and the link is not followed.
    globalThis.Office = office(true);
    assert.equal(press(), 1);
    assert.deepEqual(opened, [a.attributes.href]);
    assert.deepEqual(asked.at(-1), ["OpenBrowserWindowApi", "1.1"]);
    // An API that throws leaves the plain link.
    assert.equal(
      t.openInBrowser("http://127.0.0.1:1/figures/", {
        context: {
          requirements: { isSetSupported: () => true },
          ui: {
            openBrowserWindow() {
              throw new Error("no");
            },
          },
        },
      }),
      false,
    );
  } finally {
    if (before === undefined) delete globalThis.Office;
    else globalThis.Office = before;
  }
});

test("figures_changed makes the pane read the review file again; an event it does not know is ignored", async () => {
  const waiting = comment("R1.3", { ...OUTSIDE, figure_changes: [figure("proposed")] });
  const revision = {
    now: { comments: [waiting], ledger: [], external: [], snapshot: {} },
  };
  const t = await reviewPane(revision);
  assert.equal(t.stateOf("R1.3"), "1 figure or file waits for your decision");

  // An event of a newer daemon that this pane has no branch for: no request,
  // no error, the row as it was. (A pane from before figures_changed takes
  // that event the same way.)
  revision.now = {
    ...revision.now,
    comments: [{ ...waiting, figure_changes: [figure("adopted")] }],
  };
  const sentBefore = t.sent.length;
  t.pane.onEvent({ type: "review_event", event: "something_new", ref: "R1.3" });
  await tick();
  assert.equal(t.sent.length, sentBefore);
  assert.equal(t.stateOf("R1.3"), "1 figure or file waits for your decision");
  assert.equal(t.$("review-message").textContent, "");

  // Adopted on the figures page: the row follows without Refresh.
  t.pane.onEvent({ type: "review_event", event: "figures_changed", ref: "R1.3" });
  await tick();
  assert.deepEqual(t.sent.slice(sentBefore), ["review_load"]);
  assert.equal(t.stateOf("R1.3"), "All changes decided");
});

// ---- the daemon tells the pane -----------------------------------------------

test("a change on the figures page is told to the Word pane of that manuscript, and to no other", async () => {
  const w = await figureWorld({ runPy: fakeRunPy() });
  try {
    const other = await w.openPane("/elsewhere/other.docx");
    const events = () => w.wordPane.frames.filter((f) => f.type === "review_event");
    const p = await w.propose({});
    assert.equal(p.status, 200);
    const id = p.json.id;
    await tick(30);
    assert.deepEqual(events(), [{ type: "review_event", event: "figures_changed", ref: "R1.1" }]);

    const post = (name) => w.call(name, { body: { doc: w.doc, ref: "R1.1", id } });
    for (const name of ["adopt", "reopen", "reject", "reopen"]) {
      assert.equal((await post(name)).status, 200, name);
    }
    await tick(30);
    assert.equal(events().length, 5);
    assert.ok(events().every((e) => e.event === "figures_changed" && e.ref === "R1.1"));

    // A request that changed nothing tells nothing: a refused one, a read.
    assert.equal((await post("reopen")).status, 409);
    assert.equal((await w.call("state", { query: { doc: w.doc } })).status, 200);
    assert.equal((await w.call("adopt", { token: null, body: { doc: w.doc } })).status, 401);
    await tick(30);
    assert.equal(events().length, 5);
    assert.deepEqual(other.frames, [], "the pane of another document hears nothing");
    // The list of manuscripts does not give out the names of the panes.
    const open = (await w.call("manuscripts")).json.manuscripts;
    assert.deepEqual(Object.keys(open[0]).sort(), ["doc", "path"]);
  } finally {
    await w.cleanup();
  }
});

// ---- the page itself, on a fake DOM ------------------------------------------

const PAGE_COMMENTS = ["R1.1", "R1.2", "R2.1"].map((ref) => ({
  ref,
  reviewer: Number(ref[1]),
  section: null,
  text: `Comment ${ref}.`,
  requirements: [],
  figures: [],
}));
const OPEN = [
  { doc: "C:\\work\\first.docx", path: "/mnt/c/work/first.docx" },
  { doc: "C:\\work\\My paper #2\\paper.docx", path: "/mnt/c/work/My paper #2/paper.docx" },
];

const figuresPage = (search, { manuscripts = OPEN } = {}) =>
  openPage(search, { manuscripts, comments: PAGE_COMMENTS });

test("the figures page opens at the manuscript and the comment of the link", async () => {
  const t = await figuresPage(`?${new URLSearchParams({ doc: OPEN[1].doc, ref: "R1.2" })}`);
  // The second manuscript of the list is the chosen one, and its state was asked for.
  assert.equal(t.$("manuscript").value, "1");
  const state = t.asked.find((u) => u.includes("/api/state?"));
  assert.equal(new URL(state, "http://x").searchParams.get("doc"), OPEN[1].doc);
  assert.equal(t.$("link-note").hidden, true);
  // The comment of the link is scrolled to, once, and marked; no other is.
  assert.equal(t.cards.length, 3);
  assert.equal(t.card("R1.2").scrolled, 1);
  assert.ok(t.card("R1.2").className.split(" ").includes("linked"));
  for (const ref of ["R1.1", "R2.1"]) {
    assert.equal(t.card(ref).scrolled, 0);
    assert.equal(t.card(ref).className.includes("linked"), false);
  }
  // Nothing but elements in the list (a null would show as the word "null").
  assert.deepEqual(
    t
      .$("comments")
      .children.filter((n) => n.tagName === "#text")
      .map((n) => n.textContent),
    [],
  );
});

test("without a link the page opens as before: the first manuscript, no comment marked", async () => {
  const t = await figuresPage("");
  assert.equal(t.$("manuscript").value, "0");
  assert.equal(t.$("link-note").hidden, true);
  assert.equal(t.cards.filter((c) => c.scrolled || c.className.includes("linked")).length, 0);
});

test("a link to a manuscript that is not open, or to a comment it does not have, says so in plain words", async () => {
  const gone = await figuresPage(
    `?${new URLSearchParams({ doc: "C:\\work\\closed paper.docx", ref: "R1.2" })}`,
  );
  assert.equal(gone.$("link-note").hidden, false);
  assert.match(
    gone.$("link-note").textContent,
    /^The manuscript of this link is not open in a Word pane: closed paper\.docx\. Open it in Word with the Sidewrite pane, then press Reload\. This page shows first\.docx\.$/,
  );
  // The comment of the link is not looked for in another manuscript.
  assert.equal(gone.cards.filter((c) => c.scrolled || c.className.includes("linked")).length, 0);

  const none = await figuresPage(`?${new URLSearchParams({ doc: "/x/p.docx" })}`, {
    manuscripts: [],
  });
  assert.match(none.$("link-note").textContent, /not open in a Word pane: p\.docx\. .* Reload\.$/);
  assert.equal(none.$("empty").hidden, false);

  const noComment = await figuresPage(`?${new URLSearchParams({ doc: OPEN[0].doc, ref: "R9.9" })}`);
  assert.equal(noComment.$("manuscript").value, "0");
  assert.equal(
    noComment.$("link-note").textContent,
    "The comment of this link is not in this manuscript: R9.9.",
  );
  assert.equal(noComment.cards.filter((c) => c.scrolled).length, 0);
});
