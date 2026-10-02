// Behavior of the Response tab glue (response-pane.js) on a minimal fake DOM,
// as tests/review-pane-dom.test.mjs does for the Review tab. No Office.js:
// `word` and `sendRequest` are stubs; a job answers { accepted, job_id } and
// its result comes as a response_event.
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
const byClass = (node, cls) =>
  findAll(node, (n) =>
    String(n.attributes.class ?? "")
      .split(" ")
      .includes(cls),
  );
const button = (node, text) =>
  findAll(node, (n) => n.tagName === "BUTTON" && n.textContent === text)[0];
const tick = () => new Promise((r) => setTimeout(r, 5));

const SECTION = "2. Materials and Methods";
const LETTER = {
  location_line: "[Revised in {section}, {lines}]",
  removed_line: "[Removed from {section}, {lines}]",
  line_one: "line {a}",
  line_range: "lines {a}-{b}",
  lang: "English",
  gloss_enabled: false,
  gloss_lang: null,
};
const PROFILE = {
  ok: true,
  name: "generic",
  path: "/repo/profiles/journal_profiles/generic.json",
  profile: {
    rationale_lang: "English",
    colors: { response_body: null, response_revised: null, response_kr: null },
    response_letter: LETTER,
  },
  overrides: {},
  available: ["example-elsevier", "generic"],
  problems: [],
};
const CITATION = {
  change_ids: ["R1.1-1"],
  kind: "revised",
  section: SECTION,
  paragraph_index: 1,
  quote: "The temperature was held at 30 °C.",
  lines: null,
  lines_pdf: null,
  problem: null,
};
const revisionOf = (comments, more = {}) => ({ comments, ledger: [], ...more });
const C1 = (more = {}) => ({
  ref: "R1.1",
  reviewer: 1,
  text: "Please state the unit.",
  requirements: [{ id: "q1", text: "state the unit" }],
  changes: [{ id: "R1.1-1", state: "accepted", rationale: { requirement_ids: ["q1"] } }],
  ...more,
});
const C2 = (more = {}) => ({
  ref: "R1.2",
  reviewer: 1,
  text: "Name the strain.",
  requirements: [],
  changes: [],
  ...more,
});
const SNAPSHOT = { paragraphs: [{ index: 0, id: "P0", text: SECTION }], comments: [] };
const HEADINGS = [{ index: 0, level: 1, text: SECTION }];
const JOBS = new Set([
  "profile_propose",
  "response_draft",
  "response_translate",
  "response_locate",
  "response_check",
  "response_export",
]);

// respond: { <type>: (payload) => result }. A job's result is delivered as
// job_done right after the request is accepted, unless `hold` names the type:
// then it waits for t.finish(type).
async function setup({
  revision = revisionOf([C1(), C2()]),
  respond = {},
  word = {},
  providersReady,
  hold = [],
  missing = [],
  usedLabel,
} = {}) {
  globalThis.document = fakeDocument(missing);
  const { initResponsePane } = await import("../taskpane/shared/response-pane.js");
  const requests = [];
  const wordCalls = [];
  const held = new Map(); // type -> () => deliver
  let n = 0;
  let pane = null;
  const sendRequest = async (type, payload) => {
    requests.push({ type, payload });
    const answer = () =>
      respond[type]
        ? respond[type](payload)
        : type === "profile_get"
          ? PROFILE
          : { ok: true, revision };
    if (!JOBS.has(type)) return answer();
    const job_id = `j${++n}`;
    const deliver = () =>
      pane.onEvent({ type: "response_event", event: "job_done", job_id, result: answer() });
    if (hold.includes(type)) held.set(type, deliver);
    else setTimeout(deliver, 0);
    return { ok: true, accepted: true, job_id };
  };
  pane = initResponsePane({
    sendRequest,
    providersReady,
    usedLabel,
    word: {
      collectReviewSnapshot: async () => (wordCalls.push("snapshot"), SNAPSHOT),
      collectHeadings: async () => (wordCalls.push("headings"), HEADINGS),
      exportRenderPdf: async () => (
        wordCalls.push("pdf"),
        { pdf_base64: "UERG", byte_length: 3, slice_count: 1 }
      ),
      ...word,
    },
  });
  await pane.reload();
  const $ = (id) => document.getElementById(id);
  return {
    pane,
    requests,
    wordCalls,
    $,
    sent: (type) => requests.filter((r) => r.type === type),
    cards: () => byClass($("response-comments"), "response-card"),
    card: (ref) =>
      byClass($("response-comments"), "response-card").find(
        (c) => c.attributes["data-ref"] === ref,
      ),
    message: () => $("response-message").textContent,
    finish: (type) => held.get(type)(),
  };
}

test("reload asks for the profile and the saved revision; a card shows the comment and its answer box", async () => {
  const revision = revisionOf([
    C1({
      response: { text: "We now state the unit.", stale: ["lines_old"], citations: [CITATION] },
    }),
    C2(),
  ]);
  const t = await setup({ revision });
  assert.deepEqual(
    t.requests.map((r) => r.type),
    ["profile_get", "review_load"],
  );
  assert.deepEqual(t.wordCalls, [], "nothing is read from Word on a plain reload");
  assert.equal(t.cards().length, 2);
  const card = t.card("R1.1");
  assert.equal(byClass(card, "response-comment")[0].textContent, "Please state the unit.");
  assert.equal(byClass(card, "response-text")[0].value, "We now state the unit.");
  assert.equal(
    byClass(card, "response-stale")[0].textContent,
    "Line numbers are older than the manuscript",
  );
  assert.equal(byClass(card, "response-citation-line")[0].textContent, `[Revised in ${SECTION}]`);
  assert.equal(
    byClass(card, "response-citation-quote")[0].textContent,
    "The temperature was held at 30 °C.",
  );
  assert.equal(byClass(t.card("R1.2"), "response-text")[0].value, "");
  assert.equal(t.$("response-profile").value, "generic");
  // One thing to do next.
  assert.equal(
    t.$("response-next").textContent,
    "Next: write or draft the 1 missing response (or decline the comment).",
  );
});

test("Excel pane: without #response-root nothing happens", async () => {
  const t = await setup({ missing: ["response-root"] });
  assert.deepEqual(t.requests, []);
  t.pane.onEvent({ type: "response_event", event: "job_done", job_id: "x", result: { ok: true } });
});

test("Draft reads the manuscript, sends response_draft, and the draft lands in the answer box", async () => {
  const drafted = revisionOf([
    C1({ response: { text: "We now state the unit.", stale: [], citations: [CITATION] } }),
    C2(),
  ]);
  const t = await setup({
    respond: {
      response_draft: () => ({ ok: true, revision: drafted, checks: [], warnings: [] }),
    },
  });
  // Text typed and not saved is replaced by the draft that was asked for.
  const box = byClass(t.card("R1.1"), "response-text")[0];
  box.value = "half a sentence";
  await box.fire("input");
  await button(t.card("R1.1"), "Draft").click();
  assert.deepEqual(t.wordCalls, ["snapshot", "headings"]);
  assert.deepEqual(t.sent("response_draft")[0].payload, {
    ref: "R1.1",
    snapshot: SNAPSHOT,
    headings: HEADINGS,
  });
  assert.equal(byClass(t.card("R1.1"), "response-text")[0].value, "We now state the unit.");
  assert.match(t.message(), /^Drafted R1\.1\./);
});

test("a draft that fails says why and leaves the typed text", async () => {
  const t = await setup({
    respond: {
      response_draft: () => ({ ok: false, error: "No model is available for drafting." }),
    },
  });
  const box = byClass(t.card("R1.2"), "response-text")[0];
  box.value = "my own words";
  await box.fire("input");
  await button(t.card("R1.2"), "Draft").click();
  assert.equal(t.message(), "No model is available for drafting.");
  assert.equal(byClass(t.card("R1.2"), "response-text")[0].value, "my own words");
});

test("Save sends the text with the document as it is now", async () => {
  const t = await setup();
  const box = byClass(t.card("R1.2"), "response-text")[0];
  box.value = "The strain is named in Section 2.";
  await box.fire("input");
  await button(t.card("R1.2"), "Save").click();
  assert.deepEqual(t.sent("response_save")[0].payload, {
    ref: "R1.2",
    text: "The strain is named in Section 2.",
    shown: [], // this card shows no citation
    snapshot: SNAPSHOT,
    headings: HEADINGS,
  });
  assert.equal(t.message(), "Saved R1.2.");

  // A card with a citation says which one it shows; the daemon judges the
  // text by that, not by what another pane stored meanwhile.
  const cited = await setup({
    revision: revisionOf([
      C1({ response: { text: "We now state the unit.", stale: [], citations: [CITATION] } }),
      C2(),
    ]),
  });
  await button(cited.card("R1.1"), "Save").click();
  assert.deepEqual(cited.sent("response_save")[0].payload.shown, [
    {
      kind: "revised",
      quote: "The temperature was held at 30 °C.",
      paragraph_index: 1,
      change_ids: ["R1.1-1"],
    },
  ]);

  // The daemon saved it with a mark: the manuscript changed under the answer.
  const stale = revisionOf([
    C1({ response: { text: "x", stale: ["quote_changed"], citations: [CITATION] } }),
    C2(),
  ]);
  const u = await setup({
    respond: {
      response_save: () => ({ ok: true, revision: stale, warnings: ["Read it, then save again."] }),
    },
  });
  await button(u.card("R1.1"), "Save").click();
  assert.equal(u.message(), "Saved R1.1. Read it, then save again.");
  assert.equal(
    byClass(u.card("R1.1"), "response-stale")[0].textContent,
    "Quote changed in the manuscript",
  );
});

test("Draft all missing: only comments without an answer that are not declined, in order, and it stops at a failure", async () => {
  const revision = revisionOf([
    C1({ response: { text: "Done.", stale: [], citations: [] } }),
    C2(),
    C2({ ref: "R1.3", declined: { reason: "Out of scope." } }),
    C2({ ref: "R2.1", reviewer: 2 }),
    C2({ ref: "R2.2", reviewer: 2 }),
  ]);
  const t = await setup({
    revision,
    respond: {
      response_draft: (p) =>
        p.ref === "R2.1" ? { ok: false, error: "the model refused" } : { ok: true, revision },
    },
  });
  await t.$("response-draft-all").click();
  assert.deepEqual(
    t.sent("response_draft").map((r) => r.payload.ref),
    ["R1.2", "R2.1"],
  );
  assert.equal(t.message(), "Stopped at R2.1 after 1 draft: the model refused");

  const ok = await setup({ revision, respond: { response_draft: () => ({ ok: true, revision }) } });
  await ok.$("response-draft-all").click();
  assert.deepEqual(
    ok.sent("response_draft").map((r) => r.payload.ref),
    ["R1.2", "R2.1", "R2.2"],
  );
  assert.equal(ok.message(), "Drafted 3 responses.");

  // With what the drafts ran on, when the pane knows it (model_used events).
  const roles = [];
  const named = await setup({
    revision,
    respond: { response_draft: () => ({ ok: true, revision }) },
    usedLabel: (role) => (roles.push(role), "Claude · Opus 5.5"),
  });
  await named.$("response-draft-all").click();
  assert.equal(named.message(), "Drafted 3 responses with Claude · Opus 5.5.");
  assert.deepEqual([...new Set(roles)], ["writing"]);
  await button(named.card("R1.2"), "Draft").click();
  assert.match(named.message(), /^Drafted R1\.2 with Claude · Opus 5\.5\./);
});

test("Update line numbers sends the PDF Word rendered; a PDF too large for one frame is not sent", async () => {
  const t = await setup({
    respond: {
      response_locate: () => ({
        ok: true,
        revision: revisionOf([C1(), C2()]),
        lines_indexed: 40,
        located: "1/1",
      }),
    },
  });
  await t.$("response-locate").click();
  assert.deepEqual(t.wordCalls, ["snapshot", "headings", "pdf"]);
  assert.deepEqual(t.sent("response_locate")[0].payload, {
    snapshot: SNAPSHOT,
    headings: HEADINGS,
    pdf_base64: "UERG",
  });
  assert.equal(t.message(), "Line numbers updated: 1/1 quotes located.");

  const big = await setup({
    word: {
      exportRenderPdf: async () => ({ pdf_base64: { length: 100 * 1024 * 1024 }, byte_length: 1 }),
    },
  });
  await big.$("response-locate").click();
  assert.equal(big.sent("response_locate").length, 0);
  assert.match(big.message(), /too large/);

  const none = await setup({
    respond: {
      response_locate: () => ({
        ok: true,
        revision: revisionOf([C1(), C2()]),
        lines_indexed: 0,
        located: "0/1",
      }),
    },
  });
  await none.$("response-locate").click();
  assert.match(none.message(), /No line numbers were found in the PDF.*Line Numbers/);
});

test("Run checks shows the summary and the rows; a later edit takes the old result off the screen", async () => {
  const checks = [
    { id: "quote_exact", level: "fail", ref: "R1.1", detail: "quote is not in paragraph 3" },
    { id: "undecided_changes", level: "warn", ref: "R1.2", detail: "1 pending" },
  ];
  const t = await setup({
    respond: {
      response_check: () => ({
        ok: true,
        revision: revisionOf([C1(), C2()]),
        checks,
        summary: { fail: 1, warn: 1 },
      }),
    },
  });
  await t.$("response-check").click();
  assert.equal(t.$("response-summary").textContent, "1 failed, 1 warning");
  assert.deepEqual(
    t.$("response-checks").children.map((n) => n.textContent),
    [
      "FAIL R1.1 quote_exact: quote is not in paragraph 3",
      "WARN R1.2 undecided_changes: 1 pending",
    ],
  );
  assert.equal(t.message(), "Checked: 1 failed, 1 warning");
  assert.equal(t.$("response-export").textContent, "Export anyway");
  // Saving an answer changes what was checked.
  await button(t.card("R1.2"), "Save").click();
  assert.equal(t.$("response-summary").textContent, "");
  assert.deepEqual(t.$("response-checks").children, []);
  assert.equal(t.$("response-export").textContent, "Export letter (.docx)");
});

test("Export: held while a check fails, the rows are shown, and the button then offers to export anyway", async () => {
  const checks = [
    { id: "answered", level: "fail", ref: "R1.2", detail: "no response and not declined" },
    { id: "lines", level: "fail", ref: "R1.1", detail: "no line range; update line numbers" },
  ];
  const t = await setup({
    respond: {
      response_export: (p) =>
        p.force
          ? {
              ok: true,
              exported: true,
              path: "/m/paper.response-2.docx",
              wanted: "/m/paper.response.docx",
              checks,
              summary: { fail: 2, warn: 0 },
              revision: revisionOf([C1(), C2()]),
            }
          : { ok: true, exported: false, checks, summary: { fail: 2, warn: 0 } },
    },
  });
  const btn = t.$("response-export");
  assert.equal(btn.textContent, "Export letter (.docx)");
  await btn.click();
  assert.deepEqual(t.sent("response_export")[0].payload, {
    snapshot: SNAPSHOT,
    headings: HEADINGS,
    format: "docx",
    force: false,
  });
  assert.match(t.message(), /2 failed/);
  assert.equal(t.$("response-checks").children.length, 2);
  assert.equal(btn.textContent, "Export anyway");
  assert.equal(t.$("response-output").textContent, "");

  await btn.click();
  assert.equal(t.sent("response_export")[1].payload.force, true);
  // The path is shown, and that the file of the usual name was kept.
  assert.equal(
    t.$("response-output").textContent,
    "Exported under a new name: /m/paper.response-2.docx (/m/paper.response.docx was already there or open in Word and was kept).",
  );
});

test("Export plain text sends format text and lists the files", async () => {
  const t = await setup({
    respond: {
      response_export: () => ({
        ok: true,
        exported: true,
        paths: ["/m/paper.response.R1.txt"],
        wanted: ["/m/paper.response.R1.txt"],
        checks: [],
        summary: { fail: 0, warn: 0 },
      }),
    },
  });
  await t.$("response-export-text").click();
  assert.equal(t.sent("response_export")[0].payload.format, "text");
  assert.equal(t.$("response-output").textContent, "Exported: /m/paper.response.R1.txt");
});

test("a loaded letter is only proposed; Apply to this paper saves the overrides and the earlier answers", async () => {
  const proposal = {
    ok: true,
    overrides: { colors: { response_body: "C00000", response_revised: "0070C0" } },
    evidence: { colors: { C00000: { paragraphs: 5 }, "0070C0": { paragraphs: 9 } } },
    previous: { "R1.1": "We stated the unit." },
  };
  const t = await setup({ respond: { profile_propose: () => proposal } });
  const input = t.$("response-letter-file");
  input.files = [
    { name: "old.docx", size: 3, arrayBuffer: async () => Uint8Array.of(80, 75, 3).buffer },
  ];
  await input.fire("change");
  assert.equal(t.sent("profile_propose")[0].payload.docx_base64, "UEsD");
  assert.deepEqual(
    byClass(t.$("response-proposal"), "response-proposal-row").map((n) => n.textContent),
    [
      "colors.response_body: (none) -> C00000 (5 paragraphs)",
      "colors.response_revised: (none) -> 0070C0 (9 paragraphs)",
    ],
  );
  assert.equal(t.$("response-proposal-apply").hidden, false);
  assert.equal(t.sent("profile_set").length, 0, "nothing is saved before Apply");
  assert.equal(t.sent("response_import_previous").length, 0);
  assert.match(
    t.message(),
    /^Read the letter: 2 settings differ, 1 earlier answer found\. Nothing is saved/,
  );

  await t.$("response-proposal-apply").click();
  assert.deepEqual(t.sent("profile_set")[0].payload, { overrides: proposal.overrides });
  assert.deepEqual(t.sent("response_import_previous")[0].payload, { previous: proposal.previous });
  assert.equal(t.$("response-proposal-apply").hidden, true);
  assert.match(
    t.message(),
    /^Applied to this paper: 2 settings, 1 earlier answer kept for drafting\./,
  );
});

test("a letter that is no docx or too large is not sent", async () => {
  const t = await setup();
  const input = t.$("response-letter-file");
  input.files = [{ name: "old.doc", size: 3, arrayBuffer: async () => new ArrayBuffer(3) }];
  await input.fire("change");
  input.files = [
    { name: "old.docx", size: 80 * 1024 * 1024, arrayBuffer: async () => new ArrayBuffer(1) },
  ];
  await input.fire("change");
  assert.equal(t.sent("profile_propose").length, 0);
  assert.match(t.message(), /too large/);
});

test("Save for this paper sends the profile name and the language form as overrides", async () => {
  const t = await setup({
    respond: {
      profile_set: () => ({ ...PROFILE, name: "example-elsevier", revision: revisionOf([]) }),
    },
  });
  t.$("response-profile").value = "example-elsevier";
  t.$("response-lang").value = " Korean ";
  t.$("response-rationale-lang").value = "";
  t.$("response-gloss").checked = true;
  t.$("response-gloss-lang").value = "Korean";
  await t.$("response-profile-save").click();
  assert.deepEqual(t.sent("profile_set")[0].payload, {
    journal_profile: "example-elsevier",
    overrides: {
      rationale_lang: null,
      response_letter: { lang: "Korean", gloss_enabled: true, gloss_lang: "Korean" },
    },
  });
  assert.equal(t.message(), "Saved for this paper: profile example-elsevier.");
});

test("a profile that cannot be read is told, and another can still be picked", async () => {
  const t = await setup({
    respond: {
      profile_get: () => ({
        ...PROFILE,
        name: "mine",
        path: null,
        profile: null,
        problems: ["profile mine: not found"],
      }),
    },
  });
  assert.equal(t.message(), "Profile problem: profile mine: not found");
  assert.equal(t.$("response-profile").children.length, 2);
  assert.equal(t.cards().length, 2);
});

test("undecided changes are marked on their comment; open requests can be declined with a reason", async () => {
  const revision = revisionOf([
    C1({
      requirements: [
        { id: "q1", text: "state the unit" },
        { id: "q2", text: "give the tolerance" },
      ],
      changes: [
        { id: "R1.1-1", state: "accepted", rationale: { requirement_ids: ["q1"] } },
        { id: "R1.1-2", state: "pending", rationale: { requirement_ids: ["q2"] } },
        { id: "R1.1-3", state: "unlocated", rationale: {} },
      ],
      response: { text: "Done.", stale: [], citations: [], declined_requirements: [] },
    }),
    C2(),
  ]);
  const t = await setup({ revision });
  const card = t.card("R1.1");
  assert.equal(
    byClass(card, "response-undecided")[0].textContent,
    "Decision left in Word: 1 pending, 1 not found. Only accepted changes are cited.",
  );
  assert.equal(byClass(t.card("R1.2"), "response-undecided").length, 0);
  const req = byClass(card, "response-req")[0];
  assert.match(
    req.children[0].textContent,
    /^q2 give the tolerance: no accepted change \(a change is pending in Word\)/,
  );

  // No reason, no decline.
  await button(req, "Decline").click();
  assert.equal(t.sent("response_save").length, 0);
  assert.equal(t.message(), "Write the reason first: it is printed in the letter.");

  const reason = findAll(req, (n) => n.tagName === "INPUT")[0];
  reason.value = "The tolerance is not known.";
  await reason.fire("input");
  await button(byClass(t.card("R1.1"), "response-req")[0], "Decline").click();
  assert.deepEqual(t.sent("response_save")[0].payload, {
    ref: "R1.1",
    declined_requirements: [{ id: "q2", reason: "The tolerance is not known." }],
  });
  assert.equal(t.message(), "Declined q2 of R1.1.");
});

test("Decline this comment needs a reason; a declined comment can be opened again", async () => {
  const t = await setup();
  const card = t.card("R1.2");
  await button(card, "Decline this comment").click();
  assert.equal(t.sent("response_save").length, 0);
  const reason = byClass(card, "response-decline-reason")[0];
  reason.value = "Outside the scope of the study.";
  await reason.fire("input");
  await button(t.card("R1.2"), "Decline this comment").click();
  assert.deepEqual(t.sent("response_save")[0].payload, {
    ref: "R1.2",
    declined: { reason: "Outside the scope of the study." },
  });

  const declined = await setup({
    revision: revisionOf([C1(), C2({ declined: { reason: "Outside the scope." } })]),
  });
  const c = declined.card("R1.2");
  assert.equal(byClass(c, "response-declined")[0].textContent, "Declined: Outside the scope.");
  await button(c, "Undo decline").click();
  assert.deepEqual(declined.sent("response_save")[0].payload, { ref: "R1.2", declined: null });
});

test("every action waits for the daemon to confirm the stored provider", async () => {
  let confirm;
  let pending = new Promise((r) => (confirm = r));
  const t = await setup({
    providersReady: () => pending,
    respond: {
      response_draft: () => ({ ok: false, error: "draft was asked" }),
      response_check: () => ({ ok: false, error: "check was asked" }),
      response_export: () => ({ ok: false, error: "export was asked" }),
    },
  });
  const before = t.requests.length;
  const clicks = [
    button(t.card("R1.1"), "Draft").click(),
    t.$("response-draft-all").click(),
    t.$("response-check").click(),
    t.$("response-export").click(),
    t.$("response-export-text").click(),
    t.$("response-locate").click(),
    t.$("response-refresh").click(),
  ];
  await tick();
  assert.equal(t.message(), "Waiting for the daemon to confirm the stored provider choice...");
  assert.equal(t.requests.length, before, "held, not sent");
  assert.deepEqual(t.wordCalls, []);
  pending = null;
  confirm();
  await Promise.all(clicks);
  assert.ok(t.sent("response_draft").length >= 1);
  assert.equal(t.sent("response_check").length, 1);
});

test("no listener of the pane goes around the gate (source)", async () => {
  const src = await readFile(
    new URL("../taskpane/shared/response-pane.js", import.meta.url),
    "utf8",
  );
  const listeners = [...src.matchAll(/addEventListener\(\s*"(\w+)",\s*([^\n]*)/g)];
  assert.ok(listeners.length >= 8, `${listeners.length} listeners found`);
  for (const [, ev, handler] of listeners) {
    // Typing into a box only keeps the text; everything else is an action.
    if (ev === "input") continue;
    assert.match(handler, /^guard\(/, `${ev}: ${handler}`);
  }
  // The buttons of the cards are made by action(), whose listener is guarded
  // (counted above), and nothing else in the file sends a request on its own.
  assert.equal((src.match(/\.on(click|change)\s*=/g) ?? []).length, 0);
});

test("reconnect while a job runs: the wait ends with a message, the buttons work again, and the state is asked for anew", async () => {
  const drafted = revisionOf([
    C1({ response: { text: "We now state the unit.", stale: [], citations: [] } }),
    C2(),
  ]);
  let revision = revisionOf([C1(), C2()]);
  const t = await setup({
    hold: ["response_draft"],
    respond: {
      review_load: () => ({ ok: true, revision }),
      response_state: () => ({ ok: true, revision }),
      response_draft: () => ({ ok: true, revision: drafted }),
    },
  });
  const clicked = button(t.card("R1.1"), "Draft").click();
  await tick();
  assert.equal(t.message(), "Drafting R1.1...");
  assert.equal(t.$("response-check").disabled, true, "one job at a time");

  // The socket dropped and came back (taskpane.js calls reload on welcome).
  // The daemon finished the draft meanwhile and saved it.
  revision = drafted;
  const before = t.requests.length;
  await t.pane.reload();
  await clicked;
  assert.deepEqual(
    t.requests.slice(before).map((r) => r.type),
    ["profile_get", "review_load", "response_state"],
  );
  assert.match(t.message(), /connection to the daemon was lost/);
  assert.equal(t.$("response-check").disabled, false);
  assert.equal(button(t.card("R1.1"), "Draft").disabled, false);
  assert.equal(byClass(t.card("R1.1"), "response-text")[0].value, "We now state the unit.");
});

test("a job that ends after the reconnect is not lost: its result is shown", async () => {
  const drafted = revisionOf([
    C1({ response: { text: "Late but here.", stale: [], citations: [] } }),
    C2(),
  ]);
  const t = await setup({
    hold: ["response_draft"],
    respond: { response_draft: () => ({ ok: true, revision: drafted }) },
  });
  const clicked = button(t.card("R1.1"), "Draft").click();
  await tick();
  await t.pane.reload();
  await clicked;
  assert.equal(byClass(t.card("R1.1"), "response-text")[0].value, "");
  t.finish("response_draft"); // the same pane key: the daemon's event arrives after all
  assert.equal(byClass(t.card("R1.1"), "response-text")[0].value, "Late but here.");
  assert.match(t.message(), /finished after the connection came back/);
});

// ---- the page and taskpane.js (source) -------------------------------------------
const read = (rel) => readFile(new URL(`../${rel}`, import.meta.url), "utf8");

test("the Word page has the Response tab and every element the pane asks for", async () => {
  const html = await read("taskpane/word/index.html");
  const pane = await read("taskpane/shared/response-pane.js");
  assert.match(html, /<button class="tab" data-tab="response">Response<\/button>/);
  assert.match(html, /<section class="tab-content" data-tab="response" hidden>/);
  const ids = new Set([...pane.matchAll(/["'`](response-[a-z-]+)["'`]/g)].map((m) => m[1]));
  // Ids, not the classes of the cards.
  const asked = [...ids].filter((id) =>
    new RegExp(`\\$\\("${id}"\\)|^\\s+"${id}",$`, "m").test(pane),
  );
  assert.ok(asked.length >= 15, `${asked.length} ids found`);
  for (const id of ["response-root", ...asked])
    assert.equal(html.split(`id="${id}"`).length - 1, 1, `id="${id}" once in index.html`);
  // The Excel page has no Response tab.
  assert.doesNotMatch(await read("taskpane/excel/index.html"), /response-root/);
  // Advanced settings are folded away; the next step is on top.
  assert.match(html, /<details[^>]*id="response-profile-details"/);
  assert.ok(html.indexOf('id="response-next"') < html.indexOf('id="response-profile-details"'));
});

test("taskpane.js starts the pane behind the provider gate and hands it its events and reconnects", async () => {
  const src = await read("taskpane/shared/taskpane.js");
  assert.match(src, /import \{ initResponsePane \} from "\.\/response-pane\.js";/);
  const at = src.indexOf("responsePane = initResponsePane({");
  assert.ok(at > 0);
  const block = src.slice(at, src.indexOf("});", at));
  assert.match(block, /providersReady: \(\) => providersPane\?\.whenRestored\(\) \?\? null/);
  for (const fn of ["collectReviewSnapshot", "collectHeadings", "exportRenderPdf"])
    assert.match(block, new RegExp(`\\b${fn}\\b`), fn);
  assert.match(src, /case "response_event":\s*responsePane\?\.onEvent\(msg\);/);
  // On welcome, next to the Review tab's reload.
  assert.match(src, /reviewPane\?\.reload\(\);\s*responsePane\?\.reload\(\);/);
});

test("the strings of the Response tab are English", async () => {
  for (const rel of ["taskpane/shared/response-pane.js", "taskpane/shared/response-view.js"])
    assert.doesNotMatch(await read(rel), /[ᄀ-ᇿ㄰-㆏가-힣]/, rel);
  const html = await read("taskpane/word/index.html");
  const section = html.slice(
    html.indexOf('data-tab="response" hidden'),
    html.indexOf("<!-- Setup tab content -->"),
  );
  assert.doesNotMatch(section, /[ᄀ-ᇿ㄰-㆏가-힣]/);
});

test("an answer without its translation says so on the card; Translate asks for one", async () => {
  const withGloss = (on) => ({
    ...PROFILE,
    profile: {
      ...PROFILE.profile,
      response_letter: { ...LETTER, gloss_enabled: on, gloss_lang: on ? "German" : null },
    },
  });
  const answer = (gloss) =>
    C1({ response: { text: "We now state the unit.", gloss, stale: [], citations: [CITATION] } });
  const translated = revisionOf([answer("Wir nennen die Einheit."), C2()]);
  const t = await setup({
    revision: revisionOf([answer(""), C2()]),
    respond: {
      profile_get: () => withGloss(true),
      response_translate: () => ({ ok: true, revision: translated }),
    },
  });
  const card = t.card("R1.1");
  assert.equal(
    byClass(card, "response-gloss-missing")[0].textContent,
    "No German translation of this answer; the letter prints the answer without one.",
  );
  assert.equal(byClass(card, "response-gloss").length, 0);
  // A comment with no answer has nothing to translate.
  assert.equal(byClass(t.card("R1.2"), "response-gloss-missing").length, 0);
  assert.equal(button(t.card("R1.2"), "Translate"), undefined);

  await button(card, "Translate").click();
  assert.deepEqual(t.sent("response_translate")[0].payload, { ref: "R1.1" });
  assert.deepEqual(t.wordCalls, [], "the document is not read for a translation");
  assert.equal(byClass(t.card("R1.1"), "response-gloss")[0].textContent, "Wir nennen die Einheit.");
  assert.equal(byClass(t.card("R1.1"), "response-gloss-missing").length, 0);
  assert.equal(button(t.card("R1.1"), "Translate"), undefined);
  assert.equal(t.message(), "Translated R1.1.");

  // The second language off: no line, no button.
  const off = await setup({ revision: revisionOf([answer(""), C2()]) });
  assert.equal(byClass(off.card("R1.1"), "response-gloss-missing").length, 0);
  assert.equal(button(off.card("R1.1"), "Translate"), undefined);
});

// A request answered by a figure file (the figures page), on the card.
test("a card counts an adopted figure change as the answer, and follows a decision taken on the figures page", async () => {
  const withFigure = (state) =>
    revisionOf([
      C1({
        changes: [],
        figure_changes: [{ id: "R1.1-f1", state, requirement_ids: ["q1"] }],
        response: { text: "Done.", stale: [], citations: [], declined_requirements: [] },
      }),
    ]);
  let now = withFigure("proposed");
  const t = await setup({
    revision: now,
    respond: { review_load: () => ({ ok: true, revision: now }) },
  });
  const open = () => byClass(t.card("R1.1"), "response-req").map((r) => r.children[0].textContent);
  assert.deepEqual(open(), [
    "q1 state the unit: no accepted change (a figure or file waits for your decision on the figures page)",
  ]);

  // Adopted on the figures page: the daemon tells the pane, the card follows.
  now = withFigure("adopted");
  const loads = t.sent("review_load").length;
  t.pane.onEvent({ type: "review_event", event: "figures_changed", ref: "R1.1" });
  await tick();
  assert.equal(t.sent("review_load").length, loads + 1);
  assert.deepEqual(open(), []);
  assert.equal(t.wordCalls.length, 0, "no call into Word for it");

  // Rejected: the request is open again, with nothing waiting.
  now = withFigure("rejected");
  t.pane.onEvent({ type: "review_event", event: "figures_changed", ref: "R1.1" });
  await tick();
  assert.deepEqual(open(), ["q1 state the unit: no accepted change"]);

  // taskpane.js hands the event of the Review tab to this tab as well.
  const src = await read("taskpane/shared/taskpane.js");
  assert.match(
    src,
    /case "review_event":\s*reviewPane\?\.onEvent\(msg\);[^}]*?if \(msg\.event === "figures_changed"\) responsePane\?\.onEvent\(msg\);/,
  );
  // And gives the Review tab the document and the address for its link.
  const at = src.indexOf("reviewPane = initReviewPane({");
  const block = src.slice(at, src.indexOf("word: {", at));
  assert.match(block, /docUrl: \(\) => activeDocUrl,/);
  assert.match(block, /daemonOrigin: \(\) => window\.location\.origin,/);
  const html = await read("taskpane/word/index.html");
  assert.equal(html.split('id="review-figures-all"').length - 1, 1);
});
