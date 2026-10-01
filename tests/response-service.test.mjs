import { test } from "node:test";
import assert from "node:assert/strict";
import {
  mkdir,
  mkdtemp,
  readFile,
  readdir,
  readlink,
  rm,
  symlink,
  writeFile,
} from "node:fs/promises";
import { existsSync } from "node:fs";
import { tmpdir } from "node:os";
import { basename, join } from "node:path";
import { ledgerReply } from "./fixtures/ledger-reply.mjs";
import { createResponseService, RESPONSE_JOBS } from "../daemon/response-service.mjs";
import {
  applyStates,
  loadRevision,
  saveRevision,
  emptyRevision,
  withRevision,
} from "../daemon/revision-store.mjs";
import { classifyRevision } from "../daemon/revision-classify.mjs";
import { markStale, snapshotHash, viewText } from "../daemon/response-quotes.mjs";
import { runPy as realRunPy } from "../daemon/py-tool.mjs";

const ENV = { CC_OFFICE_WSL_PATHS: "0", CC_OFFICE_CONFIG_DIR: "/nonexistent-config" };
const SECTION = "2. Materials and Methods";
const BEFORE =
  "Flasks were shaken. The incubation temperature was held at 30 degrees. Samples were drawn daily.";
const AFTER =
  "Flasks were shaken. The incubation temperature was held at 30 °C. Samples were drawn daily.";
const QUOTE = "The incubation temperature was held at 30 °C.";

// The manuscript as the pane sends it: the change accepted, or still pending.
const snapshotOf = (paragraph = { text: AFTER }) => ({
  paragraphs: [
    { index: 0, id: "P0", text: SECTION, w_ins: 0, w_del: 0 },
    { index: 1, id: "P1", w_ins: 0, w_del: 0, ...paragraph },
  ],
  comments: [],
});
const ACCEPTED = snapshotOf();
const PENDING = snapshotOf({
  text: BEFORE.replace("30 degrees", "30 °C30 degrees"),
  reviewed_text: AFTER,
  w_ins: 1,
  w_del: 1,
});
const HEADINGS = [{ index: 0, level: 1, text: SECTION }];
const CHANGE = {
  id: "R1.1-1",
  word_comment_id: null,
  paragraph_index: 1,
  paragraph_id: "P1",
  kind: "replace",
  find: "30 degrees",
  text: "30 °C",
  before: BEFORE,
  after: AFTER,
  rationale: {
    what: "Unit written as °C.",
    why: "The reviewer asked for it.",
    addresses: "unit",
    requirement_ids: ["q1"],
  },
  context: { left: "was held at ", right: ". Samples were" },
  created_at: "2026-01-01T00:00:00.000Z",
};
const comments = (over = {}) => [
  {
    ref: "R1.1",
    reviewer: 1,
    text: "Please state the unit.",
    requirements: [{ id: "q1", text: "state the unit" }],
    changes: [CHANGE],
    ...over,
  },
];

// Fake Python tools, by script name. `files` makes the docx writer behave as
// the real one does about names that are taken.
function fakePy({ lines = {}, docx = null, verify = null, read = null } = {}) {
  const calls = [];
  const run = async (script, args) => {
    const name = basename(script);
    calls.push([name, args]);
    const arg = (flag) => args[args.indexOf(flag) + 1];
    if (name === "pdf-line-map.py") {
      const requests = JSON.parse(await readFile(args[2], "utf8"));
      const results = requests.map((r) => ({
        id: r.id,
        match:
          lines[r.id] === null
            ? null
            : {
                start: { page: 1, line_no: 9 },
                end: { page: 1, line_no: 9 },
                match_count: 1,
                ...lines[r.id],
              },
      }));
      return { ok: true, code: 0, json: { lines_indexed: 40, monotonic: true, results } };
    }
    if (name === "response_docx.py") {
      if (docx) return docx(args);
      const out = arg("--out");
      if (existsSync(out) && !args.includes("--overwrite"))
        return { ok: false, code: 4, json: { ok: false, code: "exists" } };
      await writeFile(out, "docx");
      const paragraphs = JSON.parse(await readFile(arg("--model"), "utf8")).paragraphs.length;
      return { ok: true, code: 0, json: { ok: true, path: out, paragraphs } };
    }
    if (name === "response_verify.py")
      return verify ?? { ok: true, code: 0, json: { ok: true, checked: {}, failures: [] } };
    if (name === "response_read.py")
      return (
        read ?? { ok: true, code: 0, json: { ok: true, overrides: {}, evidence: {}, previous: {} } }
      );
    throw new Error(`unexpected tool ${name}`);
  };
  return { run, calls };
}

async function setup({ env = ENV, job = null, py = fakePy(), seed = comments(), ...more } = {}) {
  const dir = await mkdtemp(join(tmpdir(), "response-svc-"));
  const doc = join(dir, "paper.docx");
  const revPath = join(dir, "paper.revision.json");
  if (seed)
    await saveRevision(revPath, {
      ...emptyRevision(doc),
      comments: seed,
      ...(more.revision ?? {}),
    });
  const events = [];
  const waiters = new Map();
  const lines = [];
  const pinned = [];
  const svc = createResponseService({
    getContext: () => ({ activeDoc: doc }),
    oneshotJob: async (key, opts) => {
      pinned.push([key, opts ?? null]);
      return job ?? { run: null, model: null };
    },
    sendEvent: (key, obj) => {
      events.push(obj);
      if (obj.event === "job_done") waiters.get(obj.job_id)?.(obj);
    },
    runPy: py.run,
    log: (s) => lines.push(s),
    env,
    ...(more.service ?? {}),
  });
  const ask = (type, payload = {}) =>
    new Promise((resolve) => {
      const reply = (r) => {
        if (!r.accepted) return resolve(r);
        waiters.set(r.job_id, (e) => resolve({ request_id: r.request_id, ...e.result }));
      };
      svc.handlers[type]({ type, request_id: "q1", ...payload }, reply, "k", "word");
    });
  return {
    dir,
    doc,
    revPath,
    svc,
    ask,
    events,
    lines,
    pinned,
    py,
    rev: () => loadRevision(revPath, doc),
    files: async () => (await readdir(dir)).sort(),
    cleanup: () => rm(dir, { recursive: true, force: true }),
  };
}
const WITH_DOC = { snapshot: ACCEPTED, headings: HEADINGS };
const PDF = Buffer.from("%PDF-1.7\nfake\n").toString("base64");
const model = (answer) => {
  const requests = [];
  return {
    requests,
    job: {
      model: "ollama:qwen",
      run: async (req) => {
        requests.push(req);
        return typeof answer === "function" ? answer(req) : answer;
      },
    },
  };
};
const draftOk = { ok: true, json: { response: " We now state the unit. " } };

// A comment answered, located and passing every check.
async function ready(t) {
  await t.ask("response_save", { ref: "R1.1", text: "We now state the unit.", ...WITH_DOC });
  const located = await t.ask("response_locate", { ...WITH_DOC, pdf_base64: PDF });
  assert.equal(located.ok, true, located.error);
}

test("jobs are the requests that call a model or Python; the others answer at once", () => {
  assert.deepEqual([...RESPONSE_JOBS].sort(), [
    "profile_propose",
    "response_check",
    "response_draft",
    "response_export",
    "response_locate",
    "response_translate",
  ]);
});

test("profile_get and profile_set: generic by default, overrides saved with the paper, bad values refused", async () => {
  const t = await setup();
  const got = await t.ask("profile_get");
  assert.deepEqual([got.ok, got.name, got.overrides, got.problems], [true, "generic", {}, []]);
  assert.ok(got.available.includes("generic") && got.available.includes("example-elsevier"));
  assert.equal(got.profile.response_letter.location_line, "[Revised in {section}, {lines}]");

  const set = await t.ask("profile_set", {
    journal_profile: "example-elsevier",
    overrides: { colors: { response_body: "AA0000" }, rationale_lang: "German" },
  });
  assert.deepEqual(
    [set.ok, set.name, set.profile.colors.response_body, set.profile.rationale_lang],
    [true, "example-elsevier", "AA0000", "German"],
  );
  const rev = await t.rev();
  assert.equal(rev.journal_profile, "example-elsevier");
  assert.deepEqual(rev.profile_overrides, {
    colors: { response_body: "AA0000" },
    rationale_lang: "German",
  });

  const red = await t.ask("profile_set", { overrides: { colors: { response_body: "red" } } });
  assert.equal(red.ok, false);
  assert.match(red.error, /colors\.response_body/);
  const missing = await t.ask("profile_set", { journal_profile: "no-such-journal" });
  assert.match(missing.error, /No profile "no-such-journal"/);
  const path = await t.ask("profile_set", { journal_profile: "../x" });
  assert.match(path.error, /profile name/);
  // Nothing of the refused requests was saved.
  assert.deepEqual((await t.rev()).profile_overrides, rev.profile_overrides);
  assert.equal((await t.rev()).journal_profile, "example-elsevier");
  // A null removes a key.
  await t.ask("profile_set", { overrides: { rationale_lang: null } });
  assert.deepEqual((await t.rev()).profile_overrides, { colors: { response_body: "AA0000" } });
  await t.cleanup();
});

test("profile_get names what is wrong with a profile that cannot be read, and still answers", async () => {
  const t = await setup({ revision: { journal_profile: "gone" } });
  const got = await t.ask("profile_get");
  assert.deepEqual([got.ok, got.name, got.profile], [true, "gone", null]);
  assert.match(got.problems[0], /No profile "gone"/);
  assert.ok(got.available.includes("generic"));
  await t.cleanup();
});

test("response_state: states, citations and headings are computed from the snapshot of the pane", async () => {
  const t = await setup();
  // Before the change is accepted nothing is cited.
  let r = await t.ask("response_state", { snapshot: PENDING, headings: HEADINGS });
  assert.equal(r.ok, true, r.error);
  assert.deepEqual(
    [r.revision.comments[0].changes[0].state, r.revision.comments[0].response],
    ["pending", null],
  );
  r = await t.ask("response_state", WITH_DOC);
  const c = r.revision.comments[0];
  assert.equal(c.changes[0].state, "accepted");
  assert.deepEqual(
    c.response.citations.map((x) => [x.kind, x.section, x.quote, x.change_ids, x.lines]),
    [["revised", SECTION, QUOTE, ["R1.1-1"], null]],
  );
  // A heading the snapshot does not hold at that index is not taken from the pane.
  r = await t.ask("response_state", {
    snapshot: ACCEPTED,
    headings: [
      { index: 0, level: 1, text: "A section name the pane made up" },
      { index: 7, level: 1, text: "Nowhere" },
      { index: "0", level: 1, text: SECTION },
    ],
  });
  assert.equal(r.revision.comments[0].response.citations[0].section, null);
  assert.ok(t.lines.some((l) => /\[response\] headings dropped=3/.test(l)));
  const bad = await t.ask("response_state", { headings: HEADINGS });
  assert.deepEqual([bad.ok, /snapshot/.test(bad.error)], [false, true]);
  await t.cleanup();
});

test("response_draft goes to the pane's provider; the answer is saved as the model's, with a receipt", async () => {
  const m = model(draftOk);
  const t = await setup({ job: m.job });
  const r = await t.ask("response_draft", { ref: "R1.1", ...WITH_DOC });
  assert.equal(r.ok, true, r.error);
  // One job handle, asked for with the drafting default for Claude.
  assert.deepEqual(t.pinned, [["k", { claudeModel: "sonnet" }]]);
  assert.equal(m.requests.length, 1);
  assert.equal("model" in m.requests[0], false);
  assert.match(m.requests[0].prompt, /\[R1\.1-1\] replaced <<<30 degrees>>> with <<<30 °C>>>/);
  assert.match(m.requests[0].prompt, /Write in English\./);
  const c = r.revision.comments[0];
  assert.deepEqual(
    [c.response.text, c.response.source, c.response.citations[0].quote, c.response.stale],
    ["We now state the unit.", "model", QUOTE, []],
  );
  const receipts = (await readFile(join(t.dir, "paper.decisions.jsonl"), "utf8"))
    .trim()
    .split("\n")
    .map((l) => JSON.parse(l));
  assert.deepEqual(
    receipts.map((x) => [x.kind, x.model, x.output.ref, x.output.words, x.output.change_ids]),
    [["response_draft", "ollama:qwen", "R1.1", 5, ["R1.1-1"]]],
  );
  // No prose in the receipt file.
  assert.ok(!JSON.stringify(receipts).includes("state the unit"));
  assert.ok(
    t.lines.some((l) =>
      /\[response\] draft R1\.1 lang=English model=ollama:qwen words=5 citations=1/.test(l),
    ),
  );
  // The line range is still missing: the checks of this comment say so.
  assert.deepEqual(
    r.checks.map((x) => [x.id, x.ref]),
    [["lines", "R1.1"]],
  );
  await t.cleanup();
});

test("response_draft: the Claude default can be set by the environment; no model, no draft", async () => {
  const m = model(draftOk);
  const env = { ...ENV, CC_OFFICE_RESPONSE_MODEL: "opus" };
  const t = await setup({ job: m.job, env });
  await t.ask("response_draft", { ref: "R1.1", ...WITH_DOC });
  assert.deepEqual(t.pinned, [["k", { claudeModel: "opus" }]]);
  await t.cleanup();

  const none = await setup();
  const r = await none.ask("response_draft", { ref: "R1.1", ...WITH_DOC });
  assert.deepEqual(
    [r.ok, r.error],
    [false, "No model is available for drafting. Pick a provider in Setup."],
  );
  await none.cleanup();
});

test("response_draft: a refused or malformed answer saves nothing and says why", async () => {
  for (const [answer, why] of [
    [
      { ok: false, error: "No API key for Groq. Enter one in Setup, Providers." },
      /No API key for Groq/,
    ],
    [{ ok: true, json: { reply: "x" } }, /no response text/],
    [{ ok: true, text: "plain words" }, /no response text/],
  ]) {
    const t = await setup({ job: model(answer).job });
    const r = await t.ask("response_draft", { ref: "R1.1", ...WITH_DOC });
    assert.equal(r.ok, false);
    assert.match(r.error, why);
    assert.equal((await t.rev()).comments[0].response?.text ?? "", "");
    assert.equal(existsSync(join(t.dir, "paper.decisions.jsonl")), false);
    await t.cleanup();
  }
  const t = await setup({ job: model(draftOk).job });
  const unknown = await t.ask("response_draft", { ref: "R9.9", ...WITH_DOC });
  assert.match(unknown.error, /No reviewer comment R9\.9/);
  await t.cleanup();
});

test("a comment whose changes are all undecided is drafted as nothing changed and cites nothing", async () => {
  const m = model({ ok: true, json: { response: "We will consider this." } });
  const t = await setup({ job: m.job });
  const r = await t.ask("response_draft", { ref: "R1.1", snapshot: PENDING, headings: HEADINGS });
  assert.equal(r.ok, true, r.error);
  const prompt = m.requests[0].prompt;
  assert.match(prompt, /the only changes you may report\):\n- \(none\)\n/);
  assert.match(
    prompt,
    /\[R1\.1-1\] not decided in Word yet; do not claim: replaced <<<30 degrees>>> with <<<30 °C>>>/,
  );
  const c = r.revision.comments[0];
  assert.deepEqual([c.changes[0].state, c.response.citations], ["pending", []]);
  // The request is neither cited nor declined: the letter cannot go out like this.
  assert.ok(r.checks.some((x) => x.id === "requirements" && /only pending changes/.test(x.detail)));
  await t.cleanup();
});

test("response_draft: a response-scope forbidden expression comes back as a check row", async () => {
  const ledger = [
    {
      id: "L1",
      kind: "forbid",
      text: "no 'unfortunately' in the letter",
      scope: "response",
      check: { type: "regex", pattern: "\\bunfortunately\\b", flags: "i" },
    },
  ];
  const m = model({ ok: true, json: { response: "Unfortunately we had the wrong unit." } });
  const t = await setup({ job: m.job, revision: { ledger } });
  const r = await t.ask("response_draft", { ref: "R1.1", ...WITH_DOC });
  assert.ok(r.checks.some((x) => x.id === "ledger" && x.level === "fail" && x.ref === "R1.1"));
  assert.match(m.requests[0].prompt, /- \[L1\] no 'unfortunately' in the letter/);
  await t.cleanup();
});

test("response_draft writes the translation only when the profile asks for one", async () => {
  const m = model((req) =>
    req.expectJson === false ? { ok: true, text: " Wir nennen die Einheit. " } : draftOk,
  );
  const t = await setup({
    job: m.job,
    revision: {
      profile_overrides: { response_letter: { gloss_enabled: true, gloss_lang: "German" } },
    },
  });
  const r = await t.ask("response_draft", { ref: "R1.1", ...WITH_DOC });
  assert.equal(r.revision.comments[0].response.gloss, "Wir nennen die Einheit.");
  assert.deepEqual(
    m.requests.map((x) => x.expectJson ?? true),
    [true, false],
  );
  assert.match(m.requests[1].system, /Translate into German/);
  assert.equal(m.requests[1].prompt, "We now state the unit.");
  await t.cleanup();
});

test("response_save: the author's text, a declined comment and declined requests", async () => {
  const seed = [
    ...comments({
      requirements: [
        { id: "q1", text: "state the unit" },
        { id: "q2", text: "name the supplier" },
      ],
    }),
    { ref: "E.1", reviewer: 0, text: "Shorten the title." },
  ];
  const t = await setup({ seed });
  let r = await t.ask("response_save", { ref: "R1.1", text: "We state the unit." });
  assert.deepEqual(
    [r.ok, r.revision.comments[0].response.text, r.revision.comments[0].response.source],
    [true, "We state the unit.", "author"],
  );
  r = await t.ask("response_save", {
    ref: "R1.1",
    declined_requirements: [{ id: "q2", reason: "The supplier was not recorded." }],
  });
  assert.deepEqual(r.revision.comments[0].response.declined_requirements, [
    { id: "q2", reason: "The supplier was not recorded." },
  ]);
  assert.equal(r.revision.comments[0].response.text, "We state the unit.");
  r = await t.ask("response_save", {
    ref: "E.1",
    declined: { reason: "The title is at the limit." },
  });
  assert.deepEqual(r.revision.comments[1].declined, { reason: "The title is at the limit." });
  r = await t.ask("response_save", { ref: "E.1", declined: null });
  assert.equal(r.revision.comments[1].declined, null);

  // Closed side: a request that is not there, a blank reason, the same request twice.
  for (const [list, why] of [
    [[{ id: "q9", reason: "x" }], /No request q9 in R1\.1/],
    [[{ id: "q2", reason: "  " }], /needs a reason/],
    [
      [
        { id: "q2", reason: "a" },
        { id: "q2", reason: "b" },
      ],
      /twice/,
    ],
    ["q2", /list/],
  ]) {
    const bad = await t.ask("response_save", { ref: "R1.1", declined_requirements: list });
    assert.equal(bad.ok, false);
    assert.match(bad.error, why);
  }
  assert.match(
    (await t.ask("response_save", { ref: "R7.1", text: "x" })).error,
    /No reviewer comment R7\.1/,
  );
  assert.equal((await t.rev()).comments[0].response.declined_requirements.length, 1);
  // Citations cannot be sent in: they are computed.
  r = await t.ask("response_save", {
    ref: "R1.1",
    text: "Again.",
    citations: [{ quote: "made up" }],
  });
  assert.deepEqual(r.revision.comments[0].response.citations, []);
  await t.cleanup();
});

test("an answer goes stale when the changes it was written against are no longer the accepted ones", async () => {
  const t = await setup();
  // Written while the change was still undecided in Word.
  let r = await t.ask("response_save", {
    ref: "R1.1",
    text: "We will look into the unit.",
    snapshot: PENDING,
    headings: HEADINGS,
  });
  assert.deepEqual(r.revision.comments[0].response.stale, []);
  // The author then accepts the change: the answer no longer fits.
  r = await t.ask("response_state", WITH_DOC);
  assert.deepEqual(r.revision.comments[0].response.stale, ["change_state"]);
  // The mark stays over the next refresh, until the answer is written again.
  r = await t.ask("response_state", WITH_DOC);
  assert.deepEqual(r.revision.comments[0].response.stale, ["change_state"]);
  r = await t.ask("response_save", { ref: "R1.1", text: "We now state the unit.", ...WITH_DOC });
  assert.deepEqual(r.revision.comments[0].response.stale, []);
  // A quoted sentence edited in Word afterwards.
  const edited = snapshotOf({ text: AFTER.replace("was held at", "was kept at") });
  r = await t.ask("response_state", { snapshot: edited, headings: HEADINGS });
  assert.deepEqual(r.revision.comments[0].response.stale, ["quote_changed"]);
  assert.match(r.revision.comments[0].response.citations[0].quote, /was kept at 30 °C/);
  await t.cleanup();
});

test("response_locate: the PDF is kept next to the manuscript and its line numbers land on the citations", async () => {
  const t = await setup();
  await t.ask("response_save", { ref: "R1.1", text: "We now state the unit.", ...WITH_DOC });
  const r = await t.ask("response_locate", { ...WITH_DOC, pdf_base64: PDF });
  assert.equal(r.ok, true, r.error);
  assert.deepEqual(
    [r.lines_indexed, r.located, r.pdf_path],
    [40, "1/1", join(t.dir, "paper.render.pdf")],
  );
  assert.ok((await readFile(r.pdf_path)).equals(Buffer.from(PDF, "base64")));
  const rev = await t.rev();
  assert.deepEqual(rev.comments[0].response.citations[0].lines, {
    start: 9,
    end: 9,
    start_page: 1,
    end_page: 1,
    match_count: 1,
  });
  assert.equal(rev.response_render.snapshot_hash, snapshotHash(ACCEPTED));
  assert.equal(rev.comments[0].response.citations[0].lines_pdf, rev.response_render.pdf_sha);
  assert.match(rev.response_render.pdf_sha, /^[0-9a-f]{64}$/);
  // The request file went to a temp folder that is gone again.
  const [, args] = t.py.calls.find(([name]) => name === "pdf-line-map.py");
  assert.deepEqual([args[0], args[1]], [r.pdf_path, "--quotes-json"]);
  assert.equal(existsSync(args[2]), false);
  assert.ok(
    t.lines.some((l) => /\[response\] locate pdf=\d+ lines_indexed=40 located=1\/1/.test(l)),
  );
  // After the manuscript text changed the range is old until the PDF is read again.
  const edited = snapshotOf({ text: AFTER.replace("daily", "every day") });
  const state = await t.ask("response_state", { snapshot: edited, headings: HEADINGS });
  assert.deepEqual(state.revision.comments[0].response.stale, ["lines_old"]);

  const notPdf = await t.ask("response_locate", {
    ...WITH_DOC,
    pdf_base64: Buffer.from("PK docx").toString("base64"),
  });
  assert.deepEqual([notPdf.ok, /PDF/.test(notPdf.error)], [false, true]);
  await t.cleanup();
});

test("response_check: rows and a summary; an LLM rule of the response scope is asked once for all answers", async () => {
  const t = await setup();
  await ready(t);
  let r = await t.ask("response_check", WITH_DOC);
  assert.deepEqual([r.ok, r.checks, r.summary], [true, [], { fail: 0, warn: 0 }]);
  assert.ok(t.lines.some((l) => l === "[response] check fail=0 warn=0"));
  assert.deepEqual(t.pinned, []); // no LLM rule, no model asked
  await t.cleanup();

  const ledger = [
    {
      id: "L1",
      kind: "instruction",
      text: "Be courteous.",
      scope: "response",
      check: { type: "llm" },
    },
    { id: "L2", kind: "instruction", text: "Manuscript rule.", check: { type: "llm" } },
  ];
  // The model answers in the form the ledger check asks for: a verdict for
  // every line it was given (the answers, one line each).
  const m = model((req) =>
    ledgerReply(req, [{ index: 0, text: "We now state the unit." }], {
      violations: [{ paragraph_index: 0, quote: "state the unit", reason: "curt" }],
    }),
  );
  const u = await setup({ job: m.job, revision: { ledger } });
  await ready(u);
  r = await u.ask("response_check", WITH_DOC);
  assert.equal(m.requests.length, 1);
  assert.match(m.requests[0].prompt, /Be courteous\./);
  assert.deepEqual(
    r.checks.map((x) => [x.id, x.level, x.ref, x.detail]),
    [["ledger", "fail", "R1.1", "L1: state the unit (curt)"]],
  );
  assert.deepEqual(r.summary, { fail: 1, warn: 0 });
  await u.cleanup();

  // No model to ask: the rule is not passed over in silence.
  const n = await setup({ revision: { ledger } });
  await ready(n);
  r = await n.ask("response_check", WITH_DOC);
  assert.deepEqual(
    r.checks.map((x) => [x.id, x.level, x.ref]),
    [["ledger", "fail", null]],
  );
  await n.cleanup();
});

test("response_export holds the letter back while a check fails, and writes it with force", async () => {
  const t = await setup();
  await t.ask("response_save", { ref: "R1.1", text: "We now state the unit.", ...WITH_DOC });
  let r = await t.ask("response_export", { ...WITH_DOC, format: "docx" });
  assert.deepEqual([r.ok, r.exported, r.summary.fail], [true, false, 1]);
  assert.deepEqual(
    r.checks.map((x) => x.id),
    ["lines"],
  );
  assert.deepEqual(await t.files(), ["paper.revision.json"]);
  r = await t.ask("response_export", { ...WITH_DOC, format: "docx", force: true });
  assert.deepEqual([r.exported, basename(r.path)], [true, "paper.response.docx"]);
  assert.deepEqual((await t.rev()).response_export.fail, 1);
  await t.cleanup();
});

test("response_export writes the docx, its model and the manuscript paragraph list, and verifies the file", async () => {
  const t = await setup();
  await ready(t);
  const r = await t.ask("response_export", { ...WITH_DOC, format: "docx" });
  assert.equal(r.ok, true, r.error);
  assert.deepEqual([r.exported, r.verify.ok, r.summary], [true, true, { fail: 0, warn: 0 }]);
  const stem = join(t.dir, "paper.response");
  assert.deepEqual(r.paths, {
    docx: `${stem}.docx`,
    model: `${stem}.model.json`,
    manuscript: `${stem}.manuscript.json`,
  });
  assert.deepEqual(
    JSON.parse(await readFile(r.paths.manuscript, "utf8")),
    ACCEPTED.paragraphs.map(viewText),
  );
  const letter = JSON.parse(await readFile(r.paths.model, "utf8"));
  assert.deepEqual(
    letter.paragraphs.map((p) => [p.kind, p.text]),
    [
      ["reviewer_head", "Reviewer #1"],
      ["comment_head", "Comment 1"],
      ["comment", "Please state the unit."],
      ["response", "We now state the unit."],
      ["location", `[Revised in ${SECTION}, line 9]`],
      ["quote", QUOTE],
    ],
  );
  const tools = t.py.calls.filter(([name]) => name.startsWith("response_"));
  assert.deepEqual(
    tools.map(([name]) => name),
    ["response_docx.py", "response_verify.py"],
  );
  assert.deepEqual(tools[0][1], ["--model", r.paths.model, "--out", r.paths.docx]);
  assert.deepEqual(tools[1][1], [
    "--docx",
    r.paths.docx,
    "--model",
    r.paths.model,
    "--manuscript",
    r.paths.manuscript,
  ]);
  const saved = (await t.rev()).response_export;
  assert.deepEqual([saved.path, saved.fail, saved.warn], [r.paths.docx, 0, 0]);
  assert.ok(
    t.lines.some((l) => l === `[response] export path=${r.paths.docx} paragraphs=6 verify=ok`),
  );
  await t.cleanup();
});

test("response_export never replaces a letter that is there: the next free name, unless overwrite is asked", async () => {
  const t = await setup();
  await ready(t);
  const first = await t.ask("response_export", { ...WITH_DOC, format: "docx" });
  await writeFile(first.path, "edited by the author in Word");
  const second = await t.ask("response_export", { ...WITH_DOC, format: "docx" });
  assert.equal(basename(second.path), "paper.response-2.docx");
  // The usual name is told with the one used, so the pane can say a file was kept.
  assert.deepEqual([first.wanted, second.wanted], [first.path, first.path]);
  assert.deepEqual(
    [basename(second.paths.model), basename(second.paths.manuscript)],
    ["paper.response-2.model.json", "paper.response-2.manuscript.json"],
  );
  assert.equal(await readFile(first.path, "utf8"), "edited by the author in Word");
  // A file of the next name that someone else put there is skipped as well.
  await writeFile(join(t.dir, "paper.response-3.docx"), "mine");
  const third = await t.ask("response_export", { ...WITH_DOC, format: "docx" });
  assert.equal(basename(third.path), "paper.response-4.docx");
  assert.equal(await readFile(join(t.dir, "paper.response-3.docx"), "utf8"), "mine");
  // Asked to: the first name again, replaced.
  const again = await t.ask("response_export", { ...WITH_DOC, format: "docx", overwrite: true });
  assert.equal(basename(again.path), "paper.response.docx");
  assert.equal(await readFile(first.path, "utf8"), "docx");
  assert.ok(t.py.calls.at(-2)[1].includes("--overwrite"));
  await t.cleanup();
});

test("response_export: a name taken between the look and the write, a locked file, a failed verification", async () => {
  // The writer says "exists" for the first name although the service saw none.
  let n = 0;
  const racing = fakePy({
    docx: async (args) => {
      const out = args[args.indexOf("--out") + 1];
      if (++n === 1) return { ok: false, code: 4, json: { ok: false, code: "exists" } };
      await writeFile(out, "docx");
      return { ok: true, code: 0, json: { ok: true, path: out, paragraphs: 6 } };
    },
  });
  const t = await setup({ py: racing });
  await ready(t);
  const r = await t.ask("response_export", { ...WITH_DOC, format: "docx" });
  assert.equal(basename(r.path), "paper.response-2.docx");
  // The sidecars of the name that was lost are not left behind.
  assert.deepEqual(
    (await t.files()).filter((f) => f.includes(".response")),
    ["paper.response-2.docx", "paper.response-2.manuscript.json", "paper.response-2.model.json"],
  );
  await t.cleanup();

  // Open in Word while overwrite was asked: a new name, the open file untouched.
  const locked = fakePy({
    docx: async (args) => {
      const out = args[args.indexOf("--out") + 1];
      if (args.includes("--overwrite"))
        return { ok: false, code: 3, json: { ok: false, code: "locked", error: "in use" } };
      await writeFile(out, "docx");
      return { ok: true, code: 0, json: { ok: true, path: out, paragraphs: 6 } };
    },
  });
  const u = await setup({ py: locked });
  await ready(u);
  await writeFile(join(u.dir, "paper.response.docx"), "open in Word");
  const l = await u.ask("response_export", { ...WITH_DOC, format: "docx", overwrite: true });
  assert.deepEqual([l.ok, basename(l.path)], [true, "paper.response-2.docx"]);
  assert.equal(await readFile(join(u.dir, "paper.response.docx"), "utf8"), "open in Word");
  await u.cleanup();

  // The file on disk is not what the model says: no letter is left as if it were good.
  const failing = fakePy({
    verify: {
      ok: false,
      code: 1,
      json: { ok: false, checked: {}, failures: [{ axis: "colour", index: 3, detail: "x" }] },
    },
  });
  const v = await setup({ py: failing });
  await ready(v);
  const bad = await v.ask("response_export", { ...WITH_DOC, format: "docx" });
  assert.equal(bad.ok, false);
  assert.match(bad.error, /failed verification \(colour\)/);
  assert.deepEqual(
    (await v.files()).filter((f) => f.includes(".response")),
    [],
  );
  assert.equal((await v.rev()).response_export, null);
  await v.cleanup();
});

test("response_export as text: one file per reviewer and one for the editor, never over an old one", async () => {
  const seed = [
    ...comments(),
    {
      ref: "R2.1",
      reviewer: 2,
      text: "Name the antifoam.",
      declined: { reason: "None was used." },
    },
    {
      ref: "E.1",
      reviewer: 0,
      text: "Shorten the title.",
      declined: { reason: "It is at the limit." },
    },
  ];
  const t = await setup({ seed });
  await ready(t);
  const r = await t.ask("response_export", { ...WITH_DOC, format: "text" });
  assert.equal(r.ok, true, r.error);
  assert.deepEqual(
    r.paths.map((p) => basename(p)),
    ["paper.response.E.txt", "paper.response.R1.txt", "paper.response.R2.txt"],
  );
  assert.equal(
    await readFile(r.paths[1], "utf8"),
    `Comment 1\nPlease state the unit.\nWe now state the unit.\n[Revised in ${SECTION}, line 9]\n${QUOTE}\n`,
  );
  assert.equal(
    await readFile(r.paths[2], "utf8"),
    "Comment 1\nName the antifoam.\nNone was used.\n[No change made]\n",
  );
  assert.equal(
    t.py.calls.some(([name]) => name.startsWith("response_")),
    false,
  );
  const again = await t.ask("response_export", { ...WITH_DOC, format: "text" });
  assert.deepEqual(
    again.paths.map((p) => basename(p)),
    ["paper.response.E-2.txt", "paper.response.R1-2.txt", "paper.response.R2-2.txt"],
  );
  assert.deepEqual([r.wanted, again.wanted], [r.paths, r.paths]);
  const unknown = await t.ask("response_export", { ...WITH_DOC, format: "pdf" });
  assert.match(unknown.error, /format/);
  await t.cleanup();
});

test("without Python the export says so", async () => {
  const none = {
    run: async () => ({ ok: false, code: null, error: "python3 not found" }),
    calls: [],
  };
  const t = await setup({ py: none });
  await t.ask("response_save", { ref: "R1.1", text: "We now state the unit.", ...WITH_DOC });
  const r = await t.ask("response_export", { ...WITH_DOC, format: "docx", force: true });
  assert.deepEqual([r.ok, r.error], [false, "python3 not found"]);
  assert.deepEqual(
    (await t.files()).filter((f) => f.includes(".response")),
    [],
  );
  const l = await t.ask("response_locate", { ...WITH_DOC, pdf_base64: PDF });
  assert.match(l.error, /python3 not found/);
  await t.cleanup();
});

test("profile_propose reads the letter and proposes; nothing is saved", async () => {
  const read = {
    ok: true,
    code: 0,
    json: {
      ok: true,
      overrides: { colors: { response_body: "C00000" } },
      evidence: { paragraphs: 3 },
      previous: { "R1.1": "We added the unit." },
    },
  };
  const t = await setup({ py: fakePy({ read }) });
  const before = await readFile(t.revPath, "utf8");
  const r = await t.ask("profile_propose", {
    docx_base64: Buffer.from("PK letter").toString("base64"),
    filename: "../../old letter.docx",
  });
  assert.deepEqual(
    [r.ok, r.overrides, r.evidence, r.previous],
    [true, read.json.overrides, read.json.evidence, read.json.previous],
  );
  assert.equal(await readFile(t.revPath, "utf8"), before);
  const [, args] = t.py.calls[0];
  assert.deepEqual([args[0], args[2]], ["--docx", "--profile"]);
  // Both temp files are gone, and the letter never went by the name the pane gave.
  assert.equal(existsSync(args[1]) || existsSync(args[3]), false);
  assert.equal(args[1].includes("old letter"), false);
  const bad = await setup({
    py: fakePy({
      read: { ok: false, code: 1, json: { ok: false, error: "File is not a zip file" } },
    }),
  });
  const b = await bad.ask("profile_propose", { docx_base64: "eA==", filename: "x.docx" });
  assert.deepEqual([b.ok, b.error], [false, "File is not a zip file"]);
  assert.match((await bad.ask("profile_propose", { filename: "x.docx" })).error, /No letter/);
  await t.cleanup();
  await bad.cleanup();
});

test("response_import_previous keeps the earlier answers as `previous`, for comments that exist", async () => {
  const t = await setup();
  const r = await t.ask("response_import_previous", {
    previous: { "R1.1": "We added the unit in round one.", "R4.4": "gone", "R1.2": 7 },
  });
  assert.deepEqual([r.ok, r.imported, r.unknown], [true, ["R1.1"], ["R4.4", "R1.2"]]);
  const c = r.revision.comments[0];
  assert.deepEqual([c.response.previous, c.response.text], ["We added the unit in round one.", ""]);
  assert.match((await t.ask("response_import_previous", { previous: [] })).error, /previous/);
  await t.cleanup();
});

test("a pane without a local document gets the same answer for every request", async () => {
  const svc = createResponseService({
    getContext: () => ({ activeDoc: null }),
    env: ENV,
    log: () => {},
  });
  for (const type of ["profile_get", "response_state", "response_save"]) {
    const replies = [];
    await svc.handlers[type](
      { request_id: "q", ...WITH_DOC, ref: "R1.1" },
      (r) => replies.push(r),
      "k",
      "word",
    );
    assert.deepEqual(
      [replies[0].ok, /Save the document/.test(replies[0].error)],
      [false, true],
      type,
    );
  }
});

test("with the real Python tools: the letter is written, verified, and read back", async () => {
  const t = await setup({ py: { run: realRunPy, calls: [] } });
  await t.ask("response_save", { ref: "R1.1", text: "We now state the unit.", ...WITH_DOC });
  const r = await t.ask("response_export", { ...WITH_DOC, format: "docx", force: true });
  assert.equal(r.ok, true, r.error);
  assert.deepEqual(
    [r.exported, r.verify.ok, r.verify.checked],
    [true, true, { paragraphs: 6, quotes: 1, locations: 1 }],
  );
  const again = await t.ask("response_export", { ...WITH_DOC, format: "docx", force: true });
  assert.equal(basename(again.path), "paper.response-2.docx");
  const proposal = await t.ask("profile_propose", {
    docx_base64: (
      await readFile(new URL("./fixtures/prior_response.docx", import.meta.url))
    ).toString("base64"),
    filename: "prior_response.docx",
  });
  assert.deepEqual(proposal.overrides.colors, {
    response_body: "C00000",
    response_revised: "0070C0",
  });
  assert.equal(Object.keys(proposal.previous).length, 5);
  await t.cleanup();
});

test("the headings a request brought are logged by number and level, never by text", async () => {
  const t = await setup();
  await t.ask("response_state", {
    snapshot: ACCEPTED,
    headings: [...HEADINGS, { index: 1, level: 2, text: "not what paragraph 1 says" }],
  });
  assert.ok(t.lines.includes("[response] headings n=1 levels=1:1 dropped=1"), t.lines.join("\n"));
  await t.ask("response_state", { snapshot: ACCEPTED, headings: [] });
  assert.ok(t.lines.includes("[response] headings n=0 levels=- dropped=0"));
  assert.equal(
    t.lines.some((l) => l.startsWith("[response] headings") && l.includes(SECTION)),
    false,
  );
  // With no heading the citation has no section, and the letter is not held back for it.
  await ready(t);
  const r = await t.ask("response_check", { snapshot: ACCEPTED, headings: [] });
  assert.deepEqual(r.summary, { fail: 0, warn: 1 });
  assert.deepEqual(
    r.checks.map((x) => [x.id, x.level, x.ref]),
    [["section", "warn", null]],
  );
  await t.cleanup();
});

// ---- the marks of an answer that is out of date -----------------------------------

// A model that answers only when told to, so the manuscript can change while
// it writes.
function heldModel(answer = draftOk) {
  let release;
  const gate = new Promise((r) => (release = r));
  const m = model(async () => {
    await gate;
    return answer;
  });
  const asked = async () => {
    while (!m.requests.length) await new Promise((r) => setTimeout(r, 2));
  };
  return { ...m, release, asked };
}
const EDITED = snapshotOf({ text: AFTER.replace("30 °C", "32 °C") });
const staleRows = (rows) => rows.filter((x) => x.id === "stale").map((x) => x.detail.split(":")[0]);

test("response_draft: the manuscript changed while the model wrote, so the draft is saved as out of date", async () => {
  const m = heldModel({ ok: true, json: { response: "The temperature is now given as 30 °C." } });
  const t = await setup({ job: m.job });
  const drafting = t.ask("response_draft", { ref: "R1.1", ...WITH_DOC });
  await m.asked();
  // The author edits the sentence in Word and a pane reads the document again.
  const seen = await t.ask("response_state", { snapshot: EDITED, headings: HEADINGS });
  assert.match(seen.revision.comments[0].response.citations[0].quote, /held at 32 °C/);
  m.release();
  const r = await drafting;
  assert.equal(r.ok, true, r.error);
  const c = r.revision.comments[0];
  // The answer is kept (it cost a model call), with the marks of what moved.
  assert.equal(c.response.text, "The temperature is now given as 30 °C.");
  assert.match(c.response.citations[0].quote, /held at 32 °C/);
  assert.deepEqual(c.response.stale, ["quote_changed", "change_state"]);
  assert.deepEqual(staleRows(r.checks), ["quote_changed", "change_state"]);
  assert.match(r.warnings.join(" "), /manuscript changed while .* was drafted/);
  assert.ok(t.lines.includes("[response] draft R1.1 input changed: quote_changed,change_state"));

  // Updating the line numbers does not make it pass: the marks stay.
  const located = await t.ask("response_locate", {
    snapshot: EDITED,
    headings: HEADINGS,
    pdf_base64: PDF,
  });
  assert.deepEqual(located.revision.comments[0].response.stale, ["quote_changed", "change_state"]);
  const checked = await t.ask("response_check", { snapshot: EDITED, headings: HEADINGS });
  assert.deepEqual(staleRows(checked.checks), ["quote_changed", "change_state"]);
  assert.deepEqual(
    checked.checks.filter((x) => x.level === "fail").map((x) => x.id),
    ["stale", "stale"],
  );
  const held = await t.ask("response_export", { snapshot: EDITED, headings: HEADINGS });
  assert.equal(held.exported, false);
  await t.cleanup();
});

test("response_draft: a change decided while the model wrote marks the draft; a refresh that changed nothing does not", async () => {
  // Drafted while the change was pending ("nothing was revised"), accepted meanwhile.
  const m = heldModel();
  const t = await setup({ job: m.job });
  const drafting = t.ask("response_draft", { ref: "R1.1", snapshot: PENDING, headings: HEADINGS });
  await m.asked();
  await t.ask("response_state", WITH_DOC);
  m.release();
  const r = await drafting;
  assert.deepEqual(r.revision.comments[0].response.stale, ["quote_changed", "change_state"]);
  await t.cleanup();

  // The same document read again while the model wrote: nothing to mark.
  const m2 = heldModel();
  const t2 = await setup({ job: m2.job });
  const again = t2.ask("response_draft", { ref: "R1.1", ...WITH_DOC });
  await m2.asked();
  await t2.ask("response_state", WITH_DOC);
  m2.release();
  const r2 = await again;
  assert.deepEqual([r2.revision.comments[0].response.stale, r2.warnings], [[], []]);
  assert.equal(
    t2.lines.some((l) => l.includes("input changed")),
    false,
  );
  await t2.cleanup();
});

test("response_draft: a mark the Review tab's refresh set while the model wrote is kept", async () => {
  // The Review refresh marks the saved answers and leaves their citations as
  // they are (review-service.mjs afterRefresh), so the stored quote is the
  // same before and after: the mark itself tells that the document moved.
  const m = heldModel();
  const t = await setup({ job: m.job });
  await ready(t);
  const drafting = t.ask("response_draft", { ref: "R1.1", ...WITH_DOC });
  await m.asked();
  const KEPT = snapshotOf({ text: AFTER.replace("was held at", "was kept at") });
  const marked = await withRevision(t.revPath, t.doc, (rev) =>
    markStale(applyStates(rev, classifyRevision(rev, KEPT)), KEPT),
  );
  assert.deepEqual(marked.comments[0].response.stale, ["quote_changed", "lines_old"]);
  assert.equal(marked.comments[0].response.citations[0].quote, QUOTE);
  m.release();
  const r = await drafting;
  assert.deepEqual(r.revision.comments[0].response.stale, ["quote_changed", "lines_old"]);
  assert.match(r.warnings.join(" "), /manuscript changed while/);
  await t.cleanup();
});

test("response_draft: an answer already marked when the draft was asked for, and the sentence edited once more meanwhile", async () => {
  const reviewRefresh = (t, snap) =>
    withRevision(t.revPath, t.doc, (rev) =>
      markStale(applyStates(rev, classifyRevision(rev, snap)), snap),
    );
  const KEPT = snapshotOf({ text: AFTER.replace("was held at", "was kept at") });
  const AGAIN = snapshotOf({ text: AFTER.replace("was held at", "was always kept at") });
  const start = async () => {
    const m = heldModel();
    const t = await setup({ job: m.job });
    await ready(t);
    // The sentence was edited: the answer is marked, and a redraft is asked for.
    const seen = await t.ask("response_state", { snapshot: KEPT, headings: HEADINGS });
    assert.deepEqual(seen.revision.comments[0].response.stale, ["quote_changed", "lines_old"]);
    const drafting = t.ask("response_draft", { ref: "R1.1", snapshot: KEPT, headings: HEADINGS });
    await m.asked();
    return { m, t, drafting };
  };

  // Edited again while the model writes; the Review refresh sees it. The mark
  // was there before and is there after, and the citation is not replaced.
  const a = await start();
  const marked = await reviewRefresh(a.t, AGAIN);
  assert.deepEqual(marked.comments[0].response.stale, ["quote_changed", "lines_old"]);
  assert.match(marked.comments[0].response.citations[0].quote, /was kept at 30/);
  a.m.release();
  let r = await a.drafting;
  assert.deepEqual(r.revision.comments[0].response.stale, ["quote_changed", "lines_old"]);
  assert.match(r.warnings.join(" "), /manuscript changed while/);
  await a.t.cleanup();

  // The same refresh with the document the draft was asked for: the draft
  // answers the mark, as it does with no refresh at all.
  const b = await start();
  await reviewRefresh(b.t, KEPT);
  b.m.release();
  r = await b.drafting;
  assert.deepEqual([r.revision.comments[0].response.stale, r.warnings], [["lines_old"], []]);
  await b.t.cleanup();

  // Another paragraph changed meanwhile: not this answer's business.
  const c = await start();
  await reviewRefresh(c.t, {
    ...KEPT,
    paragraphs: [...KEPT.paragraphs, { index: 2, id: "P2", text: "A new last paragraph." }],
  });
  c.m.release();
  r = await c.drafting;
  assert.deepEqual([r.revision.comments[0].response.stale, r.warnings], [["lines_old"], []]);
  await c.t.cleanup();
});

// The marks quote_changed and change_state on every way an answer is saved.
// What must hold: reading the line numbers again never makes an answer pass
// that was written for another sentence.
const EDITED_DOC = { snapshot: EDITED, headings: HEADINGS };
const staleOf = (r) => r.revision.comments[0].response.stale;
async function stillHeld(t, doc = EDITED_DOC) {
  const located = await t.ask("response_locate", { ...doc, pdf_base64: PDF });
  assert.deepEqual(staleOf(located), ["quote_changed"], "locate keeps the mark");
  const checked = await t.ask("response_check", doc);
  assert.deepEqual(
    checked.checks.filter((x) => x.level === "fail").map((x) => x.id),
    ["stale"],
  );
  const held = await t.ask("response_export", doc);
  assert.equal(held.exported, false);
}

test("response_save: the sentence the card showed was edited before the save, so the answer is saved as out of date", async () => {
  const t = await setup();
  await ready(t); // the card shows the sentence with 30 °C
  // The author edits the sentence in Word, then saves an answer typed before that.
  let r = await t.ask("response_save", {
    ref: "R1.1",
    text: "The temperature is now given as 30 °C.",
    ...EDITED_DOC,
  });
  assert.match(r.revision.comments[0].response.citations[0].quote, /held at 32 °C/);
  assert.deepEqual(staleOf(r), ["quote_changed", "lines_old"]);
  assert.match(r.warnings.join(" "), /manuscript changed under this answer/);
  await stillHeld(t);
  // Read against the sentence as it is now and saved again: the mark goes.
  r = await t.ask("response_save", {
    ref: "R1.1",
    text: "The temperature is now given as 32 °C.",
    ...EDITED_DOC,
  });
  assert.deepEqual([staleOf(r), r.warnings], [[], []]);
  r = await t.ask("response_check", EDITED_DOC);
  // (One warning is left: the number 30 of the earlier text is not in the new one.)
  assert.deepEqual(
    r.checks.map((x) => [x.id, x.level]),
    [["lost_facts", "warn"]],
  );
  await t.cleanup();

  // A cited change that was rejected before the save.
  const u = await setup();
  await ready(u);
  const REJECTED = { snapshot: snapshotOf({ text: BEFORE }), headings: HEADINGS };
  r = await u.ask("response_save", { ref: "R1.1", text: "We now state the unit.", ...REJECTED });
  assert.deepEqual(r.revision.comments[0].response.citations, []);
  assert.ok(staleOf(r).includes("change_state"), JSON.stringify(staleOf(r)));
  await u.cleanup();

  // The first answer to a comment, with no citation shown yet: nothing moved under it.
  const v = await setup();
  r = await v.ask("response_save", { ref: "R1.1", text: "We now state the unit.", ...WITH_DOC });
  assert.deepEqual([staleOf(r), r.warnings], [[], []]);
  await v.cleanup();
});

test("response_save without the document, and a redraft in another language: neither lets a stale answer pass", async () => {
  // Saved without a snapshot: the citations are not touched, so the next
  // read of the document marks the answer.
  const t = await setup();
  await ready(t);
  let r = await t.ask("response_save", { ref: "R1.1", text: "The temperature is 30 °C." });
  assert.deepEqual(staleOf(r), []);
  assert.equal(r.revision.comments[0].response.citations[0].quote, QUOTE);
  r = await t.ask("response_state", EDITED_DOC);
  assert.deepEqual(staleOf(r), ["quote_changed", "lines_old"]);
  await stillHeld(t);

  // Changing the language of the letter does not touch the marks; the
  // redraft that follows is a draft like any other, written for the
  // sentence as it is now.
  const m = model(draftOk);
  const u = await setup({ job: m.job });
  await ready(u);
  await u.ask("response_state", EDITED_DOC);
  r = await u.ask("profile_set", { overrides: { response_letter: { lang: "German" } } });
  assert.deepEqual(staleOf(r), ["quote_changed", "lines_old"]);
  await stillHeld(u);
  r = await u.ask("response_draft", { ref: "R1.1", ...EDITED_DOC });
  assert.match(m.requests[0].prompt, /Write in German\./);
  assert.deepEqual(staleOf(r), []);
  await u.cleanup();
  await t.cleanup();
});

test("response_check: an LLM rule reads the reasons for declining too, as the letter prints them", async () => {
  const REASON = "The supplier was not recorded, sadly.";
  const paras = [
    { index: 0, text: "We now state the unit." },
    { index: 1, text: REASON },
  ];
  const m = model((req) =>
    ledgerReply(req, paras, {
      violations: [{ paragraph_index: 1, quote: "sadly", reason: "not courteous" }],
    }),
  );
  const t = await setup({
    job: m.job,
    seed: comments({
      requirements: [
        { id: "q1", text: "state the unit" },
        { id: "q2", text: "name the supplier" },
      ],
    }),
    revision: {
      ledger: [
        {
          id: "L1",
          kind: "instruction",
          text: "Be courteous.",
          scope: "response",
          check: { type: "llm" },
        },
      ],
    },
  });
  await ready(t);
  await t.ask("response_save", {
    ref: "R1.1",
    declined_requirements: [{ id: "q2", reason: REASON }],
  });
  const r = await t.ask("response_check", WITH_DOC);
  assert.equal(m.requests.length, 1);
  assert.ok(m.requests[0].prompt.includes(REASON), m.requests[0].prompt);
  assert.deepEqual(
    r.checks.map((x) => [x.id, x.level, x.ref, x.detail]),
    [["ledger", "fail", "R1.1", "L1: sadly (not courteous)"]],
  );
  await t.cleanup();
});

// ---- the translation paragraph follows the answer text --------------------------

const GLOSS_ON = {
  profile_overrides: { response_letter: { gloss_enabled: true, gloss_lang: "German" } },
};
// Translates by a table, so a translation can be told from its answer.
const GERMAN = {
  "We now state the unit.": "Wir nennen die Einheit.",
  "We now state the unit and the supplier.": "Wir nennen die Einheit und den Hersteller.",
};
const translator = (held = null) =>
  model(async (req) => {
    if (req.expectJson !== false) return draftOk;
    await held;
    return { ok: true, text: GERMAN[req.prompt] ?? "" };
  });
const glossOf = (r) => r.revision.comments[0].response.gloss;
const translationRows = (r) => r.checks.filter((x) => x.id === "translation");

test("response_save: an answer changed by hand drops its translation, and the check says so", async () => {
  const m = translator();
  const t = await setup({ job: m.job, revision: GLOSS_ON });
  let r = await t.ask("response_draft", { ref: "R1.1", ...WITH_DOC });
  assert.equal(glossOf(r), "Wir nennen die Einheit.");
  // Saved as it is: the translation still fits.
  r = await t.ask("response_save", { ref: "R1.1", text: "We now state the unit.", ...WITH_DOC });
  assert.equal(glossOf(r), "Wir nennen die Einheit.");
  // A decision about a request is not a change of the text.
  r = await t.ask("response_save", { ref: "R1.1", declined_requirements: [] });
  assert.equal(glossOf(r), "Wir nennen die Einheit.");
  // The author adds the supplier: the old translation no longer says what the answer says.
  r = await t.ask("response_save", {
    ref: "R1.1",
    text: "We now state the unit and the supplier.",
    ...WITH_DOC,
  });
  assert.equal(glossOf(r), "");
  await t.ask("response_locate", { ...WITH_DOC, pdf_base64: PDF });
  r = await t.ask("response_check", WITH_DOC);
  assert.deepEqual(
    r.checks.map((x) => [x.id, x.level, x.ref]),
    [["translation", "warn", "R1.1"]],
  );
  assert.match(r.checks[0].detail, /no German translation/);
  // The letter is written without a translation of another text.
  r = await t.ask("response_export", WITH_DOC);
  assert.equal(r.exported, true, r.error);
  const letter = JSON.parse(await readFile(r.paths.model, "utf8"));
  assert.deepEqual(
    letter.paragraphs.filter((p) => p.kind === "gloss"),
    [],
  );
  assert.equal(translationRows(r).length, 1);

  // Translate: the answer as saved, nothing else rewritten.
  r = await t.ask("response_translate", { ref: "R1.1" });
  assert.equal(r.ok, true, r.error);
  assert.equal(glossOf(r), "Wir nennen die Einheit und den Hersteller.");
  assert.equal(r.revision.comments[0].response.text, "We now state the unit and the supplier.");
  assert.equal(r.revision.comments[0].response.source, "author");
  assert.match(m.requests.at(-1).system, /Translate into German/);
  r = await t.ask("response_check", WITH_DOC);
  assert.deepEqual(r.checks, []);
  assert.ok(t.lines.some((l) => /\[response\] translate R1\.1 lang=German/.test(l)));
  await t.cleanup();
});

test("no translation row without the second language; an answer never translated has one", async () => {
  // Written by hand with the second language on: there never was a translation.
  const on = await setup({ revision: GLOSS_ON });
  await ready(on);
  let r = await on.ask("response_check", WITH_DOC);
  assert.deepEqual(
    r.checks.map((x) => [x.id, x.level, x.ref]),
    [["translation", "warn", "R1.1"]],
  );
  await on.cleanup();
  // Off: nothing to tell, whatever the answer holds.
  const off = await setup();
  await ready(off);
  r = await off.ask("response_check", WITH_DOC);
  assert.deepEqual(r.checks, []);
  await off.cleanup();
});

test("response_translate: what it needs, and an answer changed while it ran", async () => {
  const none = await setup({ revision: GLOSS_ON });
  await ready(none);
  let r = await none.ask("response_translate", { ref: "R1.1" });
  assert.deepEqual(
    [r.ok, r.error],
    [false, "No model is available for translating. Pick a provider in Setup."],
  );
  await none.cleanup();

  const off = await setup({ job: translator().job });
  await ready(off);
  r = await off.ask("response_translate", { ref: "R1.1" });
  assert.match(r.error, /second language is off/);
  r = await off.ask("response_translate", { ref: "R9.9" });
  assert.match(r.error, /No reviewer comment R9\.9/);
  await off.cleanup();

  const empty = await setup({ job: translator().job, revision: GLOSS_ON });
  r = await empty.ask("response_translate", { ref: "R1.1" });
  assert.match(r.error, /no answer to translate/i);
  await empty.cleanup();

  // The model returns nothing: an error, nothing saved.
  const mute = await setup({ job: translator().job, revision: GLOSS_ON });
  await mute.ask("response_save", { ref: "R1.1", text: "A text the table does not hold." });
  r = await mute.ask("response_translate", { ref: "R1.1" });
  assert.match(r.error, /translation failed/);
  await mute.cleanup();

  // The text is saved again while the model translates the old one.
  let release;
  const m = translator(new Promise((res) => (release = res)));
  const t = await setup({ job: m.job, revision: GLOSS_ON });
  await ready(t);
  const translating = t.ask("response_translate", { ref: "R1.1" });
  while (!m.requests.length) await new Promise((res) => setTimeout(res, 2));
  await t.ask("response_save", { ref: "R1.1", text: "We now state the unit and the supplier." });
  release();
  r = await translating;
  assert.match(r.error, /changed while it was translated/);
  assert.equal((await t.rev()).comments[0].response.gloss, "");
  await t.cleanup();
});

// ---- two exports of one manuscript at the same time -----------------------------

// The Python tools as they are about files: the writer refuses a name that is
// taken, the verifier needs the model and the manuscript list it is given.
// peak(): how many tool calls ran at the same time at most.
function strictPy() {
  const py = fakePy();
  let active = 0;
  let peak = 0;
  const run = async (script, args) => {
    peak = Math.max(peak, ++active);
    try {
      await new Promise((res) => setTimeout(res, 5));
      return await tool(script, args);
    } finally {
      active -= 1;
    }
  };
  const tool = async (script, args) => {
    if (basename(script) === "response_verify.py") {
      const missing = ["--docx", "--model", "--manuscript"]
        .map((flag) => args[args.indexOf(flag) + 1])
        .filter((p) => !existsSync(p));
      if (missing.length)
        return {
          ok: false,
          code: 1,
          json: { ok: false, error: `missing ${missing.map((p) => basename(p)).join(", ")}` },
        };
    }
    return py.run(script, args);
  };
  return { run, calls: py.calls, peak: () => peak };
}

test("response_export: two panes export the same manuscript at once and both letters are whole", async () => {
  // Both requests are held at the same point (the model of an LLM rule) and
  // let go together, so both look for a free name before either writes.
  let open;
  const gate = new Promise((res) => (open = res));
  let waiting = 0;
  const t = await setup({
    py: strictPy(),
    revision: {
      ledger: [
        {
          id: "L1",
          kind: "instruction",
          text: "Be kind.",
          scope: "response",
          check: { type: "llm" },
        },
      ],
    },
    service: {
      oneshotJob: async () => {
        waiting += 1;
        await gate;
        return { run: null, model: null };
      },
    },
  });
  await ready(t);
  const ask = () => t.ask("response_export", { ...WITH_DOC, format: "docx", force: true });
  const both = [ask(), ask()];
  while (waiting < 2) await new Promise((res) => setTimeout(res, 2));
  open();
  const [a, b] = await Promise.all(both);
  assert.deepEqual([a.ok, a.error, b.ok, b.error], [true, undefined, true, undefined]);
  assert.deepEqual([basename(a.path), basename(b.path)].sort(), [
    "paper.response-2.docx",
    "paper.response.docx",
  ]);
  // Each letter has its own model and manuscript list, and nothing else is left.
  assert.deepEqual(
    (await t.files()).filter((f) => f.includes(".response")),
    [
      "paper.response-2.docx",
      "paper.response-2.manuscript.json",
      "paper.response-2.model.json",
      "paper.response.docx",
      "paper.response.manuscript.json",
      "paper.response.model.json",
    ],
  );
  for (const r of [a, b])
    assert.equal(JSON.parse(await readFile(r.paths.model, "utf8")).paragraphs.length, 6);
  // One after the other: the second began to write when the first was done.
  assert.equal(t.py.peak(), 1);
  // Plain text for the portal goes through the same queue.
  const texts = await Promise.all([
    t.ask("response_export", { ...WITH_DOC, format: "text", force: true }),
    t.ask("response_export", { ...WITH_DOC, format: "text", force: true }),
  ]);
  assert.deepEqual(texts.map((x) => basename(x.paths[0])).sort(), [
    "paper.response.R1-2.txt",
    "paper.response.R1.txt",
  ]);
  await t.cleanup();
});

test("response_export removes only files it made: a file that takes the name after the look is kept", async () => {
  // A link without a target: the look for a free name does not see it, the
  // write does. It stands for a file another program puts there in between.
  const t = await setup({ py: strictPy() });
  await ready(t);
  const foreign = join(t.dir, "paper.response.manuscript.json");
  await symlink(join(t.dir, "nowhere.json"), foreign);
  const r = await t.ask("response_export", { ...WITH_DOC, format: "docx" });
  assert.equal(r.ok, true, r.error);
  assert.equal(basename(r.path), "paper.response-2.docx");
  // The foreign file is as it was; our model of the first name is not left behind.
  assert.equal(await readlink(foreign), join(t.dir, "nowhere.json"));
  assert.deepEqual((await readdir(t.dir)).filter((f) => f.includes(".response")).sort(), [
    "paper.response-2.docx",
    "paper.response-2.manuscript.json",
    "paper.response-2.model.json",
    "paper.response.manuscript.json",
  ]);
  await t.cleanup();
});

// An earlier letter under the plain name, each file with a content of its own.
const OLD = {
  "paper.response.docx": "old docx",
  "paper.response.model.json": '{"old":"model"}\n',
  "paper.response.manuscript.json": '["old manuscript"]\n',
};
async function withOldLetter(py) {
  const t = await setup({ py });
  await ready(t);
  for (const [name, text] of Object.entries(OLD)) await writeFile(join(t.dir, name), text);
  const plain = async () =>
    Object.fromEntries(
      await Promise.all(
        Object.keys(OLD).map(async (name) => [name, await readFile(join(t.dir, name), "utf8")]),
      ),
    );
  // Every file the export left, the daemon's own two aside.
  const left = async () =>
    (await t.files()).filter((f) => !["paper.revision.json", "paper.render.pdf"].includes(f));
  return { t, plain, left };
}
const docxTool = (answer) =>
  fakePy({
    docx: async (args) => {
      const out = args[args.indexOf("--out") + 1];
      const r = args.includes("--overwrite") ? answer : null;
      if (r) return r;
      await writeFile(out, "docx");
      return { ok: true, code: 0, json: { ok: true, path: out, paragraphs: 6 } };
    },
  });

test("response_export with overwrite: the old letter and its sidecars stay a set, wherever the write stops", async () => {
  // The docx is open in Word: nothing of the old set is touched, a new set is written.
  const locked = await withOldLetter(
    docxTool({ ok: false, code: 3, json: { ok: false, code: "locked", error: "in use" } }),
  );
  let r = await locked.t.ask("response_export", { ...WITH_DOC, format: "docx", overwrite: true });
  assert.deepEqual([r.ok, basename(r.path)], [true, "paper.response-2.docx"]);
  assert.deepEqual(await locked.plain(), OLD);
  assert.deepEqual(await locked.left(), [
    "paper.response-2.docx",
    "paper.response-2.manuscript.json",
    "paper.response-2.model.json",
    "paper.response.docx",
    "paper.response.manuscript.json",
    "paper.response.model.json",
  ]);
  await locked.t.cleanup();

  // The writer fails, or cannot run at all: the old set as it was, nothing else.
  for (const answer of [
    { ok: false, code: 1, json: { ok: false, error: "disk full" } },
    { ok: false, code: null, json: null, error: "python3 not found" },
  ]) {
    const failed = await withOldLetter(docxTool(answer));
    r = await failed.t.ask("response_export", { ...WITH_DOC, format: "docx", overwrite: true });
    assert.equal(r.ok, false);
    assert.match(r.error, /disk full|python3 not found/);
    assert.deepEqual(await failed.plain(), OLD);
    assert.deepEqual(await failed.left(), Object.keys(OLD).sort());
    await failed.t.cleanup();
  }

  // Written: all three are the new ones.
  const done = await withOldLetter(fakePy());
  r = await done.t.ask("response_export", { ...WITH_DOC, format: "docx", overwrite: true });
  assert.deepEqual([r.ok, basename(r.path)], [true, "paper.response.docx"]);
  const now = await done.plain();
  assert.equal(now["paper.response.docx"], "docx");
  assert.equal(JSON.parse(now["paper.response.model.json"]).paragraphs.length, 6);
  assert.deepEqual(
    JSON.parse(now["paper.response.manuscript.json"]),
    ACCEPTED.paragraphs.map(viewText),
  );
  assert.deepEqual(await done.left(), Object.keys(OLD).sort());
  await done.t.cleanup();

  // Written, and the verifier says no: the new docx is there with its own
  // sidecars (the error names the file), not with the old ones.
  const bad = await withOldLetter(
    fakePy({ verify: { ok: false, code: 1, json: { ok: false, failures: [{ axis: "colour" }] } } }),
  );
  r = await bad.t.ask("response_export", { ...WITH_DOC, format: "docx", overwrite: true });
  assert.match(r.error, /failed verification \(colour\)\. The file is at .*paper\.response\.docx/);
  const kept = await bad.plain();
  assert.equal(kept["paper.response.docx"], "docx");
  assert.equal(JSON.parse(kept["paper.response.model.json"]).paragraphs.length, 6);
  assert.notEqual(kept["paper.response.manuscript.json"], OLD["paper.response.manuscript.json"]);
  assert.deepEqual(await bad.left(), Object.keys(OLD).sort());
  await bad.t.cleanup();
});
