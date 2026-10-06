// What scripts/e2e/replay.mjs computes without Word (replay-lib.mjs), and
// that what it asks of the driver and the Windows scripts exists. The run
// itself needs Word and does not run here. Every text is invented.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { runInNewContext } from "node:vm";
import path from "node:path";
import { cleanupDecision } from "../scripts/e2e/judge.mjs";
import { execWithInteropRetry } from "../scripts/e2e/wsl-interop.mjs";
import {
  changesReport,
  extractBatchesFromLog,
  fallbacksFromLog,
  modelFailures,
  restoredFromLog,
  rewritesFromLog,
  followFromLog,
  staysFromLog,
  refusalsFromLog,
  keepRefPattern,
  ledgerFromLines,
  otherModels,
  parseOnly,
  readModelLine,
  responsesReport,
  ruleLines,
  summaryText,
  sweepSelection,
  newWords,
  blankWordTitle,
  turnFromLog,
  turnResult,
} from "../scripts/e2e/replay-lib.mjs";
import { LedgerItemSchema } from "../daemon/revision-store.mjs";

const read = (p) => readFileSync(new URL(p, import.meta.url), "utf8");

test("replay: refused, unknown and lingering quits never terminate Word", () => {
  const runner = read("../scripts/e2e/replay.mjs");
  const source = runner.slice(
    runner.indexOf("function quitWord("),
    runner.indexOf("function withStrays("),
  );
  for (const quit of [
    { ok: false, refused: "foreign_documents", outside: ["C:\\user\\notes.docx"] },
    { ok: false, refused: "ownership_unproven" },
    { ok: false, error: "COM timeout" },
    { ok: true, documents: ["C:\\run\\anchor.docx"], exited: false },
    { ok: true, documents: ["C:\\run\\anchor.docx"], exited: true },
  ]) {
    const record = { word: { preexisting: [] }, cleanup: {} };
    runInNewContext(`${source}; quitWord(42);`, {
      record,
      cleanupDecision,
      DOC: "anchor.docx",
      RUN: "C:\\run",
      winPath: (p) => p,
      powershell: () => quit,
      wordPids: () => (quit.exited ? [] : [42]),
      execFileSync: () => assert.fail("process termination attempted"),
      console: { error() {} },
    });
    assert.equal(!!record.cleanup.failed, !quit.exited);
    assert.equal(record.cleanup.left_running?.pid, quit.exited ? undefined : 42);
    assert.equal(record.cleanup.quit, quit);
  }
});

test("replay: an open-check quit refusal survives without a stop fallback", () => {
  const runner = read("../scripts/e2e/replay.mjs");
  const source = runner.slice(
    runner.indexOf("function wordOpens("),
    runner.indexOf("const embed ="),
  );
  const refused = { ok: true, pid: 42, quit: { ok: false, refused: "foreign_documents" } };
  const record = { word: {}, cleanup: {} };
  runInNewContext(`let checks = 0; ${source}; wordOpens("copy.docx");`, {
    record,
    RUN: "run",
    REPO: "repo",
    path,
    fs: { existsSync: () => true },
    winPath: (p) => p,
    withStrays: (_step, run) => run(),
    powershell: () => refused,
    wordPids: () => [42],
    save() {},
    console: { error() {} },
    stopOwn: () => assert.fail("refusal overridden"),
  });
  assert.equal(record.cleanup.failed, true);
  assert.equal(record.word.open_checks[0].quit.refused, "foreign_documents");
  assert.equal(record.word.open_checks[0].left_running.pid, 42);
});

test("--only: a list of refs, or nothing; anything else is refused", () => {
  assert.deepEqual(parseOnly(undefined), []);
  assert.deepEqual(parseOnly(" R1.2, E.1 ,R12.30"), ["R1.2", "E.1", "R12.30"]);
  assert.throws(() => parseOnly("R1"), /refs like R1\.2/);
  assert.throws(() => parseOnly("R1.2;R1.3"), /refs like/);
});

test("the draft rows that are kept: the refs of --only, or every reviewer and editor comment", () => {
  const all = new RegExp(keepRefPattern([]));
  for (const ref of ["R1.1", "R2.14", "E.1"]) assert.ok(all.test(ref), ref);
  for (const ref of ["X.1", "R1", "R1.1b", "?"]) assert.ok(!all.test(ref), ref);
  const some = new RegExp(keepRefPattern(["R1.2", "E.1"]));
  assert.deepEqual(
    ["R1.2", "E.1", "R1.20", "R1x2", "R11.2"].map((r) => some.test(r)),
    [true, true, false, false, false],
  );
});

test("instruction and forbid lines become ledger rules the model checks, in the form the pane step fills in", () => {
  assert.deepEqual(ruleLines("\n# a note\n  Keep the unit. \r\n\nNo abbreviations.\n"), [
    "Keep the unit.",
    "No abbreviations.",
  ]);
  const ledger = ledgerFromLines({
    instructions: "Keep the unit.\n",
    forbid: "# none yet\nnovel\n",
  });
  assert.deepEqual(
    ledger.items.map((i) => [i.kind, i.text, i.scope, i.check.type]),
    [
      ["instruction", "Keep the unit.", "manuscript", "llm"],
      ["forbid", "novel", "manuscript", "llm"],
    ],
  );
  for (const it of ledger.items)
    assert.deepEqual(Object.keys(it.check).sort(), ["expect", "flags", "pattern", "type"]);
  // The store takes such a rule (the pane sends the type only).
  for (const it of ledger.items)
    LedgerItemSchema.parse({
      id: "L1",
      kind: it.kind,
      text: it.text,
      scope: it.scope,
      check: { type: "llm" },
    });
  assert.deepEqual(ledgerFromLines({}).items, []);
});

const LOG = [
  "[agent] init session abc (model: claude-opus-x)",
  "[review] turn done R1.1",
  "[agent] init session def (model: claude-sonnet-y)",
  "[review] turn done R1.2 failed: Claude usage limit reached.",
  "[review] turn done R1.20",
].join("\n");

test("the model and the end of a turn are read from the daemon log", () => {
  assert.deepEqual(turnFromLog(LOG, "R1.1"), {
    model: "claude-sonnet-y",
    ended: true,
    ok: true,
    error: null,
  });
  assert.deepEqual(turnFromLog(LOG, "R1.2"), {
    model: "claude-sonnet-y",
    ended: true,
    ok: false,
    error: "Claude usage limit reached.",
  });
  // R1.2 is not R1.20, and a ref the log does not name has no end.
  assert.equal(turnFromLog("[review] turn done R1.20", "R1.2").ended, false);
  assert.deepEqual(turnFromLog("", "R9.9"), { model: null, ended: false, ok: false, error: null });
});

test("how a turn ended: ok, failed, timeout, or no turn at all", () => {
  const ok = turnResult("R1.1", {
    step: { ok: true, seconds: 41 },
    log: "[agent] init session a (model: m1)\n[review] turn done R1.1",
    seconds: 41,
  });
  assert.deepEqual(ok, {
    outcome: "ok",
    reason: null,
    seconds: 41,
    model: "m1",
    model_line: null,
    other_model: false,
    declined: null,
    fallbacks: [],
    stays: [],
    rewrites: [],
    refused: {},
  });
  // The pane's line right after the turn says what the edits ran on.
  const said = (line) =>
    turnResult("R1.1", { step: { ok: true }, log: "[review] turn done R1.1", seconds: 1, line });
  assert.deepEqual([said(LINE).model_line, said(LINE).other_model], ["Claude · Opus 9", false]);
  assert.deepEqual(
    [said(DRIFT).model_line, said(DRIFT).other_model],
    ["asked for Claude · Opus 9, ran on Claude · Sonnet 9", true],
  );
  // The session was started for an earlier comment: the model comes from the caller.
  assert.equal(
    turnResult("R1.1", {
      step: { ok: true },
      log: "[review] turn done R1.1",
      seconds: 1,
      model: "m0",
    }).model,
    "m0",
  );
  const failed = turnResult("R1.2", {
    step: { ok: true, seconds: 9 },
    log: "[review] turn done R1.2 failed: The model declined this request.",
    seconds: 9,
  });
  assert.deepEqual([failed.outcome, failed.reason], ["failed", "The model declined this request."]);
  const timeout = turnResult("R1.3", {
    step: { ok: false, error: "turn R1.3 still running after 600000 ms" },
    log: "",
    seconds: 600,
  });
  assert.deepEqual(
    [timeout.outcome, timeout.reason],
    ["timeout", "turn R1.3 still running after 600000 ms"],
  );
  const broken = turnResult("R1.4", {
    step: { ok: false, error: "locator.click: Timeout 30000ms exceeded" },
    log: "",
    seconds: 30,
  });
  assert.equal(broken.outcome, "failed");
  const none = turnResult("R1.5", { step: { ok: true }, log: "", seconds: 2 });
  assert.equal(none.outcome, "no_turn");
});

test("what the guards refused, the rewrites and a model replaced after a refusal are read from the log of the turn", () => {
  const log = [
    "[review] number needs author R1.1 count=2",
    "[review] number needs author R1.1 count=1",
    "[review] number needs author R1.10 count=1",
    "[review] blocked R1.1 by=L2",
    "[review] refused R1.1: task pane without a cap",
    "[review] refused R1.1: forbid rule(s) L4 do not compile",
    "[review] refused R1.1: something new",
    "[review] not a follow place R1.1 paragraph=5",
    "[review] not a follow place R1.1 paragraph=?",
    "[review] blocked R2.1 by=L2,L3",
    "[review] rewrite blocked R2.1 by=L2",
    "[review] rewrite R2.1 paragraph=12 parts=4 applied=4",
    "[review] rewrite R2.1 paragraph=14 parts=3 applied=0 stopped=ParagraphNotPlain",
    "[review] rewrite R2.1 paragraph=15 parts=3 applied=0 stopped=refused",
    "[review] rewrite R2.1 paragraph=16 parts=5 applied=2 stopped=StaleTarget",
    "[review] rewrite R2.10 paragraph=1 parts=1 applied=1",
    "[agent] model refusal fallback original=claude-opus-9 fallback=claude-opus-8 scope=session category=cyber",
    "[agent] model stays on claude-opus-8: claude-opus-9 declined 2 turns in a row (category=bio)",
    "[agent] model stays on claude-opus-8: claude-opus-9 declined 3 turns in a row (category=-)",
    "[agent] model restored to opus after the fallback to claude-opus-8",
  ].join("\n");
  assert.deepEqual(refusalsFromLog(log, "R1.1"), {
    NumberNeedsAuthor: 2,
    ForbidBlocked: 1,
    ForbidRuleBroken: 1,
    TaskPaneOutdated: 1,
    Refused: 1,
    NotAFollowPlace: 2,
  });
  // A rewrite that stopped: before its first part with the code of what
  // stopped it, after one or more parts as RewriteStopped.
  assert.deepEqual(refusalsFromLog(log, "R2.1"), {
    ForbidBlocked: 1,
    RewriteBlocked: 1,
    ParagraphNotPlain: 1,
    RewriteStopped: 2,
  });
  assert.deepEqual(refusalsFromLog(log, "R3.1"), {});
  assert.deepEqual(rewritesFromLog(log, "R2.1"), [
    { paragraph: 12, parts: 4, applied: 4, stopped: null },
    { paragraph: 14, parts: 3, applied: 0, stopped: "ParagraphNotPlain" },
    { paragraph: 15, parts: 3, applied: 0, stopped: "refused" },
    { paragraph: 16, parts: 5, applied: 2, stopped: "StaleTarget" },
  ]);
  assert.deepEqual(fallbacksFromLog(log), [
    { original: "claude-opus-9", fallback: "claude-opus-8" },
  ]);
  assert.deepEqual(fallbacksFromLog(""), []);
  assert.deepEqual(staysFromLog(log), [
    { fallback: "claude-opus-8", original: "claude-opus-9", count: 2, category: "bio" },
    { fallback: "claude-opus-8", original: "claude-opus-9", count: 3, category: null },
  ]);
  assert.equal(restoredFromLog(log), 1);
  assert.deepEqual([staysFromLog(""), restoredFromLog(undefined)], [[], 0]);
  const t = turnResult("R1.1", {
    step: { ok: true },
    log: `${log}\n[review] turn done R1.1`,
    seconds: 5,
  });
  assert.deepEqual(
    [t.outcome, t.refused.NumberNeedsAuthor, t.fallbacks.length, t.stays.length, t.rewrites],
    ["ok", 2, 1, 2, []],
  );
});

test("the batches of the request extraction are read from the log: labelled, unlabelled, failed", () => {
  const log = [
    "[review] requirement extract batch 1/3 comments=7 sentences=38 labelled=7/7",
    "[review] requirement extract batch 2/3 comments=6 sentences=40 labelled=4/6 unlabelled=R2.1,R2.4",
    "[review] requirement extract batch 3/3 comments=2 sentences=9 failed: request timed out after 120 s",
    "[review] requirement extract batch 9/9 comments=x",
  ].join("\n");
  assert.deepEqual(extractBatchesFromLog(log), [
    {
      batch: 1,
      of: 3,
      comments: 7,
      sentences: 38,
      ok: true,
      error: null,
      labelled: 7,
      unlabelled: [],
    },
    {
      batch: 2,
      of: 3,
      comments: 6,
      sentences: 40,
      ok: true,
      error: null,
      labelled: 4,
      unlabelled: ["R2.1", "R2.4"],
    },
    {
      batch: 3,
      of: 3,
      comments: 2,
      sentences: 9,
      ok: false,
      error: "request timed out after 120 s",
      labelled: null,
      unlabelled: [],
    },
  ]);
  assert.deepEqual(extractBatchesFromLog(""), []);
});

// The lines read above are the ones the daemon writes, and the codes are the
// ones its tool results carry: a rename there fails here and does not turn
// the counts of a replay into zeros.
test("the log lines read by the replay are the ones the daemon writes", () => {
  const tools = read("../daemon/review-tools.mjs");
  for (const line of [
    "`[review] number needs author ${turn.ref} count=${numbers.length}`",
    '`[review] blocked ${turn.ref} by=${open.join(",")}`',
    '`[review] refused ${turn.ref}: forbid rule(s) ${ids.join(",")} do not compile`',
    '`[review] refused ${turn.ref}: forbid rule(s) ${bad.join(",")} do not compile`',
    "`[review] refused ${turn.ref}: task pane without ${TARGET_HASH_CAP}`",
    '`[review] rewrite blocked ${turn.ref} by=${hits.join(",")}`',
    "`[review] rewrite ${turn.ref} paragraph=${index} parts=${parts.length} applied=${applied}`",
    '`${stopped ? ` stopped=${stopped.code ?? "refused"}` : ""}`',
    '`[review] not a follow place ${turn.ref} paragraph=${index ?? "?"}`',
  ])
    assert.ok(tools.includes(line), line);
  assert.equal((tools.match(/`\[review\] refused /g) ?? []).length, 3);
  for (const code of [
    "NumberNeedsAuthor",
    "ForbidRuleBroken",
    "TaskPaneOutdated",
    "RewriteBlocked",
    "RewriteStopped",
    "RewriteNotExpressible",
    "NotAFollowPlace",
  ])
    assert.ok(tools.includes(`code: "${code}"`), code);
  assert.ok(read("../taskpane/shared/tools-word.js").includes('code: "ParagraphNotPlain"'));
  assert.ok(
    read("../daemon/model-list.mjs").includes(
      "`[agent] model refusal fallback original=${msg.original_model} fallback=${msg.fallback_model} scope=",
    ),
  );
  const daemon = read("../daemon/index.mjs");
  for (const line of [
    '`[agent] model stays on ${d.fallback}: ${d.original} declined ${d.count} turns in a row (category=${d.category ?? "-"})`',
    "`[agent] model restored to ${modelArgFor(key)} after the fallback to ${live.fellBackTo}`",
  ])
    assert.ok(daemon.includes(line), line);
  const split = read("../daemon/review-split.mjs");
  for (const line of [
    "`[review] requirement extract batch ${n + 1}/${batches.length} comments=${batch.length}`",
    "` sentences=${batch.reduce((k, it) => k + it.sentences.length, 0)}`",
    "`${head} failed: ${r.error}`",
    "`${head} labelled=${batch.length - left.length}/${batch.length}`",
    '`${left.length ? ` unlabelled=${left.map((it) => it.ref).join(",")}` : ""}`',
  ])
    assert.ok(split.includes(line), line);
  assert.match(read("../daemon/review-service.mjs"), /`\[review\] turn done \$\{turn\.ref\}/);
});

test("what fails the run: another model in the pane's line without a reason", () => {
  assert.deepEqual(modelFailures(LINE), []);
  assert.deepEqual(modelFailures(STAY), []);
  assert.deepEqual(modelFailures(DRIFT), [
    "edits: asked for Claude · Opus 9, ran on Claude · Sonnet 9",
  ]);
  // The last turn of the run was the one the model was replaced in: no turn
  // came after it that could start on the chosen model again.
  assert.deepEqual(modelFailures(DRIFT, { fallbacks: [{ original: "a", fallback: "b" }] }), []);
  assert.deepEqual(modelFailures(DRIFT, { fallbacks: [] }).length, 1);
  // Writing or a judgment on another model is not excused by that.
  const writing =
    "Edits: Claude · Opus 9   Writing: asked for Claude · Opus 9, ran on Claude · Sonnet 9   Judgments: Claude · Haiku 9";
  assert.deepEqual(modelFailures(writing, { fallbacks: [{}] }), [
    "writing: asked for Claude · Opus 9, ran on Claude · Sonnet 9",
  ]);
  assert.deepEqual(modelFailures(undefined), []);
});

// The chosen model declines the manuscript and the document stays on another.
const STAY =
  "Edits: Claude · Opus 5 (Opus 5.5 declines this manuscript: bio)   Writing: Claude · Opus 9   Judgments: Claude · Haiku 9";

test("a document that stays on another model is read as that, not as another model without a reason", async () => {
  const e = readModelLine(STAY).edits;
  assert.deepEqual(e, {
    text: "Claude · Opus 5 (Opus 5.5 declines this manuscript: bio)",
    used: true,
    other: false,
    failed: false,
    declined: { on: "Claude · Opus 5", original: "Opus 5.5", category: "bio" },
  });
  assert.deepEqual(readModelLine("Edits: A · B (C declines this manuscript)").edits.declined, {
    on: "A · B",
    original: "C",
    category: null,
  });
  // The words are the ones the pane writes (model-line.js roleText).
  const { modelLineText } = await import("../taskpane/shared/model-line.js");
  const choice = { provider: "claude", model: "opus" };
  const line = modelLineText(
    {
      turn: {
        used: {
          provider: "claude",
          asked: "opus",
          ran: "claude-sonnet-9",
          declined: { stay: true, original: "opus", fallback: "sonnet", category: "bio" },
        },
        choice,
      },
      writing: { used: null, choice },
      oneshot: { used: null, choice },
    },
    null,
  );
  const d = readModelLine(line).edits;
  assert.equal(d.other, false);
  assert.equal(d.declined?.category, "bio");
  assert.deepEqual(modelFailures(line), []);
  const t = turnResult("R1.1", {
    step: { ok: true },
    log: "[review] turn done R1.1",
    seconds: 1,
    line,
  });
  assert.deepEqual([t.other_model, t.declined?.category], [false, "bio"]);
});

const LINE =
  "Edits: Claude · Opus 9   Writing: Claude · Opus 9 (not used yet)   Judgments: Claude · Haiku 9";
const DRIFT =
  "Edits: asked for Claude · Opus 9, ran on Claude · Sonnet 9   Writing: Claude · Opus 9 (last request failed)   Judgments: Claude · Haiku 9 (next: Claude · Sonnet 9)";

test("the pane's model line is read apart: used or not, another model, a failed request", () => {
  assert.deepEqual(readModelLine(LINE), {
    edits: { text: "Claude · Opus 9", used: true, other: false, failed: false, declined: null },
    writing: {
      text: "Claude · Opus 9 (not used yet)",
      used: false,
      other: false,
      failed: false,
      declined: null,
    },
    judgments: {
      text: "Claude · Haiku 9",
      used: true,
      other: false,
      failed: false,
      declined: null,
    },
  });
  const d = readModelLine(DRIFT);
  assert.deepEqual(
    [d.edits.other, d.writing.failed, d.judgments.other, d.judgments.used],
    [true, true, false, true],
  );
  assert.deepEqual(otherModels(LINE), []);
  assert.deepEqual(otherModels(DRIFT), [
    "edits: asked for Claude · Opus 9, ran on Claude · Sonnet 9",
  ]);
  assert.deepEqual(readModelLine(""), {});
  assert.deepEqual(otherModels(undefined), []);
});

// The words of the line are the ones the pane writes (model-line.js).
test("the model line the pane writes is the one this reads", async () => {
  const { modelLineText } = await import("../taskpane/shared/model-line.js");
  const line = modelLineText(
    {
      turn: {
        used: { provider: "claude", asked: "opus", ran: "claude-sonnet-9" },
        choice: { provider: "claude", model: "opus" },
      },
      writing: { used: null, choice: { provider: "claude", model: "opus" } },
      oneshot: {
        used: { provider: "ollama", asked: "tiny", ran: "tiny", ok: false },
        choice: { provider: "ollama", model: "tiny" },
      },
    },
    null,
  );
  const r = readModelLine(line);
  assert.deepEqual(Object.keys(r), ["edits", "writing", "judgments"]);
  assert.deepEqual([r.edits.other, r.writing.used, r.judgments.failed], [true, false, true]);
});

const change = (id, more = {}) => ({
  id,
  paragraph_index: 4,
  kind: "replace",
  before: "The value was high.",
  after: "The value was 12 mM.",
  rationale: {
    what: "gave the value",
    why: "the reviewer asked",
    addresses: "state the value",
    requirement_ids: ["q1"],
  },
  state: "pending",
  ...more,
});
const REVISION = {
  comments: [
    {
      ref: "R1.1",
      reviewer: 1,
      section: "Major comments",
      text: "Please state the value and the unit.",
      requirements: [
        { id: "q1", text: "state the value" },
        { id: "q2", text: "state the unit" },
        { id: "q3", text: "add the raw data as a supplement" },
        { id: "q4", text: "say why the flask was sealed" },
      ],
      questions: [
        {
          id: "a1",
          requirement_id: "q2",
          question: "Which unit did the reading have?",
          why: "the text gives none",
          asked_at: "2026-01-01T00:00:00Z",
          answer: "",
          answered_at: null,
        },
      ],
      requirements_outside: [{ id: "q3", why: "a supplementary file", at: "2026-01-01T00:00:00Z" }],
      changes: [change("R1.1-1"), change("R1.1-2", { paragraph_index: null, state: "unlocated" })],
      blocked: [{ id: "b1" }],
      response: {
        text: "We now state the value.",
        source: "model",
        citations: [
          {
            change_ids: ["R1.1-1"],
            kind: "revised",
            section: "Methods",
            paragraph_index: 4,
            quote: "The value was 12 mM.",
            lines: { start: 3, end: 3 },
            problem: null,
          },
        ],
      },
    },
    { ref: "R1.2", reviewer: 1, text: "Fine.", requirements: [], changes: [], response: null },
  ],
};
const HEADINGS = [
  { index: 0, level: 1, text: "Abstract" },
  { index: 3, level: 1, text: "Methods" },
  { index: 9, level: 1, text: "Results" },
];

test("changes.json: per comment its changes with heading and rationale, and the requirements no change names", () => {
  const turns = { "R1.1": { outcome: "ok", reason: null, seconds: 41, model: "m1" } };
  const report = changesReport(REVISION, HEADINGS, turns);
  assert.deepEqual(
    report.map((c) => c.ref),
    ["R1.1", "R1.2"],
  );
  const [a, b] = report;
  assert.deepEqual(a.changes[0], {
    change_id: "R1.1-1",
    paragraph_index: 4,
    heading: "Methods",
    kind: "replace",
    before: "The value was high.",
    after: "The value was 12 mM.",
    what: "gave the value",
    why: "the reviewer asked",
    addresses: "state the value",
    requirement_ids: ["q1"],
    state: "pending",
    rewrite_id: null,
    from_review: [],
  });
  assert.deepEqual([a.changes[1].heading, a.changes[1].state], [null, "unlocated"]);
  // Every request with its state: a change names it, a question holds it
  // up, it was declared outside this file, or nobody spoke of it.
  assert.deepEqual(
    a.requirements.map((q) => [q.id, q.state]),
    [
      ["q1", "changed"],
      ["q2", "needs_author"],
      ["q3", "outside_this_file"],
      ["q4", "open"],
    ],
  );
  assert.equal(a.requirements[2].why, "a supplementary file");
  assert.deepEqual(a.unmet_requirements, ["q2", "q3", "q4"]);
  assert.deepEqual(a.questions, [
    {
      id: "a1",
      requirement_id: "q2",
      question: "Which unit did the reading have?",
      why: "the text gives none",
      answer: "",
    },
  ]);
  assert.deepEqual(a.outside, [{ requirement_id: "q3", why: "a supplementary file" }]);
  assert.deepEqual([a.blocked_edits, a.turn, a.section], [1, turns["R1.1"], "Major comments"]);
  assert.deepEqual(
    [
      b.changes,
      b.unmet_requirements,
      b.turn,
      b.section,
      b.questions,
      b.outside,
      b.still_present,
      b.refused,
    ],
    [[], [], null, null, [], [], [], {}],
  );
  assert.deepEqual(
    [b.named_terms, b.same_term_at, b.rewrites, b.paragraph_findings, b.rewrite_turns],
    [[], [], [], [], []],
  );
  // What the guards refused in the turn of the comment.
  const guarded = changesReport(REVISION, HEADINGS, {
    "R1.1": { outcome: "ok", refused: { NumberNeedsAuthor: 2 } },
  });
  assert.deepEqual(guarded[0].refused, { NumberNeedsAuthor: 2 });
  assert.deepEqual(changesReport({}, [], {}), []);
});

test("changes.json: a word a change took out that other paragraphs still hold", () => {
  const revision = {
    snapshot: {
      paragraphs: [
        "The effect was convergent in every line.",
        "The effect was shared in every line.",
        "A convergent pattern is discussed below.",
        "Nothing of it here.",
        "Nor here.",
      ],
      ids: null,
    },
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "Do not call it convergent.",
        requirements: [],
        // The comment and the reason of the change both name the word
        // (residue.js claimTerms); a word only reworded away gives no row.
        changes: [
          change("R1.1-1", {
            find: "convergent",
            text: "shared",
            paragraph_index: 1,
            rationale: {
              what: "reworded",
              why: "convergent is not shown",
              addresses: "the wording",
            },
          }),
        ],
      },
    ],
  };
  const [c] = changesReport(revision, [], {});
  assert.deepEqual(c.still_present, [
    { term: "convergent", change_ids: ["R1.1-1"], places: [0, 2] },
  ]);
});

// A comment that names a term, a paragraph written anew by a rewrite, a
// change with a quantity from the review, and the paragraph check.
const STACKED = {
  snapshot: {
    paragraphs: [
      "Figure 2. A convergent pattern in the invented lines.",
      "The effect was shared in every line of the invented set.",
      "Nothing of it here.",
    ],
    ids: ["a1", "a2", "a3"],
  },
  consistency: {
    checked_at: "2026-01-01T00:00:00Z",
    model: "m1",
    paragraphs: [{ index: 1, id: "a2", change_ids: ["R1.1-1", "R1.1-2", "R1.1-3"], many: true }],
    findings: [
      {
        paragraph_index: 1,
        paragraph_id: "a2",
        kind: "claim_left",
        detail: "the claim is made once more",
        change_ids: ["R1.1-2"],
      },
      {
        paragraph_index: 2,
        paragraph_id: "a3",
        kind: "open_question",
        detail: "x",
        change_ids: [],
      },
    ],
    warnings: [],
  },
  comments: [
    {
      ref: "R1.1",
      reviewer: 1,
      text: 'The term "convergent" is not justified.',
      requirements: [],
      named_terms: [],
      same_term_at: ["a1"],
      changes: [
        change("R1.1-1", { paragraph_index: 1 }),
        change("R1.1-2", { paragraph_index: 1, rewrite_id: "R1.1-wabc", from_review: ["12 mM"] }),
        change("R1.1-3", { paragraph_index: 1, rewrite_id: "R1.1-wabc" }),
      ],
    },
  ],
};
const REWRITE_TURN = {
  ref: "R1.1",
  paragraph_index: 1,
  outcome: "ok",
  seconds: 20,
  rewrites: [{ paragraph: 1, parts: 2, applied: 2, stopped: null }],
  refused: { RewriteBlocked: 1 },
};

test("changes.json: named terms and where they stand, rewrites, words from the review, paragraph findings", () => {
  const [c] = changesReport(STACKED, [], {}, [REWRITE_TURN, { ref: "R9.9" }]);
  assert.deepEqual(c.named_terms, [
    { term: "convergent", places: [{ index: 0, id: "a1", kind: "caption" }] },
  ]);
  assert.deepEqual(c.same_term_at, ["a1"]);
  assert.deepEqual(c.rewrites, [
    { rewrite_id: "R1.1-wabc", paragraph_index: 1, change_ids: ["R1.1-2", "R1.1-3"] },
  ]);
  assert.deepEqual(
    c.changes.map((x) => [x.rewrite_id, x.from_review]),
    [
      [null, []],
      ["R1.1-wabc", ["12 mM"]],
      ["R1.1-wabc", []],
    ],
  );
  // The findings on a change of this comment; the others are in
  // paragraph_check.json only.
  assert.deepEqual(c.paragraph_findings, [
    {
      kind: "claim_left",
      paragraph_index: 1,
      paragraph_id: "a2",
      detail: "the claim is made once more",
      change_ids: ["R1.1-2"],
    },
  ]);
  assert.deepEqual(c.rewrite_turns, [REWRITE_TURN]);
});

test("responses.json: the answer and what it cites, for the comments that have one", () => {
  assert.deepEqual(responsesReport(REVISION), [
    {
      ref: "R1.1",
      text: "We now state the value.",
      source: "model",
      citations: [
        {
          change_ids: ["R1.1-1"],
          kind: "revised",
          section: "Methods",
          paragraph_index: 4,
          quote: "The value was 12 mM.",
          problem: null,
        },
      ],
    },
  ]);
});

test("summary.txt: counts, states, refusals, models and times; no text of the manuscript", () => {
  const turn = (more) => ({
    reason: null,
    model: "m1",
    model_line: "Claude · Opus 9",
    other_model: false,
    fallbacks: [],
    refused: {},
    ...more,
  });
  const last = turn({
    outcome: "timeout",
    reason: "x",
    seconds: 600,
    other_model: true,
    fallbacks: [{ original: "o", fallback: "f" }],
  });
  const report = changesReport(REVISION, HEADINGS, {
    "R1.1": turn({
      outcome: "ok",
      seconds: 41,
      refused: { NumberNeedsAuthor: 2, ForbidBlocked: 1 },
    }),
    "R1.2": last,
  });
  const text = summaryText(report, {
    seconds: 700,
    saved: true,
    models: { asked: { edits: "opus", writing: null, judgments: "haiku" }, line: LINE },
    source: "opened as it is",
    responses: 1,
    paragraphCheck: {
      paragraphs: [{}, {}, {}],
      findings: [{ kind: "conflict" }, { kind: "conflict" }],
    },
  });
  assert.equal(
    text,
    [
      "comments: 2",
      "turns run: 2",
      "changes: 2",
      "comments without a change: 1",
      "request extraction: no batch in the log",
      "comments with extracted requests: 1; without: 1",
      "requests by state: changed 1, needs_author 1, outside_this_file 1, open 1",
      "requirements no change names: 3",
      "questions for the author: 1",
      "requests declared outside this file: 1",
      "terms still present elsewhere: 0",
      "terms the reviewers named: 0 in 0 comments; still in the text: 0",
      "changes with words from the review: 0",
      "paragraphs written anew: 0 rewrites of 0 paragraphs, 0 parts recorded; stopped: 0",
      "rewrite turns run: 0",
      "follow-through turns run: 0",
      "follow-through: changes 0 in 0 paragraphs, questions 0",
      "edits a guard refused: NumberNeedsAuthor 2, ForbidBlocked 1",
      "turns that did not end well: 1 (R1.2 timeout)",
      "models asked for: edits opus, writing (the pane's own), judgments haiku",
      `models that ran, as the pane says: ${LINE}`,
      "models that ran, from the daemon log: m1",
      "model replaced after a refusal: 1 times in 1 turns (R1.2); restored 0 times",
      "turns on another model than asked: 1 (R1.2)",
      "another model than asked: no",
      "paragraph check: 2 findings in 3 paragraphs (conflict 2)",
      "turn time: 641 s",
      "total time: 700 s",
      "copy saved with its tracked changes: yes",
      "manuscript: opened as it is",
      "responses drafted: 1",
      "",
    ].join("\n"),
  );
  assert.ok(!text.includes("value"));
  assert.ok(!text.includes("unit"));
  // Another model in the line at the end, without a reason, fails the run;
  // not the edits when the last turn was the one the model was replaced in.
  const models = (line) => ({ asked: { edits: "opus" }, line });
  assert.match(
    summaryText(report, { seconds: 1, saved: true, models: models(DRIFT) }),
    /another model than asked: FAILED \(edits: asked for Claude · Opus 9, ran on Claude · Sonnet 9\)\n/,
  );
  assert.match(
    summaryText(report, { seconds: 1, saved: true, models: models(DRIFT), lastTurn: last }),
    /another model than asked: no\n/,
  );
  // A document that stays on another model: said, and no failure. The words
  // come from the pane's line, or from the log when the line does not say it.
  const stay = summaryText(report, { seconds: 1, saved: true, models: models(STAY), restored: 1 });
  assert.match(
    stay,
    /; restored 1 times\nedits stayed on Claude · Opus 5 because Opus 5\.5 declines this manuscript \(bio\)\nturns on another model than asked: 1 \(R1\.2\)\nanother model than asked: no\n/,
  );
  assert.match(
    summaryText(report, {
      seconds: 1,
      saved: true,
      models: models(LINE),
      stays: [{ fallback: "f1", original: "o1", count: 2, category: null }],
    }),
    /\nedits stayed on f1 because o1 declines this manuscript \(-\)\n/,
  );
  assert.doesNotMatch(text, /edits stayed on/);
  // A run in which no request was extracted does not claim that none is unmet.
  const none = changesReport(
    { comments: [{ ref: "R1.1", reviewer: 1, text: "x", requirements: [], changes: [] }] },
    [],
    {},
  );
  const bare = summaryText(none, {
    seconds: 3,
    saved: false,
    models: { asked: { edits: "opus" }, line: "" },
  });
  assert.match(bare, /comments with extracted requests: 0; without: 1\n/);
  assert.match(bare, /requirements no change names: not measured \(no requests were extracted\)\n/);
  assert.match(bare, /paragraph check: not run\n/);
  assert.match(bare, /copy saved with its tracked changes: NO\n$/);
  assert.match(bare, /models that ran, as the pane says: -\n/);
});

test("summary.txt: the batches of the extraction, named terms, rewrites and the kinds of the paragraph check", () => {
  const report = changesReport(
    STACKED,
    [],
    { "R1.1": { outcome: "ok", seconds: 30, rewrites: [], refused: { NumberNeedsAuthor: 1 } } },
    [
      REWRITE_TURN,
      {
        ref: "R1.1",
        paragraph_index: 4,
        outcome: "timeout",
        seconds: 5,
        rewrites: [{ paragraph: 4, parts: 3, applied: 1, stopped: "StaleTarget" }],
        refused: { RewriteStopped: 1 },
      },
    ],
  );
  const text = summaryText(report, {
    seconds: 60,
    saved: true,
    models: { asked: { edits: "opus" }, line: LINE },
    paragraphCheck: STACKED.consistency,
    extraction: [
      {
        batch: 1,
        of: 2,
        comments: 7,
        sentences: 38,
        ok: true,
        error: null,
        labelled: 6,
        unlabelled: ["R2.1"],
      },
      {
        batch: 2,
        of: 2,
        comments: 2,
        sentences: 9,
        ok: false,
        error: "timed out",
        labelled: null,
        unlabelled: [],
      },
    ],
  });
  for (const line of [
    "request extraction: 2 batches, ok 1, failed 1",
    "  batch 1/2: 7 comments, 38 sentences, labelled 6/7, unlabelled R2.1",
    "  batch 2/2: 2 comments, 9 sentences, FAILED: timed out",
    "terms the reviewers named: 1 in 1 comments; still in the text: 1",
    "changes with words from the review: 1",
    "paragraphs written anew: 1 rewrites of 1 paragraphs, 2 parts recorded; stopped: 1",
    "rewrite turns run: 2",
    "edits a guard refused: NumberNeedsAuthor 1, RewriteBlocked 1, RewriteStopped 1",
    "turns that did not end well: 1 (R1.1 rewrite p4 timeout)",
    "paragraph check: 2 findings in 1 paragraphs (claim_left 1, open_question 1)",
    "turn time: 55 s",
  ])
    assert.ok(text.includes(`\n${line}\n`), line);
  // No term, no word of the review and no finding text is in it.
  for (const word of ["convergent", "12 mM", "once more"]) assert.ok(!text.includes(word), word);
});

// --follow-through: a second turn of a comment for the places its changes
// did not touch. A made-up bridge inspection report.
const FOLLOWED = {
  snapshot: {
    paragraphs: [
      "Deck survey",
      "Rust was reduced on every span (Fig. 2).",
      "The coating held.",
      "Figure 2. Rust cover per span.",
    ],
    ids: ["k0", "k1", "k2", "k3"],
  },
  comments: [
    {
      ref: "R1.1",
      reviewer: 1,
      text: "Too strong.",
      requirements: [],
      changes: [
        change("R1.1-1", { paragraph_index: 1 }),
        change("R1.1-2", { paragraph_index: 3 }),
        change("R1.1-3", { paragraph_index: 3 }),
      ],
      questions: [{ id: "a1", question: "Which span?", why: "Not said." }],
      follow_through: {
        at: "2026-01-02T00:00:00.000Z",
        // p9: an id the snapshot does not hold is read as an index.
        candidate_ids: ["k0", "k3", "k2", "p9", "gone"],
        change_ids: ["R1.1-2", "R1.1-3"],
        questions: 1,
        seen_change_ids: ["R1.1-1", "R1.1-2", "R1.1-3"],
      },
    },
    { ref: "R1.2", reviewer: 1, text: "Name the supplier.", requirements: [], changes: [] },
  ],
};
const FOLLOW_TURN = {
  ref: "R1.1",
  outcome: "ok",
  seconds: 40,
  refused: { NumberNeedsAuthor: 1 },
  follow: { candidates: 5, changes: 2, questions: 1 },
};

test("the end of a follow-through turn is read from the log of the turn, as the daemon writes it", () => {
  const log = [
    "[review] follow R1.1 candidates=5 changes=2 questions=1",
    "[review] follow R1.10 candidates=9 changes=9 questions=9",
    "[review] turn done R1.1",
  ].join("\n");
  assert.deepEqual(followFromLog(log, "R1.1"), { candidates: 5, changes: 2, questions: 1 });
  assert.equal(followFromLog(log, "R1.2"), null);
  const t = turnResult("R1.1", { step: { ok: true }, log, seconds: 40 });
  assert.deepEqual(t.follow, { candidates: 5, changes: 2, questions: 1 });
  // A turn that was no follow-through turn has no such key.
  const plain = turnResult("R1.2", { step: { ok: true }, log: "[review] turn done R1.2" });
  assert.ok(!("follow" in plain));
  assert.ok(
    read("../daemon/review-service.mjs").includes(
      "`[review] follow ${turn.ref} candidates=${candidate_ids.length} changes=${added.changes} questions=${added.questions}`",
    ),
  );
});

test("changes.json and summary.txt: what the follow-through turns did", () => {
  const report = changesReport(
    FOLLOWED,
    [],
    { "R1.1": { outcome: "ok", seconds: 30 } },
    [],
    [FOLLOW_TURN, { ref: "R9.9" }],
  );
  assert.deepEqual(report[0].follow, {
    candidates: [0, 3, 2, 9],
    changes: ["R1.1-2", "R1.1-3"],
    questions: 1,
  });
  assert.deepEqual(report[0].follow_turns, [FOLLOW_TURN]);
  assert.deepEqual([report[1].follow, report[1].follow_turns], [null, []]);
  const text = summaryText(report, { seconds: 90, saved: true, models: { asked: {}, line: "" } });
  for (const line of [
    "follow-through turns run: 1",
    "follow-through: changes 2 in 1 paragraphs, questions 1",
    "edits a guard refused: NumberNeedsAuthor 1",
    "turns that did not end well: 0",
    "turn time: 70 s",
  ])
    assert.ok(text.includes(`\n${line}\n`), `${line}\n${text}`);
  // A turn that did not end well is named; a run without the flag says 0.
  const bad = changesReport(FOLLOWED, [], {}, [], [{ ...FOLLOW_TURN, outcome: "timeout" }]);
  assert.ok(
    summaryText(bad, { seconds: 1, saved: true, models: {} }).includes(
      "\nturns that did not end well: 1 (R1.1 follow-through timeout)\n",
    ),
  );
  const off = summaryText(changesReport({ comments: [FOLLOWED.comments[1]] }), {
    seconds: 1,
    saved: true,
    models: {},
  });
  assert.ok(off.includes("\nfollow-through turns run: 0\n"));
  assert.ok(off.includes("\nfollow-through: changes 0 in 0 paragraphs, questions 0\n"));
});

test("replay.mjs asks only for pane steps and Windows scripts that exist, and names no path of its own", () => {
  const runner = read("../scripts/e2e/replay.mjs");
  const base = read("../scripts/e2e/pane-driver.cjs");
  const extra = read("../scripts/e2e/replay-driver.cjs");
  const steps = [...new Set([...runner.matchAll(/\bpane\(\s*"([a-z-]+)"/g)].map((m) => m[1]))];
  assert.ok(steps.length >= 10, steps.join(","));
  const has = (src, s) => new RegExp(`\\n  (?:async )?"?${s}"?(?:\\(page|: )`).test(src);
  for (const s of steps) assert.ok(has(base, s) || has(extra, s), `pane step ${s}`);
  for (const s of ["set-model", "models-line", "headings", "stop-turn"])
    assert.ok(has(extra, s), s);
  // The paragraph check runs after the last turn and its refresh, and its
  // result is written out.
  assert.ok(steps.includes("check-paragraphs"));
  // --rewrite-stacked: off unless asked for; the rewrite turns come after
  // the paragraph check, each button once, and a refresh follows them.
  assert.match(runner, /"rewrite-stacked": \{ type: "boolean", default: false \}/);
  const stacked = runner.indexOf('if (opt["rewrite-stacked"]) {');
  assert.ok(runner.indexOf('pane("check-paragraphs"') < stacked);
  assert.ok(stacked < runner.indexOf('pane("rewrite-paragraph", { nth }'));
  assert.ok(runner.indexOf('pane("rewrite-paragraph"') < runner.lastIndexOf('pane("refresh")'));
  assert.equal((runner.match(/pane\("rewrite-paragraph"/g) ?? []).length, 1);
  // --follow-through: off unless asked for; the turns come after the turns
  // of the comments and their refresh and before the paragraph check, which
  // then reads what they left; one per comment the card offers the check
  // for (follow-places.js followState), and a refresh follows them.
  assert.match(runner, /"follow-through": \{ type: "boolean", default: false \}/);
  const follow = runner.indexOf('if (opt["follow-through"]) {');
  assert.ok(runner.indexOf('pane("refresh")') < follow);
  assert.ok(follow < runner.indexOf('pane("follow-through", { ref }'));
  assert.ok(runner.indexOf('pane("follow-through"') < runner.indexOf('pane("check-paragraphs"'));
  assert.equal((runner.match(/pane\("follow-through"/g) ?? []).length, 1);
  const between = runner.slice(follow, runner.indexOf('pane("check-paragraphs"'));
  assert.match(between, /followState\(/);
  assert.match(between, /pane\("refresh"\)/);
  assert.match(base, /"Check other places"/);
  assert.match(runner, /changesReport\(revision, headings, turns, rewriteTurns, followTurns\)/);
  // The exit code follows the pane's line (modelFailures), nothing else.
  assert.match(
    runner,
    /const otherModel = \(\) => modelFailures\(modelsLine, lastTurn\)\.length > 0;/,
  );
  assert.ok(runner.lastIndexOf('pane("propose"') < runner.indexOf('pane("refresh")'));
  assert.ok(runner.indexOf('pane("refresh")') < runner.indexOf('pane("check-paragraphs"'));
  assert.match(runner, /"paragraph_check\.json"/);
  assert.match(runner, /paragraphCheck: revision\.consistency/);
  // The models are chosen in Setup, not in the menu of the chat box (it is
  // not shown on the Review tab).
  assert.doesNotMatch(extra, /selectOption\("#composer-model"/);
  assert.match(extra, /selectOption\("#provider-turn-model"/);
  const html = read("../taskpane/word/index.html");
  for (const id of ["provider-turn", "provider-turn-model", "models-line", "stop-agent"])
    assert.ok(html.includes(`id="${id}"`), id);
  // The files copied to the run folder are in the repository.
  for (const [, name] of runner.matchAll(/"((?:[a-z-]+)\.(?:cjs|ps1))"/g))
    assert.doesNotThrow(() => read(`../scripts/e2e/${name}`), name);
  // The copy is saved only by the script that checks the folder, and Word is
  // closed only through the quit of word-com.ps1.
  const save = read("../scripts/e2e/replay-save.ps1");
  assert.match(save, /\$d = \[SwRot\]::Find\(\$Doc\)/);
  assert.match(save, /StartsWith\(\$allow/);
  assert.ok(save.indexOf("StartsWith($allow") < save.indexOf("$d.Save()"));
  // eslint-disable-next-line no-control-regex
  assert.doesNotMatch(save, /[^\x00-\x7f]/);
  assert.match(runner, /"-Action",\s*"quit"/);
  // The open check: a Word of its own (word-own.ps1), nothing opened or
  // saved outside the run folder, the given file never written.
  const open = read("../scripts/e2e/replay-open.ps1");
  // eslint-disable-next-line no-control-regex
  assert.doesNotMatch(open, /[^\x00-\x7f]/);
  assert.match(open, /word-own\.ps1/);
  assert.match(open, /QuitOwnWord \$app \$anchorDoc \$p \$ours/);
  assert.ok(open.indexOf("(Under $File)") < open.indexOf("OpenReadOnly $app $File"));
  assert.ok(open.indexOf("(Under $RepairTo)") < open.indexOf("SaveAs2("));
  // The plain open is the one call of word-own.ps1 that the Word judge of
  // the marked copies uses too: one place, one meaning of "Word opens it".
  const own = read("../scripts/e2e/word-own.ps1");
  assert.match(
    own,
    /function OpenReadOnly\(\$app, \[string\]\$path\) \{\s*return \$app\.Documents\.Open\(\$path, \$false, \$true, \$false\)\s*\}/,
  );
  const judge = read("../scripts/e2e/marked-diff-word.ps1");
  assert.match(judge, /\$doc = OpenReadOnly \$app \$job\.path/);
  assert.doesNotMatch(judge, /Documents\.Open\(/);
  // The only Documents.Open written out here is the one with Open and Repair.
  assert.equal((open.match(/Documents\.Open\(/g) ?? []).length, 1);
  assert.match(open, /Documents\.Open\(\$File, \$false, \$true, \$false, (?:\$m, ){8}\$true\)/);
  // An open that gives nothing usable is not "opens".
  assert.match(open, /\$opened\.FullName -ne \$File/);
  assert.match(open, /\$count -lt 1/);
  // PowerShell does not tell $Doc from $doc (that made the first open check
  // pass on a file Word cannot open): in the scripts of the replay every
  // variable has one spelling.
  for (const src of [open, save]) {
    const spellings = new Map();
    for (const [, name] of src.matchAll(/\$([A-Za-z_][A-Za-z0-9_]*)/g)) {
      const key = name.toLowerCase();
      if (!spellings.has(key)) spellings.set(key, new Set());
      spellings.get(key).add(name);
    }
    for (const [key, set] of spellings)
      assert.equal(set.size, 1, `$${key} is written as ${[...set].join(" and ")}`);
  }
  assert.equal((open.match(/SaveAs2\(/g) ?? []).length, 2);
  assert.doesNotMatch(open, /\$doc\.Save\(\)/);
  // The runner copies the manuscript first and gives Word only copies; a
  // file Word must repair stops the run unless --repair is given.
  assert.ok(
    runner.indexOf("fs.copyFileSync(opt.manuscript, SOURCE)") <
      runner.indexOf("prepareCopy()", runner.indexOf("async function setup")),
  );
  // What is asked about is the file the run opens (the add-in tagged on),
  // also after a repair; and the resident daemon is paused before it.
  const prep = runner.slice(
    runner.indexOf("function prepareCopy()"),
    runner.indexOf("// ---- the run"),
  );
  assert.deepEqual(
    [...prep.matchAll(/wordOpens\(([^)]*)\)/g)].map((m) => m[1]),
    ["DOC", "SOURCE, REPAIRED", "DOC"],
  );
  assert.deepEqual(
    [...prep.matchAll(/embed\((\w+)\)/g)].map((m) => m[1]),
    ["SOURCE", "REPAIRED"],
  );
  const setupAt = runner.indexOf("async function setup");
  assert.ok(runner.indexOf("pauseResident()", setupAt) < runner.indexOf("prepareCopy()", setupAt));
  assert.ok(
    runner.indexOf("prepareCopy()", setupAt) < runner.indexOf("await startDaemon()", setupAt),
  );
  // Read in two places only: the check of its name, and that copy.
  assert.equal((runner.match(/opt\.manuscript/g) ?? []).length, 2);
  assert.match(runner, /run again with --repair/);
  assert.match(runner, /pauseResident\(\)/);
  // No input path is written into the code: they come as arguments.
  for (const src of [runner, extra, save, read("../scripts/e2e/replay-compare.py")])
    assert.doesNotMatch(src, /\/mnt\/[a-z]\/|[A-Z]:\\\\(?!Program Files)/);
});

// The sweep at the end of a run. Pids, times and titles are invented.
const word = (pid, started, title = "Document1 - Word") => ({ pid, started, title });
const T = "2026-10-05T10:05:00.1234567Z";

test("sweep: a concurrent user-created blank Word recorded as a stray survives", () => {
  const { close, skip } = sweepSelection({
    processes: [word(4172, T), word(900, "2026-10-05T09:00:00Z", "thesis.docx - Word")],
    strays: newWords([word(900, T)], [word(900, T), word(4172, T)]),
    preexisting: [900],
  });
  assert.deepEqual(
    close.map((p) => p.pid),
    [],
  );
  assert.deepEqual(
    skip.map((p) => [p.pid, p.reason]),
    [
      [4172, "left open: ownership unproven"],
      [900, "ran before the run"],
    ],
  );
  // Korean Word, and a stray with no window at all.
  assert.equal(
    sweepSelection({ processes: [word(5, T, "문서1 - Word")], strays: [word(5, T)] }).close.length,
    0,
  );
  assert.equal(
    sweepSelection({ processes: [word(6, T, "")], strays: [word(6, T)] }).close.length,
    0,
  );
});

test("sweep: a Word that ran before the run is never closed, also when it is a stray", () => {
  const { close, skip } = sweepSelection({
    processes: [word(4172, T)],
    strays: [word(4172, T)],
    preexisting: [4172],
  });
  assert.deepEqual(close, []);
  assert.deepEqual(
    skip.map((p) => p.reason),
    ["ran before the run"],
  );
});

test("sweep: a Word that started during the run but is no recorded stray is left open", () => {
  // The user opened Word during the run, also a blank one: not the run's.
  const { close, skip } = sweepSelection({
    processes: [word(7, T), word(8, T, "notes.docx - Word")],
    strays: [],
    preexisting: [],
  });
  assert.deepEqual(close, []);
  assert.deepEqual(
    skip.map((p) => [p.pid, p.reason]),
    [
      [7, "left open: not started by this run"],
      [8, "left open: not started by this run"],
    ],
  );
});

test("sweep: a stray whose title names a real document is left open", () => {
  const { close, skip } = sweepSelection({
    processes: [word(9, T, "thesis.docx - Word"), word(10, T, "Document1 notes - Word")],
    strays: [word(9, T), word(10, T)],
  });
  assert.deepEqual(close, []);
  assert.deepEqual(
    skip.map((p) => [p.pid, p.reason]),
    [
      [9, "left open: ownership unproven"],
      [10, "left open: ownership unproven"],
    ],
  );
});

test("sweep: kept, a reused pid or an unknown start time is left open", () => {
  const { close, skip } = sweepSelection({
    processes: [word(1, T), word(2, "2026-10-05T11:00:00Z"), word(3, null)],
    strays: [word(1, T), word(2, T), word(3, null)],
    keep: [1],
  });
  assert.deepEqual(close, []);
  assert.deepEqual(
    skip.map((p) => [p.pid, p.reason]),
    [
      [1, "kept by the run"],
      [2, "left open: not started by this run"],
      [3, "left open: not started by this run"],
    ],
  );
});

test("sweep: new Words of a step are the ones not listed before it, without the step's own", () => {
  const before = [word(900, "2026-10-05T09:00:00Z", "thesis.docx - Word")];
  const after = [...before, word(21400, T, "anchor.docx - Word"), word(4172, T)];
  assert.deepEqual(
    newWords(before, after, 21400).map((p) => p.pid),
    [4172],
  );
  assert.deepEqual(newWords(before, before), []);
  assert.ok(blankWordTitle("Document12 - Word"));
  assert.ok(!blankWordTitle("thesis.docx - Word"));
});

test("replay: the sweep runs last, reports only, and the Word steps record their strays", () => {
  const runner = read("../scripts/e2e/replay.mjs");
  const sweep = runner.slice(
    runner.indexOf("function sweepWord()"),
    runner.indexOf("// One run of replay-open.ps1"),
  );
  assert.doesNotMatch(runner, /CloseMainWindow|Stop-Process|stopOwn/);
  assert.doesNotMatch(sweep, /Stop-Process|-Force|\.Kill\(|taskkill/);
  assert.match(sweep, /strays: record\.word\.strays \?\? \[\]/);
  // The Word kept by --keep-on-fail and one left for a refused quit stay.
  assert.match(sweep, /record\.cleanup\.kept\?\.word_pid, record\.cleanup\.left_running\?\.pid/);
  const last = runner.lastIndexOf("sweepWord();");
  assert.ok(last > runner.lastIndexOf("quitWord(record.word.launch?.pid)"));
  assert.ok(last < runner.indexOf("record.finished_at ="));
  // Every step that starts or binds Word is listed around.
  assert.match(runner, /withStrays\(\s*`open check \$\{checks\}`/);
  assert.match(runner, /withStrays\(\s*"word launch"/);
  assert.match(runner, /withStrays\(\s*"save copy"/);
  assert.equal((runner.match(/withStrays\(/g) ?? []).length, 4);
  const opens = runner.slice(
    runner.indexOf("function wordOpens("),
    runner.indexOf("const embed ="),
  );
  assert.ok(opens.indexOf("withStrays(") < opens.indexOf('"replay-open.ps1"'));
  assert.ok(opens.indexOf("withStrays(") < opens.indexOf("open_checks ??= []"));
});

// The leak of an open check: BindToMoniker on a document that Word has not
// registered yet starts a WINWORD of its own. The Word scripts look in the
// Running Object Table only.
test("Word scripts find a document in the Running Object Table only, never by a bind that starts Word", () => {
  const code = (src) => src.replace(/^\s*#.*$/gm, "");
  for (const name of ["word-own.ps1", "word-com.ps1", "replay-save.ps1"]) {
    const src = read(`../scripts/e2e/${name}`);
    assert.doesNotMatch(
      code(src),
      /BindToMoniker|GetActiveObject|New-Object\s+-ComObject|CreateObject/i,
      name,
    );
    assert.match(src, /\. \(Join-Path \$PSScriptRoot "word-rot\.ps1"\)/, name);
    assert.match(code(src), /\[SwRot\]::Find\(/, name);
  }
  const rot = read("../scripts/e2e/word-rot.ps1");
  assert.match(rot, /rot\.EnumRunning/);
  assert.match(rot, /rot\.GetObject\(/);
  assert.doesNotMatch(code(rot), /BindToMoniker|CoCreateInstance|BindToObject/);
  // Every runner that copies a script dot-sourcing word-rot.ps1 copies it too.
  for (const runner of [
    "replay.mjs",
    "run-acceptance.mjs",
    "marked-diff-word-check.mjs",
    "word-math-probe.mjs",
  ])
    assert.match(read(`../scripts/e2e/${runner}`), /"word-rot\.ps1"/, runner);
});

// The Word of a run is sent behind the user's windows without the focus and
// is never minimized: with Word minimized the edits of a turn lost their
// rationale comments (acceptance run of 2026-10-05, stage 3). Only the
// process the run started is touched, by its id.
test("the Word windows a run starts go to the back of the z-order, never minimized", () => {
  const code = (src) => src.replace(/^\s*#.*$/gm, "");
  const back = code(read("../scripts/e2e/word-back.ps1"));
  // HWND_BOTTOM (1), SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE (0x13).
  assert.match(back, /SetWindowPos\(\$h, \[System\.IntPtr\]1, 0, 0, 0, 0, 0x13\)/);
  assert.doesNotMatch(back, /ShowWindow|CloseMainWindow|Stop-Process|SW_MIN|SetForegroundWindow/i);
  // No list of every Word: one process, by the id it was given.
  assert.doesNotMatch(back, /Get-Process\s+WINWORD/i);
  assert.match(back, /Get-Process -Id \$WatchProcessId/);
  for (const name of ["word-launch.ps1", "word-own.ps1"]) {
    const src = code(read(`../scripts/e2e/${name}`));
    assert.match(src, /\. \(Join-Path \$PSScriptRoot "word-back\.ps1"\)/, name);
    assert.match(src, /SendWordToBack \$p\.Id/, name);
    assert.match(src, /StartWordBackWatch \$p\.Id/, name);
  }
  // Only once this run's own Word is known: not a Word that ran before.
  assert.match(
    code(read("../scripts/e2e/word-launch.ps1")),
    /if \(-not \$p\.HasExited -and \$before -notcontains \$p\.Id\) \{\s*\. \(Join-Path/,
  );
  for (const runner of [
    "replay.mjs",
    "run-acceptance.mjs",
    "marked-diff-word-check.mjs",
    "word-math-probe.mjs",
  ])
    assert.match(read(`../scripts/e2e/${runner}`), /"word-back\.ps1"/, runner);
});

// Parse check of the changed PowerShell scripts. No Word: the parser only
// reads the files. Runs where powershell.exe is reachable (WSL on Windows).
const winPs = (() => {
  try {
    execWithInteropRetry(
      () => execFileSync("powershell.exe", ["-NoProfile", "-Command", "exit 0"], { stdio: "pipe" }),
      { what: "powershell.exe", log: () => {} },
    );
    return true;
  } catch {
    return false;
  }
})();
test(
  "the changed PowerShell scripts parse",
  { skip: !winPs && "powershell.exe not reachable" },
  () => {
    // powershell.exe started from WSL can fail to start under load
    // (UtilAcceptVsock accept4 failed 110): tried once more (wsl-interop.mjs).
    const ps = (command) =>
      execWithInteropRetry(
        () =>
          execFileSync("powershell.exe", ["-NoProfile", "-Command", command], { encoding: "utf8" }),
        { what: "powershell.exe", log: () => {} },
      ).trim();
    const win = (name) =>
      execFileSync(
        "wslpath",
        ["-w", fileURLToPath(new URL(`../scripts/e2e/${name}`, import.meta.url))],
        { encoding: "utf8" },
      ).trim();
    for (const name of [
      "word-rot.ps1",
      "word-own.ps1",
      "word-com.ps1",
      "replay-save.ps1",
      "word-back.ps1",
      "word-launch.ps1",
    ]) {
      const out = ps(
        `$e = $null; [void][System.Management.Automation.Language.Parser]::ParseFile('${win(name)}', [ref]$null, [ref]$e); ($e | ForEach-Object { $_.Message }) -join '; '`,
      );
      assert.equal(out, "", `${name}: ${out}`);
    }
    // The C# of word-rot.ps1 and word-back.ps1 compiles. Dot-sourcing them
    // only adds the types: nothing is called, no window or Word is touched.
    assert.equal(ps(`. '${win("word-rot.ps1")}'; [bool]('SwRot' -as [type])`), "True");
    assert.equal(ps(`. '${win("word-back.ps1")}'; [bool]('SwBack.User32' -as [type])`), "True");
  },
);
