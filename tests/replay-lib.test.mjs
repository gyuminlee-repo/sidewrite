// What scripts/e2e/replay.mjs computes without Word (replay-lib.mjs), and
// that what it asks of the driver and the Windows scripts exists. The run
// itself needs Word and does not run here. Every text is invented.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import {
  changesReport,
  keepRefPattern,
  ledgerFromLines,
  otherModels,
  parseOnly,
  readModelLine,
  responsesReport,
  ruleLines,
  summaryText,
  turnFromLog,
  turnResult,
} from "../scripts/e2e/replay-lib.mjs";
import { LedgerItemSchema } from "../daemon/revision-store.mjs";

const read = (p) => readFileSync(new URL(p, import.meta.url), "utf8");

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

const LINE =
  "Edits: Claude · Opus 9   Writing: Claude · Opus 9 (not used yet)   Judgments: Claude · Haiku 9";
const DRIFT =
  "Edits: asked for Claude · Opus 9, ran on Claude · Sonnet 9   Writing: Claude · Opus 9 (last request failed)   Judgments: Claude · Haiku 9 (next: Claude · Sonnet 9)";

test("the pane's model line is read apart: used or not, another model, a failed request", () => {
  assert.deepEqual(readModelLine(LINE), {
    edits: { text: "Claude · Opus 9", used: true, other: false, failed: false },
    writing: { text: "Claude · Opus 9 (not used yet)", used: false, other: false, failed: false },
    judgments: { text: "Claude · Haiku 9", used: true, other: false, failed: false },
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
      ],
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
  });
  assert.deepEqual([a.changes[1].heading, a.changes[1].state], [null, "unlocated"]);
  assert.deepEqual(a.unmet_requirements, ["q2"]);
  assert.deepEqual([a.blocked_edits, a.turn, a.section], [1, turns["R1.1"], "Major comments"]);
  assert.deepEqual([b.changes, b.unmet_requirements, b.turn, b.section], [[], [], null, null]);
  assert.deepEqual(changesReport({}, [], {}), []);
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

test("summary.txt: counts, times and models, no text of the manuscript", () => {
  const turn = (more) => ({
    reason: null,
    model: "m1",
    model_line: "Claude · Opus 9",
    other_model: false,
    ...more,
  });
  const report = changesReport(REVISION, HEADINGS, {
    "R1.1": turn({ outcome: "ok", seconds: 41 }),
    "R1.2": turn({ outcome: "timeout", reason: "x", seconds: 600 }),
  });
  const text = summaryText(report, {
    seconds: 700,
    saved: true,
    models: { asked: { edits: "opus", writing: null, judgments: "haiku" }, line: LINE },
    source: "opened as it is",
    responses: 1,
  });
  assert.equal(
    text,
    [
      "comments: 2",
      "turns run: 2",
      "changes: 2",
      "comments without a change: 1",
      "requirements no change names: 1",
      "turns that did not end well: 1 (R1.2 timeout)",
      "models asked for: edits opus, writing (the pane's own), judgments haiku",
      `models that ran, as the pane says: ${LINE}`,
      "models that ran, from the daemon log: m1",
      "another model than asked: no",
      "turn time: 641 s",
      "total time: 700 s",
      "copy saved with its tracked changes: yes",
      "manuscript: opened as it is",
      "responses drafted: 1",
      "",
    ].join("\n"),
  );
  assert.ok(!text.includes("value"));
  // Another model than the one asked for is said as a failure: at the end of
  // the run, or after one turn.
  const drift = summaryText(report, {
    seconds: 1,
    saved: true,
    models: { asked: { edits: "opus" }, line: DRIFT },
  });
  assert.match(
    drift,
    /another model than asked: FAILED \(edits: asked for Claude · Opus 9, ran on Claude · Sonnet 9\)\n/,
  );
  const one = changesReport(REVISION, HEADINGS, {
    "R1.1": turn({ outcome: "ok", seconds: 1, other_model: true }),
  });
  assert.match(
    summaryText(one, { seconds: 1, saved: true, models: { asked: { edits: "opus" }, line: LINE } }),
    /another model than asked: FAILED \(turn R1\.1\)\n/,
  );
  const bare = summaryText([], {
    seconds: 3,
    saved: false,
    models: { asked: { edits: "opus" }, line: "" },
  });
  assert.match(bare, /copy saved with its tracked changes: NO\n$/);
  assert.match(bare, /models that ran, as the pane says: -\n/);
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
  assert.match(save, /BindToMoniker/);
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
  assert.match(open, /QuitOwnWord \$app \$d \$p \$ours/);
  assert.ok(open.indexOf("(Under $Doc)") < open.indexOf("Documents.Open("));
  assert.ok(open.indexOf("(Under $RepairTo)") < open.indexOf("SaveAs2("));
  assert.equal((open.match(/SaveAs2\(/g) ?? []).length, 2);
  assert.doesNotMatch(open, /\$doc\.Save\(\)/);
  // The runner copies the manuscript first and gives Word only copies; a
  // file Word must repair stops the run unless --repair is given.
  assert.ok(
    runner.indexOf("fs.copyFileSync(opt.manuscript, SOURCE)") <
      runner.indexOf("openCheck()", runner.indexOf("async function setup")),
  );
  // Read in two places only: the check of its name, and that copy.
  assert.equal((runner.match(/opt\.manuscript/g) ?? []).length, 2);
  assert.match(runner, /run again with --repair/);
  assert.match(runner, /pauseResident\(\)/);
  // No input path is written into the code: they come as arguments.
  for (const src of [runner, extra, save, read("../scripts/e2e/replay-compare.py")])
    assert.doesNotMatch(src, /\/mnt\/[a-z]\/|[A-Z]:\\\\(?!Program Files)/);
});
