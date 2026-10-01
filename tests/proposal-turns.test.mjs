// A proposal turn (daemon/proposal-turns.mjs): one request to a provider
// without tools, then every proposal checked and applied through the same
// office_tracked_edit path as a Claude turn. The provider and Word are
// stand-ins; the review service and its review file are real (a temp
// folder). Invented text.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createReviewService } from "../daemon/review-service.mjs";
import { createProposalTurns, proposalTurnsFor } from "../daemon/proposal-turns.mjs";
import { PROPOSAL_JSON_SCHEMA } from "../daemon/proposal.mjs";
import { applyTrackedEdit } from "../daemon/review-tools.mjs";
import {
  loadRevision,
  revisionPathFor,
  setSnapshot,
  withRevision,
} from "../daemon/revision-store.mjs";

const ENV = { CC_OFFICE_WSL_PATHS: "0" };
const PARAGRAPHS = [
  "Methods",
  "Cells were grown at 30 degrees for two days.",
  "The cell-free extract was kept on ice.",
  "Growth was followed by optical density.",
];
const COMMENT = {
  ref: "R1.1",
  reviewer: 1,
  text: "State the unit, name the lysate, and say how growth was followed.",
  requirements: ["state the unit"], // numbered q1 by the review file
};
const P = (over = {}) => ({
  paragraph_id: "p1",
  find: "30 degrees",
  op: "replace",
  text: "30 °C",
  what: "단위를 적었다",
  why: "리뷰어 요청",
  addresses: "단위",
  requirement_ids: ["q1"],
  ...over,
});
const THREE = [
  P(),
  P({
    paragraph_id: "p2",
    find: "cell-free extract",
    text: "clarified lysate",
    requirement_ids: [],
  }),
  P({
    paragraph_id: "p3",
    find: "optical density",
    op: "insert_after",
    text: " at 600 nm",
    requirement_ids: [],
  }),
];
const reply = (proposals, note = "") => ({
  ok: true,
  text: JSON.stringify({ proposals, note }),
  json: { proposals, note },
});
const CODEX = { id: "codex", model: "gpt-6.1-sol" };
const deferred = () => {
  let resolve;
  const promise = new Promise((r) => (resolve = r));
  return { promise, resolve };
};

// keys: the panes; each has a document of its own in one temp folder.
async function setup({
  generate = async () => reply(THREE),
  word = null,
  keys = ["k"],
  ledger = [],
  over = {},
  depsOver = null,
} = {}) {
  const dir = await mkdtemp(join(tmpdir(), "proposal-turns-"));
  const docs = Object.fromEntries(keys.map((k) => [k, join(dir, `${k}.docx`)]));
  const service = []; // review_event objects of the review service
  const lines = [];
  const svc = createReviewService({
    getContext: (key) => ({ activeDoc: docs[key], caps: ["target_hash"] }),
    callTool: null,
    pushTurn: () => {},
    sendEvent: (key, obj) => service.push([key, obj]),
    log: (s) => lines.push(s),
    env: ENV,
  });
  const ask = (key, type, payload = {}) =>
    new Promise((resolve) =>
      svc.handlers[type]({ type, request_id: "q", ...payload }, resolve, key, "word"),
    );
  const PAST = new Date(Date.now() - 60_000);
  for (const key of keys) {
    await ask(key, "review_save_comments", { comments: [COMMENT] });
    for (const item of ledger) await ask(key, "ledger_add", { item });
    await withRevision(revisionPathFor(docs[key]), docs[key], (rev) =>
      setSnapshot(rev, PARAGRAPHS, PAST),
    );
  }
  const out = []; // what the task pane is sent, in order
  const ended = [];
  const waiters = [];
  const asked = []; // requests to the provider
  const wordCalls = [];
  const applied = []; // arguments applyTrackedEdit was called with
  let n = 0;
  const turns = createProposalTurns({
    handleFor: (choice) => ({
      provider: choice.id,
      model: `${choice.id}:${choice.model}`,
      run: (req) => (asked.push(req), generate(req, asked.length)),
    }),
    trackedEditDeps: depsOver
      ? (key, d) => depsOver(svc.trackedEditDeps(key, d))
      : svc.trackedEditDeps,
    call: async (name, args, key) => {
      wordCalls.push([key, name, args]);
      n += 1;
      if (word) return word(args, n, key);
      return {
        ok: true,
        word_comment_id: 100 + n,
        paragraph_index: args.expected_paragraph?.index ?? null,
        before: "b",
        after: "a",
        target_checked: true,
      };
    },
    applyEdit: (...a) => (applied.push(a), applyTrackedEdit(...a)),
    sendText: (text, key) => out.push([key, "text", text]),
    sendEvent: (e, key) => out.push([key, e.event, e]),
    endTurn: (key, outcome) => {
      ended.push([key, outcome]);
      waiters.splice(0).forEach((w) => w());
    },
    log: (s) => lines.push(s),
    ...over,
  });
  const at = (key) => svc.trackedEditDeps(key).begin();
  const start = (key = "k", source = {}, choice = CODEX) =>
    turns.start(key, "word", { ref: "R1.1" }, { at: at(key), selected: [], choice, ...source });
  const turnEnded = (count = 1) =>
    new Promise((resolve) => {
      const check = () => (ended.length >= count ? resolve() : waiters.push(check));
      check();
    });
  const texts = (key = "k") => out.filter((o) => o[0] === key && o[1] === "text").map((o) => o[2]);
  const events = (key = "k") => out.filter((o) => o[0] === key && o[1] !== "text").map((o) => o[1]);
  const revision = (key = "k") => loadRevision(revisionPathFor(docs[key]), docs[key]);
  const decisions = async (key = "k") =>
    (await readFile(join(dir, `${key}.decisions.jsonl`), "utf8"))
      .trim()
      .split("\n")
      .map((l) => JSON.parse(l));
  const batch = async (key = "k") =>
    (await decisions(key)).filter((d) => d.kind === "proposal_batch");
  return {
    dir,
    docs,
    svc,
    turns,
    start,
    turnEnded,
    out,
    ended,
    asked,
    wordCalls,
    applied,
    service,
    lines,
    texts,
    events,
    revision,
    decisions,
    batch,
    cleanup: () => rm(dir, { recursive: true, force: true }),
  };
}

test("three proposals are applied in order; the pane sees the same events as in a Claude turn", async () => {
  const t = await setup();
  try {
    t.start();
    await t.turnEnded();
    assert.deepEqual(t.ended, [["k", { ok: true }]]);
    assert.deepEqual(t.events(), [
      "tool_use_announce",
      "tool_use_announce",
      "tool_use_announce",
      "turn_complete",
    ]);
    assert.deepEqual(t.out.at(-1)[2], { event: "turn_complete", subtype: "success" });
    const announced = t.out.filter((o) => o[1] === "tool_use_announce").map((o) => o[2]);
    assert.deepEqual(
      announced.map((e) => [e.tool, e.input.find, e.input.comment_ref, e.input.paragraph_hint]),
      [
        ["mcp__office__office_tracked_edit", "30 degrees", "R1.1", "p1"],
        ["mcp__office__office_tracked_edit", "cell-free extract", "R1.1", "p2"],
        ["mcp__office__office_tracked_edit", "optical density", "R1.1", "p3"],
      ],
    );
    assert.deepEqual(
      t.wordCalls.map((c) => [c[1], c[2].find, c[2].expected_paragraph.index]),
      [
        ["office_tracked_edit", "30 degrees", 1],
        ["office_tracked_edit", "cell-free extract", 2],
        ["office_tracked_edit", "optical density", 3],
      ],
    );
    assert.equal(t.wordCalls[2][2].insert_after, " at 600 nm");
    const rev = await t.revision();
    assert.deepEqual(
      rev.comments[0].changes.map((c) => [c.id, c.kind]),
      [
        ["R1.1-1", "replace"],
        ["R1.1-2", "replace"],
        ["R1.1-3", "insert"],
      ],
    );
    assert.deepEqual(
      t.service.map((s) => s[1].event),
      ["change_recorded", "change_recorded", "change_recorded"],
    );
    assert.deepEqual(t.texts(), ["3 proposals received.\n", "3 proposals: 3 applied.\n"]);
    // One request, with the system text, the reply form and a signal.
    assert.equal(t.asked.length, 1);
    assert.deepEqual(t.asked[0].schema, PROPOSAL_JSON_SCHEMA);
    assert.ok(t.asked[0].system.includes("Reply with one JSON object"));
    assert.ok(t.asked[0].prompt.includes("[p1] Cells were grown at 30 degrees for two days."));
    assert.ok(t.asked[0].signal instanceof AbortSignal);
    assert.equal(t.turns.running("k"), false);
  } finally {
    await t.cleanup();
  }
});

test("no override ever reaches applyTrackedEdit, and the arguments are the daemon's own", async () => {
  const sneaky = THREE.map((p) => ({
    ...p,
    override: { blocked_by: ["L1"] },
    comment_ref: "R9.9",
    expected_before_hash: "x",
  }));
  const t = await setup({ generate: async () => reply(sneaky) });
  try {
    t.start();
    await t.turnEnded();
    assert.equal(t.applied.length, 3);
    for (const call of t.applied) {
      assert.equal(call.length, 3, "deps, args, turn: no fourth argument");
      assert.deepEqual(
        Object.keys(call[1]).sort(),
        [
          "comment_ref",
          "find",
          call[1].insert_after === undefined ? "replace" : "insert_after",
          "paragraph_hint",
          "rationale",
        ].sort(),
      );
      assert.equal(call[1].comment_ref, "R1.1");
      assert.deepEqual(call[2], { ref: "R1.1" });
    }
  } finally {
    await t.cleanup();
  }
});

test("a proposal the forbid list blocks stops the turn: the next one is not sent to Word", async () => {
  const NOVEL = {
    kind: "forbid",
    text: "novel 금지",
    check: { type: "regex", pattern: String.raw`\bnovel\b`, flags: "i" },
  };
  const list = [THREE[0], { ...THREE[1], text: "novel lysate" }, THREE[2]];
  const t = await setup({ generate: async () => reply(list), ledger: [NOVEL] });
  try {
    t.start();
    await t.turnEnded();
    assert.equal(t.wordCalls.length, 1, "P1 only");
    assert.deepEqual(
      t.service.map((s) => s[1].event),
      ["change_recorded", "edit_blocked"],
    );
    assert.deepEqual(t.events(), [
      "tool_use_announce",
      "tool_use_announce",
      "error",
      "turn_complete",
    ]);
    assert.deepEqual(t.out.at(-1)[2], { event: "turn_complete", subtype: "error" });
    const [, outcome] = t.ended[0];
    assert.equal(outcome.ok, false);
    assert.match(outcome.error, /^The turn stopped at P2 \(Blocked\)\. .*L1/);
    assert.deepEqual(t.texts().slice(1), [
      "P3 not applied: the turn stopped at P2 (Blocked).\n",
      "3 proposals: 1 applied, 1 failed, 1 not applied.\n",
    ]);
    const [b] = await t.batch();
    assert.deepEqual(
      b.output.proposals.map((p) => [p.id, p.outcome, p.code]),
      [
        ["P1", "applied", null],
        ["P2", "failed", "Blocked"],
        ["P3", "not_applied", null],
      ],
    );
    assert.equal(b.output.stopped_at, "P2");
  } finally {
    await t.cleanup();
  }
});

test("StaleTarget, an unanswered call, a change that cannot be recorded and a missing comment each stop the turn", async () => {
  const cases = [
    {
      code: "StaleTarget",
      outcome: "failed",
      word: (args, n) =>
        n === 2
          ? { ok: false, code: "StaleTarget", paragraph_index: 2 }
          : {
              ok: true,
              word_comment_id: n,
              paragraph_index: 1,
              before: "b",
              after: "a",
              target_checked: true,
            },
      hint: /Press Refresh status, then propose again/,
      events: ["change_recorded", "stale_target"],
    },
    {
      code: "OutcomeUnknown",
      outcome: "unknown",
      word: (args, n) => {
        if (n === 2) {
          throw Object.assign(new Error("tool call timed out"), {
            code: "ToolTimeout",
            toolCallId: "c1",
          });
        }
        return {
          ok: true,
          word_comment_id: n,
          paragraph_index: 1,
          before: "b",
          after: "a",
          target_checked: true,
        };
      },
      hint: /may or may not be in the document/,
      events: ["change_recorded", "edit_blocked"],
    },
    {
      code: "CommentMissing",
      outcome: "applied_no_comment",
      word: (args, n) => ({
        ok: n !== 2,
        ...(n === 2 ? { applied: true, code: "CommentInsertFailed" } : {}),
        word_comment_id: n === 2 ? null : n,
        paragraph_index: n,
        before: "b",
        after: "a",
        target_checked: true,
      }),
      hint: /Add rationale comment/,
      events: ["change_recorded", "change_recorded"],
    },
  ];
  for (const c of cases) {
    const t = await setup({ word: c.word });
    try {
      t.start();
      await t.turnEnded();
      assert.equal(t.wordCalls.length, 2, `${c.code}: P3 is not sent to Word`);
      const [, outcome] = t.ended[0];
      assert.equal(outcome.ok, false, c.code);
      assert.ok(outcome.error.startsWith(`The turn stopped at P2 (${c.code}). `), outcome.error);
      assert.match(outcome.error, c.hint);
      assert.deepEqual(
        t.service.map((s) => s[1].event),
        c.events,
        c.code,
      );
      const [b] = await t.batch();
      assert.deepEqual(
        b.output.proposals.map((p) => p.outcome),
        ["applied", c.outcome, "not_applied"],
        c.code,
      );
      assert.deepEqual([b.output.stopped_at, b.output.proposals[1].code], ["P2", c.code]);
    } finally {
      await t.cleanup();
    }
  }
  // Word applied the edit and the review file could not be written.
  let edits = 0;
  const t = await setup({
    depsOver: (deps) => ({
      ...deps,
      onEdit: async (a) => {
        edits += 1;
        if (edits === 2) throw new Error("disk full");
        return deps.onEdit(a);
      },
    }),
  });
  try {
    t.start();
    await t.turnEnded();
    assert.equal(t.wordCalls.length, 2);
    assert.match(
      t.ended[0][1].error,
      /^The turn stopped at P2 \(AppliedNotRecorded\)\. .*Refresh status/,
    );
    assert.ok(t.service.some((s) => s[1].event === "record_failed"));
  } finally {
    await t.cleanup();
  }
});

test("a reply in the wrong form is asked for once more; twice wrong fails the turn with nothing applied", async () => {
  const bad = { ok: true, text: "{}", json: { proposals: [{ find: "x" }] } };
  const once = await setup({ generate: async (req, n) => (n === 1 ? bad : reply(THREE)) });
  try {
    once.start();
    await once.turnEnded();
    assert.equal(once.asked.length, 2);
    assert.equal(once.asked[1].system, once.asked[0].system);
    assert.ok(once.asked[1].prompt.startsWith(once.asked[0].prompt));
    const added = once.asked[1].prompt.slice(once.asked[0].prompt.length);
    assert.match(
      added,
      /^\n\nYour reply was not usable: proposals\.0\.paragraph_id: .*\. Reply again with only the JSON object\.$/,
    );
    assert.deepEqual(once.ended, [["k", { ok: true }]]);
    assert.equal((await once.batch())[0].output.retried, true);
  } finally {
    await once.cleanup();
  }
  const twice = await setup({
    generate: async (req, n) =>
      n === 1 ? { ok: false, error: "reply was not JSON", text: "Sure! Here you go" } : bad,
  });
  try {
    twice.start();
    await twice.turnEnded();
    assert.equal(twice.asked.length, 2, "one retry, not more");
    assert.match(twice.asked[1].prompt, /Your reply was not usable: the reply was not JSON\./);
    assert.equal(twice.wordCalls.length, 0);
    assert.equal(twice.ended[0][1].ok, false);
    assert.match(
      twice.ended[0][1].error,
      /^Codex CLI \(ChatGPT sign-in\) answered twice in a form that cannot be used \(proposals\.0\.paragraph_id: .*\)\. No change was made\.$/,
    );
    const [b] = await twice.batch();
    assert.deepEqual([b.output.retried, b.output.proposals], [true, []]);
  } finally {
    await twice.cleanup();
  }
});

test("a generation that used a tool, or failed, applies nothing", async () => {
  const tool = await setup({
    generate: async () => ({
      ok: false,
      tool_event: "codex_shell",
      error:
        "Codex CLI used a tool (codex_shell) in a run that must not use any. Its reply was discarded.",
    }),
  });
  try {
    tool.start();
    await tool.turnEnded();
    assert.equal(tool.asked.length, 1, "not asked again");
    assert.equal(tool.wordCalls.length, 0);
    assert.deepEqual(tool.ended[0][1], {
      ok: false,
      error:
        "Codex CLI used a tool (codex_shell) in a run that must not use any. Its reply was discarded. No change was made.",
      code: "ToolEvent",
    });
    assert.deepEqual((await tool.batch())[0].output.tool_event, "codex_shell");
  } finally {
    await tool.cleanup();
  }
  const failed = await setup({
    generate: async () => ({ ok: false, error: "codex exited 1: boom" }),
  });
  try {
    failed.start();
    await failed.turnEnded();
    assert.equal(failed.asked.length, 1);
    assert.deepEqual(failed.ended[0][1], {
      ok: false,
      error:
        "Codex CLI (ChatGPT sign-in) gave no proposals (codex exited 1: boom). No change was made.",
      code: "NoReply",
    });
  } finally {
    await failed.cleanup();
  }
});

test("Stop works in every stage: before the request, while the provider answers, between proposals, while Word works", async () => {
  const STOPPED = { ok: false, error: "The turn was stopped." };
  const interrupted = (t) =>
    assert.deepEqual(t.out.at(-1)[2], { event: "turn_complete", interrupted: true });

  // 1. Right after start(), before anything was awaited.
  const early = await setup();
  try {
    early.start();
    assert.equal(early.turns.running("k"), true, "registered at once");
    assert.equal(early.turns.stop("k"), true);
    await early.turnEnded();
    assert.deepEqual(early.ended, [["k", STOPPED]]);
    assert.equal(early.asked.length, 0, "the provider was not asked");
    interrupted(early);
    assert.equal(early.turns.stop("k"), false, "nothing is running any more");
    assert.equal((await early.batch()).length, 1, "a receipt all the same");
  } finally {
    await early.cleanup();
  }

  // 2. While the provider answers: its run is aborted.
  const gate = deferred();
  const during = await setup({
    generate: (req) =>
      new Promise((resolve) => {
        gate.resolve();
        req.signal.addEventListener("abort", () =>
          resolve({ ok: false, aborted: true, error: "stopped" }),
        );
      }),
  });
  try {
    during.start();
    await gate.promise;
    assert.equal(during.turns.stop("k"), true);
    await during.turnEnded();
    assert.deepEqual(during.ended, [["k", STOPPED]]);
    assert.equal(during.wordCalls.length, 0);
    interrupted(during);
  } finally {
    await during.cleanup();
  }

  // 3. Between two proposals: the first is applied and recorded, the second not started.
  let between;
  between = await setup({
    word: (args, n) => {
      if (n === 1) between.turns.stop("k");
      return {
        ok: true,
        word_comment_id: n,
        paragraph_index: 1,
        before: "b",
        after: "a",
        target_checked: true,
      };
    },
  });
  try {
    between.start();
    await between.turnEnded();
    assert.deepEqual(between.ended, [["k", STOPPED]]);
    assert.equal(between.wordCalls.length, 1);
    assert.equal((await between.revision()).comments[0].changes.length, 1);
    interrupted(between);
    const [b] = await between.batch();
    assert.deepEqual(
      b.output.proposals.map((p) => p.outcome),
      ["applied", "not_applied", "not_applied"],
    );
    assert.equal(b.output.stopped_at, "P2");
  } finally {
    await between.cleanup();
  }

  // 4. While Word works: its answer is awaited and recorded, then the turn stops.
  const inWord = deferred();
  const answer = deferred();
  const word = await setup({
    word: async (args, n) => {
      inWord.resolve();
      await answer.promise;
      return {
        ok: true,
        word_comment_id: n,
        paragraph_index: 1,
        before: "b",
        after: "a",
        target_checked: true,
      };
    },
  });
  try {
    word.start();
    await inWord.promise;
    assert.equal(word.turns.stop("k"), true);
    assert.equal(word.ended.length, 0, "the answer of Word is waited for");
    answer.resolve();
    await word.turnEnded();
    assert.deepEqual(word.ended, [["k", STOPPED]]);
    assert.deepEqual(
      word.service.map((s) => s[1].event),
      ["change_recorded"],
    );
    assert.equal(word.wordCalls.length, 1);
    assert.equal(word.events().filter((e) => e === "tool_use_announce").length, 1);
  } finally {
    await word.cleanup();
  }
});

test("a status refresh while the turn runs stops it: the proposals were made for another snapshot", async () => {
  let t;
  t = await setup({
    word: async (args, n) => {
      if (n === 1) {
        await withRevision(revisionPathFor(t.docs.k), t.docs.k, (rev) =>
          setSnapshot(rev, PARAGRAPHS, new Date()),
        );
      }
      return {
        ok: true,
        word_comment_id: n,
        paragraph_index: 1,
        before: "b",
        after: "a",
        target_checked: true,
      };
    },
  });
  try {
    t.start();
    await t.turnEnded();
    assert.equal(t.wordCalls.length, 1);
    assert.match(
      t.ended[0][1].error,
      /^The turn stopped at P2 \(SnapshotChanged\)\. The status was refreshed while the turn ran\. Propose again\.$/,
    );
    const [b] = await t.batch();
    assert.deepEqual(
      b.output.proposals.map((p) => [p.outcome, p.code]),
      [
        ["applied", null],
        ["failed", "SnapshotChanged"],
        ["not_applied", null],
      ],
    );
  } finally {
    await t.cleanup();
  }
});

test("refused proposals are named in the chat and do not stop the others; one made useless by an earlier edit is skipped", async () => {
  const list = [
    P(),
    P({ paragraph_id: "p9" }), // not shown
    P({ find: "Cells were grown at 30 degrees", text: "Cells grew at 30 degrees" }), // gone after P1
    THREE[1],
  ];
  const t = await setup({ generate: async () => reply(list) });
  try {
    t.start();
    await t.turnEnded();
    assert.deepEqual(t.ended, [["k", { ok: true }]]);
    assert.deepEqual(
      t.wordCalls.map((c) => c[2].find),
      ["30 degrees", "cell-free extract"],
    );
    assert.deepEqual(t.texts(), [
      "4 proposals received.\n",
      "P2 rejected (UnknownParagraph): paragraph p9. It names a paragraph that was not shown, so it was not sent to Word.\n",
      "P3 skipped (SupersededInTurn): paragraph p1. A change before it in this turn already changed that text. Propose again for what is left.\n",
      "4 proposals: 2 applied, 1 rejected, 1 skipped.\n",
    ]);
    assert.ok(t.lines.includes("[proposal] rejected ref=R1.1 id=P2 code=UnknownParagraph"));
    const [b] = await t.batch();
    assert.deepEqual(
      b.output.proposals.map((p) => [p.id, p.outcome, p.code]),
      [
        ["P1", "applied", null],
        ["P2", "rejected", "UnknownParagraph"],
        ["P3", "skipped", "SupersededInTurn"],
        ["P4", "applied", null],
      ],
    );
    // A refused proposal leaves nothing in the review file.
    assert.deepEqual((await t.revision()).comments[0].blocked, []);
  } finally {
    await t.cleanup();
  }
});

test("no proposals with a note is a turn that worked; proposals that all fail the checks is one that did not", async () => {
  const none = await setup({
    generate: async () => reply([], "Section 3 is not among the paragraphs."),
  });
  try {
    none.start();
    await none.turnEnded();
    const notice =
      "0 of 1 request has a proposed change; q1 has none: Section 3 is not among the paragraphs.";
    assert.deepEqual(none.ended, [["k", { ok: true, notice }]]);
    assert.deepEqual(none.texts(), ["0 proposals received.\n", `${notice}\n`]);
    assert.deepEqual((await none.batch())[0].output.uncovered, ["q1"]);
  } finally {
    await none.cleanup();
  }
  const all = await setup({ generate: async () => reply([P({ find: "not there" })]) });
  try {
    all.start();
    await all.turnEnded();
    assert.deepEqual(all.ended, [
      ["k", { ok: false, error: "No proposal could be placed.", code: "NothingPlaced" }],
    ]);
    assert.equal(all.wordCalls.length, 0);
  } finally {
    await all.cleanup();
  }
});

test("the receipt: one proposal_batch line per turn, with ids and outcomes and no text of the manuscript or the reply", async () => {
  const t = await setup();
  try {
    t.start();
    await t.turnEnded();
    const [b, ...more] = await t.batch();
    assert.equal(more.length, 0);
    assert.equal(b.model, "codex:gpt-6.1-sol");
    assert.match(b.input_fingerprint, /^[0-9a-f]{64}$/);
    assert.deepEqual(Object.keys(b.output).sort(), [
      "proposals",
      "ref",
      "retried",
      "snapshot_taken_at",
      "stopped_at",
      "turn_id",
      "uncovered",
    ]);
    assert.deepEqual(b.output.uncovered, []);
    assert.deepEqual(b.output.proposals[0], {
      id: "P1",
      paragraph_id: "p1",
      op: "replace",
      requirement_ids: ["q1"],
      outcome: "applied",
      code: null,
      change_id: "R1.1-1",
    });
    const raw = await readFile(join(t.dir, "k.decisions.jsonl"), "utf8");
    const line = raw.split("\n").find((l) => l.includes('"proposal_batch"'));
    for (const secret of [
      "30 degrees",
      "30 °C",
      "Cells were grown",
      "clarified lysate",
      "단위를 적었다",
    ]) {
      assert.equal(line.includes(secret), false, secret);
    }
    // The log of the turn holds ids, counts and codes only.
    const log = t.lines.filter((l) => l.startsWith("[proposal]")).join("\n");
    assert.match(
      log,
      /\[proposal\] prompt ref=R1\.1 provider=codex:gpt-6\.1-sol scope=manuscript paragraphs=4 chars=\d+/,
    );
    assert.match(
      log,
      /\[proposal\] done ref=R1\.1 proposals=3 applied=3 rejected=0 skipped=0 stopped_at=-/,
    );
    for (const secret of ["30 degrees", "Cells were grown"])
      assert.equal(log.includes(secret), false);
  } finally {
    await t.cleanup();
  }
});

test("a turn that cannot be built is refused before the provider is asked", async () => {
  const t = await setup();
  try {
    // A provider with no budget for a proposal turn.
    t.start("k", {}, { id: "ollama", model: "qwen3.5:9b" });
    await t.turnEnded();
    assert.deepEqual(t.ended[0][1], {
      ok: false,
      error: "Ollama (local) does not run proposal turns.",
      code: "NotOffered",
    });
    // A selection the snapshot does not have.
    t.start("k", { selected: [{ index: 40 }] });
    await t.turnEnded(2);
    assert.match(
      t.ended[1][1].error,
      /^The selected paragraphs are not in the last status refresh/,
    );
    // Only the selected paragraph is shown, and a proposal for another one is refused.
    t.start("k", { selected: [{ index: 2, text: "ignored" }] });
    await t.turnEnded(3);
    assert.equal(t.asked.length, 1);
    assert.ok(t.asked[0].prompt.endsWith("[p2] The cell-free extract was kept on ice."));
    assert.deepEqual(
      t.wordCalls.map((c) => c[2].find),
      ["cell-free extract"],
    );
  } finally {
    await t.cleanup();
  }
  // Too long for the provider: the count is said and nothing is sent.
  const small = await setup({ over: { budgetOf: () => 500 } });
  try {
    small.start();
    await small.turnEnded();
    assert.equal(small.asked.length, 0);
    assert.match(
      small.ended[0][1].error,
      /^The manuscript is too long for Codex CLI \(ChatGPT sign-in\) in one request \(\d+ characters; limit 500\)\. Select the paragraphs/,
    );
  } finally {
    await small.cleanup();
  }
});

test("two panes run side by side; a second turn on one pane is refused; shutdown stops and waits for all", async () => {
  const t = await setup({
    keys: ["k", "k2"],
    generate: (req) =>
      new Promise((resolve) => {
        const stop = () => resolve({ ok: false, aborted: true, error: "stopped" });
        if (req.signal.aborted) stop();
        else req.signal.addEventListener("abort", stop);
      }),
  });
  try {
    t.start("k");
    t.start("k2");
    assert.deepEqual([t.turns.running("k"), t.turns.running("k2")], [true, true]);
    await assert.rejects(t.start("k"), /already running/);
    await assert.rejects(t.turns.start("k2", "word", {}, {}), /needs the comment/);
    // Stopping one pane leaves the other running.
    assert.equal(t.turns.stop("k"), true);
    await t.turnEnded(1);
    assert.deepEqual(
      t.ended.map((e) => e[0]),
      ["k"],
    );
    assert.equal(t.turns.running("k2"), true);
    await t.turns.shutdown();
    assert.deepEqual(
      t.ended.map((e) => e[0]),
      ["k", "k2"],
    );
    assert.equal(t.turns.running("k2"), false);
  } finally {
    await t.cleanup();
  }
});

// A request of the comment that no change answers is said when the turn
// ends, in the chat and with the end of the turn (the Review tab shows it):
// a turn that ends as a success with a request left open would be missed.
test("a request no change answers is named at the end of a turn that worked", async () => {
  const two = [THREE[1], THREE[2]]; // neither names q1
  const t = await setup({ generate: async () => reply(two) });
  try {
    t.start();
    await t.turnEnded();
    const notice = "0 of 1 request has a proposed change; q1 has none.";
    assert.deepEqual(t.ended, [["k", { ok: true, notice }]]);
    assert.deepEqual(t.texts(), [
      "2 proposals received.\n",
      "2 proposals: 2 applied.\n",
      `${notice}\n`,
    ]);
    assert.deepEqual((await t.batch())[0].output.uncovered, ["q1"]);
  } finally {
    await t.cleanup();
  }
  // A note with proposals and every request answered: the note is shown alone.
  const noted = await setup({ generate: async () => reply(THREE, "The unit was assumed.") });
  try {
    noted.start();
    await noted.turnEnded();
    assert.deepEqual(noted.ended, [["k", { ok: true }]]);
    assert.deepEqual(noted.texts(), [
      "3 proposals received.\n",
      "3 proposals: 3 applied.\n",
      "The unit was assumed.\n",
    ]);
  } finally {
    await noted.cleanup();
  }
  // A turn that failed or was stopped says why it ended, not this.
  const failed = await setup({ generate: async () => reply([P({ find: "not there" })]) });
  try {
    failed.start();
    await failed.turnEnded();
    assert.deepEqual(failed.ended, [
      ["k", { ok: false, error: "No proposal could be placed.", code: "NothingPlaced" }],
    ]);
  } finally {
    await failed.cleanup();
  }
});

// What a turn would send, without sending it (Preview prompt in the Review
// tab): the same system text and prompt as the turn itself.
test("preview gives the request of the turn and asks nobody", async () => {
  const t = await setup();
  try {
    const at = t.svc.trackedEditDeps("k").begin();
    const p = await t.turns.preview("k", { ref: "R1.1" }, { at, selected: [], choice: CODEX });
    assert.equal(p.ok, true);
    assert.equal(t.asked.length, 0);
    assert.equal(t.out.length, 0, "nothing goes to the pane");
    t.start();
    await t.turnEnded();
    // The same but for the fence of the comment, which every prompt takes anew.
    const fence = (s) => /\n<<<([0-9a-f]{12})\n/.exec(s)[1];
    assert.notEqual(fence(p.prompt), fence(t.asked[0].prompt));
    assert.deepEqual(
      [p.system, p.prompt.replaceAll(fence(p.prompt), "F")],
      [t.asked[0].system, t.asked[0].prompt.replaceAll(fence(t.asked[0].prompt), "F")],
    );
    const sel = await t.turns.preview(
      "k",
      { ref: "R1.1" },
      { at, selected: [{ index: 3 }], choice: CODEX },
    );
    assert.equal(sel.scope, "selection");
    // A refusal is the answer, in the words of the turn.
    const no = await t.turns.preview(
      "k",
      { ref: "R1.1" },
      { at, selected: [], choice: { id: "ollama", model: "qwen3.5:9b" } },
    );
    assert.deepEqual(no, {
      ok: false,
      code: "NotOffered",
      message: "Ollama (local) does not run proposal turns.",
    });
    const gone = await t.turns.preview("k", { ref: "R7.7" }, { at, selected: [], choice: CODEX });
    assert.deepEqual(gone, { ok: false, code: "NoComment", message: "No reviewer comment R7.7." });
  } finally {
    await t.cleanup();
  }
});

// The one place where the runner meets the bridge and the review service
// (index.mjs calls it). The bridge takes (text, key) and (event, key); the
// review service sends its own events as (key, obj). A swapped pair would
// send a turn's text to no pane.
test("proposalTurnsFor: the runner's lines, events and Word calls reach the bridge in its argument order", async () => {
  const dir = await mkdtemp(join(tmpdir(), "proposal-wiring-"));
  const doc = join(dir, "k.docx");
  const seen = { text: [], event: [], tool: [], review: [], ended: [] };
  const svc = createReviewService({
    getContext: () => ({ activeDoc: doc, caps: ["target_hash"] }),
    callTool: null,
    pushTurn: () => {},
    sendEvent: (key, obj) => seen.review.push([key, obj.event]),
    log: () => {},
    env: { ...ENV, CC_OFFICE_REVIEW_PROMPT_FORBID: "0" },
  });
  const ask = (type, payload) =>
    new Promise((resolve) =>
      svc.handlers[type]({ type, request_id: "q", ...payload }, resolve, "k", "word"),
    );
  try {
    await ask("review_save_comments", { comments: [COMMENT] });
    await ask("ledger_add", {
      item: {
        kind: "forbid",
        text: "novel 금지",
        check: { type: "regex", pattern: "novel", flags: "i" },
      },
    });
    await withRevision(revisionPathFor(doc), doc, (rev) =>
      setSnapshot(rev, PARAGRAPHS, new Date(Date.now() - 60_000)),
    );
    const asked = [];
    let done;
    const ended = new Promise((r) => (done = r));
    const turns = proposalTurnsFor({
      bridge: {
        sendAssistantText: (...a) => seen.text.push(a),
        sendAssistantEvent: (...a) => seen.event.push(a),
        callTaskpaneTool: async (...a) => {
          seen.tool.push(a);
          return {
            ok: true,
            word_comment_id: 7,
            paragraph_index: 1,
            before: "b",
            after: "a",
            target_checked: true,
          };
        },
      },
      review: svc,
      router: {
        handleFor: (choice) => ({
          model: `codex:${choice.model}`,
          run: async (req) => (asked.push(req), reply([THREE[0]])),
        }),
      },
      endTurn: (...a) => (seen.ended.push(a), done()),
      env: { CC_OFFICE_PROPOSAL_TIMEOUT_MS: "1234" },
      log: () => {},
    });
    await turns.start(
      "k",
      "word",
      { ref: "R1.1" },
      { at: svc.at("k"), selected: [], choice: CODEX },
    );
    await ended;
    assert.deepEqual(seen.text, [
      ["1 proposal received.\n", "k"],
      ["1 proposal: 1 applied.\n", "k"],
    ]);
    assert.deepEqual(
      seen.event.map(([e, key]) => [e.event, key]),
      [
        ["tool_use_announce", "k"],
        ["turn_complete", "k"],
      ],
    );
    assert.deepEqual(
      seen.tool.map(([name, args, key]) => [name, args.find, key]),
      [["office_tracked_edit", "30 degrees", "k"]],
    );
    assert.deepEqual(seen.review, [["k", "change_recorded"]]);
    assert.deepEqual(seen.ended, [["k", { ok: true }]]);
    // The time limit of the environment and the forbid switch of the review
    // service (one switch for both kinds of turn) reach the request.
    assert.equal(asked[0].timeoutMs, 1234);
    assert.equal(
      asked[0].prompt.includes("novel"),
      false,
      "forbid rules left out as for a Claude turn",
    );
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

// What the author reads for a proposal that was not sent to Word: the code,
// the paragraph, and what to do next.
test("every refusal of a proposal says what happened and what to do", async () => {
  const list = [
    P({ find: "not there" }),
    P({ find: "e" }),
    P({ requirement_ids: ["q7"] }),
    P({ text: "30 degrees" }),
    P({ op: "insert_after", text: "" }),
    P(),
    P(),
  ];
  const t = await setup({ generate: async () => reply(list) });
  try {
    t.start();
    await t.turnEnded();
    assert.deepEqual(t.texts().slice(1, -1), [
      "P1 rejected (FindNotInParagraph): paragraph p1. Its text is not in that paragraph, so it was not sent to Word. Propose again if the change is still needed.\n",
      "P2 rejected (FindAmbiguous): paragraph p1. Its text is in that paragraph more than once, so it was not sent to Word. Propose again if the change is still needed.\n",
      "P3 rejected (UnknownRequirement): paragraph p1. It names a request the comment does not list, so it was not sent to Word. Propose again if the change is still needed.\n",
      "P4 rejected (NoChange): paragraph p1. It changes nothing.\n",
      "P5 rejected (EmptyInsert): paragraph p1. It inserts nothing.\n",
      "P7 rejected (Duplicate): paragraph p1. A proposal before it changes the same text.\n",
    ]);
    assert.equal(t.wordCalls.length, 1, "P6 alone goes to Word");
  } finally {
    await t.cleanup();
  }
});

// Stop is answered when it is pressed. Between the check at the head of a
// proposal and the call to Word lie file reads and a file write (the attempt
// record): a Stop in that stretch must still keep the edit from Word.
test("a Stop after a proposal was taken up and before its Word call sends nothing to Word", async () => {
  let t = null;
  t = await setup({
    depsOver: (deps) => ({
      ...deps,
      onAttempt: async (a) => {
        assert.equal(t.turns.stop("k"), true);
        return deps.onAttempt(a);
      },
    }),
  });
  try {
    t.start();
    await t.turnEnded();
    assert.equal(t.wordCalls.length, 0, "no edit reaches Word after the Stop was answered");
    assert.deepEqual(t.ended, [["k", { ok: false, error: "The turn was stopped." }]]);
    assert.deepEqual(t.out.at(-1)[2], { event: "turn_complete", interrupted: true });
    const rev = await t.revision();
    assert.deepEqual([rev.comments[0].changes, rev.comments[0].blocked], [[], []]);
    const [b] = await t.batch();
    assert.deepEqual(
      b.output.proposals.map((p) => p.outcome),
      ["not_applied", "not_applied", "not_applied"],
    );
    assert.equal(b.output.stopped_at, "P1");
  } finally {
    await t.cleanup();
  }
});

// The daemon log holds ids, counts and codes of a proposal turn, no text of
// the manuscript. A failure is said twice: to the author in words (which may
// quote the manuscript: Word's own message, a CLI's stderr), and to the log
// as a code. index.mjs hands the outcome of a turn to turnEnded, which logs it.
test("the failure of a turn goes to the log as a code, not as the sentence the author reads", async () => {
  const cases = [
    {
      over: {
        word: () => ({ ok: false, code: "NotFound", message: '"30 degrees" not found' }),
      },
      error: /^The turn stopped at P1 \(NotFound\)\. "30 degrees" not found$/,
      code: "P1:NotFound",
    },
    {
      over: {
        generate: async () => ({
          ok: false,
          error: "codex exited 1: Cells were grown at 30 degrees for two days.",
        }),
      },
      error: /gave no proposals \(codex exited 1: Cells were grown at 30 degrees/,
      code: "NoReply",
    },
    {
      over: {
        word: () => {
          throw new Error('Paragraph "Cells were grown at 30 degrees" is locked');
        },
      },
      error: /^The turn stopped at P1 \(ToolError\)\. Paragraph "Cells were grown/,
      code: "P1:ToolError",
    },
  ];
  for (const c of cases) {
    const t = await setup(c.over);
    try {
      t.start();
      await t.turnEnded();
      const [, outcome] = t.ended[0];
      assert.match(outcome.error, c.error);
      assert.equal(outcome.code, c.code);
      t.svc.turnEnded("k", { ref: "R1.1" }, outcome);
      assert.ok(t.lines.includes(`[review] turn done R1.1 failed: ${c.code}`), c.code);
      const log = t.lines.join("\n");
      for (const text of ["30 degrees", "Cells were grown", "two days"]) {
        assert.equal(log.includes(text), false, `${c.code}: ${text}`);
      }
      // The author still reads the sentence: turn_done carries it.
      assert.match(t.service.at(-1)[1].error, c.error);
    } finally {
      await t.cleanup();
    }
  }
});

// Word searches the text of a paragraph as it stands with its tracked
// changes: text a tracked change deleted is still found (assumed from
// tools-word.js, which reads the reviewed text apart from `.text`; to be
// measured in Word). The daemon counts in the reviewed text. A proposal
// whose find is once in the reviewed paragraph and also in the text an
// earlier change deleted there would come back from Word as Ambiguous and
// end the turn: it is passed over instead, and the proposals after it run.
test("a proposal whose text is also in text an earlier change deleted is skipped, not sent to Word", async () => {
  const list = [
    P({
      paragraph_id: "p2",
      find: "cell-free extract",
      text: "clarified cell-free lysate",
      requirement_ids: [],
    }),
    // Once in the paragraph before and after the first change, and once more
    // in what that change deleted.
    P({ paragraph_id: "p2", find: "cell-free", text: "cell free", requirement_ids: [] }),
    THREE[0],
  ];
  // A stand-in Word that searches as Word is taken to: deleted text counts.
  const deleted = [];
  const word = (args, n) => {
    const hits = deleted.filter(
      (d) => d.index === args.expected_paragraph.index && d.text.includes(args.find),
    );
    if (hits.length) {
      return {
        ok: false,
        code: "Ambiguous",
        hit_count: 2,
        message: `"${args.find}" occurs 2 times`,
      };
    }
    if (typeof args.replace === "string") {
      deleted.push({ index: args.expected_paragraph.index, text: args.narrow?.find ?? args.find });
    }
    return {
      ok: true,
      word_comment_id: n,
      paragraph_index: args.expected_paragraph.index,
      before: "b",
      after: "a",
      target_checked: true,
    };
  };
  const t = await setup({ generate: async () => reply(list), word });
  try {
    t.start();
    await t.turnEnded();
    assert.deepEqual(
      t.wordCalls.map((c) => c[2].find),
      ["cell-free extract", "30 degrees"],
      "the second proposal is not sent",
    );
    assert.deepEqual(t.ended, [["k", { ok: true }]]);
    assert.ok(
      t
        .texts()
        .includes(
          "P2 skipped (FoundInDeletedText): paragraph p2. Its text is also in text that a change before it deleted, where Word still finds it. Accept or reject that change in Word, press Refresh status, then propose again.\n",
        ),
    );
    const [b] = await t.batch();
    assert.deepEqual(
      b.output.proposals.map((p) => [p.outcome, p.code]),
      [
        ["applied", null],
        ["skipped", "FoundInDeletedText"],
        ["applied", null],
      ],
    );
  } finally {
    await t.cleanup();
  }
});
