import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdir, mkdtemp, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createReviewService } from "../daemon/review-service.mjs";
import { applyTrackedEdit } from "../daemon/review-tools.mjs";
import { loadRevision, saveRevision } from "../daemon/revision-store.mjs";
import { REQUIREMENT_SYSTEM } from "../daemon/review-split.mjs";
import { ledgerReply } from "./fixtures/ledger-reply.mjs";

const ENV = { CC_OFFICE_WSL_PATHS: "0" };

async function setup({ env = ENV, oneshot = null, callResult, callImpl = null, ...more } = {}) {
  const dir = await mkdtemp(join(tmpdir(), "review-svc-"));
  const doc = join(dir, "paper.docx");
  const events = [];
  const jobWaiters = new Map(); // job_id -> resolve, for ask()
  const lines = [];
  const pushed = [];
  const calls = [];
  const svc = createReviewService({
    getContext: () => ({ activeDoc: doc, caps: ["target_hash"] }),
    callTool: async (name, args) => {
      calls.push([name, args]);
      if (callImpl) return callImpl(name, args, calls.length);
      return (
        callResult ?? { ok: true, word_comment_id: 5, paragraph_index: 2, before: "b", after: "a" }
      );
    },
    pushTurn: (key, host, text, turn) => pushed.push({ key, host, text, turn }),
    oneshot,
    sendEvent: (key, obj) => {
      events.push(obj);
      if (obj.event === "job_done") jobWaiters.get(obj.job_id)?.(obj);
    },
    log: (s) => lines.push(s),
    env,
    ...more,
  });
  // A long job (ledger_check, review_split) answers { accepted, job_id } and
  // sends its result in a job_done event; ask() resolves with that result.
  const ask = (type, payload = {}) =>
    new Promise((resolve) => {
      const reply = (r) => {
        if (!r.accepted) return resolve(r);
        jobWaiters.set(r.job_id, (e) => resolve({ request_id: r.request_id, ...e.result }));
      };
      svc.handlers[type]({ type, request_id: "q1", ...payload }, reply, "k", "word");
    });
  return {
    dir,
    doc,
    svc,
    ask,
    events,
    lines,
    pushed,
    calls,
    cleanup: () => rm(dir, { recursive: true, force: true }),
  };
}

const readDecisions = async (dir) =>
  (await readFile(join(dir, "paper.decisions.jsonl"), "utf8"))
    .trim()
    .split("\n")
    .map((l) => JSON.parse(l));

const COMMENTS = [
  { ref: "R1.1", reviewer: 1, text: "Use Celsius." },
  { ref: "R1.3", reviewer: 1, text: "Say how growth was followed; call the substrate novel." },
];
const NOVEL = {
  kind: "forbid",
  text: "novel 금지",
  check: { type: "regex", pattern: String.raw`\bnovel\b`, flags: "i" },
};
const IN_VITRO = {
  kind: "instruction",
  text: "in vitro 하이픈 금지",
  check: { type: "regex", pattern: String.raw`\bin-vitro\b`, flags: "i" },
};

test("every reply echoes request_id and <type>_result; errors come back as ok:false", async () => {
  const t = await setup();
  try {
    const r = await t.ask("review_load");
    assert.deepEqual([r.type, r.request_id, r.ok], ["review_load_result", "q1", true]);
    assert.equal(r.path, join(t.dir, "paper.revision.json"));
    const bad = await t.ask("review_set_status", { ref: "R9.9", status: "done" });
    assert.deepEqual([bad.type, bad.ok], ["review_set_status_result", false]);
    assert.match(bad.error, /No reviewer comment R9\.9/);
  } finally {
    await t.cleanup();
  }
});

test("unsaved document: a clear error, nothing written", async () => {
  const svc = createReviewService({
    getContext: () => ({ activeDoc: null }),
    callTool: null,
    pushTurn: null,
    env: ENV,
  });
  const r = await new Promise((resolve) =>
    svc.handlers.review_load({ request_id: "x" }, resolve, "k", "word"),
  );
  assert.equal(r.ok, false);
  assert.match(r.error, /Save the document/);
});

test("split (regex) then save; the file is written next to the manuscript", async () => {
  const t = await setup();
  try {
    const s = await t.ask("review_split", {
      text: "Reviewer #1\nComment 1: Use Celsius.\nComment 2: Rename it.",
    });
    assert.equal(s.method, "regex");
    assert.deepEqual(
      s.comments.map((c) => c.ref),
      ["R1.1", "R1.2"],
    );
    const saved = await t.ask("review_save_comments", { comments: COMMENTS });
    assert.equal(saved.ok, true);
    assert.deepEqual(await readdir(t.dir), ["paper.revision.json"]);
    const onDisk = JSON.parse(await readFile(join(t.dir, "paper.revision.json"), "utf8"));
    assert.equal(onDisk.comments.length, 2);
  } finally {
    await t.cleanup();
  }
});

test("a regex split extracts requirements once; saving edited ones asks no model", async () => {
  const calls = [];
  const t = await setup({
    oneshotModel: "haiku",
    oneshot: async (req) => {
      calls.push(req);
      return {
        ok: true,
        json: {
          sentences: [
            { ref: "R1.1", s: 1, type: "other" },
            { ref: "R1.1", s: 2, type: "request" },
            { ref: "R1.2", s: 1, type: "request" },
          ],
        },
      };
    },
  });
  try {
    const s = await t.ask("review_split", {
      text: "Reviewer #1\nComment 1: Nice study. Please use Celsius.\nComment 2: Rename it.",
    });
    assert.equal(s.method, "regex");
    assert.equal(calls.length, 1);
    assert.equal(calls[0].system, REQUIREMENT_SYSTEM);
    assert.deepEqual(
      s.comments.map((c) => [c.ref, c.text, c.requirements]),
      [
        ["R1.1", "Nice study. Please use Celsius.", [{ id: "q1", text: "Please use Celsius." }]],
        ["R1.2", "Rename it.", [{ id: "q1", text: "Rename it." }]],
      ],
    );
    const [row] = await readDecisions(t.dir);
    assert.deepEqual(
      [row.kind, row.model, row.output],
      ["requirement_extract", "haiku", { requests: { "R1.1": 1, "R1.2": 1 } }],
    );
    const raw = await readFile(join(t.dir, "paper.decisions.jsonl"), "utf8");
    assert.ok(!raw.includes("Celsius"));
    const saved = await t.ask("review_save_comments", {
      comments: [{ ref: "R1.1", reviewer: 1, text: "Use Celsius.", requirements: ["use °C"] }],
    });
    assert.equal(calls.length, 1);
    assert.deepEqual(saved.revision.comments[0].requirements, [{ id: "q1", text: "use °C" }]);
  } finally {
    await t.cleanup();
  }
});

test("a bad requirement reply warns and keeps the split", async () => {
  const t = await setup({ oneshot: async () => ({ ok: true, json: { sentences: [] } }) });
  try {
    const s = await t.ask("review_split", {
      text: "Reviewer #1\nComment 1: Use Celsius.\nComment 2: Rename it.",
    });
    assert.equal(s.method, "regex");
    assert.deepEqual(
      s.comments.map((c) => [c.ref, c.requirements]),
      [
        ["R1.1", []],
        ["R1.2", []],
      ],
    );
    assert.ok(
      s.warnings.includes("model could not label requirement sentences; add them by hand"),
      s.warnings.join(" | "),
    );
  } finally {
    await t.cleanup();
  }
});

test("propose: ledger goes into the prompt, C3 log line, turn pushed with the ref", async () => {
  const t = await setup();
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    await t.ask("ledger_add", { item: IN_VITRO });
    await t.ask("ledger_add", { item: NOVEL });
    const r = await t.ask("review_propose", {
      ref: "R1.3",
      paragraphs: [{ id: "P1", index: 18, text: "Growth reached a plateau." }],
    });
    assert.deepEqual([r.ok, r.started], [true, true]);
    assert.ok(!Number.isNaN(Date.parse(r.started_at)), "L9: the turn start time comes back");
    assert.equal(t.pushed.length, 1);
    assert.deepEqual(t.pushed[0].turn, { ref: "R1.3" });
    assert.match(t.pushed[0].text, /- \[L1\] in vitro 하이픈 금지/);
    assert.match(t.pushed[0].text, /- \[L2\] novel 금지/);
    assert.ok(
      t.lines.includes(
        "[review] prompt R1.3 instructions=L1 forbid=L2 omitted_forbid=- paragraphs=1",
      ),
    );
    assert.equal(r.revision.comments[1].status, "in_progress");
  } finally {
    await t.cleanup();
  }
});

test("CC_OFFICE_REVIEW_PROMPT_FORBID=0 omits the forbid list (acceptance-test knob)", async () => {
  const t = await setup({ env: { ...ENV, CC_OFFICE_REVIEW_PROMPT_FORBID: "0" } });
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    await t.ask("ledger_add", { item: NOVEL });
    await t.ask("review_propose", { ref: "R1.3" });
    assert.ok(!t.pushed[0].text.includes("[L1]"));
    assert.ok(
      t.lines.includes(
        "[review] prompt R1.3 instructions=- forbid=- omitted_forbid=L1 paragraphs=0",
      ),
    );
  } finally {
    await t.cleanup();
  }
});

test("preview logs the same line with (preview) and pushes nothing", async () => {
  const t = await setup();
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    const r = await t.ask("review_preview", { ref: "R1.1" });
    assert.match(r.prompt, /^\[Review turn R1\.1\]/);
    assert.equal(t.pushed.length, 0);
    assert.ok(t.lines.some((l) => l.endsWith("(preview)")));
  } finally {
    await t.cleanup();
  }
});

test("tracked-edit deps record the change, push an event, and block forbidden text", async () => {
  const t = await setup();
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    await t.ask("ledger_add", { item: NOVEL });
    const deps = {
      ...t.svc.trackedEditDeps("k"),
      call: async () => ({
        ok: true,
        word_comment_id: 44,
        paragraph_index: 9,
        before: "b",
        after: "a",
      }),
    };
    const R = { what: "w", why: "y", addresses: "a" };
    const ok = await applyTrackedEdit(
      deps,
      { comment_ref: "R1.1", find: "30 degrees", replace: "30 °C", rationale: R },
      { ref: "R1.1" },
    );
    assert.equal(ok.change_id, "R1.1-1");
    const blocked = await applyTrackedEdit(
      deps,
      { comment_ref: "R1.3", find: "substrate", replace: "novel substrate", rationale: R },
      { ref: "R1.3" },
    );
    assert.deepEqual(blocked.blocked_by, ["L1"]);
    assert.deepEqual(
      t.events.map((e) => e.event),
      ["change_recorded", "edit_blocked"],
    );
    const rev = (await t.ask("review_load")).revision;
    assert.equal(rev.comments[0].changes[0].word_comment_id, "44");
    assert.equal(rev.comments[1].blocked[0].id, "R1.3-b1");
  } finally {
    await t.cleanup();
  }
});

test("override applies the stored edit through the task pane and removes it from blocked", async () => {
  const t = await setup();
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    await t.ask("ledger_add", { item: NOVEL });
    const deps = { ...t.svc.trackedEditDeps("k"), call: async () => ({ ok: true }) };
    const R = { what: "w", why: "y", addresses: "a" };
    await applyTrackedEdit(
      deps,
      { comment_ref: "R1.3", find: "substrate", replace: "novel substrate", rationale: R },
      { ref: "R1.3" },
    );
    const r = await t.ask("review_override", { ref: "R1.3", blocked_id: "R1.3-b1" });
    assert.deepEqual([r.ok, r.applied], [true, true]);
    assert.equal(t.calls.length, 1);
    assert.equal(t.calls[0][1].comment_text, "[R1.3] w / y (금지 규칙 L1 예외 승인) / a");
    assert.equal(r.revision.comments[1].blocked.length, 0);
    assert.deepEqual(r.revision.comments[1].changes[0].override_of, ["L1"]);
  } finally {
    await t.cleanup();
  }
});

test("refresh classifies, stores states, and logs counts", async () => {
  const t = await setup();
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    const deps = {
      ...t.svc.trackedEditDeps("k"),
      call: async () => ({
        ok: true,
        word_comment_id: 44,
        paragraph_index: 9,
        before: "held at 30 degrees",
        after: "held at 30 °C",
      }),
    };
    await applyTrackedEdit(
      deps,
      {
        comment_ref: "R1.1",
        find: "30 degrees",
        replace: "30 °C",
        rationale: { what: "w", why: "y", addresses: "a" },
      },
      { ref: "R1.1" },
    );
    const snapshot = {
      paragraphs: [{ index: 9, id: null, text: "held at 30 °C\u0005\u0005", w_ins: 0, w_del: 0 }],
      comments: [
        { id: "44", paragraph_index: 9, content: "[R1.1] w / y / a" },
        { id: "45", paragraph_index: 1, content: "Check units later" },
      ],
    };
    const r = await t.ask("review_refresh", { snapshot });
    assert.deepEqual(
      r.rows.map((x) => [x.change_id, x.state]),
      [["R1.1-1", "accepted"]],
    );
    assert.equal(r.revision.comments[0].changes[0].state, "accepted");
    assert.ok(
      t.lines.includes(
        "[review] refresh changes=1 tool_comments=1 user_comments=1 states=pending:0,accepted:1,rejected:0,user_edited:0,unlocated:0 external=0",
      ),
    );
  } finally {
    await t.cleanup();
  }
});

test("ledger_check updates status and violation counts", async () => {
  const t = await setup();
  try {
    await t.ask("ledger_add", { item: IN_VITRO });
    const r = await t.ask("ledger_check", {
      snapshot: {
        paragraphs: [{ index: 3, text: "measured in-vitro", reviewed_text: "measured in-vitro" }],
      },
    });
    assert.deepEqual(
      r.outcomes.map((x) => [x.id, x.status, x.violations.length]),
      [["L1", "violations", 1]],
    );
    assert.equal(r.revision.ledger[0].violation_count, 1);
    assert.ok(t.lines.includes("[review] ledger check L1 status=violations violations=1"));
  } finally {
    await t.cleanup();
  }
});

test("ledger_check marks response-scope items skipped, still catches manuscript items", async () => {
  const t = await setup();
  try {
    await t.ask("ledger_add", { item: IN_VITRO });
    await t.ask("ledger_add", {
      item: {
        kind: "forbid",
        text: "no in-vitro in the response letter",
        scope: "response",
        check: { type: "regex", pattern: String.raw`\bin-vitro\b`, flags: "i" },
      },
    });
    const snapshot = {
      paragraphs: [{ index: 3, text: "measured in-vitro", reviewed_text: "measured in-vitro" }],
    };
    for (const full of [false, true]) {
      const r = await t.ask("ledger_check", { snapshot, full });
      assert.deepEqual(
        r.outcomes.map((x) => [x.id, x.status, x.skipped ?? null, x.violations.length]),
        [
          ["L1", "violations", null, 1],
          ["L2", "skipped", "scope", 0],
        ],
      );
      const stored = r.revision.ledger.find((i) => i.id === "L2");
      assert.equal(stored.status, "skipped");
      assert.equal(stored.violation_count, 0);
    }
    assert.ok(t.lines.includes("[review] ledger check L2 status=skipped violations=0"));
  } finally {
    await t.cleanup();
  }
});

test("comments_backup writes <stem>.comments_backup.json, never overwrites, checks the count", async () => {
  const t = await setup();
  try {
    const backup = { version: 1, count: 2, comments: [{ id: "1" }, { id: "2" }] };
    const a = await t.ask("comments_backup", { backup });
    assert.deepEqual([a.ok, a.count, a.path], [true, 2, join(t.dir, "paper.comments_backup.json")]);
    const b = await t.ask("comments_backup", { backup });
    assert.notEqual(b.path, a.path);
    assert.equal(JSON.parse(await readFile(b.path, "utf8")).comments.length, 2);
    const bad = await t.ask("comments_backup", { backup: { ...backup, count: 3 } });
    assert.equal(bad.ok, false);
    const threads = {
      version: 1,
      count: 1,
      reply_count: 2,
      comments: [{ id: "1", replies: [{ content: "a" }, { content: "b" }] }],
    };
    const c = await t.ask("comments_backup", { backup: threads });
    assert.deepEqual([c.ok, c.count, c.reply_count], [true, 1, 2]);
    const badReplies = await t.ask("comments_backup", { backup: { ...threads, reply_count: 1 } });
    assert.equal(badReplies.ok, false);
  } finally {
    await t.cleanup();
  }
});

test("turnEnded pushes turn_done", async () => {
  const t = await setup();
  try {
    t.svc.turnEnded("k", { ref: "R1.1" });
    assert.deepEqual(t.events.at(-1), {
      type: "review_event",
      event: "turn_done",
      ref: "R1.1",
      ok: true,
    });
    assert.equal(t.lines.at(-1), "[review] turn done R1.1");
    // How the turn ended goes with turn_done itself: an error event sent
    // before or after it is not what the pane's Propose all decides on.
    t.svc.turnEnded("k", { ref: "R1.2" }, { ok: false, error: "Claude usage limit reached." });
    assert.deepEqual(t.events.at(-1), {
      type: "review_event",
      event: "turn_done",
      ref: "R1.2",
      ok: false,
      error: "Claude usage limit reached.",
    });
    assert.equal(t.lines.at(-1), "[review] turn done R1.2 failed: Claude usage limit reached.");
  } finally {
    await t.cleanup();
  }
});

// ---- L4, L5, L10 -------------------------------------------------------------

const para = (index, text) => ({ index, id: null, text, reviewed_text: text, w_ins: 0, w_del: 0 });
const DOC = [
  "2. Materials and Methods",
  "Cultures were grown at 30 °C. The headspace was refreshed once per day with 20 mL of air.",
  "Enzyme activity was measured over five minutes.",
];

test("L4 refresh stores a snapshot; the next refresh records outside changes", async () => {
  const t = await setup();
  try {
    const first = await t.ask("review_refresh", {
      snapshot: { paragraphs: DOC.map((x, i) => para(i, x)), comments: [] },
    });
    assert.deepEqual(first.external, []);
    assert.deepEqual(first.revision.snapshot.paragraphs, DOC);
    const edited = [DOC[0], "Cultures were grown at 30 °C.", DOC[2]];
    const r = await t.ask("review_refresh", {
      snapshot: { paragraphs: edited.map((x, i) => para(i, x)), comments: [] },
    });
    assert.deepEqual(
      r.external.map((x) => [x.id, x.kind, x.paragraph_index, x.lost_facts]),
      [["X1", "modified", 1, ["20 mL"]]],
    );
    assert.ok(
      t.lines.includes("[review] external X1 kind=modified paragraph=1 lost=1 facts=20 mL"),
    );
    assert.match(
      t.lines.find((l) => l.startsWith("[review] refresh ") && l.endsWith("external=1")),
      /external=1$/,
    );
    assert.deepEqual(r.revision.snapshot.paragraphs, edited);
    const d = await t.ask("review_external_dismiss", { id: "X1" });
    assert.deepEqual(d.revision.external, []);
  } finally {
    await t.cleanup();
  }
});

// The refresh joins each row to its recorded change, so a tool edit
// explains only its own range and a later hand edit in the same paragraph
// is still recorded.
test("L4 refresh records a hand edit next to an accepted tool change", async () => {
  const t = await setup();
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    const before = "Cells grew rapidly and were harvested after 24 h.";
    const slow = "Cells grew slowly and were harvested after 24 h.";
    const deps = {
      ...t.svc.trackedEditDeps("k"),
      call: async () => ({
        ok: true,
        word_comment_id: 44,
        paragraph_index: 1,
        before,
        after: slow,
      }),
    };
    await applyTrackedEdit(
      deps,
      {
        comment_ref: "R1.1",
        find: "rapidly",
        replace: "slowly",
        rationale: { what: "w", why: "y", addresses: "a" },
      },
      { ref: "R1.1" },
    );
    const comments = [{ id: "44", paragraph_index: 1, content: "[R1.1] w / y / a" }];
    const pending = {
      index: 1,
      id: null,
      text: "Cells grew slowlyrapidly and were harvested after 24 h.",
      reviewed_text: slow,
      w_ins: 1,
      w_del: 1,
    };
    const first = await t.ask("review_refresh", {
      snapshot: { paragraphs: [para(0, DOC[0]), pending], comments },
    });
    assert.deepEqual(first.external, []);
    const hand = "Cells grew slowly and were harvested after 48 h.";
    const r = await t.ask("review_refresh", {
      snapshot: { paragraphs: [para(0, DOC[0]), para(1, hand)], comments },
    });
    assert.deepEqual(
      r.rows.map((x) => x.state),
      ["accepted"],
    );
    assert.deepEqual(
      r.external.map((x) => [x.kind, x.paragraph_index, x.lost_facts]),
      [["modified", 1, ["24 h"]]],
    );
    const again = await t.ask("review_refresh", {
      snapshot: { paragraphs: [para(0, DOC[0]), para(1, hand)], comments },
    });
    assert.deepEqual(again.external, []);
  } finally {
    await t.cleanup();
  }
});

test("L5 requirements: saved with the comment list, edited, and given to the tracked edit", async () => {
  const t = await setup();
  try {
    await t.ask("review_save_comments", {
      comments: [{ ...COMMENTS[1], requirements: ["say how growth was followed"] }],
    });
    const r = await t.ask("review_set_requirements", {
      ref: "R1.3",
      requirements: ["say how growth was followed", "describe the substrate as novel"],
    });
    assert.deepEqual(
      r.revision.comments[0].requirements.map((x) => x.id),
      ["q1", "q2"],
    );
    await t.ask("review_refresh", {
      snapshot: { paragraphs: DOC.map((x, i) => para(i, x)), comments: [] },
    });
    const ctx = await t.svc.trackedEditDeps("k").getReviewContext("R1.3");
    assert.deepEqual(
      [ctx.requirements.length, ctx.comment_text, ctx.paragraphs],
      [2, COMMENTS[1].text, DOC],
    );
  } finally {
    await t.cleanup();
  }
});

test("L10 ledger_check reads changed paragraphs only; full:true reads all", async () => {
  const t = await setup();
  try {
    await t.ask("ledger_add", { item: IN_VITRO });
    const paras = [para(0, "measured in-vitro"), para(1, "grown in flasks")];
    await t.ask("ledger_check", { snapshot: { paragraphs: paras } });
    assert.ok(t.lines.includes("[review] ledger scope L1 mode=full paragraphs=2/2"));
    const next = [paras[0], para(1, "grown in-vitro too")];
    const r = await t.ask("ledger_check", { snapshot: { paragraphs: next } });
    assert.ok(t.lines.includes("[review] ledger scope L1 mode=incremental paragraphs=1/2"));
    assert.ok(t.lines.includes("[review] ledger check L1 status=violations violations=2"));
    assert.equal(r.revision.ledger[0].violation_count, 2);
    await t.ask("ledger_check", { snapshot: { paragraphs: next }, full: true });
    assert.ok(t.lines.includes("[review] ledger scope L1 mode=full paragraphs=2/2"));
    assert.equal(t.lines.filter((l) => l.includes("mode=full")).length, 2);
  } finally {
    await t.cleanup();
  }
});

const noAnswer = (id, code = "ToolTimeout") =>
  Object.assign(new Error(`Tool office_tracked_edit timed out`), { code, toolCallId: id });

// A tracked edit whose call timed out may still be in Word. When its result
// arrives late (bridge onLateToolResult) the change is recorded once.
test("a late applied result of a timed-out tracked edit is recorded once; not applied or unknown ids are not", async () => {
  const t = await setup();
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    const R = { what: "w", why: "y", addresses: "a" };
    let n = 0;
    const deps = {
      ...t.svc.trackedEditDeps("k"),
      call: async () => {
        throw noAnswer(`x${++n}`);
      },
    };
    const edit = { comment_ref: "R1.1", find: "30 degrees", replace: "30 °C", rationale: R };
    const r1 = await applyTrackedEdit(deps, edit, { ref: "R1.1" });
    assert.deepEqual([r1.ok, r1.code, r1.tool_call_id], [false, "OutcomeUnknown", "x1"]);
    await applyTrackedEdit(deps, edit, { ref: "R1.1" });

    const applied = { ok: true, word_comment_id: 44, paragraph_index: 9, before: "b", after: "a" };
    const late = (id, extra) =>
      t.svc.lateToolResult("k", { id, name: "office_tracked_edit", ok: true, ...extra });
    await late("x1", { result: applied });
    await late("x1", { result: applied }); // a second copy
    await late("x2", { result: { ok: false, code: "FindNotFound" } });
    await late("x9", { result: applied });

    const rev = (await t.ask("review_load")).revision;
    assert.equal(rev.comments[0].changes.length, 1);
    assert.equal(rev.comments[0].changes[0].word_comment_id, "44");
    assert.ok(t.lines.includes("[review] late result recorded x1"));
    assert.ok(t.lines.some((l) => l.startsWith("[review] late result not applied x2")));
    // Each unanswered edit is first written to the file (edit_blocked makes
    // the pane reload); both are settled by their late answers.
    assert.deepEqual(
      t.events.map((e) => e.event),
      ["edit_blocked", "edit_blocked", "change_recorded", "late_result_recorded"],
    );
    assert.equal(rev.comments[0].blocked.length, 0);
  } finally {
    await t.cleanup();
  }
});

// review_override whose call got no answer: the edit may be in Word. The
// blocked item turns outcome_unknown and a plain second override is refused.
async function unknownOverride(t) {
  await t.ask("review_save_comments", { comments: COMMENTS });
  await t.ask("ledger_add", { item: NOVEL });
  await t.ask("review_refresh", {
    snapshot: {
      paragraphs: [
        { index: 0, id: "p0", text: "Title", w_ins: 0, w_del: 0 },
        { index: 1, id: "p1", text: "Cells grew on the substrate at 30 C.", w_ins: 0, w_del: 0 },
      ],
      comments: [],
    },
  });
  const deps = { ...t.svc.trackedEditDeps("k"), call: async () => ({ ok: true }) };
  const R = { what: "w", why: "y", addresses: "a" };
  await applyTrackedEdit(
    deps,
    { comment_ref: "R1.3", find: "the substrate", replace: "the novel substrate", rationale: R },
    { ref: "R1.3" },
  );
  return t.ask("review_override", { ref: "R1.3", blocked_id: "R1.3-b1" });
}

test("override with no answer marks the item outcome_unknown and refuses a second override", async () => {
  const t = await setup({
    callImpl: (name, args, n) => {
      if (n === 1) throw noAnswer("x1");
      return { ok: true, word_comment_id: 3, paragraph_index: 1, before: "b", after: "a" };
    },
  });
  try {
    const r = await unknownOverride(t);
    assert.deepEqual([r.ok, r.applied, r.outcome_unknown], [true, false, true]);
    const b = r.revision.comments[1].blocked[0];
    assert.equal(b.state, "outcome_unknown");
    assert.equal(b.attempt.tool_call_id, "x1");
    assert.ok(t.lines.includes("[review] override outcome unknown R1.3-b1 id=x1"));

    const again = await t.ask("review_override", { ref: "R1.3", blocked_id: "R1.3-b1" });
    assert.equal(again.ok, false);
    assert.match(again.error, /Refresh status/);
    assert.equal(t.calls.length, 1);

    // The explicit re-apply (pressed twice in the task pane) goes through.
    const forced = await t.ask("review_override", {
      ref: "R1.3",
      blocked_id: "R1.3-b1",
      force: true,
    });
    assert.deepEqual([forced.ok, forced.applied], [true, true]);
    assert.equal(t.calls.length, 2);
    assert.equal(forced.revision.comments[1].blocked.length, 0);
  } finally {
    await t.cleanup();
  }
});

test("a late applied result of an unknown override records the change and releases the item", async () => {
  const t = await setup({
    callImpl: () => {
      throw noAnswer("x1");
    },
  });
  try {
    await unknownOverride(t);
    await t.svc.lateToolResult("k", {
      id: "x1",
      name: "office_tracked_edit",
      ok: true,
      result: { ok: true, word_comment_id: 8, paragraph_index: 1, before: "b", after: "a" },
    });
    const rev = (await t.ask("review_load")).revision;
    assert.equal(rev.comments[1].blocked.length, 0);
    assert.equal(rev.comments[1].changes.length, 1);
    assert.deepEqual(rev.comments[1].changes[0].override_of, ["L1"]);
    assert.ok(t.lines.includes("[review] late result recorded x1"));
  } finally {
    await t.cleanup();
  }
});

test("refresh that finds the override applied records it once and releases the item", async () => {
  const t = await setup({
    callImpl: () => {
      throw noAnswer("x1");
    },
  });
  try {
    await unknownOverride(t);
    const comment = "[R1.3] w / y (금지 규칙 L1 예외 승인) / a";
    const r = await t.ask("review_refresh", {
      snapshot: {
        paragraphs: [
          { index: 0, id: "p0", text: "Title", w_ins: 0, w_del: 0 },
          {
            index: 1,
            id: "p1",
            text: "Cells grew on the novel substrate at 30 C.\u0005\u0005",
            w_ins: 0,
            w_del: 0,
          },
        ],
        comments: [{ id: "12", paragraph_index: 1, content: comment }],
      },
    });
    const c = r.revision.comments[1];
    assert.equal(c.blocked.length, 0);
    assert.equal(c.changes.length, 1);
    assert.equal(c.changes[0].word_comment_id, "12");
    assert.deepEqual(c.changes[0].override_of, ["L1"]);
    assert.equal(c.changes[0].state, "accepted");
    assert.deepEqual(r.external, []);
    // The late answer of the same call comes afterwards: not recorded again.
    await t.svc.lateToolResult("k", {
      id: "x1",
      name: "office_tracked_edit",
      ok: true,
      result: { ok: true, word_comment_id: 12, paragraph_index: 1, before: "b", after: "a" },
    });
    const rev = (await t.ask("review_load")).revision;
    assert.equal(rev.comments[1].changes.length, 1);
  } finally {
    await t.cleanup();
  }
});

test("refresh that does not find the override applied keeps the item outcome_unknown", async () => {
  const t = await setup({
    callImpl: () => {
      throw noAnswer("x1");
    },
  });
  try {
    await unknownOverride(t);
    const r = await t.ask("review_refresh", {
      snapshot: {
        paragraphs: [
          { index: 0, id: "p0", text: "Title", w_ins: 0, w_del: 0 },
          { index: 1, id: "p1", text: "Cells grew on the substrate at 30 C.", w_ins: 0, w_del: 0 },
        ],
        comments: [],
      },
    });
    const c = r.revision.comments[1];
    assert.equal(c.blocked[0].state, "outcome_unknown");
    assert.equal(c.changes.length, 0);
  } finally {
    await t.cleanup();
  }
});

// Word shows a pending tracked replace as inserted then deleted text.
test("refresh that finds the override pending records it and releases the item", async () => {
  const t = await setup({
    callImpl: () => {
      throw noAnswer("x1");
    },
  });
  try {
    await unknownOverride(t);
    const r = await t.ask("review_refresh", {
      snapshot: {
        paragraphs: [
          { index: 0, id: "p0", text: "Title", w_ins: 0, w_del: 0 },
          {
            index: 1,
            id: "p1",
            text: "Cells grew on the novel substratethe substrate at 30 C.",
            w_ins: 1,
            w_del: 1,
          },
        ],
        comments: [],
      },
    });
    const c = r.revision.comments[1];
    assert.equal(c.blocked.length, 0);
    assert.equal(c.changes.length, 1);
    assert.equal(c.changes[0].state, "pending");
    assert.equal(c.changes[0].word_comment_id, null);
    assert.equal(c.changes[0].comment_missing, "CommentNotFound");
    assert.ok(t.lines.some((l) => /R1\.3-b1 found applied in refresh .* state=pending/.test(l)));
  } finally {
    await t.cleanup();
  }
});

// Two overrides of the same blocked edit that overlap (the pane gave up on the
// first after its request timeout and the user approved again) apply it once.
test("overlapping overrides of one blocked edit call the tool once; the second is told it is applying", async () => {
  let release;
  let started;
  const gate = new Promise((r) => (release = r));
  const calling = new Promise((r) => (started = r));
  const t = await setup({
    callImpl: async () => {
      started();
      await gate;
      return { ok: true, word_comment_id: 3, paragraph_index: 1, before: "b", after: "a" };
    },
  });
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    await t.ask("ledger_add", { item: NOVEL });
    const deps = { ...t.svc.trackedEditDeps("k"), call: async () => ({ ok: true }) };
    const R = { what: "w", why: "y", addresses: "a" };
    await applyTrackedEdit(
      deps,
      { comment_ref: "R1.3", find: "substrate", replace: "novel substrate", rationale: R },
      { ref: "R1.3" },
    );
    const first = t.ask("review_override", { ref: "R1.3", blocked_id: "R1.3-b1" });
    const second = t.ask("review_override", { ref: "R1.3", blocked_id: "R1.3-b1", force: true });
    await calling;
    const during = (await t.ask("review_load")).revision.comments[1].blocked[0];
    release();
    const [r1, r2] = await Promise.all([first, second]);
    assert.deepEqual([r1.ok, r1.applied], [true, true]);
    assert.equal(r2.ok, false);
    assert.match(r2.error, /already being applied/);
    assert.equal(during.state, "applying");
    assert.equal(t.calls.length, 1);
    const rev = (await t.ask("review_load")).revision;
    assert.equal(rev.comments[1].changes.length, 1);
    assert.equal(rev.comments[1].blocked.length, 0);
  } finally {
    await t.cleanup();
  }
});

test("an override the tool refuses or that throws goes back to blocked", async () => {
  let n = 0;
  const t = await setup({
    callImpl: async () => {
      if (++n === 1) return { ok: false, code: "FindNotFound", message: "not found" };
      throw new Error("pane crashed");
    },
  });
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    await t.ask("ledger_add", { item: NOVEL });
    const deps = { ...t.svc.trackedEditDeps("k"), call: async () => ({ ok: true }) };
    const R = { what: "w", why: "y", addresses: "a" };
    await applyTrackedEdit(
      deps,
      { comment_ref: "R1.3", find: "substrate", replace: "novel substrate", rationale: R },
      { ref: "R1.3" },
    );
    const r1 = await t.ask("review_override", { ref: "R1.3", blocked_id: "R1.3-b1" });
    assert.deepEqual([r1.ok, r1.applied], [true, false]);
    assert.equal(r1.revision.comments[1].blocked[0].state, "blocked");
    const r2 = await t.ask("review_override", { ref: "R1.3", blocked_id: "R1.3-b1" });
    assert.equal(r2.ok, false);
    assert.match(r2.error, /pane crashed/);
    const rev = (await t.ask("review_load")).revision;
    assert.equal(rev.comments[1].blocked[0].state, "blocked");
    assert.equal(t.calls.length, 2);
  } finally {
    await t.cleanup();
  }
});

// A daemon that stopped mid-override leaves "applying" in the file. The next
// daemon did not send it, so load and refresh turn it into outcome_unknown.
test("an applying item left by a stopped daemon becomes outcome_unknown on load and refresh", async () => {
  let started;
  const calling = new Promise((r) => (started = r));
  const t = await setup({
    callImpl: () => {
      started();
      return new Promise(() => {});
    },
  });
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    await t.ask("ledger_add", { item: NOVEL });
    const deps = { ...t.svc.trackedEditDeps("k"), call: async () => ({ ok: true }) };
    const R = { what: "w", why: "y", addresses: "a" };
    await applyTrackedEdit(
      deps,
      { comment_ref: "R1.3", find: "substrate", replace: "novel substrate", rationale: R },
      { ref: "R1.3" },
    );
    t.ask("review_override", { ref: "R1.3", blocked_id: "R1.3-b1" }); // never answers
    await calling;
    const stuck = (await t.ask("review_load")).revision.comments[1].blocked[0];
    assert.equal(stuck.state, "applying"); // this daemon is still applying it

    const next = createReviewService({
      getContext: () => ({ activeDoc: t.doc, caps: ["target_hash"] }),
      callTool: async () => ({ ok: true }),
      pushTurn: () => {},
      log: () => {},
      env: ENV,
    });
    const ask2 = (type, payload = {}) =>
      new Promise((resolve) =>
        next.handlers[type]({ type, request_id: "q2", ...payload }, resolve, "k", "word"),
      );
    // An override sent before any load gets the Refresh hint, not "applying".
    const early = await ask2("review_override", { ref: "R1.3", blocked_id: "R1.3-b1" });
    assert.equal(early.ok, false);
    assert.match(early.error, /Refresh status/);
    const loaded = await ask2("review_load");
    assert.equal(loaded.revision.comments[1].blocked[0].state, "outcome_unknown");
    // The same through refresh: put applying back in the file, then refresh.
    const saved = JSON.parse(await readFile(loaded.path, "utf8"));
    saved.comments[1].blocked[0].state = "applying";
    await writeFile(loaded.path, JSON.stringify(saved));
    const refreshed = await ask2("review_refresh", {
      snapshot: { paragraphs: [{ index: 0, id: "p0", text: "Title" }], comments: [] },
    });
    assert.equal(refreshed.ok, true);
    assert.equal(refreshed.revision.comments[1].blocked[0].state, "outcome_unknown");
    const plain = await ask2("review_override", { ref: "R1.3", blocked_id: "R1.3-b1" });
    assert.equal(plain.ok, false);
    assert.match(plain.error, /Refresh status/);
  } finally {
    await t.cleanup();
  }
});

// Two backups that overlap each keep their own file: the second may not
// overwrite the first between its existence check and its write.
test("overlapping comments_backup requests write two files, neither overwritten", async () => {
  const t = await setup();
  try {
    const one = { version: 1, count: 1, comments: [{ id: "1" }] };
    const two = { version: 1, count: 2, comments: [{ id: "1" }, { id: "2" }] };
    const [a, b] = await Promise.all([
      t.ask("comments_backup", { backup: one }),
      t.ask("comments_backup", { backup: two }),
    ]);
    assert.deepEqual([a.ok, b.ok], [true, true]);
    assert.notEqual(a.path, b.path);
    assert.equal(JSON.parse(await readFile(a.path, "utf8")).count, 1);
    assert.equal(JSON.parse(await readFile(b.path, "utf8")).count, 2);
  } finally {
    await t.cleanup();
  }
});

test("ledger_check stores an unsure LLM verdict on the item", async () => {
  const t = await setup({
    oneshot: async (req) =>
      ledgerReply(req, [{ index: 3, text: "a" }], {
        unsure: [{ paragraph_index: 3, reason: "cannot tell" }],
      }),
  });
  try {
    await t.ask("ledger_add", { item: { kind: "instruction", text: "t", check: { type: "llm" } } });
    const r = await t.ask("ledger_check", {
      snapshot: { paragraphs: [{ index: 3, text: "a", reviewed_text: "a" }] },
    });
    const item = r.revision.ledger[0];
    assert.deepEqual([item.status, item.unsure.length, item.checked_hashes], ["unsure", 1, []]);
    assert.ok(t.lines.includes("[review] ledger check L1 status=unsure violations=0 unsure=1"));
  } finally {
    await t.cleanup();
  }
});

test("receipts: an LLM split, LLM ledger verdicts and requirement links land in <stem>.decisions.jsonl", async () => {
  const REVIEW = 'Please say how "secret growth phrase" was followed.';
  let call = 0;
  const t = await setup({
    oneshotModel: "haiku",
    getModel: () => "sonnet",
    oneshot: async (req) => {
      call += 1;
      if (call === 1) return { ok: true, json: { lines: [{ n: 1, type: "comment_start" }] } };
      if (call === 2) {
        assert.equal(req.system, REQUIREMENT_SYSTEM);
        return { ok: true, json: { sentences: [{ ref: "R1.1", s: 1, type: "request" }] } };
      }
      return ledgerReply(
        req,
        [
          { index: 3, text: "some hidden words" },
          { index: 4, text: "maybe" },
        ],
        {
          violations: [{ paragraph_index: 3, quote: "hidden words", reason: "r" }],
          unsure: [{ paragraph_index: 4, reason: "?" }],
        },
      );
    },
  });
  try {
    await t.ask("review_split", { text: REVIEW });
    await t.ask("ledger_add", {
      item: { kind: "instruction", text: "secret rule", check: { type: "llm" } },
    });
    await t.ask("ledger_check", {
      snapshot: {
        paragraphs: [
          { index: 3, text: "some hidden words", reviewed_text: "some hidden words" },
          { index: 4, text: "maybe", reviewed_text: "maybe" },
          { index: 5, text: "fine", reviewed_text: "fine" },
        ],
      },
    });
    await t.ask("review_save_comments", { comments: COMMENTS });
    await t.ask("review_set_requirements", { ref: "R1.1", requirements: ["use Celsius"] });
    const deps = {
      ...t.svc.trackedEditDeps("k"),
      call: async () => ({
        ok: true,
        word_comment_id: 4,
        paragraph_index: 1,
        before: "b",
        after: "a",
      }),
    };
    await applyTrackedEdit(
      deps,
      {
        comment_ref: "R1.1",
        find: "30 degrees",
        replace: "30 °C",
        rationale: { what: "w", why: "y", addresses: "a", requirement_ids: ["q1"] },
      },
      { ref: "R1.1" },
    );
    const raw = await readFile(join(t.dir, "paper.decisions.jsonl"), "utf8");
    for (const secret of ["secret growth phrase", "secret rule", "hidden words", "30 degrees"])
      assert.ok(!raw.includes(secret), secret);
    const rows = await readDecisions(t.dir);
    assert.deepEqual(
      rows.map((r) => [r.kind, r.model, r.expected]),
      [
        ["comment_split", "haiku", null],
        ["requirement_extract", "haiku", null],
        ["ledger_check", "haiku", null],
        ["ledger_check", "haiku", null],
        ["ledger_check", "haiku", null],
        ["requirement_link", "sonnet", null],
      ],
    );
    for (const r of rows) {
      assert.match(r.input_fingerprint, /^[0-9a-f]{64}$/);
      assert.ok(!Number.isNaN(Date.parse(r.ts)));
    }
    assert.deepEqual(rows[0].output.refs, ["R1.1"]);
    assert.equal(rows[0].output.method, "llm");
    assert.deepEqual(rows[0].output.labels, {
      editor: 0,
      reviewer_head: 0,
      comment_start: 1,
      continue: 0,
      drop: 0,
    });
    assert.deepEqual(
      rows.slice(2, 5).map((r) => [r.output.item_id, r.output.paragraph_index, r.output.verdict]),
      [
        ["L1", 3, "violation"],
        ["L1", 4, "unsure"],
        ["L1", 5, "ok"],
      ],
    );
    assert.deepEqual(rows[1].output, { requests: { "R1.1": 1 } });
    assert.deepEqual(rows[5].output, {
      change_id: "R1.1-1",
      comment_ref: "R1.1",
      requirement_ids: ["q1"],
    });
  } finally {
    await t.cleanup();
  }
});

test("receipts: a failed LLM ledger batch is an error verdict; a failed write only logs", async () => {
  const t = await setup({ oneshot: async () => ({ ok: false, error: "limit" }) });
  try {
    await t.ask("ledger_add", { item: { kind: "instruction", text: "t", check: { type: "llm" } } });
    const r = await t.ask("ledger_check", {
      snapshot: { paragraphs: [{ index: 2, text: "a", reviewed_text: "a" }] },
    });
    assert.equal(r.revision.ledger[0].status, "error");
    const [row] = await readDecisions(t.dir);
    assert.deepEqual([row.output.verdict, row.output.error, row.model], ["error", "limit", null]);
    // A directory where the receipt file should be: the append fails.
    await rm(join(t.dir, "paper.decisions.jsonl"));
    await mkdir(join(t.dir, "paper.decisions.jsonl"));
    const again = await t.ask("ledger_check", {
      snapshot: { paragraphs: [{ index: 2, text: "a", reviewed_text: "a" }] },
      full: true,
    });
    assert.equal(again.ok, true);
    assert.ok(t.lines.some((l) => l.startsWith("[review] decision log failed")));
  } finally {
    await t.cleanup();
  }
});

// A rule edited while an LLM check of it waits: the old verdict and its
// checked_hashes must not land on the new rule, or the next incremental
// check skips every paragraph and misses the new rule's violations.
test("ledger_check drops the result of a rule edited during the check; the next incremental check runs in full", async () => {
  let release;
  const gate = new Promise((r) => (release = r));
  const t = await setup({
    oneshot: async (req) => {
      await gate;
      return ledgerReply(req, []);
    },
  });
  try {
    await t.ask("ledger_add", { item: { kind: "forbid", text: "t", check: { type: "llm" } } });
    const paras = [para(0, "a novel substrate"), para(1, "grown in flasks")];
    const pending = t.ask("ledger_check", { snapshot: { paragraphs: paras } });
    await new Promise((r) => setImmediate(r));
    await t.ask("ledger_update", {
      id: "L1",
      patch: { text: NOVEL.text, check: NOVEL.check },
    });
    release();
    const first = await pending;
    assert.equal(first.ok, true);
    const item = first.revision.ledger[0];
    assert.deepEqual([item.status, item.checked_hashes], ["unchecked", null]);
    assert.ok(t.lines.includes("[review] ledger check L1 dropped: rule changed during the check"));
    const next = await t.ask("ledger_check", { snapshot: { paragraphs: paras } });
    assert.equal(next.revision.ledger[0].status, "violations");
    assert.ok(t.lines.includes("[review] ledger scope L1 mode=full paragraphs=2/2"));
  } finally {
    await t.cleanup();
  }
});

test("ledger_check skips an item removed during the check and still saves the others", async () => {
  let release;
  const gate = new Promise((r) => (release = r));
  const t = await setup({
    oneshot: async (req) => {
      await gate;
      return ledgerReply(req, []);
    },
  });
  try {
    await t.ask("ledger_add", { item: IN_VITRO });
    await t.ask("ledger_add", { item: { kind: "instruction", text: "t", check: { type: "llm" } } });
    const paras = [para(0, "measured in-vitro")];
    const pending = t.ask("ledger_check", { snapshot: { paragraphs: paras } });
    await new Promise((r) => setImmediate(r));
    await t.ask("ledger_remove", { id: "L2" });
    release();
    const r = await pending;
    assert.equal(r.ok, true);
    assert.deepEqual(
      r.revision.ledger.map((i) => [i.id, i.status]),
      [["L1", "violations"]],
    );
    assert.ok(t.lines.includes("[review] ledger check L2 dropped: rule removed during the check"));
  } finally {
    await t.cleanup();
  }
});

test("ledger_check: a stale check that ends after a newer one leaves the newer result", async () => {
  let release;
  const gate = new Promise((r) => (release = r));
  let n = 0;
  const t = await setup({
    oneshot: async (req) => {
      if (++n === 1) {
        await gate;
        return ledgerReply(req, []);
      }
      return ledgerReply(req, [para(0, "a novel substrate")], {
        violations: [{ paragraph_index: 0, quote: "novel", reason: "r" }],
      });
    },
  });
  try {
    await t.ask("ledger_add", { item: { kind: "forbid", text: "t", check: { type: "llm" } } });
    const paras = [para(0, "a novel substrate")];
    const stale = t.ask("ledger_check", { snapshot: { paragraphs: paras } });
    await new Promise((r) => setImmediate(r));
    await t.ask("ledger_update", { id: "L1", patch: { text: "novel 금지" } });
    const fresh = await t.ask("ledger_check", { snapshot: { paragraphs: paras } });
    assert.equal(fresh.revision.ledger[0].status, "violations");
    release();
    const r = await stale;
    assert.deepEqual(
      [r.revision.ledger[0].status, r.revision.ledger[0].violation_count],
      ["violations", 1],
    );
  } finally {
    await t.cleanup();
  }
});

// Codex review 8: two checks of the same rule on two snapshots overlap. The
// rule fingerprint is the same, so only the check generation tells which one
// is the latest: the result of the check started last is kept, in either
// order of completion.
async function overlappingChecks(order) {
  const gates = { old: null, fresh: null };
  const opened = { old: new Promise((r) => (gates.old = r)), fresh: null };
  opened.fresh = new Promise((r) => (gates.fresh = r));
  const t = await setup({
    oneshot: async (req) => {
      if (req.prompt.includes("novel")) {
        await opened.old;
        return ledgerReply(req, [para(0, "a novel substrate")], {
          violations: [{ paragraph_index: 0, quote: "novel", reason: "r" }],
        });
      }
      await opened.fresh;
      return ledgerReply(req, []);
    },
  });
  try {
    await t.ask("ledger_add", { item: { kind: "forbid", text: "t", check: { type: "llm" } } });
    const old = t.ask("ledger_check", { snapshot: { paragraphs: [para(0, "a novel substrate")] } });
    await new Promise((r) => setImmediate(r));
    const fresh = t.ask("ledger_check", {
      snapshot: { paragraphs: [para(0, "a known substrate")] },
    });
    await new Promise((r) => setImmediate(r));
    const [first, second] = order === "old-last" ? ["fresh", "old"] : ["old", "fresh"];
    gates[first]();
    await (first === "old" ? old : fresh);
    gates[second]();
    await (second === "old" ? old : fresh);
    const saved = (await t.ask("review_load")).revision.ledger[0];
    assert.deepEqual([saved.status, saved.violation_count], ["ok", 0], `${order}: latest kept`);
    if (order === "old-last") {
      assert.ok(t.lines.includes("[review] ledger check L1 dropped: a newer check started"));
    }
  } finally {
    await t.cleanup();
  }
}
test("ledger_check: an older check that ends last does not replace the latest result", () =>
  overlappingChecks("old-last"));
test("ledger_check: an older check that ends first is replaced by the latest result", () =>
  overlappingChecks("old-first"));

// Codex review 8: a long LLM job used to answer only when done, and the task
// pane stopped waiting after 120 s. The request now answers at once with a
// job id and the result comes in a job_done event, however long it takes.
// Waits for a condition the job reaches after its file writes.
async function until(pred, ms = 5000) {
  for (const end = Date.now() + ms; !pred(); ) {
    if (Date.now() > end) throw new Error("condition not reached");
    await new Promise((r) => setTimeout(r, 5));
  }
}

test("ledger_check and review_split answer with a job id at once and send the result in job_done", async () => {
  let release;
  const gate = new Promise((r) => (release = r));
  const t = await setup({
    oneshot: async (req) => {
      await gate;
      return ledgerReply(req, []);
    },
  });
  try {
    await t.ask("ledger_add", { item: { kind: "forbid", text: "t", check: { type: "llm" } } });
    const replies = [];
    const msg = {
      type: "ledger_check",
      request_id: "q9",
      snapshot: { paragraphs: [para(0, "a")] },
    };
    await t.svc.handlers.ledger_check(msg, (r) => replies.push(r), "k", "word");
    assert.equal(replies.length, 1, "answered while the model still works");
    const { job_id, ...rest } = replies[0];
    assert.equal(typeof job_id, "string");
    assert.deepEqual(rest, {
      type: "ledger_check_result",
      request_id: "q9",
      ok: true,
      accepted: true,
    });
    assert.ok(!t.events.some((e) => e.event === "job_done"), "no result before the model answers");
    release();
    await until(() => t.events.some((e) => e.event === "job_done"));
    const done = t.events.find((e) => e.event === "job_done");
    assert.equal(done.type, "review_event");
    assert.equal(done.job_id, job_id);
    assert.equal(done.result.ok, true);
    assert.deepEqual(
      done.result.outcomes.map((o) => [o.id, o.status]),
      [["L1", "ok"]],
    );

    const split = [];
    await t.svc.handlers.review_split(
      { type: "review_split", request_id: "q10", text: "" },
      (r) => split.push(r),
      "k",
      "word",
    );
    assert.equal(split[0].accepted, true);
    await until(() => t.events.filter((e) => e.event === "job_done").length >= 2);
    const failed = t.events.filter((e) => e.event === "job_done")[1];
    assert.equal(failed.job_id, split[0].job_id);
    assert.equal(typeof failed.result.ok, "boolean");
  } finally {
    await t.cleanup();
  }
});

// A comment list saved while work on a comment is still open (a review turn,
// a tracked edit waiting for Word, an edit whose answer may come late, an
// exception applying or unanswered) keeps that comment and says so. A change
// Word applies afterwards is recorded under it.
const KEPT = (refs) => `Not removed because work on it is in progress: ${refs}`;
const APPLIED = { ok: true, word_comment_id: 44, paragraph_index: 9, before: "b", after: "a" };
const R_ = { what: "w", why: "y", addresses: "a" };
const EDIT = { comment_ref: "R1.1", find: "30 degrees", replace: "30 °C", rationale: R_ };
const refsOf = (rev) => rev.comments.map((c) => c.ref);

test("comment list saved while a tracked edit waits for Word: comment kept, edit recorded", async () => {
  const t = await setup();
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    let answer, called;
    const waiting = new Promise((r) => (answer = r));
    const inWord = new Promise((r) => (called = r));
    const call = () => (called(), waiting);
    const deps = { ...t.svc.trackedEditDeps("k"), call };
    const edit = applyTrackedEdit(deps, EDIT, { ref: "R1.1" }).catch((e) => e);
    await inWord; // Word has the edit and has not answered
    const saved = await t.ask("review_save_comments", { comments: [COMMENTS[1]] });
    assert.equal(saved.ok, true);
    assert.deepEqual(refsOf(saved.revision), ["R1.3", "R1.1"]);
    assert.deepEqual(saved.warnings, [KEPT("R1.1")]);
    answer(APPLIED);
    const r = await edit;
    assert.equal(r.change_id, "R1.1-1");
    const rev = (await t.ask("review_load")).revision;
    assert.equal(rev.comments.find((c) => c.ref === "R1.1").changes.length, 1);
  } finally {
    await t.cleanup();
  }
});

test("comment list saved during a review turn keeps its comment until the turn ends", async () => {
  const t = await setup();
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    await t.ask("review_propose", { ref: "R1.1" });
    const during = await t.ask("review_save_comments", { comments: [COMMENTS[1]] });
    assert.deepEqual(refsOf(during.revision), ["R1.3", "R1.1"]);
    assert.deepEqual(during.warnings, [KEPT("R1.1")]);
    t.svc.turnEnded("k", { ref: "R1.1" });
    const after = await t.ask("review_save_comments", { comments: [COMMENTS[1]] });
    assert.deepEqual(refsOf(after.revision), ["R1.3"]);
    assert.deepEqual(after.warnings, []);
  } finally {
    await t.cleanup();
  }
});

test("comment list saved while a timed-out edit may still answer: comment kept, late result recorded", async () => {
  const t = await setup();
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    const deps = {
      ...t.svc.trackedEditDeps("k"),
      call: async () => {
        throw noAnswer("x1");
      },
    };
    await applyTrackedEdit(deps, EDIT, { ref: "R1.1" });
    const saved = await t.ask("review_save_comments", { comments: [COMMENTS[1]] });
    assert.deepEqual(saved.warnings, [KEPT("R1.1")]);
    const out = await t.svc.lateToolResult("k", {
      id: "x1",
      name: "office_tracked_edit",
      ok: true,
      result: APPLIED,
    });
    assert.equal(out.change_id, "R1.1-1");
  } finally {
    await t.cleanup();
  }
});

test("comment list saved with an unanswered exception on file keeps that comment", async () => {
  const t = await setup();
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    const path = join(t.dir, "paper.revision.json");
    const rev = await loadRevision(path, t.doc);
    rev.comments[1].blocked.push({
      id: "R1.3-b1",
      args: { comment_ref: "R1.3", find: "x", replace: "novel x", rationale: R_ },
      blocked_by: ["L1"],
      created_at: "2026-01-01T00:00:00Z",
      state: "outcome_unknown",
    });
    await saveRevision(path, rev);
    const saved = await t.ask("review_save_comments", { comments: [COMMENTS[0]] });
    assert.deepEqual(refsOf(saved.revision), ["R1.1", "R1.3"]);
    assert.deepEqual(saved.warnings, [KEPT("R1.3")]);
  } finally {
    await t.cleanup();
  }
});

// Last line of defense: the comment is gone from the file anyway (written by
// hand, or by a path the protection does not see). The applied change and a
// blocked edit are recorded under a restored comment, never thrown away.
test("an edit whose comment is gone restores the comment: ref only, or the text the turn had", async () => {
  const t = await setup();
  try {
    const path = join(t.dir, "paper.revision.json");
    await t.ask("review_save_comments", { comments: COMMENTS });
    await t.ask("review_propose", { ref: "R1.1" });
    await saveRevision(path, { ...(await loadRevision(path, t.doc)), comments: [] });
    const deps = { ...t.svc.trackedEditDeps("k"), call: async () => APPLIED };
    const r = await applyTrackedEdit(deps, EDIT, { ref: "R1.1" });
    assert.equal(r.change_id, "R1.1-1");
    const blocked = await applyTrackedEdit(
      { ...deps, getLedger: async () => [{ ...NOVEL, id: "L1", scope: "manuscript" }] },
      { ...EDIT, comment_ref: "R2.4", replace: "novel" },
      { ref: "R2.4" },
    );
    assert.equal(blocked.blocked_id, "R2.4-b1");
    const rev = await loadRevision(path, t.doc);
    const r11 = rev.comments.find((c) => c.ref === "R1.1");
    assert.deepEqual([r11.text, r11.reviewer, r11.changes.length], ["Use Celsius.", 1, 1]);
    const r24 = rev.comments.find((c) => c.ref === "R2.4");
    assert.deepEqual([r24.text, r24.reviewer, r24.blocked.length], ["", 2, 1]);
    assert.ok(t.lines.includes("[review] comment R1.1 was missing; restored to keep the change"));
    assert.ok(t.lines.includes("[review] comment R2.4 was missing; restored to keep the change"));
  } finally {
    await t.cleanup();
  }
});

test("review_comments_cleaned marks the change; refresh then keeps it and reports no missing comment", async () => {
  const { invariantReport } = await import("../taskpane/shared/review-view.js");
  const t = await setup();
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    const deps = {
      ...t.svc.trackedEditDeps("k"),
      call: async () => ({
        ok: true,
        word_comment_id: 44,
        paragraph_index: 9,
        before: "held at 30 degrees",
        after: "held at 30 °C",
      }),
    };
    await applyTrackedEdit(
      deps,
      {
        comment_ref: "R1.1",
        find: "30 degrees",
        replace: "30 °C",
        rationale: { what: "w", why: "y", addresses: "a" },
      },
      { ref: "R1.1" },
    );
    const c = await t.ask("review_comments_cleaned", { ids: ["44", "99"] });
    assert.equal(c.ok, true);
    assert.equal(typeof c.revision.comments[0].changes[0].comment_cleaned_at, "string");
    const snapshot = {
      paragraphs: [{ index: 9, id: null, text: "held at 30 °C", w_ins: 0, w_del: 0 }],
      comments: [{ id: "45", paragraph_index: 1, content: "Check units later" }],
    };
    const r = await t.ask("review_refresh", { snapshot });
    assert.deepEqual(
      r.rows.map((x) => [x.change_id, x.state]),
      [["R1.1-1", "accepted"]],
    );
    const inv = invariantReport(r.revision, snapshot.comments);
    assert.deepEqual([inv.ok, inv.missing_comment], [true, []]);
    // Nothing recorded for these ids: no write.
    const none = await t.ask("review_comments_cleaned", { ids: ["7"] });
    assert.equal(none.ok, true);
  } finally {
    await t.cleanup();
  }
});

// Codex review (P2): delete L1, add another forbid rule, and it was numbered
// L1 again. The edit blocked by the old L1 still said blocked_by ["L1"], so
// approving that exception applied text the NEW rule forbids.
test("an exception approved for a deleted forbid rule does not cover a rule added later", async () => {
  const t = await setup();
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    await t.ask("ledger_add", { item: NOVEL });
    const deps = { ...t.svc.trackedEditDeps("k"), call: async () => ({ ok: true }) };
    const R = { what: "w", why: "y", addresses: "a" };
    const first = await applyTrackedEdit(
      deps,
      { comment_ref: "R1.3", find: "substrate", replace: "novel substrate", rationale: R },
      { ref: "R1.3" },
    );
    assert.deepEqual(first.blocked_by, ["L1"]);
    await t.ask("ledger_remove", { id: "L1" });
    const added = await t.ask("ledger_add", {
      item: {
        kind: "forbid",
        text: "substrate 를 novel 이라 부르지 않는다",
        check: { type: "regex", pattern: "novel substrate", flags: "i" },
      },
    });
    assert.equal(added.item.id, "L2");
    const r = await t.ask("review_override", { ref: "R1.3", blocked_id: "R1.3-b1" });
    assert.equal(r.applied, false);
    assert.deepEqual(r.result.blocked_by, ["L2"]);
    assert.equal(t.calls.length, 0); // nothing was sent to Word
    const rev = await loadRevision(join(t.dir, "paper.revision.json"), t.doc);
    assert.equal(rev.comments[1].changes.length, 0);
    assert.equal(rev.ledger_seq, 2);
  } finally {
    await t.cleanup();
  }
});

test("tracked-edit context: changes recorded after the last refresh are given with the snapshot", async () => {
  const t = await setup();
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    await t.ask("review_refresh", {
      snapshot: {
        paragraphs: [{ index: 0, id: "U0", text: "This is a new material.", reviewed_text: null }],
        comments: [],
      },
    });
    const deps = {
      ...t.svc.trackedEditDeps("k"),
      call: async () => ({ ok: true, word_comment_id: 7, paragraph_index: 0 }),
    };
    const R = { what: "w", why: "y", addresses: "a" };
    assert.deepEqual((await deps.getReviewContext("R1.1")).recent_changes, []);
    await new Promise((r) => setTimeout(r, 5));
    await applyTrackedEdit(
      deps,
      { comment_ref: "R1.1", find: "new", replace: "novel", rationale: R },
      { ref: "R1.1" },
    );
    const ctx = await deps.getReviewContext("R1.3");
    assert.deepEqual(ctx.paragraphs, ["This is a new material."]);
    assert.deepEqual(ctx.paragraph_ids, ["U0"]); // for paragraph_hint
    assert.deepEqual(
      ctx.recent_changes.map((c) => [c.id, c.find, c.text]),
      [["R1.1-1", "new", "novel"]],
    );
    // End to end: the second edit forms the forbidden phrase with the first.
    await t.ask("ledger_add", {
      item: {
        kind: "forbid",
        text: "novel substrate 금지",
        check: { type: "regex", pattern: "novel substrate", flags: "" },
      },
    });
    const r = await applyTrackedEdit(
      deps,
      { comment_ref: "R1.3", find: "material", replace: "substrate", rationale: R },
      { ref: "R1.3" },
    );
    assert.deepEqual([r.ok, r.blocked_by, r.blocked_id], [false, ["L1"], "R1.3-b1"]);
  } finally {
    await t.cleanup();
  }
});

// P15: the ledger as a JSON file next to the manuscript.
test("ledger_export writes <stem>.ledger.json; ledger_import adds its rules and skips known ones", async () => {
  const t = await setup();
  try {
    await t.ask("ledger_add", { item: NOVEL });
    await t.ask("ledger_add", { item: IN_VITRO });
    const x = await t.ask("ledger_export");
    assert.deepEqual([x.ok, x.count, x.path], [true, 2, join(t.dir, "paper.ledger.json")]);
    const file = JSON.parse(await readFile(x.path, "utf8"));
    assert.equal(file.format, "manuscript-pane-ledger");
    assert.deepEqual(
      file.items.map((i) => Object.keys(i).sort()),
      [
        ["check", "kind", "scope", "text"],
        ["check", "kind", "scope", "text"],
      ],
    );
    assert.ok(t.lines.includes(`[review] ledger export path=${x.path} count=2`));

    // Same rules again: nothing added.
    const same = await t.ask("ledger_import", { doc: file });
    assert.deepEqual([same.ok, same.added, same.skipped], [true, 0, 2]);
    assert.ok(t.lines.includes("[review] ledger import added=0 skipped=2 mode=append"));

    await t.ask("ledger_remove", { id: "L1" });
    const back = await t.ask("ledger_import", { doc: file, mode: "append" });
    assert.deepEqual([back.added, back.skipped], [1, 1]);
    assert.deepEqual(
      back.revision.ledger.map((i) => i.text),
      [IN_VITRO.text, NOVEL.text],
    );

    const bad = await t.ask("ledger_import", { doc: { items: [{ kind: "forbid" }] } });
    assert.equal(bad.ok, false);
    assert.match(bad.error, /item 1/);
    const none = await t.ask("ledger_import", { doc: null });
    assert.match(none.error, /items array/);
    assert.equal((await loadRevision(join(t.dir, "paper.revision.json"), t.doc)).ledger.length, 2);
  } finally {
    await t.cleanup();
  }
});

test("ledger_suggest is a job: candidates from the review text, one receipt, nothing stored", async () => {
  const { SUGGEST_SYSTEM } = await import("../daemon/ledger-suggest.mjs");
  const systems = [];
  const t = await setup({
    oneshot: async (req) => {
      systems.push(req.system);
      return {
        ok: true,
        json: {
          labels: [
            { i: 0, type: "none" },
            { i: 2, type: "instruction" },
          ],
        },
      };
    },
    oneshotModel: "haiku",
  });
  try {
    const text = 'Reviewer 1\nDo not use "novel".\nUse SI units throughout.';
    let first = null;
    await new Promise((resolve) =>
      t.svc.handlers.ledger_suggest(
        { type: "ledger_suggest", request_id: "q7", text },
        (r) => ((first = r), resolve()),
        "k",
        "word",
      ),
    );
    assert.deepEqual([first.accepted, typeof first.job_id], [true, "string"]);

    const r = await t.ask("ledger_suggest", { text });
    assert.deepEqual([r.ok, r.method], [true, "regex+llm"]);
    assert.deepEqual(
      r.candidates.map((c) => [c.kind, c.check.type, c.text]),
      [
        ["forbid", "regex", 'Do not use "novel".'],
        ["instruction", "llm", "Use SI units throughout."],
      ],
    );
    assert.deepEqual([...new Set(systems)], [SUGGEST_SYSTEM]);
    assert.ok(t.lines.includes("[review] ledger suggest method=regex+llm candidates=2 warnings=0"));

    await new Promise((res) => setTimeout(res, 20)); // the first job's receipt
    const raw = await readFile(join(t.dir, "paper.decisions.jsonl"), "utf8");
    assert.equal(raw.includes("novel"), false, "no review text in the receipt");
    const log = await readDecisions(t.dir);
    assert.deepEqual(
      log.map((x) => [x.kind, x.model, x.output.candidates]),
      [
        ["ledger_suggest", "haiku", 2],
        ["ledger_suggest", "haiku", 2],
      ],
    );
    assert.equal((await loadRevision(join(t.dir, "paper.revision.json"), t.doc)).ledger.length, 0);

    const empty = await t.ask("ledger_suggest", { text: "  " });
    assert.deepEqual([empty.ok, /review text/.test(empty.error)], [false, true]);
  } finally {
    await t.cleanup();
  }
});

// P14: one-shot judgments go to the provider chosen in the pane (key "k").
test("review_split uses the pane's one-shot and its model label", async () => {
  const used = [];
  const t = await setup({
    oneshotJob: async (key) => ({
      run: async () => (used.push(key), { ok: true, json: { lines: [] }, text: "{}" }),
      model: `ollama:qwen3.5:9b@${key}`,
    }),
  });
  try {
    await t.ask("review_split", { text: "Reviewer 1\nOne comment only." });
    assert.ok(used.includes("k"));
    // The comment_split and requirement_extract receipts name the routed model.
    const log = await readDecisions(t.dir);
    assert.ok(log.length >= 1);
    assert.ok(log.every((r) => r.model === "ollama:qwen3.5:9b@k"));
  } finally {
    await t.cleanup();
  }
});

test("ledger_check and ledger_suggest use the pane's one-shot too; the default stays the constructor's", async () => {
  const used = [];
  const routed = async (req) => {
    used.push(req.system.slice(0, 12));
    return { ok: false, error: "routed provider is down" };
  };
  const t = await setup({
    // The constructor's one-shot must not be asked once a route exists.
    oneshot: async () => {
      throw new Error("the default one-shot was used");
    },
    oneshotModel: "haiku",
    oneshotJob: async (key) => ({
      run: key === "k" ? routed : null,
      model: `codex:gpt-6.1-sol@${key}`,
    }),
  });
  try {
    await t.ask("ledger_add", {
      item: { kind: "instruction", text: "Use SI units", check: { type: "llm" } },
    });
    const check = await t.ask("ledger_check", {
      snapshot: { paragraphs: [{ index: 1, text: "Ten feet.", reviewed_text: "Ten feet." }] },
      full: true,
    });
    assert.equal(check.ok, true);
    const suggest = await t.ask("ledger_suggest", { text: "Reviewer 1\nUse SI units throughout." });
    assert.deepEqual([suggest.ok, suggest.method], [true, "regex"]);
    assert.match(suggest.warnings.join(" "), /routed provider is down/);
    assert.equal(used.length, 2, "one ledger check batch and one suggestion were routed");
    const log = await readDecisions(t.dir);
    assert.deepEqual([...new Set(log.map((r) => r.kind))].sort(), [
      "ledger_check",
      "ledger_suggest",
    ]);
    assert.ok(log.every((r) => r.model === "codex:gpt-6.1-sol@k"));
  } finally {
    await t.cleanup();
  }
});

// Codex review of stage 2.5: the provider of a job is taken once, when the
// job starts. A choice changed in the pane while the job runs applies to the
// next job; every request and every receipt of this one is the first choice.
test("a ledger check in several batches stays on the provider it started with, and its receipts say so", async () => {
  let choice = "ollama:qwen3.5:9b";
  const pins = [];
  const calls = [];
  const t = await setup({
    oneshot: async () => {
      throw new Error("the default one-shot was used");
    },
    // As oneshot-router.mjs pin(): the choice of this moment, kept by the handle.
    oneshotJob: async () => {
      const mine = choice;
      pins.push(mine);
      return {
        model: mine,
        run: async () => {
          calls.push(mine);
          choice = "codex:gpt-6.1-sol"; // the pane changes it while the job runs
          return { ok: true, text: "{}", json: { violations: [] } };
        },
      };
    },
  });
  try {
    for (const text of ["Use SI units", "Write numbers below ten as words"]) {
      await t.ask("ledger_add", { item: { kind: "instruction", text, check: { type: "llm" } } });
    }
    const paragraphs = Array.from({ length: 3 }, (_, i) => ({
      index: i,
      text: `Paragraph ${i} is ten feet long.`,
      reviewed_text: `Paragraph ${i} is ten feet long.`,
    }));
    const check = await t.ask("ledger_check", { snapshot: { paragraphs }, full: true });
    assert.equal(check.ok, true);
    assert.equal(pins.length, 1, "the provider was resolved once for the whole job");
    assert.ok(calls.length >= 2, `several requests in one job (${calls.length})`);
    assert.deepEqual([...new Set(calls)], ["ollama:qwen3.5:9b"]);
    const log = await readDecisions(t.dir);
    assert.ok(log.length >= 2);
    assert.deepEqual([...new Set(log.map((r) => r.model))], ["ollama:qwen3.5:9b"]);
    // The next job starts with the choice as it is then.
    await t.ask("ledger_check", { snapshot: { paragraphs }, full: true });
    assert.deepEqual(pins, ["ollama:qwen3.5:9b", "codex:gpt-6.1-sol"]);
  } finally {
    await t.cleanup();
  }
});

test("review_split: the split and the requirement extraction after it run on one provider", async () => {
  let choice = "ollama:qwen3.5:9b";
  const pins = [];
  const calls = [];
  const t = await setup({
    oneshotJob: async () => {
      const mine = choice;
      pins.push(mine);
      return {
        model: mine,
        run: async (req) => {
          calls.push(mine);
          choice = "codex:gpt-6.1-sol";
          // The split: both lines start a comment. The extraction then fails.
          if (!req.prompt.includes("\nLines:\n")) return { ok: false, error: "the model is down" };
          const lines = [1, 2].map((n) => ({ n, type: "comment_start" }));
          return { ok: true, json: { lines }, text: "{}" };
        },
      };
    },
  });
  try {
    // No reviewer heading: the regex split fails and the model is asked, then
    // the requirements of the comments are asked for.
    await t.ask("review_split", { text: "The units are missing.\nThe figure is unclear." });
    assert.equal(pins.length, 1);
    assert.ok(calls.length >= 2, `split and extraction (${calls.length})`);
    assert.deepEqual([...new Set(calls)], ["ollama:qwen3.5:9b"]);
    const log = await readDecisions(t.dir);
    assert.deepEqual([...new Set(log.map((r) => r.model))], ["ollama:qwen3.5:9b"]);
  } finally {
    await t.cleanup();
  }
});

// Codex review of stage 2.5: a review turn waits for the pane's confirmed
// provider like a judgment does. Without one it fails before anything is
// built or queued: no turn goes to a provider the pane did not confirm.
test("review_propose asks the turn gate first; a refusal starts no turn and changes nothing", async () => {
  let gate = async () => {
    throw new Error(
      "The stored provider for review turns cannot be used: Codex CLI is not available.",
    );
  };
  const t = await setup({ turnGate: (key) => gate(key) });
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    const r = await t.ask("review_propose", { ref: "R1.1" });
    assert.deepEqual(
      [r.ok, r.error],
      [false, "The stored provider for review turns cannot be used: Codex CLI is not available."],
    );
    assert.deepEqual(t.pushed, [], "no turn was queued");
    const rev = await loadRevision(join(t.dir, "paper.revision.json"), t.doc);
    assert.equal(rev.comments.find((c) => c.ref === "R1.1").status, "pending");
    // Once the provider is confirmed the same request starts the turn.
    const seen = [];
    gate = async (key) => void seen.push(key);
    const ok = await t.ask("review_propose", { ref: "R1.1" });
    assert.equal(ok.started, true);
    assert.deepEqual(seen, ["k"]);
    assert.equal(t.pushed.length, 1);
  } finally {
    await t.cleanup();
  }
});

// Codex review 14: a normal tracked edit Word applied while the connection
// dropped left nothing in the file: the attempt was only in memory, no late
// result came, and the refresh recovery read blocked exceptions only. The
// change and its rationale comment were in Word with no change record.
const P_BEFORE = "Cells grew on the substrate at 30 C.";
const P_AFTER = "Cells grew on the medium at 30 C.";
const UNANSWERED_COMMENT = "[R1.3] w / y / a";
// reviewed: the paragraph with pending changes accepted (Word reviewed text).
const snap = (text, { comments = [], marks = 0, reviewed = null } = {}) => ({
  paragraphs: [
    { index: 0, id: "p0", text: "Title", w_ins: 0, w_del: 0 },
    { index: 1, id: "p1", text, reviewed_text: reviewed, w_ins: marks, w_del: marks },
  ],
  comments,
});
// A normal edit ("the substrate" -> "the medium") that Word never answered.
async function unansweredEdit(t, { id = "x1" } = {}) {
  await t.ask("review_save_comments", { comments: COMMENTS });
  await t.ask("review_refresh", { snapshot: snap(P_BEFORE) });
  const deps = {
    ...t.svc.trackedEditDeps("k"),
    call: async () => {
      throw noAnswer(id);
    },
  };
  return applyTrackedEdit(
    deps,
    {
      comment_ref: "R1.3",
      find: "the substrate",
      replace: "the medium",
      rationale: { what: "w", why: "y", addresses: "a" },
    },
    { ref: "R1.3" },
  );
}
const r13 = async (t) => (await t.ask("review_load")).revision.comments[1];

test("an unanswered normal edit is written to the file as outcome_unknown, with what it would record", async () => {
  for (const id of ["x1", null]) {
    const t = await setup();
    try {
      const r = await unansweredEdit(t, { id });
      assert.equal(r.code, "OutcomeUnknown");
      const c = await r13(t);
      assert.equal(c.changes.length, 0);
      assert.equal(c.blocked.length, 1);
      const b = c.blocked[0];
      assert.deepEqual([b.id, b.state, b.blocked_by], ["R1.3-b1", "outcome_unknown", []]);
      assert.deepEqual([b.args.find, b.args.replace], ["the substrate", "the medium"]);
      assert.equal(b.attempt.tool_call_id, id);
      assert.equal(b.attempt.comment_text, UNANSWERED_COMMENT);
      assert.deepEqual(
        [b.attempt.change.paragraph_index, b.attempt.change.before, b.attempt.change.after],
        [1, P_BEFORE, P_AFTER],
      );
      assert.equal(c.blocked_seq, 1);
      assert.ok(t.lines.some((l) => /outcome unknown R1\.3-b1 kept for refresh/.test(l)));
    } finally {
      await t.cleanup();
    }
  }
});

test("refresh reconciles an unanswered normal edit: applied, not applied, or still unknown", async () => {
  const comment = { id: "12", paragraph_index: 1, content: UNANSWERED_COMMENT };
  const cases = [
    // [name, snapshot, changes, blocked left, state of the change]
    ["accepted text and the comment", snap(P_AFTER, { comments: [comment] }), 1, 0, "accepted"],
    [
      "pending tracked replace and the comment",
      snap("Cells grew on the mediumthe substrate at 30 C.", {
        comments: [comment],
        marks: 1,
        reviewed: P_AFTER,
      }),
      1,
      0,
      "pending",
    ],
    ["before text intact, no comment", snap(P_BEFORE), 0, 0, null],
    ["paragraph rewritten by hand, no comment", snap("Cells grew well at 30 C."), 0, 1, null],
    ["paragraph gone", snap("Something else entirely."), 0, 1, null],
  ];
  for (const [name, snapshot, changes, blocked, state] of cases) {
    const t = await setup();
    try {
      await unansweredEdit(t);
      const r = await t.ask("review_refresh", { snapshot });
      const c = r.revision.comments[1];
      assert.deepEqual([c.changes.length, c.blocked.length], [changes, blocked], name);
      if (changes) {
        assert.deepEqual(
          [c.changes[0].state, c.changes[0].word_comment_id, c.changes[0].text],
          [state, "12", "the medium"],
          name,
        );
        assert.deepEqual(r.external, [], name);
      }
      if (blocked) assert.equal(c.blocked[0].state, "outcome_unknown", name);
      if (name.startsWith("before text")) {
        assert.ok(
          t.lines.some((l) => /R1\.3-b1 not applied.*dropped/.test(l)),
          name,
        );
      }
    } finally {
      await t.cleanup();
    }
  }
});

test("a daemon that restarted reconciles the unanswered edit it finds in the file", async () => {
  for (const [text, comments, changes] of [
    [P_AFTER, [{ id: "12", paragraph_index: 1, content: UNANSWERED_COMMENT }], 1],
    [P_BEFORE, [], 0],
  ]) {
    const t = await setup();
    try {
      await unansweredEdit(t);
      // A second service on the same folder: nothing of the first in memory.
      const svc2 = createReviewService({
        getContext: () => ({ activeDoc: t.doc, caps: ["target_hash"] }),
        callTool: async () => ({ ok: true }),
        pushTurn: () => {},
        sendEvent: () => {},
        log: () => {},
        env: ENV,
      });
      const ask2 = (type, payload = {}) =>
        new Promise((resolve) =>
          svc2.handlers[type]({ type, request_id: "q", ...payload }, resolve, "k", "word"),
        );
      const before = (await ask2("review_load")).revision.comments[1];
      assert.equal(before.blocked[0].state, "outcome_unknown");
      const r = await ask2("review_refresh", { snapshot: snap(text, { comments }) });
      const c = r.revision.comments[1];
      assert.deepEqual([c.changes.length, c.blocked.length], [changes, 0]);
      // The late answer of the old call means nothing to this daemon.
      await svc2.lateToolResult("k", {
        id: "x1",
        ok: true,
        result: { ok: true, word_comment_id: 12, paragraph_index: 1, before: "b", after: "a" },
      });
      assert.equal((await ask2("review_load")).revision.comments[1].changes.length, changes);
    } finally {
      await t.cleanup();
    }
  }
});

test("a late answer settles the unanswered normal edit: recorded, or dropped when Word says not applied", async () => {
  let t = await setup();
  try {
    await unansweredEdit(t);
    await t.svc.lateToolResult("k", {
      id: "x1",
      ok: true,
      result: {
        ok: true,
        word_comment_id: 12,
        paragraph_index: 1,
        before: P_BEFORE,
        after: P_AFTER,
      },
    });
    const c = await r13(t);
    assert.deepEqual([c.changes.length, c.blocked.length], [1, 0]);
    assert.equal(c.changes[0].word_comment_id, "12");
  } finally {
    await t.cleanup();
  }
  for (const entry of [
    { id: "x1", ok: false, error: "GeneralException" },
    { id: "x1", ok: true, result: { ok: false, code: "NotFound", message: "not found" } },
  ]) {
    t = await setup();
    try {
      await unansweredEdit(t);
      await t.svc.lateToolResult("k", entry);
      const c = await r13(t);
      assert.deepEqual([c.changes.length, c.blocked.length], [0, 0]);
    } finally {
      await t.cleanup();
    }
  }
});

test("an unanswered normal edit can be applied again from the Review tab (force), like an exception", async () => {
  const t = await setup({
    callResult: {
      ok: true,
      word_comment_id: 9,
      paragraph_index: 1,
      before: P_BEFORE,
      after: P_AFTER,
    },
  });
  try {
    await unansweredEdit(t);
    const plain = await t.ask("review_override", { ref: "R1.3", blocked_id: "R1.3-b1" });
    assert.equal(plain.ok, false);
    const forced = await t.ask("review_override", {
      ref: "R1.3",
      blocked_id: "R1.3-b1",
      force: true,
    });
    assert.deepEqual([forced.ok, forced.applied], [true, true]);
    const c = forced.revision.comments[1];
    assert.deepEqual([c.changes.length, c.blocked.length], [1, 0]);
    assert.deepEqual(c.changes[0].override_of, []);
    assert.equal(t.calls[0][1].comment_text, UNANSWERED_COMMENT); // no exception note
  } finally {
    await t.cleanup();
  }
});

// Codex review 15, class audit: a handler reads the pane's document once,
// when the request arrives. What it does after an await (a model call, the
// answer of Word, a queued turn) must not read the context again: the pane
// may be gone by then, or its document saved under another name.
test("no handler reads the pane's document again after its first await", async () => {
  const ctx = { activeDoc: null, caps: ["target_hash"] };
  const hooks = { oneshot: null, call: null, push: null };
  const t = await setup({
    getContext: () => ({ ...ctx }),
    oneshot: async (req) => hooks.oneshot(req),
    callImpl: (name, args) => hooks.call(name, args),
    pushTurn: (...a) => hooks.push?.(...a),
  });
  const other = join(t.dir, "renamed.docx");
  const gone = async (file) => assert.rejects(readFile(join(t.dir, file), "utf8"), file);
  const here = () => (ctx.activeDoc = t.doc);
  try {
    here();
    await t.ask("review_save_comments", { comments: COMMENTS });
    await t.ask("ledger_add", { item: { kind: "instruction", text: "t", check: { type: "llm" } } });
    await t.ask("ledger_add", { item: NOVEL });

    // ledger_check job: the pane goes away while the model answers.
    hooks.oneshot = async (req) => {
      ctx.activeDoc = null;
      return ledgerReply(req, []);
    };
    const checked = await t.ask("ledger_check", {
      ids: ["L1"],
      snapshot: { paragraphs: [{ index: 0, text: "A paragraph.", reviewed_text: null }] },
    });
    assert.equal(checked.ok, true, checked.error);
    here();
    assert.equal((await t.ask("review_load")).revision.ledger[0].status, "ok");
    assert.equal((await readDecisions(t.dir)).filter((d) => d.kind === "ledger_check").length, 1);

    // review_split job: the document is renamed while the model answers.
    hooks.oneshot = async () => {
      ctx.activeDoc = other;
      return { ok: false, error: "x" };
    };
    const split = await t.ask("review_split", { text: "Please add error bars.\nAlso state n." });
    assert.equal(split.ok, true, split.error);
    here();
    assert.ok((await readDecisions(t.dir)).some((d) => d.kind === "comment_split"));
    await gone("renamed.decisions.jsonl");

    // review_propose: the turn is queued, then the status is written.
    hooks.push = () => (ctx.activeDoc = null);
    const proposed = await t.ask("review_propose", { ref: "R1.1" });
    assert.equal(proposed.ok, true, proposed.error);
    assert.equal(proposed.revision.comments[0].status, "in_progress");
    hooks.push = null;
    here();

    // review_override: Word answers after the document was renamed.
    const deps = { ...t.svc.trackedEditDeps("k"), call: async () => ({ ok: true }) };
    const R = { what: "w", why: "y", addresses: "a" };
    await applyTrackedEdit(
      deps,
      { comment_ref: "R1.3", find: "substrate", replace: "novel substrate", rationale: R },
      { ref: "R1.3" },
    );
    hooks.call = () => {
      ctx.activeDoc = other;
      return { ok: true, word_comment_id: 5, paragraph_index: 2, before: "b", after: "a" };
    };
    const over = await t.ask("review_override", { ref: "R1.3", blocked_id: "R1.3-b1" });
    assert.deepEqual([over.ok, over.applied], [true, true], over.error);
    here();
    const rev = (await t.ask("review_load")).revision;
    assert.deepEqual([rev.comments[1].changes.length, rev.comments[1].blocked.length], [1, 0]);
    await gone("renamed.revision.json");

    // A tracked edit of a turn: blocked, and applied, after the pane is gone.
    const turn = { ...t.svc.trackedEditDeps("k") };
    turn.call = async () => {
      ctx.activeDoc = null;
      return { ok: true, word_comment_id: 6, paragraph_index: 2, before: "b", after: "a" };
    };
    const ok = await applyTrackedEdit(
      turn,
      { comment_ref: "R1.1", find: "30 degrees", replace: "30 °C", rationale: R },
      { ref: "R1.1" },
    );
    assert.equal(ok.ok, true, ok.message);
    here();
    assert.equal((await t.ask("review_load")).revision.comments[0].changes.length, 1);
  } finally {
    await t.cleanup();
  }
});

// The receipt of a requirement link holds no review text; the section of the
// comment is part of what its fingerprint covers.
test("receipts: the section of a comment is in the fingerprinted input of a requirement link", async () => {
  const link = async (section) => {
    const t = await setup({ getModel: () => "sonnet" });
    try {
      await t.ask("review_save_comments", {
        comments: [
          {
            ref: "R1.1",
            reviewer: 1,
            text: "Use Celsius.",
            ...(section && { section }),
            requirements: ["use Celsius"],
          },
        ],
      });
      const deps = {
        ...t.svc.trackedEditDeps("k"),
        call: async () => ({
          ok: true,
          word_comment_id: 4,
          paragraph_index: 1,
          before: "b",
          after: "a",
        }),
      };
      await applyTrackedEdit(
        deps,
        {
          comment_ref: "R1.1",
          find: "30 degrees",
          replace: "30 °C",
          rationale: { what: "w", why: "y", addresses: "a", requirement_ids: ["q1"] },
        },
        { ref: "R1.1" },
      );
      const raw = await readFile(join(t.dir, "paper.decisions.jsonl"), "utf8");
      const rows = await readDecisions(t.dir);
      return { raw, row: rows.find((r) => r.kind === "requirement_link") };
    } finally {
      await t.cleanup();
    }
  };
  const plain = await link(null);
  const abstract = await link("Abstract of the paper");
  const methods = await link("Methods of the paper");
  assert.match(plain.row.input_fingerprint, /^[0-9a-f]{64}$/);
  assert.notEqual(abstract.row.input_fingerprint, plain.row.input_fingerprint);
  assert.notEqual(abstract.row.input_fingerprint, methods.row.input_fingerprint);
  assert.ok(!abstract.raw.includes("Abstract of the paper"));
});

test("review_split: the section of a comment reaches the requirement extraction and the result", async () => {
  let prompt = "";
  const t = await setup({
    oneshotModel: "haiku",
    oneshot: async (req) => {
      prompt = req.prompt;
      return {
        ok: true,
        json: {
          sentences: [
            { ref: "R1.1", s: 1, type: "request" },
            { ref: "R1.2", s: 1, type: "request" },
          ],
        },
      };
    },
  });
  try {
    const r = await t.ask("review_split", {
      text: "Reviewer 1:\nAbstract:\n1. Shorten this section.\nMethods:\n1. Expand this section.\n",
    });
    assert.deepEqual(
      r.comments.map((c) => [c.ref, c.section, c.text, c.requirements.map((q) => q.text)]),
      [
        ["R1.1", "Abstract", "Shorten this section.", ["Shorten this section."]],
        ["R1.2", "Methods", "1. Expand this section.", ["Expand this section."]],
      ],
    );
    assert.match(prompt, /\[R1\.1 section\] Abstract\n\[R1\.1 s1\] Shorten this section\./);
    const saved = await t.ask("review_save_comments", { comments: r.comments });
    assert.deepEqual(
      saved.revision.comments.map((c) => [c.ref, c.section]),
      [
        ["R1.1", "Abstract"],
        ["R1.2", "Methods"],
      ],
    );
    const preview = await t.ask("review_preview", { ref: "R1.2", paragraphs: [] });
    assert.match(
      preview.prompt,
      /comment R1\.2 \(verbatim\):\nSection of the review this comment is under: Methods\nThe comment is the text between[^\n]*\n<<<([0-9a-f]{12})\n1\. Expand this section\.\n>>>\1\n/,
    );
    assert.ok(
      t.lines.some((l) => /^\[review\] prompt R1\.2 .* section="Methods"/.test(l)),
      t.lines.join("\n"),
    );
  } finally {
    await t.cleanup();
  }
});

// L2 reads the comment as a source of new words: the section label is part
// of what the reviewer wrote, so a word of it is not flagged as unsourced.
test("the review context of a tracked edit gives the comment text with its section", async () => {
  const t = await setup();
  try {
    await t.ask("review_save_comments", {
      comments: [
        { ref: "R1.1", reviewer: 1, text: "Shorten this section.", section: "Abstract" },
        { ref: "R1.2", reviewer: 1, text: "Add units." },
      ],
    });
    const deps = t.svc.trackedEditDeps("k");
    assert.equal(
      (await deps.getReviewContext("R1.1")).comment_text,
      "Abstract\nShorten this section.",
    );
    assert.equal((await deps.getReviewContext("R1.2")).comment_text, "Add units.");
  } finally {
    await t.cleanup();
  }
});

// An empty section from the pane clears the stored label, so it is not in
// the next prompt; a save without the key (an older pane) keeps it.
test("review_save_comments: an empty section clears the label, a missing key keeps it", async () => {
  const t = await setup();
  try {
    const first = [
      { ref: "R1.1", reviewer: 1, text: "Shorten this section.", section: "Abstract" },
      { ref: "R1.2", reviewer: 1, text: "Expand this section.", section: "Methods" },
    ];
    await t.ask("review_save_comments", { comments: first });
    const saved = await t.ask("review_save_comments", {
      comments: [
        { ref: "R1.1", reviewer: 1, text: "Shorten the abstract.", section: "" },
        { ref: "R1.2", reviewer: 1, text: "Expand this section." },
      ],
    });
    assert.deepEqual(
      saved.revision.comments.map((c) => [c.ref, c.section]),
      [
        ["R1.1", undefined],
        ["R1.2", "Methods"],
      ],
    );
    const cleared = await t.ask("review_preview", { ref: "R1.1", paragraphs: [] });
    assert.ok(!cleared.prompt.includes("Section of the review"), cleared.prompt);
    const kept = await t.ask("review_preview", { ref: "R1.2", paragraphs: [] });
    assert.ok(kept.prompt.includes("Section of the review this comment is under: Methods"));
  } finally {
    await t.cleanup();
  }
});

// A review_propose whose pane has stopped waiting (it says how long it waits
// in wait_ms) must not start a turn when the gate lets it through later: the
// author has seen that request fail.
test("review_propose: a request the pane has given up starts no turn", async () => {
  let clock = 1_000;
  let gateTakes = 11_000;
  const t = await setup({
    clock: () => clock,
    turnGate: async () => {
      clock += gateTakes; // the provider was confirmed after this long
    },
  });
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    const late = await t.ask("review_propose", { ref: "R1.1", wait_ms: 10_000 });
    assert.deepEqual(
      [late.ok, late.error],
      [
        false,
        "The task pane stopped waiting for this request after 10 s, so no turn was started. Press Propose changes again.",
      ],
    );
    assert.deepEqual(t.pushed, [], "no turn was queued");
    const rev = await loadRevision(join(t.dir, "paper.revision.json"), t.doc);
    assert.equal(rev.comments.find((c) => c.ref === "R1.1").status, "pending");
    // In time: the turn starts. A pane that sends no wait_ms is not held to one.
    gateTakes = 9_000;
    assert.equal((await t.ask("review_propose", { ref: "R1.1", wait_ms: 10_000 })).started, true);
    gateTakes = 60_000;
    assert.equal((await t.ask("review_propose", { ref: "R1.3" })).started, true);
  } finally {
    await t.cleanup();
  }
});

// H4 with the job pin of stage 2.5: the second request for a batch whose
// reply was in the wrong form goes to the provider and model the check
// started with, also when the pane's choice changed in between.
test("ledger_check: the retry of a malformed reply runs on the pinned provider of the job", async () => {
  const asked = [];
  let pins = 0;
  const provider = (name, bad) => async (req) => {
    asked.push(name);
    return bad && asked.length === 1
      ? { ok: true, json: { violations: [] } }
      : ledgerReply(req, []);
  };
  const t = await setup({
    // Each job pins the choice of its start; the choice changes after the first pin.
    oneshotJob: async () =>
      ++pins === 1
        ? { run: provider("ollama:first", true), model: "ollama:first" }
        : { run: provider("claude", false), model: "haiku" },
  });
  try {
    await t.ask("ledger_add", { item: { kind: "instruction", text: "t", check: { type: "llm" } } });
    const r = await t.ask("ledger_check", {
      snapshot: { paragraphs: [{ index: 0, text: "A paragraph.", reviewed_text: null }] },
    });
    assert.equal(r.revision.ledger[0].status, "ok");
    assert.deepEqual([pins, asked], [1, ["ollama:first", "ollama:first"]]);
    const rows = (await readDecisions(t.dir)).filter((d) => d.kind === "ledger_check");
    assert.deepEqual(
      rows.map((d) => [d.model, d.output.verdict]),
      [["ollama:first", "ok"]],
    );
  } finally {
    await t.cleanup();
  }
});

// ---- Proposal turns (a provider without tools: Codex) -------------------------
// The review service asks, per request, whether the pane's provider for
// review turns proposes instead of calling tools (proposalPreview). When it
// does, the answer holds the choice and the request that would go out.
const BUILT = {
  ok: true,
  system: "SYSTEM TEXT",
  prompt: "[Proposal turn R1.1]\n[p1] A paragraph.",
  chars: 40,
  scope: "manuscript",
  shown: [{ id: "p1", index: 1 }],
  instructionIds: [],
  forbidIds: ["L1"],
  omittedForbidIds: [],
  section: null,
};
const CODEX_CHOICE = { id: "codex", model: "gpt-6.1-sol" };
const PROPOSAL_PANE = { caps: ["target_hash", "proposal_turns"] };

test("review_preview shows the request of a proposal turn, system text and prompt, as it would go out", async () => {
  const asked = [];
  let built = BUILT;
  const t = await setup({
    proposalPreview: async (key, ref, source) => {
      asked.push([key, ref, source.selected, typeof source.at.path]);
      return { choice: CODEX_CHOICE, built };
    },
  });
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    const sel = [{ id: null, index: 1, text: "A paragraph." }];
    const r = await t.ask("review_preview", { ref: "R1.1", paragraphs: sel });
    assert.equal(r.prompt, "SYSTEM TEXT\n\n[Proposal turn R1.1]\n[p1] A paragraph.");
    assert.deepEqual(asked, [["k", "R1.1", sel, "string"]]);
    assert.ok(
      t.lines.includes(
        "[review] prompt R1.1 instructions=- forbid=L1 omitted_forbid=- paragraphs=1 (preview)",
      ),
    );
    // A request that cannot be built is said in its own words.
    built = { ok: false, code: "TooLong", message: "The manuscript is too long for Codex." };
    const no = await t.ask("review_preview", { ref: "R1.1" });
    assert.deepEqual([no.ok, no.error], [false, "The manuscript is too long for Codex."]);
  } finally {
    await t.cleanup();
  }
});

test("review_propose of a proposal turn: the turn carries the choice and the selection, and no prompt", async () => {
  let built = BUILT;
  let ctx = PROPOSAL_PANE;
  let doc = null;
  const t = await setup({
    getContext: () => ({ activeDoc: doc, ...ctx }),
    proposalPreview: async () => ({ choice: CODEX_CHOICE, built }),
  });
  doc = t.doc;
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    const sel = [{ id: null, index: 1, text: "A paragraph." }];
    // A request that cannot be built starts no turn: the refusal is the answer.
    built = { ok: false, code: "NoSnapshot", message: "The manuscript has not been read yet." };
    const no = await t.ask("review_propose", { ref: "R1.1", paragraphs: sel });
    assert.deepEqual([no.ok, no.error], [false, "The manuscript has not been read yet."]);
    assert.deepEqual(t.pushed, []);
    // A task pane from before proposal turns: it has shown its author no
    // notice of what goes to Codex. Nothing is sent.
    built = BUILT;
    ctx = { caps: ["target_hash"] };
    const old = await t.ask("review_propose", { ref: "R1.1", paragraphs: sel });
    assert.deepEqual(
      [old.ok, old.error],
      [
        false,
        "This task pane is older than the daemon and does not know the proposal turns of Codex CLI (ChatGPT sign-in), so nothing was sent. Reload the task pane (close and reopen it), then propose again.",
      ],
    );
    assert.deepEqual(t.pushed, []);
    ctx = PROPOSAL_PANE;
    const r = await t.ask("review_propose", { ref: "R1.1", paragraphs: sel });
    assert.equal(r.started, true);
    assert.deepEqual(t.pushed, [
      {
        key: "k",
        host: "word",
        text: null,
        turn: { ref: "R1.1", selected: sel, choice: CODEX_CHOICE },
      },
    ]);
    assert.equal(r.revision.comments.find((c) => c.ref === "R1.1").status, "in_progress");
  } finally {
    await t.cleanup();
  }
  // A Claude turn is as before: a prompt and { ref }.
  const claude = await setup();
  try {
    await claude.ask("review_save_comments", { comments: COMMENTS });
    await claude.ask("review_propose", { ref: "R1.1" });
    assert.deepEqual(claude.pushed[0].turn, { ref: "R1.1" });
    assert.equal(typeof claude.pushed[0].text, "string");
  } finally {
    await claude.cleanup();
  }
});

// Stop pressed while the request is still on its way to a turn (the gate
// waits for the provider, the review file is read): no turn starts after it.
test("cancelStarting: a review_propose stopped before its turn started starts none", async () => {
  let open;
  const gate = new Promise((r) => (open = r));
  const t = await setup({ turnGate: () => gate });
  try {
    await t.ask("review_save_comments", { comments: COMMENTS });
    assert.equal(t.svc.cancelStarting("k"), false, "nothing is starting");
    const asked = t.ask("review_propose", { ref: "R1.1" });
    await new Promise((r) => setTimeout(r, 5));
    assert.equal(t.svc.cancelStarting("k"), true);
    assert.equal(t.svc.cancelStarting("other"), false);
    open();
    const r = await asked;
    assert.deepEqual([r.ok, r.error], [false, "The turn was stopped."]);
    assert.deepEqual(t.pushed, []);
    const rev = await loadRevision(join(t.dir, "paper.revision.json"), t.doc);
    assert.equal(rev.comments.find((c) => c.ref === "R1.1").status, "pending");
    // The next request is a new one; once its turn is queued, a Stop is the
    // runner's or the session's business.
    const again = await t.ask("review_propose", { ref: "R1.1" });
    assert.equal(again.started, true);
    assert.equal(t.svc.cancelStarting("k"), false);
  } finally {
    await t.cleanup();
  }
});

test("turnEnded passes a notice of the turn on with turn_done; the handle and the forbid switch are exported", async () => {
  const t = await setup({ env: { ...ENV, CC_OFFICE_REVIEW_PROMPT_FORBID: "0" } });
  try {
    t.svc.turnEnded(
      "k",
      { ref: "R1.1" },
      { ok: true, notice: "1 of 2 requests have a proposed change; q2 has none." },
    );
    t.svc.turnEnded("k", { ref: "R1.1" }, { ok: true });
    assert.deepEqual(
      t.events.filter((e) => e.event === "turn_done"),
      [
        {
          type: "review_event",
          event: "turn_done",
          ref: "R1.1",
          ok: true,
          notice: "1 of 2 requests have a proposed change; q2 has none.",
        },
        { type: "review_event", event: "turn_done", ref: "R1.1", ok: true },
      ],
    );
    const d = t.svc.at("k");
    assert.deepEqual(
      [d.key, d.path, d.caps],
      ["k", join(t.dir, "paper.revision.json"), ["target_hash"]],
    );
    assert.equal(t.svc.includeForbid(), false);
  } finally {
    await t.cleanup();
  }
});
