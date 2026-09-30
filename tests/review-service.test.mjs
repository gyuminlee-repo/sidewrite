import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdir, mkdtemp, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createReviewService } from "../daemon/review-service.mjs";
import { applyTrackedEdit } from "../daemon/review-tools.mjs";
import { loadRevision, saveRevision } from "../daemon/revision-store.mjs";
import { REQUIREMENT_SYSTEM } from "../daemon/review-split.mjs";

const ENV = { CC_OFFICE_WSL_PATHS: "0" };

async function setup({ env = ENV, oneshot = null, callResult, callImpl = null, ...more } = {}) {
  const dir = await mkdtemp(join(tmpdir(), "review-svc-"));
  const doc = join(dir, "paper.docx");
  const events = [];
  const lines = [];
  const pushed = [];
  const calls = [];
  const svc = createReviewService({
    getContext: () => ({ activeDoc: doc }),
    callTool: async (name, args) => {
      calls.push([name, args]);
      if (callImpl) return callImpl(name, args, calls.length);
      return (
        callResult ?? { ok: true, word_comment_id: 5, paragraph_index: 2, before: "b", after: "a" }
      );
    },
    pushTurn: (key, host, text, turn) => pushed.push({ key, host, text, turn }),
    oneshot,
    sendEvent: (key, obj) => events.push(obj),
    log: (s) => lines.push(s),
    env,
    ...more,
  });
  const ask = (type, payload = {}) =>
    new Promise((resolve) =>
      svc.handlers[type]({ type, request_id: "q1", ...payload }, resolve, "k", "word"),
    );
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
    assert.deepEqual(t.events.at(-1), { type: "review_event", event: "turn_done", ref: "R1.1" });
    assert.equal(t.lines.at(-1), "[review] turn done R1.1");
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
    assert.deepEqual(
      t.events.map((e) => e.event),
      ["change_recorded", "late_result_recorded"],
    );
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
      getContext: () => ({ activeDoc: t.doc }),
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
    oneshot: async () => ({
      ok: true,
      json: { violations: [], unsure: [{ paragraph_index: 3, reason: "cannot tell" }] },
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
      return {
        ok: true,
        json: {
          violations: [{ paragraph_index: 3, quote: "hidden words", reason: "r" }],
          unsure: [{ paragraph_index: 4, reason: "?" }],
        },
      };
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
    oneshot: async () => {
      await gate;
      return { ok: true, json: { violations: [], unsure: [] } };
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
    oneshot: async () => {
      await gate;
      return { ok: true, json: { violations: [], unsure: [] } };
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
    oneshot: async () => {
      if (++n === 1) {
        await gate;
        return { ok: true, json: { violations: [], unsure: [] } };
      }
      return {
        ok: true,
        json: { violations: [{ paragraph_index: 0, quote: "novel", reason: "r" }], unsure: [] },
      };
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
