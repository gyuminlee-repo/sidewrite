import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, readdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createReviewService } from "../daemon/review-service.mjs";
import { applyTrackedEdit } from "../daemon/review-tools.mjs";

const ENV = { CC_OFFICE_WSL_PATHS: "0" };

async function setup({ env = ENV, oneshot = null, callResult } = {}) {
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
      return (
        callResult ?? { ok: true, word_comment_id: 5, paragraph_index: 2, before: "b", after: "a" }
      );
    },
    pushTurn: (key, host, text, turn) => pushed.push({ key, host, text, turn }),
    oneshot,
    sendEvent: (key, obj) => events.push(obj),
    log: (s) => lines.push(s),
    env,
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
