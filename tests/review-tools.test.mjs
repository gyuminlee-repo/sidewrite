import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { z } from "zod";
import {
  REVIEW_READ_TOOLS,
  REVIEW_BLOCKED_TOOLS,
  REVIEW_WRITE_TOOL,
  reviewTurnToolAllowed,
  guardedCall,
  trackedEditShape,
  formatRationaleComment,
  TOOL_COMMENT_PREFIX,
  applyTrackedEdit,
} from "../daemon/review-tools.mjs";

async function registeredWordTools() {
  const src = await readFile(new URL("../daemon/office-tools.mjs", import.meta.url), "utf8");
  const m = /const wordTools = \[([\s\S]*?)\];/.exec(src);
  assert.ok(m, "wordTools array not found in office-tools.mjs");
  return m[1]
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
}

test("every registered Word tool is classified for review turns", async () => {
  const names = await registeredWordTools();
  const unclassified = names.filter(
    (n) => !REVIEW_READ_TOOLS.has(n) && !REVIEW_BLOCKED_TOOLS.has(n) && n !== REVIEW_WRITE_TOOL,
  );
  assert.deepEqual(unclassified, []);
  for (const n of [...REVIEW_READ_TOOLS, ...REVIEW_BLOCKED_TOOLS])
    assert.ok(names.includes(n), `${n} is not registered`);
  console.log(
    `classified ${names.length} Word tools: read ${REVIEW_READ_TOOLS.size}, blocked ${REVIEW_BLOCKED_TOOLS.size}`,
  );
});

test("review turn allows reads and office_tracked_edit only", () => {
  assert.equal(reviewTurnToolAllowed("office_find"), true);
  assert.equal(reviewTurnToolAllowed("office_tracked_edit"), true);
  assert.equal(reviewTurnToolAllowed("office_replace_text"), false);
  assert.equal(reviewTurnToolAllowed("office_add_comment"), false);
  assert.equal(reviewTurnToolAllowed("excel_write_range"), false);
});

test("guardedCall refuses writers only while a review turn is active", async () => {
  const calls = [];
  const raw = async (name, args) => (calls.push(name), { ok: true, name, args });
  let turn = null;
  const call = guardedCall(raw, () => turn);
  assert.equal((await call("office_replace_text", {})).ok, true, "normal turn: upstream behavior");
  turn = { ref: "R1.2" };
  await assert.rejects(call("office_replace_text", {}), /disabled during review turn R1\.2/);
  await assert.rejects(call("office_clear_comments", { all: true }), /disabled/);
  assert.equal((await call("office_read_paragraphs", {})).ok, true);
  assert.deepEqual(calls, ["office_replace_text", "office_read_paragraphs"]);
});

test("schema: find is limited to 255 chars, ref format enforced, override is not exposed", () => {
  const s = z.object(trackedEditShape);
  const good = {
    comment_ref: "R1.3",
    find: "x",
    replace: "y",
    rationale: { what: "a", why: "b", addresses: "c" },
  };
  assert.equal(s.safeParse(good).success, true);
  assert.equal(s.safeParse({ ...good, find: "x".repeat(256) }).success, false);
  assert.equal(s.safeParse({ ...good, comment_ref: "1.3" }).success, false);
  assert.equal("override" in s.parse({ ...good, override: true }), false);
});

test("rationale comment format and the tool-comment prefix", () => {
  const t = formatRationaleComment("R1.3", {
    what: "단위 표기\n수정",
    why: "리뷰어 요청",
    addresses: "첫 문장",
  });
  assert.equal(t, "[R1.3] 단위 표기 수정 / 리뷰어 요청 / 첫 문장");
  assert.ok(TOOL_COMMENT_PREFIX.test(t));
  assert.ok(!TOOL_COMMENT_PREFIX.test("Check units later [R1.3]"));
  assert.ok(!TOOL_COMMENT_PREFIX.test("[Ref 1] user note"));
});

const NOVEL = {
  id: "L2",
  kind: "forbid",
  text: "novel 금지",
  check: { type: "regex", pattern: String.raw`\bnovel\b`, flags: "i", expect: "absent" },
};
const R = { what: "무엇", why: "왜", addresses: "어디" };

function deps(result = {}) {
  const log = [];
  const d = {
    calls: [],
    edits: [],
    blocked: [],
    log: (s) => log.push(s),
    lines: log,
    getLedger: async () => [NOVEL],
    call: async (name, args) => {
      d.calls.push([name, args]);
      return {
        ok: true,
        word_comment_id: 123,
        paragraph_index: 9,
        paragraph_id: "P9",
        before: "held at 30 degrees",
        after: "held at 30 °C",
        ...result,
      };
    },
    onEdit: async ({ ref, change }) => (d.edits.push([ref, change]), { id: `${ref}-1` }),
    onBlocked: async (b) => (d.blocked.push(b), { id: `${b.ref}-b1` }),
  };
  return d;
}

test("outside a review turn or with another ref: refused, nothing applied", async () => {
  const d = deps();
  const args = { comment_ref: "R1.1", find: "30 degrees", replace: "30 °C", rationale: R };
  assert.match((await applyTrackedEdit(d, args, null)).message, /only in a review turn/);
  assert.match((await applyTrackedEdit(d, args, { ref: "R1.2" })).message, /does not match/);
  assert.equal(d.calls.length, 0);
});

test("exactly one of replace or insert_after", async () => {
  const d = deps();
  const r = await applyTrackedEdit(
    d,
    { comment_ref: "R1.1", find: "a", replace: "b", insert_after: "c", rationale: R },
    { ref: "R1.1" },
  );
  assert.match(r.message, /exactly one/);
  assert.equal(d.calls.length, 0);
});

test("replace: applied once, comment text composed, change recorded", async () => {
  const d = deps();
  const r = await applyTrackedEdit(
    d,
    { comment_ref: "R1.1", find: "30 degrees", replace: "30 °C", rationale: R },
    { ref: "R1.1" },
  );
  assert.deepEqual(
    [r.ok, r.kind, r.change_id, r.word_comment_id],
    [true, "replace", "R1.1-1", "123"],
  );
  assert.equal(d.calls.length, 1);
  assert.deepEqual(d.calls[0], [
    "office_tracked_edit",
    {
      find: "30 degrees",
      replace: "30 °C",
      narrow: { find: "degrees", replace: "°C" },
      comment_text: "[R1.1] 무엇 / 왜 / 어디",
    },
  ]);
  const [ref, change] = d.edits[0];
  assert.equal(ref, "R1.1");
  assert.deepEqual(
    [change.kind, change.text, change.before, change.after, change.paragraph_id],
    ["replace", "30 °C", "held at 30 degrees", "held at 30 °C", "P9"],
  );
  // Words around the changed range, for judging this change on its own.
  assert.deepEqual(change.context, { left: "held at ", right: "" });
  assert.match(
    d.lines[0],
    /^\[review\] change recorded R1\.1-1 kind=replace comment=123 paragraph=9 narrowed=0 lost=0 unsourced=0$/,
  );
});

test("delete and insert kinds", async () => {
  let d = deps();
  let r = await applyTrackedEdit(
    d,
    { comment_ref: "R1.2", find: "x", replace: "", rationale: R },
    { ref: "R1.2" },
  );
  assert.equal(r.kind, "delete");
  d = deps();
  r = await applyTrackedEdit(
    d,
    {
      comment_ref: "R1.3",
      find: "four days",
      insert_after: " (OD600)",
      paragraph_hint: "P18",
      rationale: R,
    },
    { ref: "R1.3" },
  );
  assert.equal(r.kind, "insert");
  assert.deepEqual(d.calls[0][1], {
    find: "four days",
    insert_after: " (OD600)",
    paragraph_hint: "P18",
    comment_text: "[R1.3] 무엇 / 왜 / 어디",
  });
});

test("forbid hit: nothing applied, blocked edit stored, message tells the model to rephrase", async () => {
  const d = deps();
  const r = await applyTrackedEdit(
    d,
    {
      comment_ref: "R1.3",
      find: "synthetic carbon source",
      replace: "novel synthetic carbon source",
      rationale: R,
    },
    { ref: "R1.3" },
  );
  assert.deepEqual([r.ok, r.blocked_by, r.blocked_id], [false, ["L2"], "R1.3-b1"]);
  assert.match(r.message, /Rephrase/);
  assert.equal(d.calls.length, 0);
  assert.equal(d.edits.length, 0);
  assert.deepEqual(d.blocked[0].args.replace, "novel synthetic carbon source");
  assert.equal(d.lines[0], "[review] blocked R1.3 by=L2");
});

test("user-approved override applies and records the exception in the rationale", async () => {
  const d = deps();
  const r = await applyTrackedEdit(
    d,
    {
      comment_ref: "R1.3",
      find: "synthetic carbon source",
      replace: "novel synthetic carbon source",
      rationale: R,
    },
    { ref: "R1.3" },
    { override: { blocked_by: ["L2"] } },
  );
  assert.equal(r.ok, true);
  assert.equal(d.calls[0][1].comment_text, "[R1.3] 무엇 / 왜 (금지 규칙 L2 예외 승인) / 어디");
  assert.deepEqual(d.edits[0][1].override_of, ["L2"]);
});

test("task pane failure is passed through and nothing is recorded", async () => {
  const d = deps({ ok: false, code: "Ambiguous", message: "3 matches" });
  const r = await applyTrackedEdit(
    d,
    { comment_ref: "R1.2", find: "cell-free extract", replace: "lysate", rationale: R },
    { ref: "R1.2" },
  );
  assert.deepEqual([r.ok, r.code], [false, "Ambiguous"]);
  assert.equal(d.edits.length, 0);
});

// ---- L1, L2, L3, L5 in the tracked-edit flow --------------------------------

const withContext = (d, ctx) => ((d.getReviewContext = async () => ctx), d);

test("L1 lost facts go on the last line of the rationale comment and into the change", async () => {
  const d = deps();
  const r = await applyTrackedEdit(
    d,
    { comment_ref: "R1.4", find: " at 4000 × g for 10 min", replace: "", rationale: R },
    { ref: "R1.4" },
  );
  assert.deepEqual(r.lost_facts, ["4000 × g", "10 min"]);
  assert.equal(
    d.calls[0][1].comment_text,
    "[R1.4] 무엇 / 왜 / 어디\n사라지는 사실: 4000 × g; 10 min",
  );
  assert.ok(TOOL_COMMENT_PREFIX.test(d.calls[0][1].comment_text));
  assert.deepEqual(d.edits[0][1].lost_facts, ["4000 × g", "10 min"]);
  assert.equal(d.lines[1], "[review] lost facts R1.4-1: 4000 × g; 10 min");
});

test("L2 new words found in neither the snapshot nor the comment are reported", async () => {
  const d = withContext(deps(), {
    requirements: [],
    comment_text: "Please state the antifoam agent.",
    paragraphs: ["Cultures were grown in sealed flasks with shaking."],
  });
  const r = await applyTrackedEdit(
    d,
    {
      comment_ref: "R1.5",
      find: "sealed flasks",
      replace: "sealed flasks with 0.1 % antifoam and Tween",
      rationale: R,
    },
    { ref: "R1.5" },
  );
  assert.deepEqual(r.unsourced_terms, ["0.1 %", "Tween"]);
  assert.match(r.warning, /0\.1 %, Tween/);
  assert.deepEqual(d.edits[0][1].unsourced_terms, ["0.1 %", "Tween"]);
  assert.ok(d.lines.includes("[review] unsourced R1.5-1: 0.1 %; Tween"));
});

test("L3 narrowed range: recorded find, text and kind follow what the task pane applied", async () => {
  let d = deps({ narrowed: true, before: "the red big dog" });
  await applyTrackedEdit(
    d,
    { comment_ref: "R1.2", find: "the red big dog", replace: "the big dog", rationale: R },
    { ref: "R1.2" },
  );
  assert.deepEqual(d.calls[0][1].narrow, { find: "red ", replace: "" });
  let ch = d.edits[0][1];
  assert.deepEqual(
    [ch.kind, ch.find, ch.text, ch.narrowed_from],
    ["delete", "red ", "", { find: "the red big dog", text: "the big dog" }],
  );
  assert.deepEqual(ch.context, { left: "the ", right: "big dog" });
  d = deps(); // the task pane could not narrow: full range recorded
  await applyTrackedEdit(
    d,
    { comment_ref: "R1.2", find: "the red big dog", replace: "the big dog", rationale: R },
    { ref: "R1.2" },
  );
  ch = d.edits[0][1];
  assert.deepEqual([ch.kind, ch.find, ch.narrowed_from], ["replace", "the red big dog", null]);
  d = deps();
  await applyTrackedEdit(
    d,
    { comment_ref: "R1.2", find: "a b a b", replace: "a c a b", rationale: R },
    { ref: "R1.2" },
  );
  assert.equal("narrow" in d.calls[0][1], false);
});

test("L5 requirement ids: unknown ids are refused, known ids go into the comment", async () => {
  const ctx = { requirements: [{ id: "q1", text: "state how growth was followed" }] };
  let d = withContext(deps(), ctx);
  const bad = await applyTrackedEdit(
    d,
    {
      comment_ref: "R1.3",
      find: "four days",
      insert_after: " (OD600)",
      rationale: { ...R, requirement_ids: ["q1", "q7"] },
    },
    { ref: "R1.3" },
  );
  assert.match(bad.message, /Unknown requirement id\(s\) q7 for R1\.3\. Valid ids: q1/);
  assert.equal(d.calls.length, 0);
  d = withContext(deps(), ctx);
  await applyTrackedEdit(
    d,
    {
      comment_ref: "R1.3",
      find: "four days",
      insert_after: " (OD600)",
      rationale: { ...R, requirement_ids: ["q1"] },
    },
    { ref: "R1.3" },
  );
  assert.equal(d.calls[0][1].comment_text, "[R1.3] 무엇 / 왜 / 어디 [q1]");
  assert.deepEqual(d.edits[0][1].rationale.requirement_ids, ["q1"]);
});

test("office_tracked_edit is registered with the Word tools", async () => {
  assert.ok((await registeredWordTools()).includes(REVIEW_WRITE_TOOL));
});
