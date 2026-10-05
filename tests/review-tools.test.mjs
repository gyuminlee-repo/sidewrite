import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { z } from "zod";
import {
  REVIEW_READ_TOOLS,
  REVIEW_BLOCKED_TOOLS,
  REVIEW_WRITE_TOOL,
  REVIEW_WRITE_TOOLS,
  REVIEW_NOTE_TOOLS,
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
    (n) =>
      !REVIEW_READ_TOOLS.has(n) &&
      !REVIEW_BLOCKED_TOOLS.has(n) &&
      !REVIEW_NOTE_TOOLS.has(n) &&
      !REVIEW_WRITE_TOOLS.has(n),
  );
  assert.deepEqual(unclassified, []);
  for (const n of [...REVIEW_READ_TOOLS, ...REVIEW_BLOCKED_TOOLS, ...REVIEW_NOTE_TOOLS])
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
  assert.equal(s.safeParse({ ...good, comment_ref: "E.1" }).success, true);
  assert.equal(s.safeParse({ ...good, comment_ref: "R1.0" }).success, true);
  assert.equal(s.safeParse({ ...good, comment_ref: "E1" }).success, false);
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
  assert.ok(TOOL_COMMENT_PREFIX.test("[E.1] editor"));
  assert.ok(!TOOL_COMMENT_PREFIX.test("[E] note"));
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

// A new quantity is no longer among them: a change that adds a number is
// refused before it gets here (tests/number-guard.test.mjs).
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
      replace: "sealed flasks with Pluronic antifoam and Tween",
      rationale: R,
    },
    { ref: "R1.5" },
  );
  assert.deepEqual(r.unsourced_terms, ["Pluronic", "Tween"]);
  assert.match(r.warning, /Pluronic, Tween/);
  assert.deepEqual(d.edits[0][1].unsourced_terms, ["Pluronic", "Tween"]);
  assert.ok(d.lines.includes("[review] unsourced R1.5-1: Pluronic; Tween"));
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

test("applied edit without a comment anchor: recorded with no comment id, model told not to retry", async () => {
  const d = deps({
    ok: false,
    code: "CommentAnchorMissing",
    applied: true,
    word_comment_id: undefined,
  });
  const r = await applyTrackedEdit(
    d,
    { comment_ref: "R1.1", find: "30 degrees", replace: "30 °C", rationale: R },
    { ref: "R1.1" },
  );
  assert.equal(d.edits.length, 1, "the applied change must be recorded");
  const change = d.edits[0][1];
  assert.deepEqual(
    [change.word_comment_id, change.comment_missing, change.before, change.after],
    [null, "CommentAnchorMissing", "held at 30 degrees", "held at 30 °C"],
  );
  assert.deepEqual([r.ok, r.applied, r.change_id, r.word_comment_id], [true, true, "R1.1-1", null]);
  assert.match(r.warning, /applied.*no rationale comment.*do not repeat/i);
  assert.ok(d.lines.some((l) => /comment missing R1\.1-1 code=CommentAnchorMissing/.test(l)));
});

test("comment insert failure after the edit applied: recorded like a missing anchor", async () => {
  const d = deps({
    ok: false,
    code: "CommentInsertFailed",
    applied: true,
    word_comment_id: undefined,
    message: "Edit applied but the rationale comment could not be added: GeneralException",
  });
  const r = await applyTrackedEdit(
    d,
    { comment_ref: "R1.1", find: "30 degrees", replace: "30 °C", rationale: R },
    { ref: "R1.1" },
  );
  assert.equal(d.edits.length, 1, "the applied change must be recorded");
  const change = d.edits[0][1];
  assert.deepEqual(
    [change.word_comment_id, change.comment_missing, change.paragraph_index, change.after],
    [null, "CommentInsertFailed", 9, "held at 30 °C"],
  );
  assert.deepEqual([r.ok, r.applied, r.comment_missing], [true, true, "CommentInsertFailed"]);
  assert.match(r.warning, /do not repeat/i);
});

test("no answer from the task pane (timeout or disconnect): outcome unknown, not retried", async () => {
  for (const code of ["ToolTimeout", "TaskpaneDisconnected"]) {
    const d = deps();
    d.call = async () => {
      throw Object.assign(new Error(`lost: ${code}`), { code });
    };
    const r = await applyTrackedEdit(
      d,
      { comment_ref: "R1.1", find: "30 degrees", replace: "30 °C", rationale: R },
      { ref: "R1.1" },
    );
    assert.deepEqual([r.ok, r.code, d.edits.length], [false, "OutcomeUnknown", 0]);
    assert.match(r.message, /may already be in the document.*do not repeat/i);
    assert.ok(d.lines.some((l) => /outcome unknown R1\.1/.test(l)));
  }
});

test("an error the task pane sent back (edit not applied) still propagates", async () => {
  const d = deps();
  d.call = async () => {
    throw new Error("`find` exceeds the 255-character limit of Word search");
  };
  await assert.rejects(
    applyTrackedEdit(
      d,
      { comment_ref: "R1.1", find: "30 degrees", replace: "30 °C", rationale: R },
      { ref: "R1.1" },
    ),
    /255-character/,
  );
  assert.equal(d.edits.length, 0);
});

test("an old task pane without before/after on the anchor failure still records the change", async () => {
  const d = deps();
  d.call = async (name, args) => (
    d.calls.push([name, args]),
    { ok: false, code: "CommentAnchorMissing", applied: true, message: "no range" }
  );
  const r = await applyTrackedEdit(
    d,
    { comment_ref: "R1.1", find: "30 degrees", replace: "30 °C", rationale: R },
    { ref: "R1.1" },
  );
  assert.equal(r.applied, true);
  assert.deepEqual(
    [d.edits[0][1].before, d.edits[0][1].after, d.edits[0][1].paragraph_index],
    ["", "", null],
  );
});

test("a failure that did not apply the edit records nothing", async () => {
  const d = deps({ ok: false, code: "NotFound", message: "not found" });
  const r = await applyTrackedEdit(
    d,
    { comment_ref: "R1.1", find: "30 degrees", replace: "30 °C", rationale: R },
    { ref: "R1.1" },
  );
  assert.deepEqual([r.ok, r.code, d.edits.length], [false, "NotFound", 0]);
});

// L5: an unlinked change is allowed and shows as an open requirement, but a
// change that answers part of a request cites it (S2-30).
test("requirement_ids: cite a request handled in part; empty only when unrelated", () => {
  const desc = trackedEditShape.rationale.shape.requirement_ids.description;
  assert.match(desc, /part of a request[^.]*still put its id/i);
  assert.match(desc, /empty only when the change answers none of them/i);
  assert.doesNotMatch(desc, /not sure/);
});

// Codex review 12: forbid "novel substrate"; "This is a novel material.";
// replacing "material" with "substrate" was applied with ok:true.
const PHRASE = {
  id: "L4",
  kind: "forbid",
  text: "novel substrate 금지",
  check: { type: "regex", pattern: "novel substrate", flags: "i", expect: "absent" },
};
const crossDeps = (ctx) => {
  const d = deps();
  d.getLedger = async () => [PHRASE];
  d.getReviewContext = async () => ({ requirements: [], comment_text: "", ...ctx });
  return d;
};

test("forbid hit across the edit boundary: blocked through the same flow, and an override applies it", async () => {
  const d = crossDeps({ paragraphs: ["This is a novel material.", "It grew slowly."] });
  const args = { comment_ref: "R1.3", find: "material", replace: "substrate", rationale: R };
  const r = await applyTrackedEdit(d, args, { ref: "R1.3" });
  assert.deepEqual([r.ok, r.blocked_by, r.blocked_id], [false, ["L4"], "R1.3-b1"]);
  assert.match(r.message, /Rephrase/);
  assert.equal(d.calls.length, 0);
  assert.equal(d.edits.length, 0);
  assert.deepEqual(d.blocked[0].blocked_by, ["L4"]);
  assert.equal(d.blocked[0].args.replace, "substrate");
  assert.equal(d.lines[0], "[review] blocked R1.3 by=L4");
  // The author approves the exception: the stored edit goes through.
  const ok = await applyTrackedEdit(
    d,
    d.blocked[0].args,
    { ref: "R1.3" },
    {
      override: { blocked_by: ["L4"] },
    },
  );
  assert.equal(ok.ok, true);
  assert.deepEqual(d.edits[0][1].override_of, ["L4"]);
  // Insertion and deletion forms.
  const ins = await applyTrackedEdit(
    crossDeps({ paragraphs: ["This is a novel."] }),
    { comment_ref: "R1.3", find: "novel", insert_after: " substrate", rationale: R },
    { ref: "R1.3" },
  );
  assert.deepEqual(ins.blocked_by, ["L4"]);
  const del = await applyTrackedEdit(
    crossDeps({ paragraphs: ["It used a novel bright substrate."] }),
    { comment_ref: "R1.3", find: "bright ", replace: "", rationale: R },
    { ref: "R1.3" },
  );
  assert.deepEqual(del.blocked_by, ["L4"]);
});

test("an edit next to or away from an existing forbidden phrase, or one that removes it, is applied", async () => {
  const paragraphs = ["A novel substrate grows. The medium was warm material."];
  for (const [find, replace] of [
    ["warm", "hot"], // elsewhere in the paragraph
    ["novel substrate", "new substrate"], // removes it
    [" grows", " grew"], // right after it
  ]) {
    const d = crossDeps({ paragraphs });
    const r = await applyTrackedEdit(
      d,
      { comment_ref: "R1.3", find, replace, rationale: R },
      { ref: "R1.3" },
    );
    assert.equal(r.ok, true, `${find} -> ${replace}`);
    assert.equal(d.blocked.length, 0);
  }
});

// The paragraph texts are those of the last refresh. A change recorded since
// then is put into them first, so two edits in a row cannot form the phrase
// (new -> novel, then material -> substrate).
test("forbid hit across the boundary of an earlier change of the same turn", async () => {
  const recent = (find, text, extra = {}) => ({
    kind: "replace",
    find,
    text,
    narrowed_from: null,
    created_at: "2026-01-01T00:00:01.000Z",
    ...extra,
  });
  const d = crossDeps({
    paragraphs: ["This is a new material."],
    recent_changes: [recent("new", "novel")],
  });
  const r = await applyTrackedEdit(
    d,
    { comment_ref: "R1.3", find: "material", replace: "substrate", rationale: R },
    { ref: "R1.3" },
  );
  assert.deepEqual([r.ok, r.blocked_by], [false, ["L4"]]);
  // A narrowed change is put in by the find/replace the model asked for.
  const { currentTexts } = await import("../daemon/review-tools.mjs");
  assert.deepEqual(
    currentTexts(
      ["Held at 30 degrees all day.", "Other."],
      [
        recent("degrees", "°C", { narrowed_from: { find: "30 degrees", text: "30 °C" } }),
        recent("all day", " long", { kind: "insert" }),
        recent("Other.", "", { kind: "delete" }),
        recent("not in the text", "x"),
      ],
    ),
    ["Held at 30 °C all day long.", ""],
  );
});

// Codex review 13: the forbid check tried every occurrence of `find`, so an
// edit in one paragraph was blocked by a phrase it would only form in another
// ("a novel material" / "a conventional material", the hint picks the second).
const TWO_PARAS = ["This is a novel material.", "This is a conventional material."];

test("forbid check reads the paragraph the hint picks, as Word does", async () => {
  const edit = (hint) => ({
    comment_ref: "R1.3",
    find: "material",
    replace: "substrate",
    ...(hint ? { paragraph_hint: hint } : {}),
    rationale: R,
  });
  for (const [ids, clean, dirty] of [
    [["U0", "U1"], "U1", "U0"], // uniqueLocalId ids
    [null, "p1", "p0"], // index ids
  ]) {
    // The hinted paragraph forms no forbidden phrase: applied, with the hint.
    let d = crossDeps({ paragraphs: TWO_PARAS, paragraph_ids: ids });
    let r = await applyTrackedEdit(d, edit(clean), { ref: "R1.3" });
    assert.equal(r.ok, true, `hint ${clean}`);
    assert.equal(d.blocked.length, 0);
    assert.equal(d.calls[0][1].paragraph_hint, clean);
    // The hinted paragraph does: blocked, and the stored edit keeps the hint.
    d = crossDeps({ paragraphs: TWO_PARAS, paragraph_ids: ids });
    r = await applyTrackedEdit(d, edit(dirty), { ref: "R1.3" });
    assert.deepEqual([r.ok, r.blocked_by], [false, ["L4"]], `hint ${dirty}`);
    assert.equal(d.blocked[0].args.paragraph_hint, dirty);
  }
  // The target cannot be told (no hint with two hits, a hint that names no
  // paragraph): every occurrence is tried, so the phrase in the first blocks.
  for (const hint of [null, "U9"]) {
    const d = crossDeps({ paragraphs: TWO_PARAS, paragraph_ids: ["U0", "U1"] });
    const r = await applyTrackedEdit(d, edit(hint), { ref: "R1.3" });
    assert.deepEqual([r.ok, r.blocked_by], [false, ["L4"]], `hint ${hint}`);
  }
});

// Class audit (Codex review 13): what else the daemon reads from "the"
// paragraph of the edit before Word answers.
test("lost facts read the sentence position in the hinted paragraph", async () => {
  // "Methods" starts a sentence in the first paragraph (not a name there)
  // and is inside one in the second (a name candidate, so a lost fact).
  const paragraphs = [
    "Methods were chosen first. Details follow.",
    "Here Methods were chosen again.",
  ];
  const lost = async (hint) => {
    const d = crossDeps({ paragraphs, paragraph_ids: ["U0", "U1"] });
    d.getLedger = async () => [];
    await applyTrackedEdit(
      d,
      {
        comment_ref: "R1.3",
        find: "Methods were chosen",
        replace: "they were chosen",
        paragraph_hint: hint,
        rationale: R,
      },
      { ref: "R1.3" },
    );
    return d.edits[0][1].lost_facts;
  };
  assert.deepEqual(await lost("U0"), []);
  assert.deepEqual(await lost("U1"), ["Methods"]);
});

test("an unanswered edit with a hint records the form of the hinted paragraph", async () => {
  const run = async (ctx, args) => {
    const d = crossDeps(ctx);
    let attempt = null;
    d.onOutcomeUnknown = async (x) => (attempt = x.attempt);
    d.call = async () => {
      throw Object.assign(new Error("lost"), { code: "ToolTimeout" });
    };
    await applyTrackedEdit(d, { comment_ref: "R1.3", rationale: R, ...args }, { ref: "R1.3" });
    return attempt.change;
  };
  const ctx = { paragraphs: TWO_PARAS, paragraph_ids: ["U0", "U1"] };
  const hinted = await run(ctx, { find: "material", replace: "medium", paragraph_hint: "U1" });
  assert.deepEqual(
    [hinted.paragraph_index, hinted.before, hinted.after],
    [1, "This is a conventional material.", "This is a conventional medium."],
  );
  assert.deepEqual(hinted.context, { left: "is a conventional ", right: "." });
  // No hint and two hits: Word refuses such an edit; nothing is guessed.
  const none = await run(ctx, { find: "material", replace: "medium" });
  assert.deepEqual(
    [none.paragraph_index, none.before, none.after, none.context],
    [null, "", "", null],
  );
  // One hit in the document: that paragraph, with or without a hint.
  const one = await run(ctx, { find: "conventional", replace: "common", paragraph_hint: "U0" });
  assert.deepEqual([one.paragraph_index, one.after], [1, "This is a common material."]);
});

test("currentTexts puts a recorded change into the paragraph Word reported", async () => {
  const { currentTexts } = await import("../daemon/review-tools.mjs");
  const ch = (extra) => ({
    kind: "replace",
    find: "material",
    text: "substrate",
    narrowed_from: null,
    ...extra,
  });
  // `find` is in both paragraphs; the change was made in the second.
  assert.deepEqual(currentTexts(TWO_PARAS, [ch({ paragraph_index: 1 })]), [
    TWO_PARAS[0],
    "This is a conventional substrate.",
  ]);
  // No usable index: only a find that is in one paragraph is put in.
  assert.deepEqual(currentTexts(TWO_PARAS, [ch({ paragraph_index: null })]), TWO_PARAS);
  assert.deepEqual(currentTexts(TWO_PARAS, [ch({ paragraph_index: 7 })]), TWO_PARAS);
  assert.deepEqual(
    currentTexts(TWO_PARAS, [ch({ find: "novel", text: "new", paragraph_index: null })]),
    ["This is a new material.", TWO_PARAS[1]],
  );
});

test("a change records whether its paragraph had an identical twin when it was made", async () => {
  const twin = "Cells were held at 30 degrees.";
  const run = async (paragraphs, hint) => {
    const d = crossDeps({ paragraphs, paragraph_ids: paragraphs.map((_, k) => `U${k}`) });
    d.getLedger = async () => [];
    d.call = async () => ({
      ok: true,
      word_comment_id: 5,
      paragraph_index: 2,
      before: twin,
      after: "Cells were held at 30 °C.",
    });
    await applyTrackedEdit(
      d,
      {
        comment_ref: "R1.3",
        find: "30 degrees",
        replace: "30 °C",
        ...(hint ? { paragraph_hint: hint } : {}),
        rationale: R,
      },
      { ref: "R1.3" },
    );
    return d.edits[0][1].duplicate_before;
  };
  assert.equal(await run([twin, "Middle.", twin], "U2"), true);
  assert.equal(await run(["Other.", "Middle.", twin], null), false);
});

// What Word answered when the rationale comment could not be added goes to
// the daemon log (step, Office error code, where, message): the change is
// recorded either way, and without this nothing says why the comment is
// missing. No text of the manuscript is in it.
test("comment errors of the task pane are logged: a failed comment, and one added on the second try", async () => {
  const errors = [
    {
      step: "insert_comment",
      code: "GeneralException",
      location: "Range.insertComment",
      message: "GeneralException",
    },
    {
      step: "insert_comment",
      code: "GeneralException",
      location: null,
      message: "GeneralException",
    },
  ];
  let d = deps({
    ok: false,
    code: "CommentInsertFailed",
    applied: true,
    message: "Edit applied but the rationale comment could not be added: GeneralException",
    comment_errors: errors,
  });
  let r = await applyTrackedEdit(
    d,
    { comment_ref: "R1.1", find: "30 degrees", replace: "30 °C", rationale: R },
    { ref: "R1.1" },
  );
  assert.deepEqual([r.ok, r.comment_missing], [true, "CommentInsertFailed"]);
  assert.ok(d.lines.includes("[review] comment missing R1.1-1 code=CommentInsertFailed"));
  assert.deepEqual(
    d.lines.filter((l) => l.startsWith("[review] comment error")),
    [
      "[review] comment error R1.1-1 try=1 step=insert_comment word=GeneralException at=Range.insertComment msg=GeneralException",
      "[review] comment error R1.1-1 try=2 step=insert_comment word=GeneralException at=- msg=GeneralException",
    ],
  );
  assert.equal("comment_errors" in d.edits[0][1], false, "not stored on the change");
  assert.equal("comment_errors" in r, false, "not sent to the model");

  d = deps({ comment_attempts: 2, comment_errors: errors.slice(0, 1) });
  r = await applyTrackedEdit(
    d,
    { comment_ref: "R1.1", find: "30 degrees", replace: "30 °C", rationale: R },
    { ref: "R1.1" },
  );
  assert.deepEqual([r.ok, r.comment_missing, r.word_comment_id], [true, undefined, "123"]);
  assert.ok(d.lines.includes("[review] comment added on try 2 R1.1-1"));
  assert.equal(d.lines.filter((l) => l.startsWith("[review] comment error")).length, 1);
});

// Acceptance run of 2026-10-05, stage 3 (Word minimized): changes were
// recorded without their comment and the review file did not say why. The
// last error of the task pane is kept on the change (comment_error), and the
// log line carries what debugInfo added and whether the task pane was hidden.
test("a change applied without its comment keeps Word's last answer; the log says more", async () => {
  const errors = [
    {
      step: "insert_comment",
      code: "GeneralException",
      location: null,
      message: "GeneralException",
    },
    {
      step: "insert_comment",
      code: "GeneralException",
      location: "Range.insertComment",
      message: "GeneralException",
      debug_message: "The window is\nnot available.",
      inner: "AccessDenied",
      hidden: true,
    },
  ];
  let d = deps({
    ok: false,
    code: "CommentInsertFailed",
    applied: true,
    message: "Edit applied but the rationale comment could not be added: GeneralException",
    comment_attempts: 3,
    comment_errors: errors,
  });
  await applyTrackedEdit(
    d,
    { comment_ref: "R1.1", find: "30 degrees", replace: "30 °C", rationale: R },
    { ref: "R1.1" },
  );
  const last =
    "step=insert_comment word=GeneralException at=Range.insertComment msg=GeneralException" +
    " debug=The window is not available. inner=AccessDenied hidden=1";
  assert.equal(d.edits[0][1].comment_error, last);
  assert.ok(d.lines.includes(`[review] comment error R1.1-1 try=2 ${last}`));
  // A comment added on a later try: nothing is missing, nothing kept.
  d = deps({ comment_attempts: 2, comment_errors: errors.slice(0, 1) });
  await applyTrackedEdit(
    d,
    { comment_ref: "R1.1", find: "30 degrees", replace: "30 °C", rationale: R },
    { ref: "R1.1" },
  );
  assert.equal(d.edits[0][1].comment_error, null);
});

test("the comment ids on record go out with a tracked edit, so the task pane reuses none of them", async () => {
  const d = deps({ target_checked: true });
  d.getReviewContext = async () => ({ comment_ids: ["11", "12"] });
  await applyTrackedEdit(
    d,
    { comment_ref: "R1.1", find: "30 degrees", replace: "30 °C", rationale: R },
    { ref: "R1.1" },
  );
  assert.deepEqual(d.calls[0][1].known_comment_ids, ["11", "12"]);
  // None on record: the argument is left out.
  const e = deps({ target_checked: true });
  await applyTrackedEdit(
    e,
    { comment_ref: "R1.1", find: "30 degrees", replace: "30 °C", rationale: R },
    { ref: "R1.1" },
  );
  assert.equal("known_comment_ids" in e.calls[0][1], false);
});
