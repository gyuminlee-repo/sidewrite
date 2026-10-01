import { test } from "node:test";
import assert from "node:assert/strict";
import * as view from "../taskpane/shared/review-view.js";
import { normalizeParaText as daemonNormalize } from "../daemon/revision-classify.mjs";
import { TOOL_COMMENT_PREFIX, formatRationaleComment } from "../daemon/review-tools.mjs";

test("task pane and daemon normalize paragraph text identically", () => {
  const corpus = [
    "plain",
    "  lead and trail  ",
    "a\u0005b\u0005\u0005c",
    "tab\tand\nnewline",
    "30 °C\u0005\u0005 and the",
    "",
    null,
    " nbsp ",
  ];
  for (const s of corpus)
    assert.equal(view.normalizeParaText(s), daemonNormalize(s), JSON.stringify(s));
});

test("task pane and daemon agree on what a tool comment is", () => {
  const samples = [
    formatRationaleComment("R2.10", { what: "a", why: "b", addresses: "c" }),
    "[R1.1]x",
    " [R3.4] y",
    "[R1] no",
    "note [R1.1]",
    "[r1.1] lower",
    "[E.1] editor",
    "[R1.0] preamble",
    "[E] no",
  ];
  for (const s of samples) assert.equal(view.isToolComment(s), TOOL_COMMENT_PREFIX.test(s), s);
});

const BEFORE =
  "Cultures were grown in a defined mineral medium in sealed flasks with shaking. For every run the incubation temperature was held at 30 degrees and the headspace was refreshed once per day.";

test("computeAfterText: replace, delete, insert", () => {
  assert.equal(
    view.computeAfterText(BEFORE, "held at 30 degrees", { replace: "maintained at 30 °C" }).after,
    BEFORE.replace("held at 30 degrees", "maintained at 30 °C"),
  );
  assert.equal(
    view.computeAfterText(BEFORE, " with shaking", { replace: "" }).after,
    BEFORE.replace(" with shaking", ""),
  );
  assert.equal(
    view.computeAfterText("Growth reached a plateau after four days.", "after four days", {
      insert_after: " (optical density at 600 nm)",
    }).after,
    "Growth reached a plateau after four days (optical density at 600 nm).",
  );
});

test("computeAfterText: deleting a middle word leaves one space, like Word after accept", () => {
  assert.equal(
    view.computeAfterText("run the incubation temperature was", "incubation", { replace: "" })
      .after,
    "run the temperature was",
  );
});

test("computeAfterText: ignores comment anchors in the paragraph text", () => {
  assert.equal(
    view.computeAfterText("A \u0005crude extract\u0005 was", "crude extract", { replace: "lysate" })
      .after,
    "A lysate was",
  );
});

test("computeAfterText: refuses missing, repeated, or empty find", () => {
  assert.match(
    view.computeAfterText(BEFORE, "absent", { replace: "x" }).error,
    /not in the paragraph/,
  );
  assert.match(view.computeAfterText("the the", "the", { replace: "x" }).error, /more than once/);
  assert.match(view.computeAfterText(BEFORE, "  ", { replace: "x" }).error, /empty/);
  assert.match(view.computeAfterText(BEFORE, "held", {}).error, /replace or insert_after/);
});

test("countRevisionMarks counts the document part of a package only", () => {
  const part = (name, body) =>
    `<pkg:part pkg:name="${name}"><pkg:xmlData>${body}</pkg:xmlData></pkg:part>`;
  const pkg = (...parts) => `<pkg:package>${parts.join("")}</pkg:package>`;
  const ins = '<w:ins w:id="1"><w:r><w:t>x</w:t></w:r></w:ins>';
  const del = '<w:del w:id="2"><w:r><w:delText>y</w:delText></w:r></w:del>';
  // An accepted insertion whose rationale comment body is still tracked.
  assert.deepEqual(
    view.countRevisionMarks(
      pkg(part("/word/document.xml", "<w:p/>"), part("/word/comments.xml", ins)),
    ),
    { w_ins: 0, w_del: 0 },
  );
  assert.deepEqual(
    view.countRevisionMarks(
      pkg(
        part("/word/comments.xml", ins + del),
        part("/word/document.xml", ins),
        part("/word/styles.xml", del),
      ),
    ),
    { w_ins: 1, w_del: 0 },
  );
  assert.deepEqual(view.countRevisionMarks(pkg(part("/word/document.xml", ins + del))), {
    w_ins: 1,
    w_del: 1,
  });
});

test("countRevisionMarks counts w:ins/w:del and not w:delText or w:instrText", () => {
  const xml =
    '<w:p><w:ins w:id="1"><w:r><w:t>new</w:t></w:r></w:ins><w:del w:id="2"><w:r><w:delText>old</w:delText></w:r></w:del>' +
    '<w:r><w:instrText>PAGE</w:instrText></w:r><w:ins w:id="3"/><w:rPr><w:del w:id="4"/></w:rPr></w:p>';
  assert.deepEqual(view.countRevisionMarks(xml), { w_ins: 2, w_del: 2 });
  assert.deepEqual(view.countRevisionMarks(null), { w_ins: 0, w_del: 0 });
});

const COMMENTS = [
  { id: "1", content: "[R1.1] a / b / c" },
  { id: "2", content: "[R1.2] a / b / c" },
  { id: "3", content: "Check units later" },
];

test("countComments and cleanupReport", () => {
  const before = view.countComments(COMMENTS);
  assert.deepEqual(before, { tool: 2, user: 1, total: 3 });
  assert.deepEqual(view.cleanupReport(before, { tool: 0, user: 1, total: 1 }), {
    ok: true,
    text: "Tool comments 2 -> 0, your comments 1 -> 1 (unchanged)",
  });
  assert.equal(view.cleanupReport(before, { tool: 0, user: 0, total: 0 }).ok, false);
  assert.equal(view.cleanupReport(before, { tool: 1, user: 1, total: 2 }).ok, false);
});

test("buildCommentsBackup keeps every comment and its count", () => {
  const b = view.buildCommentsBackup(
    COMMENTS.map((c, i) => ({
      ...c,
      author: "A",
      resolved: i === 0,
      anchor_text: "t",
      paragraph_index: i,
    })),
    { manuscript: "/m/x.docx", now: new Date("2026-01-01T00:00:00Z") },
  );
  assert.equal(b.count, 3);
  assert.equal(b.comments.length, 3);
  assert.deepEqual(b.comments[0], {
    id: "1",
    author: "A",
    content: "[R1.1] a / b / c",
    resolved: true,
    anchor_text: "t",
    paragraph_index: 0,
    replies: [],
  });
  assert.equal(b.reply_count, 0);
});

test("buildCommentsBackup keeps the replies of every thread", () => {
  const b = view.buildCommentsBackup(
    [
      {
        id: "1",
        author: "A",
        content: "Check units later",
        replies: [
          { id: "1a", author: "B", content: "Done in 2.1", created_at: "2026-01-02T00:00:00Z" },
          { id: "1b", author: "A", content: "Thanks" },
        ],
      },
      { id: "2", author: "A", content: "[R1.1] a / b / c" },
    ],
    { now: new Date("2026-01-03T00:00:00Z") },
  );
  // count stays the number of parent comments; replies are counted apart.
  assert.deepEqual([b.count, b.reply_count], [2, 2]);
  assert.deepEqual(b.comments[0].replies, [
    { id: "1a", author: "B", content: "Done in 2.1", created_at: "2026-01-02T00:00:00Z" },
    { id: "1b", author: "A", content: "Thanks", created_at: null },
  ]);
  assert.deepEqual(b.comments[1].replies, []);
});

test("compareCommentLists: the second Delete ALL press deletes only what was shown", () => {
  const shown = [
    { id: "1", author: "A", content: "Check units", resolved: false, anchor_text: "t" },
    { id: "2", author: "A", content: "[R1.1] a / b / c", resolved: false, anchor_text: "u" },
  ];
  const same = view.compareCommentLists(shown, structuredClone(shown));
  assert.deepEqual([same.same, same.text], [true, ""]);
  const added = view.compareCommentLists(shown, [...shown, { id: "3", content: "new" }]);
  assert.deepEqual([added.same, added.added, added.removed, added.edited], [false, 1, 0, 0]);
  const edited = view.compareCommentLists(shown, [
    { ...shown[0], content: "Check units twice" },
    shown[1],
  ]);
  assert.deepEqual([edited.same, edited.edited], [false, 1]);
  const resolved = view.compareCommentLists(shown, [{ ...shown[0], resolved: true }, shown[1]]);
  assert.equal(resolved.same, false);
  const replied = view.compareCommentLists(shown, [
    { ...shown[0], replies: [{ author: "B", content: "ok" }] },
    shown[1],
  ]);
  assert.deepEqual([replied.same, replied.edited], [false, 1]);
  const removed = view.compareCommentLists(shown, [shown[1]]);
  assert.deepEqual([removed.same, removed.removed], [false, 1]);
  assert.match(added.text, /1 added/);
});

test("deleteAllReport: every backed-up thread deleted, later comments kept and named", () => {
  const r = (x) => view.deleteAllReport({ shown: 3, backedUp: 3, deleted: 3, ...x });
  assert.equal(r({}).ok, true);
  assert.equal(r({ backedUp: 2 }).ok, false);
  assert.equal(r({ deleted: 4 }).ok, false);
  assert.equal(r({ deleted: 2 }).ok, false);
  assert.equal(r({ stillThere: 1 }).ok, false);
  assert.equal(r({ deleted: 2, missing: 1 }).ok, true);
  const later = r({ addedAfter: 2 });
  assert.equal(later.ok, true);
  assert.match(later.text, /2 comment\(s\) added after the backup were not deleted/);
  const changed = r({ deleted: 2, changed: 1 });
  assert.equal(changed.ok, false);
  assert.match(changed.text, /1 changed after the backup and not deleted \(press again\)/);
});

test("planBackupDelete: only unchanged backed-up threads are deleted", () => {
  const c = (id, content, extra = {}) => ({
    id,
    author: "A",
    content,
    resolved: false,
    anchor_text: "t",
    ...extra,
  });
  const backup = view.buildCommentsBackup([
    c("1", "a"),
    c("2", "b"),
    c("3", "c"),
    c("4", "d"),
  ]).comments;
  const now = [
    c("1", "a"),
    c("2", "b edited"),
    c("3", "c", { replies: [{ author: "B", content: "ok" }] }),
    c("5", "new"),
  ];
  assert.deepEqual(view.planBackupDelete(backup, now), {
    delete_ids: ["1"],
    changed_ids: ["2", "3"],
    new_ids: ["5"],
    missing_ids: ["4"],
  });
  // A reply the backup holds is not a change.
  const withReply = [c("1", "a", { replies: [{ author: "B", content: "ok", id: "9" }] })];
  const b2 = view.buildCommentsBackup(withReply).comments;
  assert.deepEqual(view.planBackupDelete(b2, withReply).delete_ids, ["1"]);
});

test("invariantReport: one comment per change, rejected insertions excepted", () => {
  const rev = {
    comments: [
      {
        ref: "R1.1",
        changes: [{ id: "R1.1-1", kind: "replace", state: "pending", word_comment_id: "1" }],
      },
      {
        ref: "R1.2",
        changes: [{ id: "R1.2-1", kind: "replace", state: "rejected", word_comment_id: "2" }],
      },
      {
        ref: "R1.3",
        changes: [{ id: "R1.3-1", kind: "insert", state: "rejected", word_comment_id: "9" }],
      },
    ],
  };
  const r = view.invariantReport(rev, COMMENTS);
  assert.deepEqual([r.ok, r.changes, r.expected_comments, r.tool_comments], [true, 3, 2, 2]);
  const bad = view.invariantReport(rev, [...COMMENTS, { id: "7", content: "[R1.1] stray" }]);
  assert.deepEqual([bad.ok, bad.unrecorded_comment_ids], [false, ["7"]]);
  const missing = view.invariantReport(rev, COMMENTS.slice(1));
  assert.deepEqual([missing.ok, missing.missing_comment], [false, ["R1.1-1"]]);
  assert.deepEqual(missing.uncommented, []);
  // A change the task pane applied without its rationale comment
  // (comment_missing, no comment id): not a comment that went missing, and
  // not ok either. Every change needs its rationale comment.
  const bare = {
    comments: [
      {
        ref: "R1.1",
        changes: [
          { id: "R1.1-1", kind: "replace", state: "pending", word_comment_id: "3" },
          {
            id: "R1.1-2",
            kind: "replace",
            state: "pending",
            word_comment_id: null,
            comment_missing: "CommentInsertFailed",
          },
        ],
      },
    ],
  };
  const u = view.invariantReport(bare, [{ id: "3", content: "[R1.1] a / b / c" }]);
  assert.deepEqual(
    [u.ok, u.missing_comment, u.uncommented, u.unrecorded_comment_ids],
    [false, [], ["R1.1-2"], []],
  );
  assert.equal(
    view.invariantText(u),
    " (MISMATCH: no rationale comment on R1.1-2: applied in Word, the comment could not be added)",
  );
  assert.equal(view.invariantText(missing), " (MISMATCH: missing R1.1-1, unrecorded -)");
  assert.equal(view.invariantText(r), ".");
  const both = view.invariantReport(bare, []);
  assert.equal(
    view.invariantText(both),
    " (MISMATCH: missing R1.1-1, unrecorded -; no rationale comment on R1.1-2: applied in Word, the comment could not be added)",
  );
});

test("ledgerItemFromForm builds regex and llm items and rejects bad input", () => {
  assert.deepEqual(
    view.ledgerItemFromForm({
      kind: "forbid",
      text: " novel 금지 ",
      scope: "",
      checkType: "regex",
      pattern: "\\bnovel\\b",
      flags: "gi",
      expect: "",
    }).item,
    {
      kind: "forbid",
      text: "novel 금지",
      scope: "manuscript",
      check: { type: "regex", pattern: "\\bnovel\\b", flags: "i", expect: "absent" },
    },
  );
  assert.deepEqual(
    view.ledgerItemFromForm({ kind: "instruction", text: "x", checkType: "llm" }).item.check,
    { type: "llm" },
  );
  assert.match(
    view.ledgerItemFromForm({ kind: "forbid", text: "", checkType: "llm" }).error,
    /instruction text/,
  );
  assert.match(
    view.ledgerItemFromForm({ kind: "forbid", text: "x", checkType: "regex", pattern: "(" }).error,
    /Invalid regex/,
  );
});

test("commentsToResolve: accepted and rejected changes with an open comment", () => {
  const rev = {
    comments: [
      {
        ref: "R1.1",
        changes: [
          { id: "R1.1-1", word_comment_id: "1" },
          { id: "R1.1-2", word_comment_id: "2" },
          { id: "R1.1-3", word_comment_id: "3" },
          { id: "R1.1-4", word_comment_id: "4" },
        ],
      },
    ],
  };
  const rows = [
    { change_id: "R1.1-1", state: "accepted" },
    { change_id: "R1.1-2", state: "rejected" },
    { change_id: "R1.1-3", state: "pending" },
    { change_id: "R1.1-4", state: "accepted" },
  ];
  const comments = [
    { id: "1", resolved: false },
    { id: "2", resolved: false },
    { id: "3", resolved: false },
    { id: "4", resolved: true },
  ];
  assert.deepEqual(view.commentsToResolve(rows, rev, comments), ["1", "2"]);
});

// ---- L1-L3 badges, L5, L6, L7, L9 -------------------------------------------

const ch = (over = {}) => ({
  id: "R1.3-1",
  kind: "replace",
  find: "synthetic carbon source",
  text: "novel substrate",
  state: "pending",
  rationale: { what: "w", why: "y", addresses: "a", requirement_ids: [] },
  lost_facts: [],
  unsourced_terms: [],
  narrowed_from: null,
  ...over,
});

test("changeFlags shows lost facts, unsourced terms and a narrowed range", () => {
  assert.deepEqual(view.changeFlags(ch()), []);
  assert.deepEqual(
    view
      .changeFlags(
        ch({
          lost_facts: ["20 mL", "Merck"],
          unsourced_terms: ["antifoam"],
          narrowed_from: { find: "a b", text: "a c" },
        }),
      )
      .map((f) => f.text),
    ["Lost: 20 mL; Merck", "No source: antifoam", "Narrowed"],
  );
});

test("L5 unaddressedRequirements: rejected changes answer nothing", () => {
  const comment = {
    requirements: [
      { id: "q1", text: "say how growth was followed" },
      { id: "q2", text: "call it novel" },
    ],
    changes: [
      ch({ rationale: { requirement_ids: ["q1"] } }),
      ch({ state: "rejected", rationale: { requirement_ids: ["q2"] } }),
    ],
  };
  assert.deepEqual(
    view.unaddressedRequirements(comment).map((r) => r.id),
    ["q2"],
  );
  assert.deepEqual(view.requirementsFromText(" a \n\n b\r\n"), ["a", "b"]);
});

test("L6 ledger drafts from rejected changes are literal regexes that match the text", () => {
  const forbid = view.ledgerDraftFromChange(ch({ text: "novel (n = 3) at 30 °C" }));
  assert.deepEqual(
    [forbid.kind, forbid.checkType, forbid.expect, forbid.text],
    ["forbid", "regex", "absent", 'Do not write "novel (n = 3) at 30 °C" (rejected R1.3-1)'],
  );
  assert.ok(new RegExp(forbid.pattern, forbid.flags).test("a novel (n = 3) at 30 °C here"));
  assert.ok(!new RegExp(forbid.pattern, forbid.flags).test("novel n = 3 at 30 °C"));
  const keep = view.ledgerDraftFromChange(
    ch({ kind: "delete", find: "[12] and (v2.1)", text: "" }),
  );
  assert.deepEqual([keep.kind, keep.expect], ["instruction", "present"]);
  assert.ok(new RegExp(keep.pattern).test("see [12] and (v2.1)."));
  // Every draft goes through the same form parser as a typed item.
  assert.ok(view.ledgerItemFromForm(forbid).item);
  assert.ok(view.ledgerItemFromForm(keep).item);
});

test("L6 ledgerFormFromItem round-trips through ledgerItemFromForm", () => {
  const item = {
    kind: "instruction",
    text: "keep it",
    scope: "manuscript",
    check: { type: "regex", pattern: "No correction", flags: "", expect: "present" },
  };
  assert.deepEqual(view.ledgerItemFromForm(view.ledgerFormFromItem(item)).item, item);
  const llm = { kind: "forbid", text: "no hype", scope: "response", check: { type: "llm" } };
  assert.deepEqual(view.ledgerItemFromForm(view.ledgerFormFromItem(llm)).item, llm);
});

test("L6 ledgerCheckFromText: a quoted phrase becomes a literal regex, else an LLM check", () => {
  assert.deepEqual(view.ledgerCheckFromText('Do not use "state-of-the-art (SOTA)"'), {
    checkType: "regex",
    pattern: "state-of-the-art \\(SOTA\\)",
    flags: "i",
  });
  assert.equal(view.ledgerCheckFromText("「novel」 금지").pattern, "novel");
  assert.equal(view.ledgerCheckFromText("Keep the tone modest").checkType, "llm");
});

test("L7 pendingDecisions gathers blocked edits, unchecked ledger items and open requirements", () => {
  const revision = {
    comments: [
      {
        ref: "R1.3",
        requirements: [{ id: "q1", text: "call it novel" }],
        changes: [],
        blocked: [{ id: "R1.3-b1", args: { replace: "novel substrate" }, blocked_by: ["L2"] }],
      },
      { ref: "R1.1", requirements: [], changes: [ch()], blocked: [] },
    ],
    ledger: [
      { id: "L1", status: "ok" },
      { id: "L2", status: "unchecked" },
      { id: "L3", status: "error" },
      { id: "L4", status: "violations" },
    ],
  };
  const d = view.pendingDecisions(revision);
  assert.equal(d.total, 4);
  assert.deepEqual(
    [...d.blocked, ...d.ledger, ...d.requirements].map((x) => x.text),
    [
      'R1.3 blocked by L2: "novel substrate"',
      "L2 not checked yet",
      "L3 check failed",
      "R1.3 q1 has no change: call it novel",
    ],
  );
  assert.equal(view.pendingDecisions({ comments: [], ledger: [] }).total, 0);
  // A normal edit Word did not answer (no forbid rule): named as such.
  const unanswered = view.pendingDecisions({
    comments: [
      {
        ref: "R1.3",
        requirements: [],
        changes: [],
        blocked: [{ id: "R1.3-b2", args: { replace: "the medium" }, blocked_by: [] }],
      },
    ],
    ledger: [],
  });
  assert.deepEqual(
    unanswered.blocked.map((x) => x.text),
    ['R1.3 edit with no answer from Word: "the medium"'],
  );
  // The same edit while Word still works on it (on file from before the
  // call, state applying): nothing to decide yet. An exception being
  // applied (a rule in blocked_by) stays listed.
  const row = (id, state, blocked_by) => ({ id, args: { replace: "x" }, blocked_by, state });
  const running = [
    row("R1.3-b3", "applying", []),
    row("R1.3-b4", "applying", ["L1"]),
    row("R1.3-b5", "outcome_unknown", []),
  ];
  assert.deepEqual(running.map(view.inFlightEdit), [true, false, false]);
  assert.deepEqual(
    view
      .pendingDecisions({
        comments: [{ ref: "R1.3", requirements: [], changes: [], blocked: running }],
        ledger: [],
      })
      .blocked.map((x) => x.id),
    ["R1.3-b4", "R1.3-b5"],
  );
});

test("ledgerStatusText names the stage that checks a non-manuscript item", () => {
  assert.equal(
    view.ledgerStatusText({ status: "violations", violation_count: 2 }),
    "2 violation(s)",
  );
  assert.equal(view.ledgerStatusText({ status: "ok" }), "ok");
  // A check that failed (a model reply that could not be used, a regex that
  // does not compile) passed nothing.
  assert.equal(
    view.ledgerStatusText({ status: "error" }),
    "check failed: nothing counts as passed",
  );
  assert.equal(
    view.ledgerStatusText({ status: "skipped", scope: "response" }),
    "Checked at the response stage",
  );
  assert.equal(
    view.ledgerStatusText({ status: "skipped", scope: "supplement" }),
    "Checked at the supplement stage",
  );
  // A skipped item waits for nothing in the manuscript review.
  assert.equal(
    view.pendingDecisions({ comments: [], ledger: [{ id: "L1", status: "skipped" }] }).total,
    0,
  );
});

test("L9 progressText: elapsed time and comments not done", () => {
  const comments = [{ status: "done" }, { status: "in_progress" }, { status: "pending" }];
  assert.equal(
    view.progressText({
      ref: "R1.3",
      startedAt: "2026-09-30T10:00:00.000Z",
      now: Date.parse("2026-09-30T10:01:05.000Z"),
      comments,
    }),
    "R1.3 running 1:05. 2 of 3 comments not done.",
  );
});

test("reviewerOfRef reads R<n>.<m> and E.<m> refs", () => {
  assert.equal(view.reviewerOfRef("R2.3"), 2);
  assert.equal(view.reviewerOfRef("R1.0"), 1);
  assert.equal(view.reviewerOfRef("E.1"), 0);
  assert.equal(view.reviewerOfRef("E1"), null);
  assert.equal(view.reviewerOfRef("1.3"), null);
});

test("a change applied without its rationale comment: warning flag and an invariant mismatch", () => {
  const missing = ch({ word_comment_id: null, comment_missing: "CommentAnchorMissing" });
  assert.deepEqual(
    view.changeFlags(missing).map((f) => [f.kind, f.text]),
    [["comment-missing", "No rationale comment"]],
  );
  const rev = {
    comments: [
      {
        ref: "R1.1",
        changes: [{ id: "R1.1-1", kind: "replace", state: "pending", word_comment_id: "1" }],
      },
      { ref: "R1.3", changes: [{ ...missing, state: "accepted" }] },
    ],
  };
  const r = view.invariantReport(rev, [{ id: "1", content: "[R1.1] a / b / c" }]);
  assert.deepEqual(
    [r.ok, r.changes, r.expected_comments, r.tool_comments, r.missing_comment, r.uncommented],
    [false, 2, 2, 1, [], ["R1.3-1"]],
  );
  assert.deepEqual(
    view.commentsToResolve([{ change_id: "R1.3-1", state: "accepted" }], rev, []),
    [],
  );
});

test("an unsure ledger item waits for the author and says how many paragraphs", () => {
  const unsure = [
    { paragraph_index: 5, h: "a", reason: "r" },
    { paragraph_index: 8, h: "b", reason: "r" },
  ];
  const d = view.pendingDecisions({
    comments: [],
    ledger: [
      { id: "L1", status: "unsure", unsure },
      { id: "L2", status: "violations", violation_count: 1, unsure: unsure.slice(0, 1) },
      { id: "L3", status: "ok", unsure: [] },
    ],
  });
  assert.deepEqual(
    d.ledger.map((x) => x.text),
    ["L1 unsure on paragraph(s) 5, 8", "L2 unsure on paragraph(s) 5"],
  );
  assert.equal(view.ledgerStatusText({ status: "unsure", unsure }), "unsure on 2 paragraph(s)");
  assert.equal(
    view.ledgerStatusText({ status: "violations", violation_count: 1, unsure: unsure.slice(0, 1) }),
    "1 violation(s), unsure on 1 paragraph(s)",
  );
});

test("invariantReport: a change whose comment the cleanup removed expects no comment", () => {
  const rev = {
    comments: [
      {
        ref: "R1.6",
        changes: [
          {
            id: "R1.6-1",
            kind: "insert",
            state: "accepted",
            word_comment_id: "60",
            comment_cleaned_at: "2026-09-30T00:00:00.000Z",
          },
          { id: "R1.6-2", kind: "replace", state: "pending", word_comment_id: "61" },
        ],
      },
    ],
  };
  const after = view.invariantReport(rev, [{ id: "61", content: "[R1.6] a / b / c" }]);
  assert.deepEqual([after.ok, after.expected_comments, after.missing_comment], [true, 1, []]);
  const lost = view.invariantReport(rev, []);
  assert.deepEqual([lost.ok, lost.missing_comment], [false, ["R1.6-2"]]);
});

test("review file: txt and md, BOM and CRLF removed, size and type checked", () => {
  assert.deepEqual(view.reviewTextFromFile({ name: "r.txt", size: 10, text: "\uFEFFa\r\nb" }), {
    ok: true,
    text: "a\nb",
  });
  assert.equal(view.reviewTextFromFile({ name: "r.MD", size: 1, text: "x" }).ok, true);
  assert.match(
    view.reviewTextFromFile({ name: "r.docx", size: 1, text: "" }).error,
    /\.txt or \.md/,
  );
  assert.match(view.reviewTextFromFile({ name: "r.txt", size: 3_000_000, text: "" }).error, /2 MB/);
  assert.match(view.reviewTextFromFile({ name: "r.txt", size: 0, text: "  " }).error, /empty/);
});

test("ledger file: JSON with an items array, anything else is an error message", () => {
  const item = { kind: "forbid", text: "no novel", check: { type: "llm" } };
  const ok = view.ledgerImportFromJson(JSON.stringify({ note: "n", items: [item] }));
  assert.deepEqual(ok, { ok: true, doc: { items: [item] }, count: 1 });
  assert.match(view.ledgerImportFromJson("{ nope").error, /not JSON/);
  assert.match(view.ledgerImportFromJson("[1]").error, /items/);
  assert.match(view.ledgerImportFromJson('{"items":{}}').error, /items/);
  assert.match(view.ledgerImportFromJson('{"items":[]}').error, /no items/);
});

test("ledger candidates: label says what the check does; the rule drops the candidate fields", () => {
  const c = {
    kind: "forbid",
    text: 'Do not use "novel".',
    scope: "manuscript",
    check: { type: "regex", pattern: "novel", flags: "i", expect: "absent" },
    why: "forbid cue",
    sentence_index: 1,
    picked: true,
  };
  assert.equal(view.candidateLabel(c), 'Forbidden, must not match /novel/i: Do not use "novel".');
  assert.equal(
    view.candidateLabel({ kind: "instruction", text: "Use SI units.", check: { type: "llm" } }),
    "Instruction, LLM check: Use SI units.",
  );
  assert.match(
    view.candidateLabel({ ...c, kind: "instruction", check: { ...c.check, expect: "present" } }),
    /^Instruction, must keep \/novel\/i/,
  );
  assert.deepEqual(Object.keys(view.candidateRule(c)), ["kind", "text", "scope", "check"]);
});

// Codex review 16: which text of a paragraph is still a tracked change, run
// by run, read from the paragraph OOXML the snapshot already fetches.
test("trackedRuns lists the tracked insertions and deletions of the document part, with their place", () => {
  const part = (name, body) =>
    `<pkg:part pkg:name="${name}"><pkg:xmlData>${body}</pkg:xmlData></pkg:part>`;
  const pkg = (...parts) => `<pkg:package>${parts.join("")}</pkg:package>`;
  const r = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
  const ins = (t, id = 1) => `<w:ins w:id="${id}" w:author="A">${r(t)}</w:ins>`;
  const del = (t, id = 2) =>
    `<w:del w:id="${id}" w:author="A"><w:r><w:delText xml:space="preserve">${t}</w:delText></w:r></w:del>`;
  const p = (...xs) => `<w:body><w:p>${xs.join("")}</w:p></w:body>`;
  const doc = (...xs) => pkg(part("/word/document.xml", p(...xs)));

  assert.deepEqual(view.trackedRuns(doc(r("Plain text only."))), []);
  // "Cells grew [ins: well ][del: slowly ]on glucose."
  assert.deepEqual(
    view.trackedRuns(doc(r("Cells grew "), ins("well "), del("slowly "), r("on glucose."))),
    [
      { kind: "ins", text: "well ", start: 11, marks: 1 },
      { kind: "del", text: "slowly ", start: 16, marks: 1 },
    ],
  );
  // Word splits one insertion into several runs: one entry.
  assert.deepEqual(
    view.trackedRuns(
      doc(r("A "), `<w:ins w:id="1">${r("new")}${r(" word")}</w:ins>`, ins(" here", 3), r(".")),
    ),
    [{ kind: "ins", text: "new word here", start: 2, marks: 2 }],
  );
  // The same inserted text twice: two entries, by place.
  assert.deepEqual(
    view.trackedRuns(doc(r("One "), ins("(n = 3) "), r("and two "), ins("(n = 3) ", 5), r("."))),
    [
      { kind: "ins", text: "(n = 3) ", start: 4, marks: 1 },
      { kind: "ins", text: "(n = 3) ", start: 20, marks: 1 },
    ],
  );
  // Only the document part: the tracked body of a rationale comment, a
  // style, a header do not count. Entities are decoded. w:delText and
  // w:instrText are no w:del / w:ins; a self-closing mark (an inserted
  // paragraph mark in rPr) wraps no text. A deleted insertion is a deletion.
  const tricky =
    r("A &amp; B ") +
    "<w:r><w:instrText>PAGE</w:instrText></w:r>" +
    '<w:r><w:rPr><w:ins w:id="9"/></w:rPr><w:t>kept</w:t></w:r>' +
    `<w:ins w:id="4"><w:del w:id="5"><w:r><w:delText>gone</w:delText></w:r></w:del></w:ins>` +
    "<w:r><w:t/></w:r>" +
    ins(" &lt;x&gt;", 6);
  assert.deepEqual(
    view.trackedRuns(
      pkg(
        part("/word/comments.xml", `<w:comment>${ins("tracked comment body")}</w:comment>`),
        part("/word/document.xml", p(tricky)),
        part("/word/styles.xml", del("x")),
      ),
    ),
    [
      { kind: "del", text: "gone", start: 10, marks: 1 },
      // The w:ins around the deletion gave no inserted text of its own.
      { kind: "ins", text: "", start: 14, marks: 1 },
      { kind: "ins", text: " <x>", start: 14, marks: 1 },
    ],
  );
  // Whitespace before a run is counted as the daemon reads the paragraph:
  // collapsed, and without the comment anchor character.
  assert.deepEqual(view.trackedRuns(doc(r("\u0005 A   B\t"), ins("C"))), [
    { kind: "ins", text: "C", start: 4, marks: 1 },
  ]);
  assert.deepEqual(view.trackedRuns(null), []);
  assert.deepEqual(view.trackedRuns("<w:p><w:r><w:t>bare</w:t></w:r></w:p>"), []);
});

test("an open requirement of a comment under a section names the section", () => {
  const d = view.pendingDecisions({
    comments: [
      {
        ref: "R1.1",
        section: "Abstract",
        requirements: [{ id: "q1", text: "Shorten this section." }],
        changes: [],
        blocked: [],
      },
      {
        ref: "R1.2",
        requirements: [{ id: "q1", text: "Add units." }],
        changes: [],
        blocked: [],
      },
    ],
    ledger: [],
  });
  assert.deepEqual(
    d.requirements.map((x) => x.text),
    [
      "R1.1 (Abstract) q1 has no change: Shorten this section.",
      "R1.2 q1 has no change: Add units.",
    ],
  );
});

// Codex review of the follow-up: a tracked run with <w:noBreakHyphen/> lost
// that character ("non-linear" read "nonlinear"), could not be found in the
// paragraph text and was dropped, so a pending insertion read accepted.
// Expected texts are what Word returned for Paragraph.text on 2026-10-01
// (Word 16, document of scripts/e2e/make_tracked_special.py, read through
// collectReviewSnapshot): no-break hyphen U+001E, soft hyphen U+001F, tab
// U+0009, line break U+000B, carriage return U+000D, a w:sym of any font "(".
test("trackedRuns keeps the run elements that are text without being w:t, as Word's paragraph text has them", () => {
  const doc = (kind, content) => {
    const tag = kind === "ins" ? "w:t" : "w:delText";
    const body = content
      .replaceAll("<T>", `<${tag} xml:space="preserve">`)
      .replaceAll("</T>", `</${tag}>`);
    return (
      '<pkg:package><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:body><w:p>' +
      '<w:r><w:t xml:space="preserve">Before </w:t></w:r>' +
      `<w:${kind} w:id="1"><w:r><w:rPr><w:b/></w:rPr>${body}</w:r></w:${kind}>` +
      '<w:r><w:t xml:space="preserve"> after.</w:t></w:r>' +
      "</w:p></w:body></pkg:xmlData></pkg:part></pkg:package>"
    );
  };
  const SYM = '<w:sym w:font="Symbol" w:char="F061"/>';
  const cases = [
    ["ins", "<T>plain words</T>", "plain words"],
    ["ins", "<T>non</T><w:noBreakHyphen/><T>linear </T>", "non\u001elinear "],
    ["ins", "<T>micro</T><w:softHyphen/><T>scope</T>", "micro\u001fscope"],
    ["ins", "<T>left</T><w:tab/><T>right</T>", "left\tright"],
    [
      "ins",
      '<T>left</T><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/><T>right</T>',
      "left\tright",
    ],
    ["ins", "<T>first</T><w:br/><T>second</T>", "first\u000bsecond"],
    ["ins", '<T>first</T><w:br w:type="textWrapping"/><T>second</T>', "first\u000bsecond"],
    ["ins", "<T>one</T><w:cr/><T>two</T>", "one\rtwo"],
    ["ins", `<T>alpha </T>${SYM}<T> end</T>`, "alpha ( end"],
    ["ins", '<T>mark </T><w:sym w:font="Wingdings" w:char="F0FC"/>', "mark ("],
    ["ins", "<T>a &amp; b  &lt;c&gt; </T>", "a & b  <c> "],
    ["ins", "<w:noBreakHyphen/>", "\u001e"],
    ["ins", "<w:instrText>PAGE</w:instrText><T>kept</T>", "kept"],
    ["del", "<T>non</T><w:noBreakHyphen/><T>linear</T>", "non\u001elinear"],
    ["del", "<T>micro</T><w:softHyphen/><T>scope</T>", "micro\u001fscope"],
    ["del", "<T>left</T><w:tab/><T>right</T>", "left\tright"],
    ["del", `<T>alpha </T>${SYM}<T> end</T>`, "alpha ( end"],
    ["del", "<w:delInstrText>PAGE</w:delInstrText><T>gone</T>", "gone"],
  ];
  for (const [kind, content, text] of cases) {
    assert.deepEqual(
      view.trackedRuns(doc(kind, content)),
      [{ kind, text, start: 7, marks: 1 }],
      content,
    );
  }
  console.log(`tracked run element corpus: ${cases.length} snippets`);

  // marks: how many w:ins / w:del elements an entry stands for. Two
  // neighbouring insertions are one entry of two marks; a tracked element
  // from which no text was read is an entry with empty text, so the daemon
  // knows something tracked is there that it cannot place.
  const two =
    '<pkg:package><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:p>' +
    '<w:ins w:id="1"><w:r><w:t>one</w:t></w:r></w:ins><w:ins w:id="2"><w:r><w:t> two</w:t></w:r></w:ins>' +
    "<w:r><w:t> plain </w:t></w:r>" +
    '<w:ins w:id="3"><w:r><w:drawing/></w:r></w:ins>' +
    '<w:del w:id="4"><w:r><w:delText>gone</w:delText></w:r></w:del>' +
    "</w:p></pkg:xmlData></pkg:part></pkg:package>";
  assert.deepEqual(view.trackedRuns(two), [
    { kind: "ins", text: "one two", start: 0, marks: 2 },
    { kind: "ins", text: "", start: 14, marks: 1 },
    { kind: "del", text: "gone", start: 14, marks: 1 },
  ]);
});

test("doneBlockers: what is still open when a comment is marked Done", () => {
  const c = (changes, requirements = []) => ({ ref: "R1.1", requirements, changes });
  const change = (state, ids = []) => ({ state, rationale: { requirement_ids: ids } });
  // Everything decided and answered, no ledger violation: nothing to ask.
  assert.deepEqual(
    view.doneBlockers(c([change("accepted", ["q1"])], [{ id: "q1", text: "a" }]), [
      { id: "L1", status: "clean", violation_count: 0 },
      { id: "L2", status: "skipped", violation_count: 0 },
    ]),
    [],
  );
  assert.deepEqual(
    view.doneBlockers(c([change("rejected"), change("user_edited")]), undefined),
    [],
  );
  // The wording of the open-requirement case is the one the tab always had.
  assert.deepEqual(view.doneBlockers(c([], [{ id: "q1", text: "state the unit" }]), []), [
    'q1 "state the unit" has no change',
  ]);
  // Changes not decided in Word, and confirmed ledger rules with violations.
  assert.deepEqual(
    view.doneBlockers(
      c([change("pending"), change("pending"), change("unlocated")], [{ id: "q1", text: "a" }]),
      [
        { id: "L1", status: "violations", violation_count: 2 },
        { id: "L2", status: "clean", violation_count: 0 },
        { id: "L3", status: "violations", violation_count: 1 },
      ],
    ),
    [
      'q1 "a" has no change',
      "3 changes are not decided in Word (2 pending, 1 not found)",
      "the ledger check has violations (L1: 2, L3: 1)",
    ],
  );
  assert.deepEqual(view.doneBlockers(c([change("unlocated")]), []), [
    "1 change is not decided in Word (1 not found)",
  ]);
});
