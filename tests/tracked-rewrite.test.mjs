// office_tracked_rewrite: a paragraph rewritten as a whole goes into Word as
// the tracked edits of the words that differ, from the last to the first,
// each checked against the paragraph as it reads then, with one rationale
// comment for the paragraph. A fake task pane applies them to an array.
// All texts here are made up.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createReviewService } from "../daemon/review-service.mjs";
import {
  applyTrackedRewrite,
  trackedRewriteShape,
  currentTexts,
  REVIEW_WRITE_TOOLS,
  reviewTurnToolAllowed,
} from "../daemon/review-tools.mjs";
import { loadRevision, byCreated } from "../daemon/revision-store.mjs";
import { paragraphHash } from "../daemon/doc-diff.mjs";
import * as view from "../taskpane/shared/review-view.js";

const PARAS = [
  "Results",
  "Gene A is essential for growth on glycerol. The mutant grew at 0.21 per hour at 30 C. These data prove that gene A is essential.",
  "The parent grew at 0.42 per hour.",
];
const NEW =
  "Deletion of gene A halved the growth rate on glycerol. The mutant grew at 0.21 per hour at 30 °C. These data show a slower growth of the mutant.";
const COMMENT = {
  ref: "R1.1",
  reviewer: 1,
  text: "The mutant grows, so gene A is not essential. Write the unit correctly.",
  requirements: ["remove the claim that gene A is essential", "write the unit correctly"],
};

async function world({
  caps = ["target_hash", "tracked_rewrite"],
  failAt = null,
  notPlain = false,
} = {}) {
  const dir = await mkdtemp(join(tmpdir(), "rewrite-"));
  const doc = join(dir, "paper.docx");
  const paras = [...PARAS];
  const calls = [];
  const lines = [];
  let commentId = 40;
  // The fake task pane: one office_tracked_edit as tools-word.js does it.
  const callTool = async (name, args) => {
    calls.push(args);
    if (failAt === calls.length) return { ok: false, code: "NotFound", message: "not found" };
    if (args.plain_only && notPlain) {
      return { ok: false, code: "ParagraphNotPlain", message: "The paragraph holds a field." };
    }
    const i = args.expected_paragraph.index;
    const before = paras[i];
    if (paragraphHash(before) !== args.expected_before_hash) {
      return { ok: false, code: "StaleTarget", paragraph_index: i };
    }
    if (before.split(args.find).length !== 2) return { ok: false, code: "Ambiguous" };
    const after = before.replace(args.find, () => args.replace);
    paras[i] = after;
    return {
      ok: true,
      word_comment_id: args.no_comment ? null : ++commentId,
      paragraph_index: i,
      paragraph_id: null,
      before,
      after,
      narrowed: false,
      target_checked: true,
      ...(args.no_comment ? { no_comment: true } : {}),
    };
  };
  const svc = createReviewService({
    getContext: () => ({ activeDoc: doc, caps }),
    callTool,
    pushTurn: () => {},
    sendEvent: () => {},
    log: (l) => lines.push(l),
    env: { CC_OFFICE_WSL_PATHS: "0" },
  });
  const ask = (type, payload = {}) =>
    new Promise((resolve) =>
      svc.handlers[type]({ type, request_id: "q", ...payload }, resolve, "k", "word"),
    );
  await ask("review_save_comments", { comments: [COMMENT] });
  await ask("review_refresh", {
    snapshot: {
      paragraphs: PARAS.map((text, index) => ({
        index,
        id: null,
        text,
        reviewed_text: text,
        w_ins: 0,
        w_del: 0,
      })),
      comments: [],
    },
  });
  const deps = { ...svc.trackedEditDeps("k"), call: callTool };
  const rev = () => loadRevision(join(dir, "paper.revision.json"));
  return {
    dir,
    svc,
    ask,
    deps,
    paras,
    calls,
    lines,
    rev,
    cleanup: () => rm(dir, { recursive: true, force: true }),
  };
}

const ARGS = {
  comment_ref: "R1.1",
  paragraph_id: "p1",
  text: NEW,
  what: "Rewrote the paragraph without the claim of essentiality.",
  why: "The mutant still grows.",
  addresses: "the claim and the unit",
  requirement_ids: ["q1", "q2"],
};
const TURN = { ref: "R1.1" };

test("the tool is a write tool of a review turn, with the paragraph and the rationale as its arguments", () => {
  assert.deepEqual([...REVIEW_WRITE_TOOLS], ["office_tracked_edit", "office_tracked_rewrite"]);
  assert.equal(reviewTurnToolAllowed("office_tracked_rewrite"), true);
  assert.deepEqual(Object.keys(trackedRewriteShape), [
    "comment_ref",
    "paragraph_id",
    "text",
    "what",
    "why",
    "addresses",
    "requirement_ids",
  ]);
});

test("a rewrite goes in as the edits of the words that differ, last first, with one comment for the paragraph", async () => {
  const t = await world();
  try {
    const r = await applyTrackedRewrite(t.deps, ARGS, TURN);
    assert.equal(r.ok, true, r.message);
    assert.equal(t.paras[1], NEW, "the fake document reads the new paragraph");
    assert.equal(t.calls.length, 3);
    assert.equal(r.parts, 3);
    // From the last to the first: each find lies before the one applied before it.
    const at = t.calls.map((c) => PARAS[1].indexOf(c.find.slice(0, 12)));
    assert.deepEqual(
      at,
      [...at].sort((a, b) => b - a),
    );
    // The first one sent carries the comment and asks for a plain paragraph;
    // the others add no comment.
    assert.match(
      t.calls[0].comment_text,
      /^\[R1\.1\] Rewrote the paragraph without the claim of essentiality\. \/ The mutant still grows\. \/ the claim and the unit \[q1,q2\]/,
    );
    assert.equal(t.calls[0].plain_only, true);
    assert.equal(t.calls[0].no_comment, undefined);
    assert.deepEqual(
      t.calls.slice(1).map((c) => [c.no_comment, c.plain_only]),
      [
        [true, undefined],
        [true, undefined],
      ],
    );
    // Every call was checked against the paragraph as it read then (the fake
    // refuses another hash), under the paragraph of the rewrite.
    assert.ok(t.calls.every((c) => c.expected_paragraph.index === 1 && c.paragraph_hint === "p1"));
    // On record: three changes, one comment id, one rewrite id.
    const changes = (await t.rev()).comments[0].changes;
    assert.equal(changes.length, 3);
    assert.deepEqual(new Set(changes.map((ch) => ch.word_comment_id)), new Set(["41"]));
    assert.equal(new Set(changes.map((ch) => ch.rewrite_id)).size, 1);
    assert.ok(changes[0].rewrite_id);
    assert.deepEqual(
      r.change_ids,
      changes.map((ch) => ch.id),
    );
    assert.deepEqual(changes[0].rationale.requirement_ids, ["q1", "q2"]);
    // The paragraph as the daemon has it now is the new one.
    const rev = await t.rev();
    assert.equal(currentTexts(rev.snapshot.paragraphs, [...changes].sort(byCreated))[1], NEW);
    assert.ok(t.lines.includes("[review] rewrite R1.1 paragraph=1 parts=3 applied=3"));
    // The report of the Review tab: three changes, one comment expected, and that is in order.
    const inv = view.invariantReport(rev, [{ id: "41", content: t.calls[0].comment_text }]);
    assert.deepEqual(
      [inv.changes, inv.expected_comments, inv.tool_comments, inv.ok],
      [3, 1, 1, true],
    );
  } finally {
    await t.cleanup();
  }
});

test("the guards run once, on the paragraph as a whole, before Word is called", async () => {
  const t = await world();
  try {
    // A number nothing gives.
    let r = await applyTrackedRewrite(t.deps, { ...ARGS, text: NEW.replace("0.21", "0.25") }, TURN);
    assert.deepEqual([r.ok, r.code, r.numbers], [false, "NumberNeedsAuthor", ["0.25"]]);
    // A number of the paragraph moved to another sentence is not a new one.
    const moved =
      "At 30 C the mutant grew at 0.21 per hour. Gene A is essential for growth on glycerol. These data prove that gene A is essential.";
    assert.deepEqual(t.calls, []);
    // A forbidden word the rewrite brings in; one the paragraph already had does not count.
    await t.ask("ledger_add", {
      item: {
        kind: "forbid",
        text: "no novel",
        check: { type: "regex", pattern: String.raw`\bnovel\b`, flags: "i" },
      },
    });
    await t.ask("ledger_add", {
      item: {
        kind: "forbid",
        text: "no prove",
        check: { type: "regex", pattern: String.raw`\bprove\b`, flags: "i" },
      },
    });
    r = await applyTrackedRewrite(
      t.deps,
      { ...ARGS, text: NEW.replace("a slower", "a novel, slower") },
      TURN,
    );
    assert.deepEqual([r.ok, r.code, r.blocked_by], [false, "RewriteBlocked", ["L1"]]);
    assert.deepEqual(t.calls, []);
    r = await applyTrackedRewrite(t.deps, { ...ARGS, text: moved }, TURN);
    assert.equal(r.ok, true, r.message);
    assert.equal(t.paras[1], moved);
  } finally {
    await t.cleanup();
  }
});

test("what the rewrite drops goes on the one comment; nothing to change, an unknown paragraph and another turn are refused", async () => {
  const t = await world();
  try {
    let r = await applyTrackedRewrite(t.deps, { ...ARGS, text: PARAS[1] }, TURN);
    assert.match(r.message, /reads the same/);
    r = await applyTrackedRewrite(t.deps, { ...ARGS, paragraph_id: "p9" }, TURN);
    assert.equal(r.code, "ParagraphNotFound");
    r = await applyTrackedRewrite(t.deps, ARGS, { ref: "R1.2" });
    assert.match(r.message, /does not match the comment of this turn/);
    r = await applyTrackedRewrite(t.deps, ARGS, null);
    assert.match(r.message, /only in a review turn/);
    r = await applyTrackedRewrite(t.deps, { ...ARGS, requirement_ids: ["q7"] }, TURN);
    assert.match(r.message, /Unknown requirement id/);
    assert.deepEqual(t.calls, []);
    // The rate is dropped: said on the comment and in the answer.
    const dropped =
      "Deletion of gene A slowed growth on glycerol at 30 C. These data prove that gene A is essential.";
    r = await applyTrackedRewrite(t.deps, { ...ARGS, text: dropped }, TURN);
    assert.equal(r.ok, true, r.message);
    assert.ok(r.lost_facts.length > 0);
    assert.match(t.calls[0].comment_text, /\n사라지는 사실: /);
    const changes = (await t.rev()).comments[0].changes;
    assert.deepEqual(changes[0].lost_facts, r.lost_facts);
    assert.ok(changes.slice(1).every((ch) => ch.lost_facts.length === 0));
  } finally {
    await t.cleanup();
  }
});

test("a part Word refuses stops the rewrite: what is in the document is on record and the answer says where it stands", async () => {
  const t = await world({ failAt: 3 });
  try {
    const r = await applyTrackedRewrite(t.deps, ARGS, TURN);
    assert.deepEqual([r.ok, r.code, r.applied_parts, r.parts], [false, "RewriteStopped", 2, 3]);
    assert.equal(t.calls.length, 3);
    assert.equal((await t.rev()).comments[0].changes.length, 2);
    assert.match(r.message, /2 of 3 parts are in the document/);
    assert.match(r.message, /Do not send the rewrite again/);
    assert.equal(r.paragraph_now, t.paras[1]);
    assert.ok(
      t.lines.includes("[review] rewrite R1.1 paragraph=1 parts=3 applied=2 stopped=NotFound"),
    );
  } finally {
    await t.cleanup();
  }
});

test("a rewrite stopped part way is recorded on its comment, shown on the card and kept by a comment-list save", async () => {
  const t = await world({ failAt: 3 });
  try {
    const r = await applyTrackedRewrite(t.deps, ARGS, TURN);
    assert.equal(r.code, "RewriteStopped");
    const c = (await t.rev()).comments[0];
    assert.equal(c.rewrites_stopped.length, 1);
    const w = c.rewrites_stopped[0];
    assert.deepEqual(
      [w.paragraph_index, w.paragraph_id, w.applied, w.parts, w.code],
      [1, "p1", 2, 3, "NotFound"],
    );
    assert.deepEqual(w.change_ids, r.change_ids);
    assert.match(w.rewrite_id, /^R1\.1-w/);
    // The log line the replay parses is as it was.
    assert.ok(
      t.lines.includes("[review] rewrite R1.1 paragraph=1 parts=3 applied=2 stopped=NotFound"),
    );
    const notes = view.commentNotices(c, null);
    assert.deepEqual(
      notes.map((n) => [n.kind, n.paragraph_index]),
      [["rewrite_stopped", 1]],
    );
    assert.match(notes[0].text, /stopped after 2 of 3 parts/);
    assert.ok(view.doneBlockers(c, []).includes("the rewrite of paragraph p1 stopped part way"));
    // A save of the comment list rebuilds each comment field by field.
    await t.ask("review_save_comments", { comments: [COMMENT] });
    assert.deepEqual((await t.rev()).comments[0].rewrites_stopped, c.rewrites_stopped);
  } finally {
    await t.cleanup();
  }
});

test("a rewrite Word refuses at its first part, or one that goes in whole, records no stop", async () => {
  for (const failAt of [1, null]) {
    const t = await world({ failAt });
    try {
      await applyTrackedRewrite(t.deps, ARGS, TURN);
      assert.equal((await t.rev()).comments[0].rewrites_stopped, undefined);
    } finally {
      await t.cleanup();
    }
  }
});

test("a paragraph that is not plain text, and a task pane that cannot do it: nothing is changed", async () => {
  let t = await world({ notPlain: true });
  try {
    const r = await applyTrackedRewrite(t.deps, ARGS, TURN);
    assert.deepEqual([r.ok, r.code, r.applied_parts], [false, "ParagraphNotPlain", 0]);
    assert.match(r.message, /office_tracked_edit/);
    assert.equal(t.calls.length, 1);
    assert.equal(t.paras[1], PARAS[1]);
    assert.equal((await t.rev()).comments[0].changes.length, 0);
  } finally {
    await t.cleanup();
  }
  t = await world({ caps: ["target_hash"] });
  try {
    const r = await applyTrackedRewrite(t.deps, ARGS, TURN);
    assert.deepEqual([r.ok, r.code], [false, "TaskPaneOutdated"]);
    assert.deepEqual(t.calls, []);
  } finally {
    await t.cleanup();
  }
});

test("changes recorded in the same instant are replayed in the order of their ids", () => {
  const ch = (id, find, text) => ({
    id,
    kind: "replace",
    find,
    text,
    paragraph_index: 0,
    created_at: "2026-01-01T00:00:00.000Z",
  });
  const changes = [ch("R1.1-10", "b c", "B"), ch("R1.1-2", "a", "b"), ch("R1.1-1", "x", "a")];
  assert.deepEqual(
    [...changes].sort(byCreated).map((c) => c.id),
    ["R1.1-1", "R1.1-2", "R1.1-10"],
  );
  assert.deepEqual(currentTexts(["x c"], [...changes].sort(byCreated)), ["B"]);
});

test("a comment shared by the parts of a rewrite is resolved when all of them are decided", () => {
  const revision = {
    comments: [
      {
        changes: [
          { id: "c1", word_comment_id: "41", rewrite_id: "w" },
          { id: "c2", word_comment_id: "41", rewrite_id: "w" },
          { id: "c3", word_comment_id: "50" },
        ],
      },
    ],
  };
  const comments = [
    { id: "41", resolved: false },
    { id: "50", resolved: false },
  ];
  const rows = (s1, s2) => [
    { change_id: "c1", state: s1 },
    { change_id: "c2", state: s2 },
    { change_id: "c3", state: "accepted" },
  ];
  assert.deepEqual(view.commentsToResolve(rows("accepted", "pending"), revision, comments), ["50"]);
  assert.deepEqual(view.commentsToResolve(rows("accepted", "rejected"), revision, comments), [
    "41",
    "50",
  ]);
  // The flag of a change row.
  assert.deepEqual(view.changeFlags({ rewrite_id: "w" }), [{ kind: "narrowed", text: "Rewrite" }]);
});
