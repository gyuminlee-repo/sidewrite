// "Check other places": a second turn of a comment that is given the places
// its changes did not touch (taskpane/shared/follow-places.js) and told to
// make them agree. All texts here are made up (a bridge inspection report).
import { plainPane } from "./helpers/plain-pane.mjs";
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { followNote, followPlan } from "../daemon/paragraph-check.mjs";
import { createReviewService } from "../daemon/review-service.mjs";
import { applyTrackedEdit, applyTrackedRewrite, askAuthor } from "../daemon/review-tools.mjs";
import {
  RevisionSchema,
  loadRevision,
  replaceComments,
  setFollowThrough,
} from "../daemon/revision-store.mjs";

const PARAGRAPHS = [
  "Deck survey",
  "Rust was eliminated on all 14 spans after recoating (Fig. 2).",
  "The coating held through the winter.",
  "Figure 2. Rust coverage per span.",
];
const COMMENTS = [
  { ref: "R1.1", reviewer: 1, text: "Saying rust was eliminated is too strong." },
  { ref: "R1.2", reviewer: 1, text: "Name the paint supplier." },
];

const T0 = "2026-01-01T00:00:00.000Z"; // the status refresh
const REV = {
  snapshot: { taken_at: T0, paragraphs: PARAGRAPHS, ids: ["k0", "k1", "k2", "k3"] },
  comments: [
    {
      ...COMMENTS[0],
      changes: [
        {
          id: "R1.1-1",
          kind: "replace",
          find: "was eliminated",
          text: "was reduced",
          paragraph_index: 1,
          paragraph_id: "k1",
          rationale: { what: "w", why: "y", addresses: "a", requirement_ids: ["q1"] },
          narrowed_from: null,
          state: "pending",
          created_at: "2026-01-02T00:00:00.000Z", // after the refresh: put into the text
        },
        {
          id: "R1.1-2",
          kind: "delete",
          find: " through the winter",
          text: "",
          paragraph_index: 2,
          paragraph_id: "k2",
          rationale: { what: "w", why: "y", addresses: "a", requirement_ids: [] },
          narrowed_from: null,
          state: "rejected",
          created_at: "2025-12-31T00:00:00.000Z", // before the refresh, which found it rejected
        },
      ],
    },
    { ...COMMENTS[1], changes: [] },
  ],
};

test("followPlan: the changes of the comment, their paragraphs as they read now, the other places", () => {
  const plan = followPlan(REV, "R1.1");
  assert.deepEqual(plan.changes, [
    {
      id: "R1.1-1",
      kind: "replace",
      before: "was eliminated",
      after: "was reduced",
      requirement_ids: ["q1"],
    },
  ]);
  assert.deepEqual(plan.changed, [
    { index: 1, id: "k1", text: "Rust was reduced on all 14 spans after recoating (Fig. 2)." },
  ]);
  assert.deepEqual(
    plan.candidates.map((p) => [p.id, p.kind, p.why, p.text]),
    [
      ["k0", "heading", "heading above a changed paragraph", "Deck survey"],
      ["k3", "caption", "legend of Fig. 2", "Figure 2. Rust coverage per span."],
      // The rejected deletion is not in the document: the paragraph is whole.
      ["k2", "body", "same section as a changed paragraph", "The coating held through the winter."],
    ],
  );
  assert.deepEqual(followPlan(REV, "R1.2").candidates, []);
});

test("followNote: the instruction a follow-through turn carries", () => {
  const note = followNote(REV, "R1.1", "F1");
  assert.equal(
    note,
    [
      "[Other places of R1.1]",
      "This turn does one thing. The changes made for R1.1 so far:",
      '- R1.1-1 (q1): "was eliminated" → "was reduced"',
      'The 3 places below were not changed by them. Each is given as [paragraph id] (what it is; why it is listed) and its text as it stands now. For each place decide whether it still states what these changes withdrew, reworded or made conditional. Where it does, change it so that it agrees with them: office_tracked_edit or office_tracked_rewrite, comment_ref "R1.1", with the requirement_ids of the change it follows. Where agreeing takes data, a number, an analysis or a citation the manuscript does not have, ask with office_ask_author and name the place. Where it already agrees, leave it. Make no other change in this turn.',
      "<<<F1",
      "[k0] (heading; heading above a changed paragraph) Deck survey",
      "[k3] (caption; legend of Fig. 2) Figure 2. Rust coverage per span.",
      "[k2] (body; same section as a changed paragraph) The coating held through the winter.",
      ">>>F1",
    ].join("\n"),
  );
  // A deletion and an insertion, and a long text cut short.
  const rev = structuredClone(REV);
  rev.comments[0].changes = [
    { ...rev.comments[0].changes[0], kind: "delete", find: "Rust ", text: "" },
    {
      ...rev.comments[0].changes[0],
      id: "R1.1-3",
      kind: "insert",
      find: "recoating",
      text: ` ${"x".repeat(300)}`,
      rationale: null,
    },
  ];
  const lines = followNote(rev, "R1.1", "F1").split("\n");
  assert.equal(lines[2], '- R1.1-1 (q1): took out "Rust "');
  assert.equal(lines[3], `- R1.1-3: after "recoating" put in " ${"x".repeat(196)}..."`);
  // Nothing to check: no note.
  assert.throws(
    () => followNote(REV, "R1.2", "F1"),
    /R1\.2 has no change in the document, so there is no other place to check/,
  );
  const alone = structuredClone(REV);
  alone.snapshot = { taken_at: T0, paragraphs: [PARAGRAPHS[1]], ids: ["k1"] };
  assert.throws(
    () => followNote(alone, "R1.1", "F1"),
    /No other place was found for the changes of R1\.1/,
  );
});

test("the field of the review file: old files load, a comment-list save keeps it", () => {
  let rev = RevisionSchema.parse({
    version: 1,
    manuscript: "m",
    updated_at: "t",
    comments: [{ ref: "R1.1", reviewer: 1, text: "c" }],
  });
  // No default: an old file saves as it was loaded.
  assert.ok(!("follow_through" in rev.comments[0]));
  const record = {
    at: "2026-01-03T00:00:00.000Z",
    candidate_ids: ["k0", "k3"],
    change_ids: ["R1.1-2"],
    questions: 1,
    seen_change_ids: ["R1.1-1", "R1.1-2"],
  };
  rev = setFollowThrough(rev, "R1.1", record);
  assert.deepEqual(rev.comments[0].follow_through, record);
  assert.equal(setFollowThrough(rev, "R9.9", record), null);
  assert.deepEqual(RevisionSchema.parse(JSON.parse(JSON.stringify(rev))), rev);
  const saved = replaceComments(rev, [{ ref: "R1.1", reviewer: 1, text: "c edited" }]);
  assert.deepEqual(saved.comments[0].follow_through, record);
});

async function setup() {
  const dir = await mkdtemp(join(tmpdir(), "follow-svc-"));
  const doc = join(dir, "paper.docx");
  const events = [];
  const lines = [];
  const pushed = [];
  const svc = createReviewService({
    getContext: () => ({ activeDoc: doc, caps: ["target_hash"] }),
    callTool: async () => ({ ok: true }),
    pushTurn: (key, host, text, turn) => pushed.push({ key, host, text, turn }),
    oneshot: null,
    sendEvent: (key, obj) => events.push(obj),
    log: (s) => lines.push(s),
    env: { CC_OFFICE_WSL_PATHS: "0" },
  });
  const ask = (type, payload = {}) =>
    new Promise((resolve) => {
      svc.handlers[type]({ type, request_id: "q1", ...plainPane(payload) }, resolve, "k", "word");
    });
  await ask("review_save_comments", { comments: COMMENTS });
  await ask("review_refresh", {
    snapshot: {
      paragraphs: PARAGRAPHS.map((text, index) => ({
        index,
        id: `p${index}`,
        text,
        reviewed_text: text,
        w_ins: 0,
        w_del: 0,
      })),
      comments: [],
    },
  });
  // Word, as the task pane answers for it: the paragraph the edit went into.
  const word = { paragraph_index: 1 };
  const deps = {
    ...svc.trackedEditDeps("k"),
    call: async () => ({
      ok: true,
      word_comment_id: 5,
      paragraph_index: word.paragraph_index,
      before: "b",
      after: "a",
      target_checked: true,
    }),
  };
  const R = { what: "w", why: "y", addresses: "a" };
  const edit = (find, replace, turn) =>
    applyTrackedEdit(deps, { comment_ref: "R1.1", find, replace, rationale: R }, turn);
  const turnDone = async (turn, outcome) => {
    const n = () => events.filter((e) => e.event === "turn_done").length;
    const before = n();
    svc.turnEnded("k", turn, outcome);
    for (let i = 0; i < 100 && n() === before; i++) await new Promise((r) => setTimeout(r, 5));
    assert.equal(n(), before + 1);
  };
  return {
    dir,
    svc,
    ask,
    lines,
    pushed,
    word,
    deps,
    edit,
    turnDone,
    load: () => loadRevision(join(dir, "paper.revision.json")),
    cleanup: () => rm(dir, { recursive: true, force: true }),
  };
}

const CAPTION = ["Rust coverage per span.", "Rust coverage on the 14 spans."];

test("review_propose with follow_through: the turn is given the other places of the comment", async () => {
  const t = await setup();
  try {
    // No change yet: no turn.
    const none = await t.ask("review_propose", {
      ref: "R1.1",
      paragraphs: [],
      follow_through: true,
    });
    assert.equal(none.ok, false);
    assert.match(none.error, /R1\.1 has no change in the document/);
    assert.equal(t.pushed.length, 0);

    const first = await t.edit("was eliminated", "was reduced", { ref: "R1.1" });
    assert.equal(first.ok, true, first.message);

    // Not with a rewrite in one turn.
    const both = await t.ask("review_propose", {
      ref: "R1.1",
      paragraphs: [],
      follow_through: true,
      rewrite_paragraph: 1,
    });
    assert.equal(both.ok, false);
    assert.match(both.error, /either writes a paragraph anew or checks the other places/);
    assert.equal(t.pushed.length, 0);

    const r = await t.ask("review_propose", { ref: "R1.1", paragraphs: [], follow_through: true });
    assert.equal(r.ok, true, r.error);
    const { text, turn } = t.pushed.at(-1);
    assert.match(
      text,
      /\n\n\[Other places of R1\.1\]\nThis turn does one thing\. The changes made for R1\.1 so far:\n- R1\.1-1: "was eliminated" → "was reduced"\nThe 3 places below/,
    );
    assert.ok(text.includes("[p3] (caption; legend of Fig. 2) Figure 2. Rust coverage per span."));
    assert.equal(turn.ref, "R1.1");
    assert.deepEqual(turn.follow.candidate_ids, ["p0", "p3", "p2"]);
    assert.deepEqual(turn.follow.sources, [
      "Rust was reduced on all 14 spans after recoating (Fig. 2).",
    ]);
    // A turn without the flag carries nothing of it.
    await t.ask("review_propose", { ref: "R1.1", paragraphs: [] });
    assert.deepEqual(t.pushed.at(-1).turn, { ref: "R1.1" });
    assert.ok(!t.pushed.at(-1).text.includes("[Other places of"));
  } finally {
    await t.cleanup();
  }
});

test("number guard: a number of a changed paragraph may go to its legend in a follow-through turn only", async () => {
  const t = await setup();
  try {
    assert.equal((await t.edit("was eliminated", "was reduced", { ref: "R1.1" })).ok, true);
    t.word.paragraph_index = 3;
    // An ordinary turn: 14 is not in the legend, the comment or an answer.
    const refused = await t.edit(...CAPTION, { ref: "R1.1" });
    assert.equal(refused.ok, false);
    assert.equal(refused.code, "NumberNeedsAuthor");
    assert.deepEqual(refused.numbers, ["14"]);

    await t.ask("review_propose", { ref: "R1.1", paragraphs: [], follow_through: true });
    const { turn } = t.pushed.at(-1);
    const ok = await t.edit(...CAPTION, turn);
    assert.equal(ok.ok, true, ok.message);
    // A number no changed paragraph of the comment holds is refused as ever.
    const other = await t.edit("Deck survey", "Deck survey of 2031", turn);
    assert.equal(other.code, "NumberNeedsAuthor");
    // The same for a paragraph written anew (office_tracked_rewrite).
    t.word.paragraph_index = 2;
    const rewrite = (with_) =>
      applyTrackedRewrite(
        { ...t.deps, paneCan: () => true },
        {
          comment_ref: "R1.1",
          paragraph_id: "p2",
          text: "The coating held on the 14 spans through the winter.",
          what: "w",
          why: "y",
          addresses: "a",
        },
        with_,
      );
    assert.equal((await rewrite({ ref: "R1.1" })).code, "NumberNeedsAuthor");
    const anew = await rewrite(turn);
    assert.equal(anew.ok, true, anew.message);
  } finally {
    await t.cleanup();
  }
});

test("the end of a follow-through turn: one log line, and the comment says the check ran", async () => {
  const t = await setup();
  try {
    assert.equal((await t.edit("was eliminated", "was reduced", { ref: "R1.1" })).ok, true);
    await t.ask("review_propose", { ref: "R1.1", paragraphs: [], follow_through: true });
    const { turn } = t.pushed.at(-1);

    // A turn that failed: logged, not marked, so it can be run again.
    await t.turnDone(turn, { ok: false, error: "Claude usage limit reached." });
    assert.ok(t.lines.includes("[review] follow R1.1 candidates=3 changes=0 questions=0"));
    assert.equal((await t.load()).comments[0].follow_through, undefined);

    t.word.paragraph_index = 3;
    assert.equal((await t.edit(...CAPTION, turn)).ok, true);
    const asked = await askAuthor(
      t.deps,
      {
        comment_ref: "R1.1",
        question: "Was the heading 'Deck survey' meant to cover the approach spans?",
        why: "The report does not say.",
      },
      turn,
    );
    assert.equal(asked.ok, true, asked.message);
    await t.turnDone(turn, { ok: true });
    assert.ok(t.lines.includes("[review] follow R1.1 candidates=3 changes=1 questions=1"));
    const c = (await t.load()).comments[0];
    assert.deepEqual(
      [
        c.follow_through.candidate_ids,
        c.follow_through.change_ids,
        c.follow_through.questions,
        c.follow_through.seen_change_ids,
      ],
      [["p0", "p3", "p2"], ["R1.1-2"], 1, ["R1.1-1", "R1.1-2"]],
    );
    assert.match(c.follow_through.at, /^\d{4}-\d\d-\d\dT/);
    // An ordinary turn leaves the record as it is and logs no such line.
    const n = t.lines.filter((l) => l.startsWith("[review] follow ")).length;
    await t.turnDone({ ref: "R1.1" }, { ok: true });
    assert.equal(t.lines.filter((l) => l.startsWith("[review] follow ")).length, n);
    assert.deepEqual((await t.load()).comments[0].follow_through, c.follow_through);
  } finally {
    await t.cleanup();
  }
});
