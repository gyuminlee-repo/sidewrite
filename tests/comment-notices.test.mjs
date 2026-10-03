// What is on record about the paragraphs a comment changed reaches its card,
// its Done check and its next turn: a finding of the last paragraph check
// that the paragraph contradicts itself or still makes a withdrawn claim,
// and a rewrite Word stopped part way. All texts here are made up (a paper
// on urban transport planning).
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { buildReviewPrompt, REVIEW_RULES, outsideRuleInPart } from "../daemon/review-prompt.mjs";
import { createReviewService } from "../daemon/review-service.mjs";
import { loadRevision, saveRevision } from "../daemon/revision-store.mjs";
import {
  commentNotices,
  paragraphFindings,
  doneBlockers,
  stoppedRewrites,
} from "../taskpane/shared/review-view.js";

const change = (id, extra = {}) => ({ id, state: "pending", ...extra });
const COMMENT = {
  ref: "R1.2",
  reviewer: 1,
  text: "If there is no evidence that network densification reduced travel times, remove the claim that it did.",
  changes: [change("R1.2-1"), change("R1.2-2", { state: "rejected" })],
};
const CONSISTENCY = {
  checked_at: "2026-10-01T10:00:00.000Z",
  model: null,
  paragraphs: [
    { index: 4, id: "p4", change_ids: ["R1.2-1", "R2.1-1"], many: false },
    { index: 7, id: "p7", change_ids: ["R2.1-2", "R2.1-3"], many: false },
    { index: 9, id: "p9", change_ids: ["R1.2-2", "R2.1-4"], many: false },
  ],
  findings: [
    {
      paragraph_index: 4,
      paragraph_id: "p4",
      kind: "contradiction",
      detail: "It says travel times fell and that they were not measured.",
      change_ids: [],
    },
    {
      paragraph_index: 4,
      paragraph_id: "p4",
      kind: "conflict",
      detail: "Two changes pull apart.",
      change_ids: ["R1.2-1"],
    },
    {
      paragraph_index: 7,
      paragraph_id: "p7",
      kind: "claim_left",
      detail: "The bus lane claim stays.",
      change_ids: [],
    },
    {
      paragraph_index: 9,
      paragraph_id: "p9",
      kind: "claim_left",
      detail: "Only a rejected change of R1.2 is here.",
      change_ids: [],
    },
  ],
  warnings: [],
};

test("paragraphFindings: a contradiction or a claim left in a paragraph the comment changed, nothing else", () => {
  const found = paragraphFindings(COMMENT, CONSISTENCY);
  assert.deepEqual(
    found.map((f) => [f.paragraph_id, f.kind]),
    [["p4", "contradiction"]],
  );
  // Another comment's paragraph, a kind not surfaced, a rejected change: none.
  assert.deepEqual(paragraphFindings({ ...COMMENT, changes: [] }, CONSISTENCY), []);
  assert.deepEqual(paragraphFindings(COMMENT, null), []);
});

test("commentNotices: the finding and a stopped rewrite, for the card; Done asks about both", () => {
  const c = {
    ...COMMENT,
    rewrites_stopped: [
      {
        rewrite_id: "R1.2-wa",
        paragraph_index: 4,
        paragraph_id: "p4",
        applied: 1,
        parts: 3,
        code: "NotFound",
        change_ids: ["R1.2-1"],
        at: "2026-10-01T09:00:00.000Z",
      },
    ],
  };
  const notes = commentNotices(c, CONSISTENCY);
  assert.deepEqual(
    notes.map((n) => [n.kind, n.paragraph_index, n.paragraph_id]),
    [
      ["finding", 4, "p4"],
      ["rewrite_stopped", 4, "p4"],
    ],
  );
  assert.match(notes[0].text, /the paragraph contradicts itself\. It says travel times fell/);
  assert.match(
    notes[1].text,
    /stopped after 1 of 3 parts \(NotFound\): R1\.2-1 is in the document/,
  );
  const left = doneBlockers(c, [], [], CONSISTENCY);
  assert.ok(
    left.includes("the paragraph check found a problem in paragraph p4 this comment changed"),
  );
  assert.ok(left.includes("the rewrite of paragraph p4 stopped part way"));
  // Without the check passed in, Done says nothing of the finding.
  assert.ok(!doneBlockers(c, [], []).some((l) => l.includes("paragraph check")));
  // A rewrite whose changes were all rejected has left nothing behind.
  const gone = { ...c, changes: [change("R1.2-1", { state: "rejected" })] };
  assert.deepEqual(stoppedRewrites(gone), []);
});

test("buildReviewPrompt: the next turn of the comment is told what is on record for its paragraphs", () => {
  const built = buildReviewPrompt({ comment: COMMENT, consistency: CONSISTENCY, fence: "F" });
  assert.match(
    built.text,
    /On record for paragraphs this comment changed \(paragraph check of 2026-10-01T10:00:00\.000Z\)/,
  );
  assert.match(built.text, /- \[p4\] The paragraph check found: the paragraph contradicts itself/);
  assert.ok(!built.text.includes("bus lane"));
  assert.ok(!built.text.includes("Two changes pull apart"));
  const none = buildReviewPrompt({ comment: COMMENT, fence: "F" });
  assert.ok(!none.text.includes("On record for paragraphs"));
});

test("rule: a request conditional on evidence the file does not give gets its change, with a question beside it", () => {
  for (const rule of [REVIEW_RULES[11], outsideRuleInPart({ main_name: "paper.docx" })]) {
    assert.match(
      rule,
      /conditional on evidence[^.]*make the change the reviewer asked for in that case \(take the claim out, or restate it as what was observed\) and also call office_ask_author with the same request id/,
    );
    assert.match(rule, /The question goes with the change, not in place of it\./);
    assert.match(rule, /office_requirement_outside with its id and `part`/);
    assert.doesNotMatch(rule, /ask whether the condition holds/);
  }
});

test("review_preview: the prompt of a turn carries the findings of the last paragraph check", async () => {
  const dir = await mkdtemp(join(tmpdir(), "notices-"));
  const doc = join(dir, "plan.docx");
  const PARAS = [
    "Results",
    "Network densification reduced travel times across the district.",
    "Ridership rose on the new lines.",
  ];
  try {
    const svc = createReviewService({
      getContext: () => ({ activeDoc: doc, caps: [] }),
      callTool: async () => ({ ok: false }),
      pushTurn: () => {},
      sendEvent: () => {},
      log: () => {},
      env: { CC_OFFICE_WSL_PATHS: "0" },
    });
    const ask = (type, payload = {}) =>
      new Promise((resolve) =>
        svc.handlers[type]({ type, request_id: "q", ...payload }, resolve, "k", "word"),
      );
    await ask("review_save_comments", {
      comments: [{ ref: "R1.2", reviewer: 1, text: COMMENT.text }],
    });
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
    const path = join(dir, "plan.revision.json");
    const rev = await loadRevision(path);
    rev.comments[0].changes = [
      {
        id: "R1.2-1",
        created_at: "2000-01-01T00:00:00.000Z",
        word_comment_id: null,
        paragraph_index: 1,
        kind: "replace",
        find: "reduced",
        text: "was followed by shorter",
        before: PARAS[1],
        after: PARAS[1],
        rationale: { what: "w", why: "y", addresses: "a" },
        state: "pending",
      },
    ];
    rev.consistency = {
      ...CONSISTENCY,
      paragraphs: [{ index: 1, id: "p1", change_ids: ["R1.2-1", "R2.1-1"], many: false }],
      findings: [{ ...CONSISTENCY.findings[0], paragraph_index: 1, paragraph_id: "p1" }],
    };
    await saveRevision(path, rev);
    const r = await ask("review_preview", { ref: "R1.2" });
    assert.ok(r.prompt, JSON.stringify(r));
    assert.match(r.prompt, /- \[p1\] The paragraph check found: the paragraph contradicts itself/);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
