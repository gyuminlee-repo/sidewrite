// Three things a judged revision showed:
//  - what is not in this file (a figure file, the supplement, the title page)
//    is declared outside this file, not put to the author as a question;
//  - a request the reviewer made conditional on evidence the manuscript does
//    not give gets the change it asks for, with a question to the author;
//  - a fact a change takes from the review and the manuscript does not have
//    is marked in its rationale comment.
// All texts here are made up.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { buildReviewPrompt, REVIEW_RULES } from "../daemon/review-prompt.mjs";
import {
  applyTrackedEdit,
  requirementOutside,
  formatRationaleComment,
} from "../daemon/review-tools.mjs";
import { createReviewService } from "../daemon/review-service.mjs";
import { loadRevision } from "../daemon/revision-store.mjs";
import * as view from "../taskpane/shared/review-view.js";

test("prompt rule: outside this file is a note, the author's records are a question, a conditional request is asked", () => {
  const rule = REVIEW_RULES.find((r) => r.startsWith("Tell apart what is not in this file"));
  assert.ok(rule);
  assert.match(
    rule,
    /a figure file, the supplementary material, the title page, the response letter/,
  );
  assert.match(rule, /office_requirement_outside[^.]*not a question/);
  assert.match(
    rule,
    /office_ask_author only for what the author's records or judgement must settle/,
  );
  // A request conditional on evidence the manuscript does not give: the
  // change the reviewer asked for is the default and the question goes with
  // it (the question alone left an unsupported claim standing).
  assert.doesNotMatch(rule, /ask whether the condition holds/);
  assert.match(
    rule,
    /conditional on evidence[^.]*is not left as it is when the manuscript does not give that evidence: make the change the reviewer asked for in that case \(take the claim out, or restate it as what was observed\) and also call office_ask_author with the same request id/,
  );
  assert.match(rule, /The question goes with the change, not in place of it\./);
  const { text } = buildReviewPrompt({
    comment: { ref: "R1.1", reviewer: 1, text: "Provide the raw curves." },
    fence: "F1",
  });
  assert.ok(text.includes(rule));
  // Appended: the rules before it keep their numbers.
  assert.ok(REVIEW_RULES.indexOf(rule) > 6);
});

test("formatRationaleComment: what came from the review alone goes on a last line", () => {
  const r = { what: "w", why: "y", addresses: "a", requirement_ids: ["q1"] };
  assert.equal(
    formatRationaleComment("R1.2", r, { fromReview: ["Pluronic F-68", "0.1 %"] }),
    "[R1.2] w / y / a [q1]\nFrom the review, not checked against the source: Pluronic F-68; 0.1 %",
  );
  assert.equal(
    formatRationaleComment("R1.2", r, { lost: ["10 min"], fromReview: ["Pluronic"] }),
    "[R1.2] w / y / a [q1]\n사라지는 사실: 10 min\nFrom the review, not checked against the source: Pluronic",
  );
  assert.equal(formatRationaleComment("R1.2", r, { fromReview: [] }), "[R1.2] w / y / a [q1]");
});

function deps(ctx) {
  const d = {
    calls: [],
    edits: [],
    log: () => {},
    getLedger: async () => [],
    getReviewContext: async () => ({
      requirements: [],
      comment_text: "The antifoam, Pluronic, should be named. It was used at 0.1 %.",
      paragraphs: ["Methods", "Cultures were grown in sealed flasks with antifoam and Tween."],
      paragraph_ids: null,
      recent_changes: [],
      author_texts: [],
      ...ctx,
    }),
    call: async (name, args) => {
      d.calls.push([name, args]);
      return {
        ok: true,
        word_comment_id: 9,
        paragraph_index: 1,
        before: "b",
        after: "a",
        target_checked: true,
      };
    },
    onEdit: async ({ change }) => {
      d.edits.push(change);
      return { id: "R1.1-1" };
    },
    paneChecksTarget: () => true,
  };
  return d;
}
const R = {
  what: "Named the antifoam.",
  why: "The reviewer says it was Pluronic.",
  addresses: "the antifoam",
};

test("a word and a number the change has from the review alone are marked in the comment and on the change", async () => {
  const d = deps();
  const r = await applyTrackedEdit(
    d,
    {
      comment_ref: "R1.1",
      find: "with antifoam",
      replace: "with 0.1 % Pluronic antifoam",
      rationale: R,
    },
    { ref: "R1.1" },
  );
  assert.equal(r.ok, true, r.message);
  assert.deepEqual(r.from_review, ["0.1 %", "Pluronic"]);
  assert.match(
    d.calls[0][1].comment_text,
    /\nFrom the review, not checked against the source: 0\.1 %; Pluronic$/,
  );
  assert.deepEqual(d.edits[0].from_review, ["0.1 %", "Pluronic"]);
  assert.match(r.warning, /From the review, not checked against the source: 0\.1 %, Pluronic\./);
  // What the manuscript has itself is not marked; what nobody has stays "unsourced".
  const own = deps();
  const r2 = await applyTrackedEdit(
    own,
    {
      comment_ref: "R1.1",
      find: "with antifoam",
      replace: "with antifoam, Tween and Brij",
      rationale: R,
    },
    { ref: "R1.1" },
  );
  assert.deepEqual(r2.from_review, []);
  assert.deepEqual(r2.unsourced_terms, ["Brij"]);
  assert.equal(own.calls[0][1].comment_text.includes("From the review"), false);
});

test("the flag of a change card names what came from the review", () => {
  const flags = view.changeFlags({
    from_review: ["Pluronic"],
    lost_facts: [],
    unsourced_terms: [],
  });
  assert.deepEqual(flags, [{ kind: "unsourced", text: "From the review: Pluronic" }]);
});

async function service(more = {}) {
  const dir = await mkdtemp(join(tmpdir(), "review-o-"));
  const doc = join(dir, "paper.docx");
  const waiters = new Map();
  const svc = createReviewService({
    getContext: () => ({ activeDoc: doc, caps: ["target_hash", "author_insertion_guard"] }),
    callTool: async () => ({ ok: true }),
    pushTurn: () => {},
    sendEvent: (key, obj) => {
      if (obj.event === "job_done") waiters.get(obj.job_id)?.(obj);
    },
    log: () => {},
    env: { CC_OFFICE_WSL_PATHS: "0" },
    ...more,
  });
  const ask = (type, payload = {}) =>
    new Promise((resolve) =>
      svc.handlers[type](
        { type, request_id: "q", ...payload },
        (r) => (r.accepted ? waiters.set(r.job_id, (e) => resolve(e.result)) : resolve(r)),
        "k",
        "word",
      ),
    );
  return { dir, svc, ask, cleanup: () => rm(dir, { recursive: true, force: true }) };
}

test("a request for a supplementary file, with the requests listed: declared outside this file under its id", async () => {
  const t = await service();
  try {
    await t.ask("review_save_comments", {
      comments: [
        {
          ref: "R2.1",
          reviewer: 2,
          text: "State the unit. Provide the raw growth curves as a supplementary file.",
          requirements: ["State the unit", "Provide the raw growth curves as a supplementary file"],
        },
      ],
    });
    const r = await requirementOutside(
      t.svc.trackedEditDeps("k"),
      {
        comment_ref: "R2.1",
        requirement_id: "q2",
        why: "A supplementary file: to be made from the growth data.",
      },
      { ref: "R2.1" },
    );
    assert.equal(r.ok, true, r.message);
    const rev = await loadRevision(join(t.dir, "paper.revision.json"));
    assert.deepEqual(
      view.requirementStates(rev.comments[0]).map((x) => [x.id, x.state]),
      [
        ["q1", "open"],
        ["q2", "outside_this_file"],
      ],
    );
    assert.equal(rev.comments[0].questions.length, 0, "not a question");
  } finally {
    await t.cleanup();
  }
});

test("receipts of the split and the requirement labels name the model the request ran on", async () => {
  const t = await service({
    oneshotJob: async () => ({
      provider: "claude",
      model: "haiku",
      used: () => ({ ran: "claude-haiku-9-20260101" }),
      run: async ({ prompt }) => ({
        ok: true,
        json: {
          sentences: [...prompt.matchAll(/^\[(R\d+\.\d+) s(\d+)\] /gm)].map((m) => ({
            ref: m[1],
            s: Number(m[2]),
            type: "request",
          })),
        },
      }),
    }),
  });
  try {
    const s = await t.ask("review_split", {
      text: "Reviewer #1\nComment 1: Use Celsius.\nComment 2: Rename it.",
    });
    assert.equal(s.ok, true, s.error);
    const lines = (await readFile(join(t.dir, "paper.decisions.jsonl"), "utf8"))
      .trim()
      .split("\n")
      .map((l) => JSON.parse(l));
    const extract = lines.find((l) => l.kind === "requirement_extract");
    assert.equal(extract.model, "claude-haiku-9-20260101");
  } finally {
    await t.cleanup();
  }
});
