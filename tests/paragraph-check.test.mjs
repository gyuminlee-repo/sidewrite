// "Check revised paragraphs": each paragraph that holds two or more changes
// is read as it will stand when all of them are accepted, by the model that
// writes the edits, for a contradiction inside it, changes of different
// comments that work against each other, and a sentence that takes what the
// reviewer said as its source. All texts here are made up.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  paragraphsToCheck,
  checkPrompt,
  readFindings,
  checkParagraphs,
  CHECK_SYSTEM,
} from "../daemon/paragraph-check.mjs";
import { createReviewService } from "../daemon/review-service.mjs";
import { loadRevision } from "../daemon/revision-store.mjs";
import * as view from "../taskpane/shared/review-view.js";

const T0 = "2026-01-01T00:00:00.000Z"; // the status refresh
const change = (id, paragraph_index, find, text, extra = {}) => ({
  id,
  word_comment_id: "w",
  paragraph_index,
  paragraph_id: null,
  kind: text === "" ? "delete" : "replace",
  find,
  text,
  before: "",
  after: "",
  rationale: { what: `what ${id}`, why: "y", addresses: "a", requirement_ids: [] },
  narrowed_from: null,
  state: "pending",
  created_at: "2025-12-31T00:00:00.000Z", // before the refresh: already in its text
  ...extra,
});
const REV = {
  comments: [
    {
      ref: "R1.1",
      reviewer: 1,
      text: "The claim that gene A is essential is not supported.",
      changes: [
        change("R1.1-1", 2, "is essential for", "was associated with"),
        change("R1.1-2", 4, "30 C", "30 °C"),
        change("R1.1-3", 5, "a", "b"),
        change("R1.1-4", 5, "c", "d"),
        change("R1.1-5", 5, "e", "f", { state: "rejected" }),
        // Recorded after the refresh: put into the text here.
        change("R1.1-6", 5, "grew slowly", "grew at half the rate", {
          created_at: "2026-01-02T00:00:00.000Z",
        }),
      ],
    },
    {
      ref: "R2.3",
      reviewer: 2,
      text: "State that the deletion was not complemented.",
      changes: [
        change("R2.3-1", 2, "", " The deletion was not complemented.", {
          kind: "insert",
          find: "growth.",
        }),
        change("R2.3-2", 9, "x", "y", { state: "unlocated" }),
      ],
    },
  ],
  snapshot: {
    taken_at: T0,
    paragraphs: [
      "Title",
      "Methods",
      "Gene A was associated with growth. The deletion was not complemented.",
      "Results",
      "Cultures were held at 30 °C.",
      "The mutant grew slowly in all media.",
    ],
    ids: null,
  },
};

test("paragraphsToCheck: paragraphs with two or more changes, as they read with every change in", () => {
  const items = paragraphsToCheck(REV);
  assert.deepEqual(
    items.map((p) => [p.index, p.id, p.changes.map((c) => c.id), p.refs, p.many]),
    [
      [2, "p2", ["R1.1-1", "R2.3-1"], ["R1.1", "R2.3"], false],
      // Three changes that count (the rejected one does not): read as a whole.
      [5, "p5", ["R1.1-3", "R1.1-4", "R1.1-6"], ["R1.1"], true],
    ],
  );
  assert.equal(
    items[0].text,
    "Gene A was associated with growth. The deletion was not complemented.",
  );
  assert.equal(items[1].text, "The mutant grew at half the rate in all media.");
  assert.deepEqual(items[0].changes[0], {
    id: "R1.1-1",
    ref: "R1.1",
    kind: "replace",
    find: "is essential for",
    text: "was associated with",
    what: "what R1.1-1",
  });
  // The task pane id of a change finds its paragraph when the index moved.
  const moved = structuredClone(REV);
  moved.snapshot.ids = ["a", "b", "c", "d", "e", "f"];
  moved.comments[0].changes[0].paragraph_id = "c";
  moved.comments[0].changes[0].paragraph_index = 0;
  moved.comments[1].changes[0].paragraph_id = "c";
  assert.deepEqual(
    paragraphsToCheck(moved).map((p) => [p.index, p.id]),
    [
      [2, "c"],
      [5, "f"],
    ],
  );
  // No status refresh yet, or nothing changed twice: nothing to read.
  assert.deepEqual(paragraphsToCheck({ comments: REV.comments, snapshot: null }), []);
  assert.deepEqual(paragraphsToCheck({ comments: [], snapshot: REV.snapshot }), []);
});

test("checkPrompt: the paragraph in a fence, its changes, the comments they answer, and the three things to look for", () => {
  const items = paragraphsToCheck(REV);
  const prompt = checkPrompt(items, REV.comments, "F1");
  assert.ok(
    prompt.includes(
      [
        "Paragraph 2 (as it reads with every change accepted):",
        "<<<F1",
        "Gene A was associated with growth. The deletion was not complemented.",
        ">>>F1",
        "Changes in paragraph 2:",
        '- R1.1-1 (for R1.1) replace "is essential for" -> "was associated with": what R1.1-1',
        '- R2.3-1 (for R2.3) insert after "growth." -> " The deletion was not complemented.": what R2.3-1',
      ].join("\n"),
    ),
    prompt,
  );
  assert.ok(
    prompt.includes(
      [
        "Reviewer comments the changes answer (what a reviewer says is not a fact of the paper):",
        "R1.1:",
        "<<<F1",
        "The claim that gene A is essential is not supported.",
        ">>>F1",
      ].join("\n"),
    ),
  );
  assert.match(prompt, /"contradiction": the paragraph states something and denies it/);
  assert.match(
    prompt,
    /"conflict": changes made for different reviewer comments work against each other/,
  );
  assert.match(
    prompt,
    /"review_as_source": a sentence gives as a fact of the paper something only the reviewer said/,
  );
  assert.match(
    prompt,
    /Return \{"findings":\[\{"p":<paragraph number>,"kind":"<kind>","detail":"<one sentence>","change_ids":\["<id>"\]\}\]\}/,
  );
  assert.match(CHECK_SYSTEM, /Reply with JSON only/);
});

test("readFindings: only paragraphs that were asked about, known kinds, and change ids of that paragraph", () => {
  const items = paragraphsToCheck(REV);
  const r = readFindings(
    {
      findings: [
        {
          p: 2,
          kind: "contradiction",
          detail: " Says essential and not essential. ",
          change_ids: ["R1.1-1", "R9.9-1"],
        },
        { p: 7, kind: "contradiction", detail: "x", change_ids: [] },
        { p: 5, kind: "style", detail: "x", change_ids: [] },
        { p: 5, kind: "conflict", detail: "", change_ids: [] },
      ],
    },
    items,
  );
  assert.deepEqual(r.findings, [
    {
      paragraph_index: 2,
      paragraph_id: "p2",
      kind: "contradiction",
      detail: "Says essential and not essential.",
      change_ids: ["R1.1-1"],
    },
  ]);
  assert.equal(r.warnings.length, 3);
  assert.deepEqual(readFindings({ findings: [] }, items), { findings: [], warnings: [] });
  assert.deepEqual(readFindings({}, items).warnings, ["model reply had no findings list"]);
});

test("checkParagraphs: one request per batch; a failed request is a warning; nothing to check asks nothing", async () => {
  const asked = [];
  const oneshot = async (req) => {
    asked.push(req);
    return {
      ok: true,
      json: {
        findings: [
          { p: 5, kind: "contradiction", detail: "Slow and fast.", change_ids: ["R1.1-6"] },
        ],
      },
    };
  };
  const r = await checkParagraphs(REV, { oneshot, fence: "F1" });
  assert.equal(asked.length, 1);
  assert.equal(asked[0].system, CHECK_SYSTEM);
  assert.equal(asked[0].prompt, checkPrompt(paragraphsToCheck(REV), REV.comments, "F1"));
  assert.deepEqual([r.asked, r.paragraphs.length, r.findings.length, r.warnings], [true, 2, 1, []]);
  // A small batch size: one request per paragraph.
  asked.length = 0;
  await checkParagraphs(REV, { oneshot, fence: "F1", batchChars: 10 });
  assert.equal(asked.length, 2);
  const failed = await checkParagraphs(REV, {
    oneshot: async () => ({ ok: false, error: "timed out" }),
    fence: "F1",
  });
  assert.deepEqual(failed.findings, []);
  assert.deepEqual(failed.warnings, ["check failed for paragraphs 2, 5: timed out"]);
  const none = await checkParagraphs({ comments: [], snapshot: REV.snapshot }, { oneshot });
  assert.deepEqual([none.asked, none.paragraphs, none.findings], [false, [], []]);
});

test("service: review_check_paragraphs asks the model of the edits, saves what it found, logs counts only", async () => {
  const dir = await mkdtemp(join(tmpdir(), "review-pc-"));
  const doc = join(dir, "paper.docx");
  const lines = [];
  const events = [];
  const waiters = new Map();
  const jobs = [];
  const shots = [];
  const svc = createReviewService({
    getContext: () => ({ activeDoc: doc, caps: ["target_hash"] }),
    callTool: async () => ({
      ok: true,
      word_comment_id: 5,
      paragraph_index: 1,
      before: "b",
      after: "a",
      target_checked: true,
    }),
    pushTurn: () => {},
    sendEvent: (key, obj) => {
      events.push(obj);
      if (obj.event === "job_done") waiters.get(obj.job_id)?.(obj);
    },
    log: (s) => lines.push(s),
    env: { CC_OFFICE_WSL_PATHS: "0" },
    getModel: () => "claude-opus-9",
    oneshotJob: async (key, options) => {
      jobs.push(options);
      return {
        model: options?.claudeModel ?? "haiku",
        run: async (req) => {
          shots.push(req);
          return {
            ok: true,
            json: {
              findings: [
                { p: 1, kind: "contradiction", detail: "Both at once.", change_ids: ["R1.1-1"] },
              ],
            },
          };
        },
      };
    },
  });
  const ask = (type, payload = {}) =>
    new Promise((resolve) => {
      svc.handlers[type](
        { type, request_id: "q", ...payload },
        (r) => {
          if (!r.accepted) return resolve(r);
          waiters.set(r.job_id, (e) => resolve(e.result));
        },
        "k",
        "word",
      );
    });
  try {
    await ask("review_save_comments", {
      comments: [{ ref: "R1.1", reviewer: 1, text: "Two things about the first paragraph." }],
    });
    const para = (text, index) => ({
      index,
      id: `p${index}`,
      text,
      reviewed_text: text,
      w_ins: 0,
      w_del: 0,
    });
    await ask("review_refresh", {
      snapshot: {
        paragraphs: ["Title", "The culture grew fast and was held warm."].map(para),
        comments: [],
      },
    });
    // Nothing changed twice yet: no model call.
    const empty = await ask("review_check_paragraphs");
    assert.equal(empty.ok, true, empty.error);
    assert.deepEqual([empty.consistency.paragraphs, empty.consistency.findings], [[], []]);
    assert.equal(shots.length, 0);

    const deps = {
      ...svc.trackedEditDeps("k"),
      call: async () => ({
        ok: true,
        word_comment_id: 5,
        paragraph_index: 1,
        before: "b",
        after: "a",
        target_checked: true,
      }),
    };
    const { applyTrackedEdit } = await import("../daemon/review-tools.mjs");
    const R = { what: "w", why: "y", addresses: "a" };
    for (const [find, replace] of [
      ["grew fast", "grew slowly"],
      ["held warm", "kept warm"],
    ]) {
      const r = await applyTrackedEdit(
        deps,
        { comment_ref: "R1.1", find, replace, rationale: R },
        { ref: "R1.1" },
      );
      assert.equal(r.ok, true, r.message);
    }
    const r = await ask("review_check_paragraphs");
    assert.equal(r.ok, true, r.error);
    // The model of the edits, through the router's role for it.
    assert.deepEqual(jobs.at(-1), { role: "turn", claudeModel: "claude-opus-9" });
    assert.equal(shots.length, 1);
    assert.ok(shots[0].timeoutMs >= 300_000);
    assert.ok(shots[0].prompt.includes("The culture grew slowly and was kept warm."));
    assert.equal(r.consistency.model, "claude-opus-9");
    assert.deepEqual(r.consistency.paragraphs, [
      { index: 1, id: "p1", change_ids: ["R1.1-1", "R1.1-2"], many: false },
    ]);
    assert.deepEqual(r.consistency.findings, [
      {
        paragraph_index: 1,
        paragraph_id: "p1",
        kind: "contradiction",
        detail: "Both at once.",
        change_ids: ["R1.1-1"],
      },
    ]);
    const rev = await loadRevision(join(dir, "paper.revision.json"));
    assert.deepEqual(rev.consistency.findings, r.consistency.findings);
    assert.deepEqual(r.revision.consistency, rev.consistency);
    assert.ok(
      lines.includes("[review] paragraph check paragraphs=1 findings=1 model=claude-opus-9"),
    );
    assert.equal(
      lines.some((l) => l.includes("Both at once") || l.includes("grew slowly")),
      false,
    );
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("view: the message after a check and the lines under the button", () => {
  const consistency = {
    checked_at: "2026-01-02T03:04:05.000Z",
    model: "claude-opus-9",
    paragraphs: [
      { index: 2, id: "p2", change_ids: ["R1.1-1", "R2.3-1"], many: false },
      { index: 5, id: "id5", change_ids: ["R1.1-3", "R1.1-4", "R1.1-6"], many: true },
    ],
    findings: [
      {
        paragraph_index: 2,
        paragraph_id: "p2",
        kind: "conflict",
        detail: "One change removes the claim the other one explains.",
        change_ids: ["R1.1-1", "R2.3-1"],
      },
      {
        paragraph_index: 5,
        paragraph_id: "id5",
        kind: "review_as_source",
        detail: "Cites the review.",
        change_ids: [],
      },
    ],
    warnings: ["check failed for paragraphs 9: timed out"],
  };
  assert.equal(
    view.consistencyMessage(consistency, "Claude · Opus 9"),
    "Checked 2 revised paragraphs with Claude · Opus 9: 2 problems. check failed for paragraphs 9: timed out",
  );
  assert.equal(
    view.consistencyMessage(
      {
        ...consistency,
        findings: [],
        warnings: [],
        paragraphs: consistency.paragraphs.slice(0, 1),
      },
      null,
    ),
    "Checked 1 revised paragraph: no problem found.",
  );
  assert.equal(
    view.consistencyMessage({ ...consistency, paragraphs: [], findings: [], warnings: [] }, null),
    "No paragraph holds two or more changes: nothing to check.",
  );
  assert.deepEqual(view.consistencyLines(consistency), [
    {
      paragraph_index: 2,
      text: "p2: changes of different comments work against each other. One change removes the claim the other one explains. (R1.1-1, R2.3-1)",
    },
    {
      paragraph_index: 5,
      text: "id5: a sentence takes the review as its source. Cites the review.",
    },
    {
      paragraph_index: 5,
      text: "id5 holds 3 changes (R1.1-3, R1.1-4, R1.1-6): read it as a whole and rewrite it if it no longer reads as one paragraph.",
    },
  ]);
  assert.deepEqual(view.consistencyLines(null), []);
});
