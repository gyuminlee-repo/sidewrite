// Invented bridge-report text. Exercise the saved review file, daemon payload
// and pane ownership guard together, without Word or a model session.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createReviewService } from "../daemon/review-service.mjs";
import { applyTrackedEdit } from "../daemon/review-tools.mjs";
import { RevisionSchema, saveRevision } from "../daemon/revision-store.mjs";
import { insertionAround } from "../taskpane/shared/tools-word.js";

const HEAD = "Crews checked the bridge. ";
const FIRST = "New bearings with ";
const SECOND = "thick steel plates were fitted.";
const ADDED = FIRST + SECOND;
const TEXT = HEAD + ADDED;
const RATIONALE = { what: "Remove repeated wording", why: "Already stated", addresses: "Edit" };
const TIME = "2026-01-01T00:00:00.000Z";

function change(text, extra = {}) {
  return {
    id: "R1.1-1",
    word_comment_id: "C1",
    paragraph_index: 1,
    paragraph_id: "U1",
    kind: "insert",
    find: HEAD,
    text,
    before: HEAD,
    after: TEXT,
    rationale: RATIONALE,
    created_at: TIME,
    state: "pending",
    ...extra,
  };
}

// Word's relations for Added ranges, including adjacency of split pieces.
function trackedParagraph(text, spans) {
  const relation = (a, b) => {
    if (a.end < b.start) return "Before";
    if (a.end === b.start) return "AdjacentBefore";
    if (a.start > b.end) return "After";
    if (a.start === b.end) return "AdjacentAfter";
    if (a.start === b.start && a.end === b.end) return "Equal";
    if (a.start <= b.start && a.end >= b.end)
      return a.start === b.start ? "ContainsStart" : a.end === b.end ? "ContainsEnd" : "Contains";
    if (a.start >= b.start && a.end <= b.end)
      return a.start === b.start ? "InsideStart" : a.end === b.end ? "InsideEnd" : "Inside";
    return a.start < b.start ? "OverlapsBefore" : "OverlapsAfter";
  };
  return {
    getTrackedChanges: () => ({
      load() {},
      items: spans.map(([start, end]) => ({
        type: "Added",
        text: text.slice(start, end),
        getRange: () => ({
          start,
          end,
          compareLocationWith: (other) => ({ value: relation({ start, end }, other) }),
        }),
      })),
    }),
  };
}

async function editFromFile(
  t,
  { comments, text = TEXT, spans, find = ADDED, reportedText = text },
) {
  const dir = await mkdtemp(join(tmpdir(), "own-insertion-"));
  t.after(() => rm(dir, { recursive: true, force: true }));
  await saveRevision(
    join(dir, "report.revision.json"),
    RevisionSchema.parse({
      version: 1,
      manuscript: "report.docx",
      updated_at: TIME,
      snapshot: {
        taken_at: "2026-01-02T00:00:00.000Z",
        paragraphs: ["Bridge report", text],
        ids: ["U0", "U1"],
        unreviewed: [],
        fields: false,
      },
      comments: comments.map((changes, i) => ({
        ref: `R1.${i + 1}`,
        reviewer: 1,
        text: "Remove repeated wording.",
        changes,
      })),
    }),
  );
  // A fresh service and turn: ownership must come from disk, not turn memory.
  const svc = createReviewService({
    getContext: () => ({
      activeDoc: join(dir, "report.docx"),
      caps: ["target_hash", "author_insertion_guard"],
    }),
    env: { CC_OFFICE_WSL_PATHS: "0" },
  });
  const prior = globalThis.Office;
  globalThis.Office = { context: { requirements: { isSetSupported: () => true } } };
  t.after(() => {
    globalThis.Office = prior;
  });
  const start = text.indexOf(find);
  assert.notEqual(start, -1);
  let around;
  let payload;
  const deps = {
    ...svc.trackedEditDeps("k"),
    call: async (name, args) => {
      payload = args;
      around = await insertionAround(
        { sync: async () => {} },
        trackedParagraph(reportedText, spans),
        { start, end: start + find.length },
        find,
        args.own_insertions,
      );
      if (around?.refuse)
        return {
          ok: false,
          code: "WholeAuthorInsertion",
          reason: around.refuse,
          applied: false,
          paragraph_index: 1,
        };
      return {
        ok: true,
        target_checked: true,
        word_comment_id: "C2",
        paragraph_index: 1,
        paragraph_id: "U1",
        before: text,
        after: text.replace(find, ""),
      };
    },
  };
  const result = await applyTrackedEdit(
    deps,
    { comment_ref: "R1.2", find, replace: "", rationale: RATIONALE },
    { ref: "R1.2" },
  );
  return { result, around, payload };
}

test("own history: another comment's accepted record still owns a pending insertion", async (t) => {
  const { result } = await editFromFile(t, {
    comments: [[change(ADDED, { state: "accepted" })], []],
    spans: [[HEAD.length, TEXT.length]],
  });
  assert.equal(result.ok, true, result.message);
});

test("own history: an earlier turn without a paragraph id survives later edits", async (t) => {
  const { result } = await editFromFile(t, {
    comments: [
      [],
      [
        change(FIRST, { paragraph_id: null, after: HEAD + FIRST }),
        change(SECOND, { id: "R1.2-2", rewrite_id: "rewrite-earlier-turn" }),
      ],
    ],
    spans: [[HEAD.length, TEXT.length]],
  });
  assert.equal(result.ok, true, result.message);
});

test("own history: a saved insertion survives a new session's paragraph ids", async (t) => {
  const { result } = await editFromFile(t, {
    comments: [[change(ADDED, { paragraph_id: "previous-session-id" })], []],
    spans: [[HEAD.length, TEXT.length]],
  });
  assert.equal(result.ok, true, result.message);
});

test("own history: Word splits a concatenation of recorded rewrite parts", async (t) => {
  const joint = HEAD.length + FIRST.length;
  for (const reverse of [false, true]) {
    const spans = [
      [HEAD.length, joint],
      [joint, joint + "thick ".length],
      [joint + "thick ".length, TEXT.length],
    ];
    const { result } = await editFromFile(t, {
      comments: [
        [change(FIRST, { rewrite_id: "rewrite-1" })],
        [change(SECOND, { id: "R1.2-1", rewrite_id: "rewrite-2" })],
      ],
      spans: reverse ? spans.reverse() : spans,
      find: "thick ",
    });
    assert.equal(result.ok, true, result.message);
  }
});

test("own history: Word merges adjacent tool insertions with whitespace drift", async (t) => {
  const { result } = await editFromFile(t, {
    comments: [
      [change("New  bearings\nwith ", { state: "accepted" })],
      [change(SECOND, { rewrite_id: "rewrite-2" })],
    ],
    spans: [[HEAD.length, TEXT.length]],
  });
  assert.equal(result.ok, true, result.message);
});

test("own history: unrecorded author text still refuses whole, takes and unplaced", async (t) => {
  const repeated = HEAD + "New thick thick plates.";
  for (const [text, find, why] of [
    [TEXT, ADDED, "whole"],
    [TEXT, TEXT, "takes"],
    [repeated, "thick thick", "unplaced"],
  ]) {
    const { result, around } = await editFromFile(t, {
      comments: [[], []],
      text,
      spans: [[HEAD.length, text.length]],
      find,
      reportedText: text.replace("thick thick", "thick\tthick"),
    });
    assert.deepEqual([result.ok, result.code], [false, "WholeAuthorInsertion"]);
    assert.equal(result.applied, false);
    assert.equal(around.refuse, why);
  }
});

test("author protection: a rejected record never owns the author's matching insertion", async (t) => {
  const { result } = await editFromFile(t, {
    comments: [[change(ADDED, { state: "rejected" })], []],
    spans: [[HEAD.length, TEXT.length]],
  });
  assert.deepEqual(
    [result.ok, result.code, result.applied],
    [false, "WholeAuthorInsertion", false],
  );
});

for (const [recorded, after, text, span] of [
  ["the ", "the ", "Together.", [4, 7]],
  ["a n", "Use a n mark.", "Banana.", [1, 3]],
]) {
  test(`author protection: ${JSON.stringify(recorded)} cannot match across word boundaries`, async (t) => {
    const { result } = await editFromFile(t, {
      comments: [[change(recorded, { after, paragraph_id: "old-session" })], []],
      text,
      spans: [span],
      find: text,
    });
    assert.deepEqual(
      [result.ok, result.code, result.applied],
      [false, "WholeAuthorInsertion", false],
    );
  });
}

test("author protection: an index-shifted short record needs its recorded neighbours", async (t) => {
  const text = "We crossed the bridge.";
  const start = text.indexOf("the ");
  const { result } = await editFromFile(t, {
    comments: [
      [change("the ", { after: "We opened the gate.", paragraph_id: "other-paragraph" })],
      [],
    ],
    text,
    spans: [[start, start + "the ".length]],
    find: "the ",
  });
  assert.deepEqual(
    [result.ok, result.code, result.applied],
    [false, "WholeAuthorInsertion", false],
  );
});

test("author protection: a changed paragraph with the same id needs recorded context", async (t) => {
  const text = "A different crew reported: " + ADDED;
  const { result } = await editFromFile(t, {
    comments: [[change(ADDED)], []],
    text,
    spans: [[text.indexOf(ADDED), text.length]],
  });
  assert.deepEqual(
    [result.ok, result.code, result.applied],
    [false, "WholeAuthorInsertion", false],
  );
});

test("author protection: a short context-free record cannot identify an insertion", async (t) => {
  const text = "We crossed the bridge.";
  const start = text.indexOf("the ");
  const { result } = await editFromFile(t, {
    comments: [[change("the ", { after: "the ", paragraph_id: null })], []],
    text,
    spans: [[start, start + "the ".length]],
    find: "the ",
  });
  assert.deepEqual(
    [result.ok, result.code, result.applied],
    [false, "WholeAuthorInsertion", false],
  );
});

test("author protection: a context-free word is not enough evidence even with the same id", async (t) => {
  const text = "Crews stacked steel.";
  const start = text.indexOf("steel");
  for (const paragraph_id of ["old-session", "U1"]) {
    const { result } = await editFromFile(t, {
      comments: [[change("steel", { after: "steel", paragraph_id })], []],
      text,
      spans: [[start, start + "steel".length]],
      find: "steel",
    });
    assert.deepEqual(
      [result.ok, result.code, result.applied],
      [false, "WholeAuthorInsertion", false],
    );
  }
});

test("author protection: a longer author insertion cannot repeat a recorded piece to gain ownership", async (t) => {
  const phrase = "New pads.";
  const text = HEAD + phrase + " " + phrase;
  const { result } = await editFromFile(t, {
    comments: [[change(phrase, { after: HEAD + phrase })], []],
    text,
    spans: [[HEAD.length, text.length]],
    find: phrase + " " + phrase,
  });
  assert.deepEqual(
    [result.ok, result.code, result.applied],
    [false, "WholeAuthorInsertion", false],
  );
});

test("author protection: extra author words around a recorded phrase stay protected", async (t) => {
  const longer = ADDED + " Additional details follow.";
  const text = HEAD + longer;
  const { result } = await editFromFile(t, {
    comments: [[change(ADDED)], []],
    text,
    spans: [[HEAD.length, text.length]],
    find: longer,
  });
  assert.deepEqual(
    [result.ok, result.code, result.applied],
    [false, "WholeAuthorInsertion", false],
  );
});

test("author protection: re-inserting an accepted phrase elsewhere is not the old tool insertion", async (t) => {
  const text = TEXT + " Another crew wrote: " + ADDED;
  const { result } = await editFromFile(t, {
    comments: [[change(ADDED, { state: "accepted" })], []],
    text,
    spans: [[text.lastIndexOf(ADDED), text.length]],
    find: "Another crew wrote: " + ADDED,
  });
  assert.deepEqual(
    [result.ok, result.code, result.applied],
    [false, "WholeAuthorInsertion", false],
  );
});

test("author protection: adjacent tool and author pieces still refuse takes", async (t) => {
  const joint = HEAD.length + FIRST.length;
  const { result, around } = await editFromFile(t, {
    comments: [[change(FIRST)], []],
    spans: [
      [HEAD.length, joint],
      [joint, TEXT.length],
    ],
  });
  assert.deepEqual(
    [result.ok, result.code, result.applied],
    [false, "WholeAuthorInsertion", false],
  );
  assert.equal(around.refuse, "takes");
});
