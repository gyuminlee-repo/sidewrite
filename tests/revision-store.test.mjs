import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, readdir, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  localDocPath,
  revisionPathFor,
  commentsBackupPathFor,
  emptyRevision,
  loadRevision,
  saveRevision,
  withRevision,
  replaceComments,
  setCommentStatus,
  addChange,
  applyStates,
  addBlocked,
  takeBlocked,
  claimBlocked,
  markBlockedUnknown,
  addLedgerItem,
  updateLedgerItem,
  removeLedgerItem,
  RevisionSchema,
  setRequirements,
  setSnapshot,
  addExternal,
  dismissExternal,
  EXTERNAL_KEEP,
} from "../daemon/revision-store.mjs";

const WSL = { CC_OFFICE_WSL_PATHS: "1" };
const OFF = { CC_OFFICE_WSL_PATHS: "0" };
const posixOnly = { skip: process.platform === "win32" };

const rationale = { what: "단위를 바꿈", why: "리뷰어 요청", addresses: "R1.1 단위 표기" };
const change = (over = {}) => ({
  word_comment_id: "11",
  paragraph_index: 9,
  paragraph_id: null,
  kind: "replace",
  find: "30 degrees",
  text: "30 \u00b0C",
  before: "held at 30 degrees",
  after: "held at 30 \u00b0C",
  rationale,
  ...over,
});
const base = () =>
  replaceComments(emptyRevision("/m/paper.docx", new Date("2026-01-01T00:00:00Z")), [
    { ref: "R1.1", reviewer: 1, text: "Use Celsius." },
    { ref: "R1.2", reviewer: 1, text: "Rename the extract." },
  ]);

test("localDocPath: WSL drive path, file URL, POSIX, and unaddressable forms", posixOnly, () => {
  assert.equal(localDocPath("D:\\m\\paper v2.docx", WSL), "/mnt/d/m/paper v2.docx");
  assert.equal(localDocPath("file:///D:/m/a%20b.docx", WSL), "/mnt/d/m/a b.docx");
  assert.equal(localDocPath("file:///home/u/a.docx", WSL), "/home/u/a.docx");
  assert.equal(localDocPath("/home/u/a.docx", WSL), "/home/u/a.docx");
  assert.equal(localDocPath("https://contoso.sharepoint.com/s/a.docx", WSL), null);
  assert.equal(localDocPath("\\\\server\\share\\a.docx", WSL), null);
  assert.equal(localDocPath(null, WSL), null);
  assert.equal(localDocPath("   ", OFF), null);
});

test("sidecar names drop the Word extension and keep the folder", () => {
  assert.equal(revisionPathFor("/mnt/d/m/paper v2.docx"), "/mnt/d/m/paper v2.revision.json");
  assert.equal(commentsBackupPathFor("/mnt/d/m/X.DOCM"), "/mnt/d/m/X.comments_backup.json");
});

test("loadRevision returns an empty revision when the file does not exist", async () => {
  const dir = await mkdtemp(join(tmpdir(), "rev-"));
  try {
    const r = await loadRevision(join(dir, "x.revision.json"), "/m/x.docx");
    assert.equal(r.version, 1);
    assert.equal(r.manuscript, "/m/x.docx");
    assert.deepEqual([r.comments, r.ledger], [[], []]);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("saveRevision writes valid JSON atomically and leaves no temp file", async () => {
  const dir = await mkdtemp(join(tmpdir(), "rev-"));
  const p = join(dir, "x.revision.json");
  try {
    await saveRevision(p, base());
    const parsed = RevisionSchema.parse(JSON.parse(await readFile(p, "utf8")));
    assert.deepEqual(
      parsed.comments.map((c) => c.ref),
      ["R1.1", "R1.2"],
    );
    assert.deepEqual(await readdir(dir), ["x.revision.json"]);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("saveRevision rejects an invalid revision and keeps the old file", async () => {
  const dir = await mkdtemp(join(tmpdir(), "rev-"));
  const p = join(dir, "x.revision.json");
  try {
    await saveRevision(p, base());
    const bad = base();
    bad.comments[0].ref = "comment one";
    await assert.rejects(saveRevision(p, bad));
    const kept = JSON.parse(await readFile(p, "utf8"));
    assert.equal(kept.comments[0].ref, "R1.1");
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("loadRevision refuses a corrupt file instead of starting over", async () => {
  const dir = await mkdtemp(join(tmpdir(), "rev-"));
  const p = join(dir, "x.revision.json");
  try {
    await writeFile(p, "{not json");
    await assert.rejects(loadRevision(p, "/m/x.docx"));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("withRevision serializes concurrent updates (no lost change)", async () => {
  const dir = await mkdtemp(join(tmpdir(), "rev-"));
  const p = join(dir, "x.revision.json");
  try {
    await saveRevision(p, base());
    await Promise.all(
      Array.from({ length: 10 }, (_, i) =>
        withRevision(
          p,
          "/m/x.docx",
          (rev) => addChange(rev, "R1.1", change({ find: `f${i}` })).revision,
        ),
      ),
    );
    const r = await loadRevision(p, "/m/x.docx");
    assert.equal(r.comments[0].changes.length, 10);
    assert.deepEqual(new Set(r.comments[0].changes.map((c) => c.id)).size, 10);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("replaceComments keeps changes of surviving refs and refuses to drop edited ones", () => {
  let rev = addChange(base(), "R1.1", change()).revision;
  rev = replaceComments(rev, [
    { ref: "R1.1", reviewer: 1, text: "Use Celsius (edited)." },
    { ref: "R2.1", reviewer: 2, text: "New." },
  ]);
  assert.equal(rev.comments[0].changes.length, 1);
  assert.equal(rev.comments[0].text, "Use Celsius (edited).");
  assert.throws(
    () => replaceComments(rev, [{ ref: "R2.1", reviewer: 2, text: "x" }]),
    /R1\.1 has 1/,
  );
  assert.throws(
    () =>
      replaceComments(rev, [
        { ref: "R1.1", reviewer: 1, text: "a" },
        { ref: "R1.1", reviewer: 1, text: "b" },
      ]),
    /Duplicate/,
  );
});

test("addChange numbers per comment and moves pending to in_progress", () => {
  let rev = base();
  let r = addChange(rev, "R1.1", change(), new Date("2026-02-02T00:00:00Z"));
  assert.equal(r.change.id, "R1.1-1");
  assert.equal(r.change.state, "pending");
  assert.equal(r.change.created_at, "2026-02-02T00:00:00.000Z");
  r = addChange(r.revision, "R1.1", change());
  assert.equal(r.change.id, "R1.1-2");
  assert.equal(r.revision.comments[0].status, "in_progress");
  assert.equal(base().comments[0].status, "pending", "input revision is not mutated");
  assert.throws(() => addChange(rev, "R9.9", change()), /No reviewer comment R9\.9/);
  assert.throws(() => addChange(rev, "R1.1", change({ kind: "move" })));
});

test("setCommentStatus and applyStates", () => {
  let rev = addChange(base(), "R1.2", change()).revision;
  rev = setCommentStatus(rev, "R1.2", "done");
  assert.equal(rev.comments[1].status, "done");
  rev = applyStates(
    rev,
    [
      { ref: "R1.2", change_id: "R1.2-1", state: "rejected", paragraph_index: 13 },
      { ref: "R1.2", change_id: "missing", state: "accepted", paragraph_index: 1 },
    ],
    new Date("2026-03-03T00:00:00Z"),
  );
  const ch = rev.comments[1].changes[0];
  assert.deepEqual(
    [ch.state, ch.paragraph_index, ch.checked_at],
    ["rejected", 13, "2026-03-03T00:00:00.000Z"],
  );
});

test("blocked edits are stored and taken back by id", () => {
  let r = addBlocked(base(), "R1.1", {
    args: { find: "x", replace: "novel x" },
    blocked_by: ["L2"],
  });
  assert.equal(r.blocked.id, "R1.1-b1");
  const t = takeBlocked(r.revision, "R1.1", "R1.1-b1");
  assert.deepEqual(t.blocked.blocked_by, ["L2"]);
  assert.equal(t.revision.comments[0].blocked.length, 0);
  assert.throws(() => takeBlocked(t.revision, "R1.1", "R1.1-b1"), /No blocked edit/);
});

test("ledger items get L-ids and are validated", () => {
  let r = addLedgerItem(base(), {
    kind: "instruction",
    text: "in vitro 는 하이픈 없이 쓴다",
    check: { type: "regex", pattern: "\\bin-(?:vitro|vivo|situ)\\b", flags: "i" },
  });
  assert.equal(r.item.id, "L1");
  assert.equal(r.item.check.expect, "absent");
  assert.equal(r.item.status, "unchecked");
  r = addLedgerItem(r.revision, { kind: "forbid", text: "novel 금지", check: { type: "llm" } });
  assert.equal(r.item.id, "L2");
  let rev = updateLedgerItem(r.revision, "L1", { status: "violations", violation_count: 2 });
  assert.equal(rev.ledger[0].violation_count, 2);
  assert.throws(() => updateLedgerItem(rev, "L1", { kind: "rule" }));
  assert.throws(() =>
    addLedgerItem(rev, {
      kind: "forbid",
      text: "x",
      check: { type: "regex", pattern: "x", flags: "g" },
    }),
  );
  rev = removeLedgerItem(rev, "L1");
  assert.deepEqual(
    rev.ledger.map((x) => x.id),
    ["L2"],
  );
  assert.equal(
    addLedgerItem(rev, { kind: "forbid", text: "y", check: { type: "llm" } }).item.id,
    "L3",
  );
  assert.throws(() => removeLedgerItem(rev, "L9"), /No ledger item/);
});

test("a revision file from before L1-L10 still loads with empty defaults", () => {
  const old = {
    version: 1,
    manuscript: "/m/paper.docx",
    updated_at: "2026-01-01T00:00:00.000Z",
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "Use Celsius.",
        changes: [{ ...change(), id: "R1.1-1", created_at: "x" }],
      },
    ],
    ledger: [{ id: "L1", kind: "forbid", text: "novel", check: { type: "llm" } }],
  };
  const rev = RevisionSchema.parse(old);
  assert.deepEqual([rev.snapshot, rev.external], [null, []]);
  assert.deepEqual(rev.comments[0].requirements, []);
  const ch = rev.comments[0].changes[0];
  assert.deepEqual(
    [ch.lost_facts, ch.unsourced_terms, ch.narrowed_from, ch.rationale.requirement_ids],
    [[], [], null, []],
  );
  assert.deepEqual([rev.ledger[0].violations, rev.ledger[0].checked_hashes], [[], null]);
});

test("L5 requirements get q-ids, keep them for the same text, and ride on the comment list", () => {
  let rev = setRequirements(base(), "R1.1", ["state the unit", " state the unit ", "use Celsius"]);
  assert.deepEqual(rev.comments[0].requirements, [
    { id: "q1", text: "state the unit" },
    { id: "q2", text: "use Celsius" },
  ]);
  rev = setRequirements(rev, "R1.1", ["use Celsius", "cite the method"]);
  assert.deepEqual(
    rev.comments[0].requirements.map((r) => `${r.id}:${r.text}`),
    ["q2:use Celsius", "q3:cite the method"],
  );
  // The comment list keeps requirements unless it brings its own.
  rev = replaceComments(rev, [{ ref: "R1.1", reviewer: 1, text: "Use Celsius." }]);
  assert.equal(rev.comments[0].requirements.length, 2);
  rev = replaceComments(rev, [
    { ref: "R1.1", reviewer: 1, text: "x", requirements: ["cite the method"] },
  ]);
  assert.deepEqual(rev.comments[0].requirements, [{ id: "q3", text: "cite the method" }]);
  assert.throws(() => setRequirements(rev, "R9.9", ["x"]), /No reviewer comment/);
});

test("L5 a requirement id that a recorded change cites is never given out again", async () => {
  const { unaddressedRequirements } = await import("../taskpane/shared/review-view.js");
  let rev = setRequirements(base(), "R1.1", ["state the unit"]);
  rev = addChange(
    rev,
    "R1.1",
    change({ rationale: { ...rationale, requirement_ids: ["q1"] } }),
  ).revision;
  // The author clears the list, then lists a new request.
  rev = setRequirements(rev, "R1.1", []);
  rev = setRequirements(rev, "R1.1", ["cite the method"]);
  assert.deepEqual(rev.comments[0].requirements, [{ id: "q2", text: "cite the method" }]);
  assert.deepEqual(
    unaddressedRequirements(rev.comments[0]).map((r) => r.id),
    ["q2"],
  );
  // Same through the comment list (Save comments) path.
  rev = setRequirements(rev, "R1.1", []);
  rev = replaceComments(rev, [
    { ref: "R1.1", reviewer: 1, text: "x", requirements: ["use Celsius"] },
    { ref: "R1.2", reviewer: 1, text: "Rename the extract." },
  ]);
  // q2 was removed with no change citing it; it is still not given out again.
  assert.deepEqual(rev.comments[0].requirements, [{ id: "q3", text: "use Celsius" }]);
  assert.deepEqual(
    unaddressedRequirements(rev.comments[0]).map((r) => r.id),
    ["q3"],
  );
});

test("L10 a changed rule drops the last check; a status patch keeps it", () => {
  let rev = addLedgerItem(base(), {
    kind: "forbid",
    text: "novel",
    check: { type: "regex", pattern: "novel", flags: "i" },
  }).revision;
  const checked = {
    status: "violations",
    violation_count: 1,
    violations: [{ paragraph_index: 3, h: "0000abcd", start: 2, match: "novel" }],
    checked_hashes: ["0000abcd"],
    checked_at: "2026-01-01T00:00:00.000Z",
  };
  rev = updateLedgerItem(rev, "L1", checked);
  assert.deepEqual(updateLedgerItem(rev, "L1", { scope: "manuscript" }).ledger[0].checked_hashes, [
    "0000abcd",
  ]);
  const edited = updateLedgerItem(rev, "L1", {
    check: { type: "regex", pattern: "novel(?:ty)?", flags: "i", expect: "absent" },
  }).ledger[0];
  assert.deepEqual(
    [edited.status, edited.violation_count, edited.violations, edited.checked_hashes],
    ["unchecked", 0, [], null],
  );
  assert.equal(updateLedgerItem(rev, "L1", { text: "novelty" }).ledger[0].checked_hashes, null);
});

test("L4 snapshot, outside changes with X-ids, cap and dismiss", () => {
  let rev = setSnapshot(base(), ["a", "b"], new Date("2026-04-04T00:00:00Z"));
  assert.deepEqual(rev.snapshot, { taken_at: "2026-04-04T00:00:00.000Z", paragraphs: ["a", "b"] });
  const item = (i) => ({
    kind: "modified",
    paragraph_index: i,
    old_index: i,
    before: "x 3 mM",
    after: "x",
    lost_facts: ["3 mM"],
  });
  let r = addExternal(rev, [item(0), item(1)]);
  assert.deepEqual(
    r.added.map((x) => x.id),
    ["X1", "X2"],
  );
  rev = RevisionSchema.parse(r.revision);
  rev = dismissExternal(rev, "X1");
  assert.deepEqual(
    rev.external.map((x) => x.id),
    ["X2"],
  );
  assert.throws(() => dismissExternal(rev, "X1"), /No outside change X1/);
  const many = Array.from({ length: EXTERNAL_KEEP + 5 }, (_, i) => item(i));
  const capped = addExternal(rev, many).revision.external;
  assert.equal(capped.length, EXTERNAL_KEEP);
  assert.equal(capped.at(-1).id, `X${EXTERNAL_KEEP + 7}`);
});

test("editor comment E.1 (reviewer 0) and preamble R1.0 are valid comments", () => {
  const rev = replaceComments(emptyRevision("/m/x.docx"), [
    { ref: "E.1", reviewer: 0, text: "Shorten the abstract." },
    { ref: "R1.0", reviewer: 1, text: "Overall fine." },
  ]);
  assert.deepEqual(
    RevisionSchema.parse(rev).comments.map((c) => [c.ref, c.reviewer]),
    [
      ["E.1", 0],
      ["R1.0", 1],
    ],
  );
});

test("a change applied without its comment is stored with a null id and the error code", () => {
  const r = addChange(
    base(),
    "R1.1",
    change({ word_comment_id: null, comment_missing: "CommentAnchorMissing" }),
  );
  assert.deepEqual(
    [r.change.word_comment_id, r.change.comment_missing],
    [null, "CommentAnchorMissing"],
  );
  assert.equal(addChange(base(), "R1.1", change()).change.comment_missing, null);
});

test("an unsure ledger check is stored with its paragraphs; a changed rule clears them", () => {
  let rev = addLedgerItem(base(), {
    kind: "instruction",
    text: "t",
    check: { type: "llm" },
  }).revision;
  assert.deepEqual(rev.ledger[0].unsure, []);
  rev = updateLedgerItem(rev, "L1", {
    status: "unsure",
    unsure: [{ paragraph_index: 5, h: "0000abcd", reason: "cannot tell" }],
  });
  assert.equal(rev.ledger[0].status, "unsure");
  assert.equal(rev.ledger[0].unsure.length, 1);
  const edited = updateLedgerItem(rev, "L1", { text: "t2" }).ledger[0];
  assert.deepEqual([edited.status, edited.unsure], ["unchecked", []]);
});

// Codex review 9: a blocked edit that cites q2 may still be approved; if q2
// were given to a new requirement, the approval would answer the wrong one.
test("L5 requirement ids that blocked, applying and unanswered edits cite are never given out again", () => {
  const cites = (ids) => ({
    find: "a",
    replace: "b",
    rationale: { ...rationale, requirement_ids: ids },
  });
  let rev = setRequirements(base(), "R1.1", ["one", "two", "three", "four"]);
  rev = addBlocked(rev, "R1.1", { args: cites(["q2"]), blocked_by: ["L1"] }).revision;
  rev = addBlocked(rev, "R1.1", { args: cites(["q3"]), blocked_by: ["L1"] }).revision;
  rev = addBlocked(rev, "R1.1", { args: cites([]), blocked_by: ["L1"] }).revision;
  rev = claimBlocked(rev, "R1.1", "R1.1-b2").revision; // applying
  rev = markBlockedUnknown(rev, "R1.1", "R1.1-b3", {
    tool_call_id: null,
    unknown_at: "2026-01-01T00:00:00Z",
    comment_text: "x",
    change: { rationale: { ...rationale, requirement_ids: ["q4"] } },
  });
  assert.deepEqual(
    rev.comments[0].blocked.map((b) => b.state ?? "blocked"),
    ["blocked", "applying", "outcome_unknown"],
  );
  rev = setRequirements(rev, "R1.1", ["one"]);
  rev = setRequirements(rev, "R1.1", ["one", "new request"]);
  assert.deepEqual(rev.comments[0].requirements, [
    { id: "q1", text: "one" },
    { id: "q5", text: "new request" },
  ]);
  // Same through the comment list (Save comments) path.
  rev = setRequirements(rev, "R1.1", ["one"]);
  rev = replaceComments(rev, [
    { ref: "R1.1", reviewer: 1, text: "x", requirements: ["one", "other request"] },
  ]);
  assert.deepEqual(
    rev.comments[0].requirements.map((r) => r.id),
    ["q1", "q6"], // q5 was listed, cited by none, and removed: still never given out again
  );
});

// Codex review (P2): ids are never given out twice. A counter in the file
// outlives the item, so a stale reference (a blocked edit, an open pane, a
// running turn that was told q2) can never mean a newer item.
test("ledger ids are not reused after the highest item is deleted", () => {
  const forbid = { kind: "forbid", text: "x", check: { type: "llm" } };
  let r = addLedgerItem(base(), forbid);
  assert.deepEqual([r.item.id, r.revision.ledger_seq], ["L1", 1]);
  let rev = removeLedgerItem(r.revision, "L1");
  r = addLedgerItem(rev, forbid);
  assert.deepEqual([r.item.id, r.revision.ledger_seq], ["L2", 2]);
  rev = removeLedgerItem(r.revision, "L2");
  assert.equal(RevisionSchema.parse(rev).ledger_seq, 2); // survives a save
  assert.equal(addLedgerItem(rev, forbid).item.id, "L3");
});

test("requirement ids and blocked-edit ids are not reused after the highest one is removed", () => {
  let rev = setRequirements(base(), "R1.1", ["one", "two"]);
  rev = setRequirements(rev, "R1.1", ["one"]); // q2 removed, cited by nothing
  rev = setRequirements(rev, "R1.1", ["one", "other"]);
  assert.deepEqual(
    rev.comments[0].requirements.map((q) => q.id),
    ["q1", "q3"],
  );
  // Through the comment list (Save comments) path too, and a text that
  // stays keeps its id.
  rev = replaceComments(rev, [{ ref: "R1.1", reviewer: 1, text: "x", requirements: ["other"] }]);
  rev = replaceComments(rev, [
    { ref: "R1.1", reviewer: 1, text: "x", requirements: ["other", "third"] },
  ]);
  assert.deepEqual(
    rev.comments[0].requirements.map((q) => q.id),
    ["q3", "q4"],
  );
  // A comment list saved without requirements keeps the counter.
  rev = replaceComments(rev, [{ ref: "R1.1", reviewer: 1, text: "y" }]);
  assert.equal(rev.comments[0].requirement_seq, 4);

  const args = { find: "a", replace: "b", rationale };
  let b = addBlocked(rev, "R1.1", { args, blocked_by: ["L1"] });
  assert.equal(b.blocked.id, "R1.1-b1");
  rev = takeBlocked(b.revision, "R1.1", "R1.1-b1").revision;
  b = addBlocked(rev, "R1.1", { args, blocked_by: ["L1"] });
  assert.equal(b.blocked.id, "R1.1-b2");
});

test("a revision file without the counters gets them on load from every id it names", async () => {
  const dir = await mkdtemp(join(tmpdir(), "rev-seq-"));
  try {
    const path = join(dir, "paper.revision.json");
    const cites = (ids) => ({
      find: "a",
      replace: "b",
      rationale: { ...rationale, requirement_ids: ids },
    });
    const old = {
      version: 1,
      manuscript: "/m/paper.docx",
      updated_at: "2026-01-01T00:00:00.000Z",
      comments: [
        {
          ref: "R1.1",
          reviewer: 1,
          text: "t",
          requirements: [{ id: "q1", text: "one" }],
          changes: [
            {
              ...change(),
              id: "R1.1-1",
              created_at: "2026-01-01T00:00:00.000Z",
              override_of: ["L4"],
            },
          ],
          // L3 and L7 are gone from the ledger; q5 from the requirements.
          blocked: [
            {
              id: "R1.1-b3",
              args: cites(["q5"]),
              blocked_by: ["L3"],
              created_at: "2026-01-01T00:00:00.000Z",
            },
            {
              id: "R1.1-b4",
              args: cites([]),
              blocked_by: ["L1"],
              created_at: "2026-01-01T00:00:00.000Z",
              state: "outcome_unknown",
              attempt: {
                tool_call_id: null,
                unknown_at: "2026-01-01T00:00:00.000Z",
                comment_text: "x",
                change: { override_of: ["L7"] },
              },
            },
          ],
        },
        { ref: "R1.2", reviewer: 1, text: "u" },
      ],
      ledger: [{ id: "L1", kind: "forbid", text: "x", check: { type: "llm" } }],
    };
    await writeFile(path, JSON.stringify(old), "utf8");
    const rev = await loadRevision(path, "/m/paper.docx");
    assert.equal(rev.ledger_seq, 7);
    assert.deepEqual(
      rev.comments.map((c) => [c.requirement_seq, c.blocked_seq]),
      [
        [5, 4],
        [0, 0],
      ],
    );
    const forbid = { kind: "forbid", text: "y", check: { type: "llm" } };
    assert.equal(addLedgerItem(rev, forbid).item.id, "L8");
    assert.equal(setRequirements(rev, "R1.1", ["one", "new"]).comments[0].requirements[1].id, "q6");
    assert.equal(
      addBlocked(rev, "R1.1", { args: cites([]), blocked_by: ["L1"] }).blocked.id,
      "R1.1-b5",
    );
    // Saved, the counters stay although the references may go later.
    const saved = await saveRevision(path, rev);
    assert.equal(JSON.parse(await readFile(path, "utf8")).ledger_seq, 7);
    assert.equal(saved.comments[0].requirement_seq, 5);
    // Not yet migrated objects (no load in between) give the same ids.
    assert.equal(addLedgerItem(RevisionSchema.parse(old), forbid).item.id, "L8");
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
