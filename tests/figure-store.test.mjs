// Figure changes in revision.json: a file from before the figure lane, the
// comment list that is saved again, and ids that are never given out twice.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, copyFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  RevisionSchema,
  loadRevision,
  saveRevision,
  replaceComments,
  addFigureChange,
  updateFigureChange,
  setFigureState,
} from "../daemon/revision-store.mjs";

const FIXTURE = new URL("./fixtures/revision_before_figures.json", import.meta.url);
const HASH = (c) => c.repeat(64);
const fileRef = (path, c) => ({
  path,
  sha256: HASH(c),
  bytes: 10,
  mtime: "2026-01-01T00:00:00.000Z",
});
const change = (extra = {}) => ({
  figure: "Figure 2",
  kind: "revised",
  requirement_ids: ["q2"],
  summary: "The control curve was added.",
  before: { source: fileRef("figures/fig2_before.svg", "a") },
  after: { source: fileRef("figures/fig2_after.svg", "b") },
  ...extra,
});
const AT = new Date("2026-02-01T00:00:00.000Z");
const old = async () => RevisionSchema.parse(JSON.parse(await readFile(FIXTURE, "utf8")));
const asList = (rev) =>
  rev.comments.map((c) => ({ ref: c.ref, reviewer: c.reviewer, text: c.text }));

test("a revision.json from before the figure lane loads, and saving it changes nothing else", async () => {
  const raw = JSON.parse(await readFile(FIXTURE, "utf8"));
  assert.equal("figure_lane" in raw, false, "the fixture is the old format");
  assert.equal("figure_changes" in raw.comments[0], false);
  const dir = await mkdtemp(join(tmpdir(), "figure-store-"));
  try {
    const path = join(dir, "paper.revision.json");
    await copyFile(FIXTURE, path);
    const rev = await loadRevision(path, "paper.docx");
    assert.equal(rev.version, 1);
    assert.equal(rev.figure_lane, null);
    for (const c of rev.comments) {
      assert.deepEqual(c.figure_changes, []);
      assert.equal(c.figure_seq, 0);
    }
    const saved = await saveRevision(path, rev, new Date(raw.updated_at));
    for (const c of saved.comments) {
      delete c.figure_changes;
      delete c.figure_seq;
      if (c.response) delete c.response.figure_ids;
    }
    delete saved.figure_lane;
    assert.deepEqual(saved, raw);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("a figure change is recorded as proposed with an id of its own, and its comment is in progress", async () => {
  const rev = RevisionSchema.parse({
    ...(await old()),
    comments: [{ ref: "R2.1", reviewer: 2, text: "x", requirements: [{ id: "q1", text: "y" }] }],
  });
  const r = addFigureChange(rev, "R2.1", change({ requirement_ids: ["q1"] }), AT);
  assert.equal(r.figure.id, "R2.1-f1");
  assert.equal(r.figure.state, "proposed");
  assert.equal(r.figure.adopted, null);
  assert.equal(r.figure.made_by, "author");
  assert.equal(r.figure.created_at, AT.toISOString());
  assert.equal(r.revision.comments[0].status, "in_progress");
  assert.equal(rev.comments[0].figure_changes.length, 0, "the input is not changed");
  // The result is what the file holds.
  assert.deepEqual(RevisionSchema.parse(r.revision).comments[0].figure_changes, [r.figure]);
});

test("a figure change may name only requests of its comment; a new figure has no old file", async () => {
  const rev = await old();
  assert.throws(
    () => addFigureChange(rev, "R1.1", change({ requirement_ids: ["q9"] })),
    /No requirement q9 under R1\.1/,
  );
  assert.throws(() => addFigureChange(rev, "R9.9", change()), /No reviewer comment R9\.9/);
  const r = addFigureChange(rev, "R1.1", change({ kind: "new", before: null }), AT);
  assert.equal(r.figure.before, null);
  assert.throws(
    () => addFigureChange(rev, "R1.1", change({ after: { source: fileRef("x.svg", "z") } })),
    /sha256/,
  );
});

test("replaceComments keeps the figure changes and the counter of a comment that stays", async () => {
  let rev = addFigureChange(await old(), "R1.1", change(), AT).revision;
  rev = addFigureChange(rev, "R1.1", change({ figure: "Figure 3" }), AT).revision;
  const before = rev.comments[0].figure_changes;
  const next = replaceComments(
    rev,
    asList(rev).map((c) => ({ ...c, text: `${c.text} (edited)` })),
  );
  assert.deepEqual(next.comments[0].figure_changes, before);
  assert.equal(next.comments[0].figure_seq, 2);
  assert.equal(next.comments[1].figure_seq, 0);
  assert.deepEqual(next.comments[1].figure_changes, []);
  RevisionSchema.parse(next);
});

test("a comment with a figure change cannot be dropped from the list", async () => {
  // R1.2 has no text change: only its figure change holds it.
  const rev = addFigureChange(await old(), "R1.2", change({ requirement_ids: [] }), AT).revision;
  assert.equal(rev.comments[1].changes.length, 0);
  const without = asList(rev).filter((c) => c.ref !== "R1.2");
  assert.throws(() => replaceComments(rev, without), /R1\.2 has 1 figure change\(s\); keep it/);
  // Without the figure change the same list is accepted.
  assert.equal(replaceComments(await old(), without).comments.length, 1);
});

test("the id of a figure change is never given out twice", async () => {
  let rev = addFigureChange(await old(), "R1.1", change(), AT).revision;
  rev = addFigureChange(rev, "R1.1", change(), AT).revision;
  assert.deepEqual(
    rev.comments[0].figure_changes.map((f) => f.id),
    ["R1.1-f1", "R1.1-f2"],
  );
  // The last one is taken out of the file by hand: its id stays used.
  rev.comments[0].figure_changes.pop();
  assert.equal(addFigureChange(rev, "R1.1", change(), AT).figure.id, "R1.1-f3");
  // Also after the comment list was saved again and the file was read again.
  const dir = await mkdtemp(join(tmpdir(), "figure-store-"));
  try {
    const path = join(dir, "paper.revision.json");
    await saveRevision(path, replaceComments(rev, asList(rev)));
    const loaded = await loadRevision(path, "paper.docx");
    assert.equal(loaded.comments[0].figure_seq, 2);
    assert.equal(addFigureChange(loaded, "R1.1", change(), AT).figure.id, "R1.1-f3");
    // A file that lost the counter gets it from the ids it still names.
    const raw = JSON.parse(await readFile(path, "utf8"));
    raw.comments[0].figure_seq = null;
    raw.comments[0].figure_changes[0].id = "R1.1-f7";
    await saveRevision(path, RevisionSchema.parse(raw));
    assert.equal((await loadRevision(path, "paper.docx")).comments[0].figure_seq, 7);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("states: proposed to adopted or rejected, and back by reopening only", async () => {
  const start = addFigureChange(await old(), "R1.1", change(), AT).revision;
  const copy = { path: "figures/submission/fig2_after.svg", sha256: HASH("b"), at: "x" };
  assert.throws(
    () => setFigureState(start, "R1.1", "R1.1-f1", "adopted"),
    /adopting needs the copy/,
  );
  const adopted = setFigureState(start, "R1.1", "R1.1-f1", "adopted", { adopted: copy }, AT);
  const f = adopted.comments[0].figure_changes[0];
  assert.equal(f.state, "adopted");
  assert.deepEqual(f.adopted, copy);
  assert.equal(f.decided_at, AT.toISOString());
  assert.throws(
    () => setFigureState(adopted, "R1.1", "R1.1-f1", "rejected"),
    /is adopted; it cannot become rejected/,
  );
  assert.throws(
    () => setFigureState(adopted, "R1.1", "R1.1-f1", "adopted", { adopted: copy }),
    /is adopted; it cannot become adopted/,
  );
  const reopened = setFigureState(adopted, "R1.1", "R1.1-f1", "proposed");
  assert.equal(reopened.comments[0].figure_changes[0].state, "proposed");
  assert.equal(reopened.comments[0].figure_changes[0].adopted, null);
  assert.equal(reopened.comments[0].figure_changes[0].decided_at, null);
  assert.throws(
    () => setFigureState(start, "R1.1", "R1.1-f1", "proposed"),
    /is proposed; it cannot become proposed/,
  );
  const rejected = setFigureState(start, "R1.1", "R1.1-f1", "rejected", {}, AT);
  assert.equal(rejected.comments[0].figure_changes[0].state, "rejected");
  assert.throws(() => setFigureState(start, "R1.1", "R1.1-f9", "rejected"), /No figure change/);
});

test("the words of a figure change are changed only while it is proposed; files and state never by an update", async () => {
  const start = addFigureChange(await old(), "R1.1", change(), AT).revision;
  const edited = updateFigureChange(start, "R1.1", "R1.1-f1", {
    figure: "Figure 2A",
    summary: "New words.",
    requirement_ids: ["q1", "q2"],
    // Not taken:
    state: "adopted",
    after: { source: fileRef("evil.svg", "c") },
    id: "R1.1-f9",
  });
  const f = edited.comments[0].figure_changes[0];
  assert.equal(f.figure, "Figure 2A");
  assert.equal(f.summary, "New words.");
  assert.deepEqual(f.requirement_ids, ["q1", "q2"]);
  assert.equal(f.state, "proposed");
  assert.equal(f.after.source.path, "figures/fig2_after.svg");
  assert.equal(f.id, "R1.1-f1");
  assert.throws(
    () => updateFigureChange(start, "R1.1", "R1.1-f1", { requirement_ids: ["q5"] }),
    /No requirement q5/,
  );
  const copy = { path: "figures/submission/fig2_after.svg", sha256: HASH("b"), at: "x" };
  const adopted = setFigureState(start, "R1.1", "R1.1-f1", "adopted", { adopted: copy });
  assert.throws(
    () => updateFigureChange(adopted, "R1.1", "R1.1-f1", { summary: "Other words." }),
    /is adopted; reopen it before changing it/,
  );
  // What the daemon computed is stored in any state.
  const render = { path: "figures/.renders/x.png", of: HASH("b"), dpi: 150, width: 8, height: 5 };
  const withRender = updateFigureChange(adopted, "R1.1", "R1.1-f1", {
    renders: { before: null, after: render },
  });
  assert.deepEqual(withRender.comments[0].figure_changes[0].renders.after, render);
});
