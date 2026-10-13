// A supplementary Word document in the revision of a manuscript, step A1:
// the two optional fields of the review file, the copy of what the manuscript
// owns (mirrorFromMain), join, leave, healing of a lost end, and the
// requests that are refused in a supplement. Invented files only.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  RevisionSchema,
  emptyRevision,
  loadRevision,
  saveRevision,
  withRevision,
  writeJsonAtomic,
  replaceComments,
  setRequirements,
  addChange,
  addBlocked,
  addQuestion,
  setRequirementOutside,
  setSameTerm,
  setCitations,
  setResponse,
  setDeclined,
  setSnapshot,
  setProfile,
  addLedgerItem,
  updateLedgerItem,
  mirrorFromMain,
  setPartOf,
  leaveMain,
  addPart,
  removePart,
} from "../daemon/revision-store.mjs";
import { linkCandidates, storedPartPath, resolvePart } from "../daemon/revision-parts.mjs";
import { createReviewService } from "../daemon/review-service.mjs";
import { createResponseService } from "../daemon/response-service.mjs";
import { createAuxService } from "../daemon/aux-service.mjs";

const ENV = { CC_OFFICE_WSL_PATHS: "0" };
const AT = new Date("2026-10-01T09:00:00.000Z");
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
async function until(fn, what, ms = 3000) {
  const end = Date.now() + ms;
  for (;;) {
    const v = await fn();
    if (v) return v;
    if (Date.now() > end) throw new Error(`timed out waiting for ${what}`);
    await sleep(10);
  }
}

const CHANGE = {
  word_comment_id: "9",
  paragraph_index: 1,
  kind: "replace",
  find: "30 degrees",
  text: "30 degrees Celsius",
  before: "Incubation was at 30 degrees.",
  after: "Incubation was at 30 degrees Celsius.",
  rationale: { what: "unit", why: "asked", addresses: "R1.1", requirement_ids: ["q2"] },
};
const NOVEL = {
  kind: "forbid",
  text: "do not write novel",
  check: { type: "regex", pattern: String.raw`\bnovel\b`, flags: "i" },
};

// The review file of a manuscript: three comments, R1.1 with the requests
// q2 and q3 (q1 was removed), four rules.
function mainRevision() {
  let rev = replaceComments(emptyRevision("/work/paper/paper.docx", AT), [
    {
      ref: "R1.1",
      reviewer: 1,
      text: "State the incubation temperature.",
      section: "Major comments",
      requirements: ["state the temperature", "give the unit"],
      named_terms: ["temperature"],
    },
    { ref: "R1.2", reviewer: 1, text: "Add the control strain to Table S3." },
    { ref: "E.1", reviewer: 0, text: "Follow the journal style." },
  ]);
  rev = setRequirements(rev, "R1.1", ["give the unit", "name the incubator"]);
  rev = addLedgerItem(rev, NOVEL).revision; // L1
  rev = addLedgerItem(rev, {
    kind: "instruction",
    text: "Abstract under 250 words",
    check: { type: "count", target: "section_words", section: "Abstract", max: 250 },
  }).revision; // L2: a count rule of the manuscript
  rev = addLedgerItem(rev, {
    kind: "instruction",
    text: "Name every strain",
    scope: "supplement",
    check: { type: "llm" },
  }).revision; // L3
  rev = addLedgerItem(rev, {
    kind: "instruction",
    text: "Thank the reviewer once",
    scope: "response",
    check: { type: "llm" },
  }).revision; // L4
  rev = setProfile(rev, { journal_profile: null, overrides: { response_lang: "en" } });
  rev = setResponse(rev, "R1.1", { text: "We added the unit.", source: "author" }, AT);
  rev = setDeclined(rev, "E.1", { reason: "Already in style." });
  return RevisionSchema.parse(rev);
}

// ---- the review file ----------------------------------------------------------

test("a review file from before the link fields loads and saves byte for byte", async () => {
  const fixture = new URL("./fixtures/revision_before_parts.json", import.meta.url);
  // The fixture was saved by the daemon of the commit before the fields
  // (38a5b3d). The repository formatter lays its arrays out anew, so the
  // bytes compared are those of the daemon's own layout of the same data
  // (writeJsonAtomic); the order of the keys is the one of that save.
  const raw = JSON.parse(await readFile(fixture, "utf8"));
  const text = JSON.stringify(raw, null, 2) + "\n";
  assert.equal("part_of" in raw || "parts" in raw, false, "the fixture is the old format");
  assert.equal("doc_state" in raw, false, "and from before the state of the document");
  const dir = await mkdtemp(join(tmpdir(), "parts-old-"));
  try {
    const path = join(dir, "paper.revision.json");
    await writeFile(path, text);
    const loaded = await loadRevision(path, "x");
    assert.equal("part_of" in loaded || "parts" in loaded || "doc_state" in loaded, false);
    await saveRevision(path, loaded, new Date(raw.updated_at));
    assert.equal(await readFile(path, "utf8"), text);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("a review file with the link fields read by the earlier schema loads, and the fields drop", () => {
  const Earlier = RevisionSchema.omit({ part_of: true, parts: true });
  const main = addPart(mainRevision(), { doc: "SI.docx", label: "SI", joined_at: "t" });
  const supp = setPartOf(emptyRevision("/work/paper/SI.docx", AT), {
    manuscript: "paper.docx",
    label: "SI",
    joined_at: "t",
  });
  for (const rev of [main, supp]) {
    const old = Earlier.parse(JSON.parse(JSON.stringify(RevisionSchema.parse(rev))));
    assert.equal("part_of" in old || "parts" in old, false);
    assert.equal(old.comments.length, rev.comments.length);
  }
});

test("the link operations: one entry per document, and the key goes with the last one", () => {
  const base = mainRevision();
  const one = addPart(base, { doc: "SI.docx", label: "SI", joined_at: "t1" });
  assert.deepEqual(one.parts, [{ doc: "SI.docx", label: "SI", joined_at: "t1" }]);
  assert.equal(addPart(one, { doc: "SI.docx", label: "SI", joined_at: "t2" }), null);
  const two = addPart(one, { doc: "extra/S2.docx", label: "S2", joined_at: "t3" });
  assert.equal(two.parts.length, 2);
  assert.equal(removePart(two, "nope.docx"), null);
  assert.equal("parts" in removePart(one, "SI.docx"), false);
  assert.equal("parts" in base, false, "the input is not changed");
  const part = setPartOf(base, { manuscript: "paper.docx", label: "SI", joined_at: "t" });
  assert.equal("part_of" in setPartOf(part, null), false);
});

test("paths of the other document: relative with / under the folder, absolute outside it", () => {
  const dir = join(tmpdir(), "paper");
  assert.equal(storedPartPath(dir, join(dir, "SI.docx")), "SI.docx");
  assert.equal(storedPartPath(dir, join(dir, "suppl", "SI.docx")), "suppl/SI.docx");
  const outside = join(tmpdir(), "other", "SI.docx");
  assert.equal(storedPartPath(dir, outside), outside);
  assert.equal(resolvePart(dir, "suppl/SI.docx"), join(dir, "suppl", "SI.docx"));
  assert.equal(resolvePart(dir, outside), outside);
});

// ---- mirrorFromMain -------------------------------------------------------------

test("mirrorFromMain: request ids are those of the manuscript, also after one was removed", () => {
  const main = mainRevision();
  const mainReqs = main.comments[0].requirements;
  assert.deepEqual(
    mainReqs.map((r) => r.id),
    ["q2", "q3"],
    "q1 was removed and q3 added in the manuscript",
  );
  const supp = mirrorFromMain(emptyRevision("/work/paper/SI.docx", AT), main);
  assert.deepEqual(supp.comments[0].requirements, mainReqs);
  assert.equal(supp.comments[0].requirement_seq, 3);
  assert.deepEqual(
    supp.comments.map((c) => [c.ref, c.reviewer, c.text, c.section, c.named_terms]),
    main.comments.map((c) => [c.ref, c.reviewer, c.text, c.section, c.named_terms]),
  );
  // What replaceComments would have made of the same list: numbered by text.
  const renumbered = replaceComments(emptyRevision("x", AT), [
    { ref: "R1.1", reviewer: 1, text: "t", requirements: mainReqs },
  ]);
  assert.deepEqual(
    renumbered.comments[0].requirements.map((r) => r.id),
    ["q1", "q2"],
  );
});

test("mirrorFromMain: what is the supplement's own stays, what the manuscript owns is replaced", () => {
  const main = mainRevision();
  let supp = mirrorFromMain(emptyRevision("/work/paper/SI.docx", AT), main);
  supp = addChange(supp, "R1.1", CHANGE, AT).revision;
  supp = addBlocked(supp, "R1.1", { args: { find: "a" }, blocked_by: ["L1"] }, AT).revision;
  supp = addQuestion(supp, "R1.2", { question: "Which strain?" }, AT).revision;
  supp = setRequirementOutside(supp, "R1.1", { id: "q3", why: "In the manuscript." }, AT);
  supp = setSameTerm(supp, "R1.1", ["p4"]);
  supp = setCitations(supp, "R1.1", [{ change_ids: ["R1.1-1"], quote: "30 degrees Celsius" }]);
  supp = setSnapshot(supp, ["Supplementary Methods", "Incubation was at 30 degrees."], AT);
  supp = { ...supp, external: [], comments: supp.comments.map((c) => ({ ...c })) };
  supp.comments[0].status = "done";
  // The manuscript changes: a new text, a comment more, another profile.
  let next = replaceComments(main, [
    { ref: "R1.1", reviewer: 1, text: "State the temperature and its unit." },
    { ref: "R1.2", reviewer: 1, text: "Add the control strain to Table S3." },
    { ref: "E.1", reviewer: 0, text: "Follow the journal style." },
    { ref: "R2.1", reviewer: 2, text: "Shorten the abstract." },
  ]);
  next = setProfile(next, { overrides: { response_lang: "de" } });
  const out = mirrorFromMain(supp, next);
  const c = out.comments[0];
  assert.equal(c.text, "State the temperature and its unit.");
  assert.equal(out.comments.length, 4);
  assert.equal(out.profile_overrides.response_lang, "de");
  // Its own:
  assert.equal(c.changes.length, 1);
  assert.equal(c.changes[0].id, "R1.1-1");
  assert.equal(c.blocked.length, 1);
  assert.equal(c.blocked_seq, 1);
  assert.equal(c.status, "done");
  assert.deepEqual(
    c.requirements_outside.map((n) => n.id),
    ["q3"],
  );
  assert.deepEqual(c.same_term_at, ["p4"]);
  assert.equal(c.response.citations[0].quote, "30 degrees Celsius");
  assert.equal(out.comments[1].questions.length, 1);
  assert.deepEqual(out.snapshot, supp.snapshot);
  // The answer and the decline of the manuscript are not copied.
  assert.equal(c.response.text, "");
  assert.equal(out.comments[2].declined, null);
});

test("mirrorFromMain: a comment the manuscript no longer lists stays only with work of its own", () => {
  const main = mainRevision();
  let supp = mirrorFromMain(emptyRevision("/work/paper/SI.docx", AT), main);
  supp = addChange(
    supp,
    "R1.2",
    { ...CHANGE, rationale: { ...CHANGE.rationale, requirement_ids: [] } },
    AT,
  ).revision;
  const fewer = { ...main, comments: main.comments.filter((c) => c.ref === "R1.1") };
  const out = mirrorFromMain(supp, fewer);
  assert.deepEqual(
    out.comments.map((c) => c.ref),
    ["R1.1", "R1.2"],
    "R1.2 holds a change, E.1 holds nothing",
  );
  assert.equal(out.comments[1].changes.length, 1);
});

test("mirrorFromMain: ledger scope and results", () => {
  const main = mainRevision();
  const supp = mirrorFromMain(emptyRevision("/work/paper/SI.docx", AT), main);
  assert.deepEqual(
    supp.ledger.map((r) => [r.id, r.scope, r.check.type]),
    [
      ["L1", "manuscript", "regex"],
      ["L3", "manuscript", "llm"], // a rule for the supplement applies here
      ["L4", "response", "llm"],
    ],
    "the count rule of the manuscript (L2) is left out",
  );
  assert.equal(supp.ledger_seq, 4);
  assert.deepEqual(
    supp.ledger.map((r) => [r.kind, r.text, r.source]),
    main.ledger.filter((r) => r.id !== "L2").map((r) => [r.kind, r.text, r.source]),
  );
  // A count rule for the supplement is kept, as a manuscript rule of this file.
  const counted = addLedgerItem(main, {
    kind: "instruction",
    text: "Supplementary Methods under 900 words",
    scope: "supplement",
    check: { type: "count", target: "section_words", section: "Supplementary Methods", max: 900 },
  }).revision;
  assert.deepEqual(
    mirrorFromMain(supp, counted)
      .ledger.filter((r) => r.check.type === "count")
      .map((r) => [r.id, r.scope]),
    [["L5", "manuscript"]],
  );
  // Results of the supplement's own checks: kept while the rule is the same.
  const result = {
    status: "violations",
    violation_count: 2,
    checked_at: "t",
    checked_hashes: ["h"],
  };
  let checked = updateLedgerItem(supp, "L1", result);
  checked = updateLedgerItem(checked, "L3", result);
  const edited = updateLedgerItem(main, "L1", { text: "never write novel" });
  const out = mirrorFromMain(checked, edited);
  const [l1, l3] = out.ledger;
  assert.deepEqual(
    [l1.text, l1.status, l1.violation_count, l1.checked_hashes],
    ["never write novel", "unchecked", 0, null],
    "the rule changed: its result is reset",
  );
  assert.deepEqual(
    [l3.status, l3.violation_count, l3.checked_hashes],
    ["violations", 2, ["h"]],
    "the rule is the same: its result stays",
  );
  // The results of the manuscript's own checks are not copied.
  const mainChecked = updateLedgerItem(main, "L1", result);
  assert.equal(mirrorFromMain(emptyRevision("x", AT), mainChecked).ledger[0].status, "unchecked");
});

test("mirrorFromMain: a second run changes nothing, also across a save", async () => {
  const main = mainRevision();
  let supp = mirrorFromMain(emptyRevision("/work/paper/SI.docx", AT), main);
  supp = addChange(supp, "R1.1", CHANGE, AT).revision;
  const once = mirrorFromMain(supp, main);
  assert.deepEqual(mirrorFromMain(once, main), once);
  const dir = await mkdtemp(join(tmpdir(), "parts-twice-"));
  try {
    const path = join(dir, "SI.revision.json");
    await saveRevision(path, once, AT);
    const loaded = await loadRevision(path, "x");
    assert.equal(JSON.stringify(mirrorFromMain(loaded, main)), JSON.stringify(loaded));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("leaveMain: the copy goes, a comment with a change of its own stays", () => {
  const main = mainRevision();
  let supp = setPartOf(emptyRevision("/work/paper/SI.docx", AT), {
    manuscript: "paper.docx",
    label: "SI",
    joined_at: "t",
  });
  supp = addChange(mirrorFromMain(supp, main), "R1.1", CHANGE, AT).revision;
  const left = leaveMain(supp);
  assert.equal("part_of" in left, false);
  assert.deepEqual(
    left.comments.map((c) => c.ref),
    ["R1.1"],
  );
  assert.deepEqual([left.ledger, left.profile_overrides], [[], {}]);
});

// ---- the services ---------------------------------------------------------------

const COMMENTS = [
  {
    ref: "R1.1",
    reviewer: 1,
    text: "State the incubation temperature.",
    requirements: ["state the temperature"],
  },
  { ref: "R1.2", reviewer: 1, text: "Add the control strain to Table S3." },
  { ref: "R2.1", reviewer: 2, text: "Do not call the method novel." },
];

// One folder, one daemon, a pane for each document: the pane key is the file
// name. paper.docx holds a review of three comments and a forbid rule.
async function world({ review = true } = {}) {
  const dir = await mkdtemp(join(tmpdir(), "parts-svc-"));
  const docOf = (name) => join(dir, name);
  const events = [];
  const waiters = new Map();
  const open = new Set(["paper.docx", "SI.docx"]);
  const deps = {
    getContext: (key) => ({
      activeDoc: docOf(key),
      caps: ["target_hash", "author_insertion_guard"],
    }),
    sendEvent: (key, obj) => {
      events.push({ key, ...obj });
      if (obj.event === "job_done") waiters.get(obj.job_id)?.(obj);
    },
    log: () => {},
    env: ENV,
  };
  const pushed = [];
  const handlers = {
    ...createReviewService({
      ...deps,
      callTool: async () => ({}),
      pushTurn: (...a) => pushed.push(a),
      listPanes: () => [...open].map((key) => ({ key, host: "word", activeDoc: docOf(key) })),
    }).handlers,
    ...createResponseService(deps).handlers,
    ...createAuxService(deps).handlers,
  };
  // The answer of a request; of a job, its job_done result.
  const ask = (key, type, payload = {}) =>
    new Promise((resolve) => {
      const reply = (r) => {
        if (!r.accepted) return resolve(r);
        waiters.set(r.job_id, (e) => resolve(e.result));
      };
      handlers[type]({ type, request_id: "q", ...payload }, reply, key, "word");
    });
  const pathOf = (name) => join(dir, name.replace(/\.docx$/, ".revision.json"));
  const file = async (name) => JSON.parse(await readFile(pathOf(name), "utf8"));
  const raw = (name) => readFile(pathOf(name), "utf8");
  const w = { dir, docOf, pathOf, ask, file, raw, events, pushed, open };
  w.cleanup = () => rm(dir, { recursive: true, force: true });
  if (review) {
    const saved = await ask("paper.docx", "review_save_comments", { comments: COMMENTS });
    assert.equal(saved.ok, true, saved.error);
    assert.equal((await ask("paper.docx", "ledger_add", { item: NOVEL })).ok, true);
  }
  return w;
}
const join1 = (w, key = "SI.docx", main = "paper.docx") =>
  w.ask(key, "review_link_join", { main: w.docOf(main) });
// A file as an earlier daemon would save it: without the link fields.
async function saveAsEarlier(w, name) {
  const Earlier = RevisionSchema.omit({ part_of: true, parts: true });
  await writeJsonAtomic(w.pathOf(name), Earlier.parse(await w.file(name)));
}

test("candidates: the review files of the folder that hold comments, never a supplement or the file itself", async () => {
  const w = await world();
  try {
    // A second paper with a review, a file without comments, a broken file.
    await w.ask("other.docx", "review_save_comments", { comments: COMMENTS.slice(0, 1) });
    await w.ask("empty.docx", "review_load");
    await writeFile(join(w.dir, "broken.revision.json"), "{ not json");
    const before = await w.ask("SI.docx", "review_link_candidates");
    assert.deepEqual(
      before.candidates.map((c) => [c.name, c.comments, c.doc]),
      [
        ["other.docx", 1, w.docOf("other.docx")],
        ["paper.docx", 3, w.docOf("paper.docx")],
      ],
    );
    assert.equal((await join1(w)).ok, true);
    // A supplement is no manuscript to join; the manuscript does not list itself.
    const forOther = await w.ask("other.docx", "review_link_candidates");
    assert.deepEqual(
      forOther.candidates.map((c) => c.name),
      ["paper.docx"],
    );
    assert.deepEqual(
      (await linkCandidates(w.docOf("paper.docx"))).map((c) => c.name),
      ["other.docx"],
    );
  } finally {
    await w.cleanup();
  }
});

test("join: both files hold the link, the supplement holds the copy, and nothing joins by itself", async () => {
  const w = await world();
  try {
    // Before: a load of the supplement finds candidates and links nothing.
    const loaded = await w.ask("SI.docx", "review_load");
    assert.deepEqual([loaded.linked.part_of, loaded.linked.parts], [null, []]);
    assert.equal(loaded.revision.comments.length, 0);
    const r = await join1(w);
    assert.equal(r.ok, true, r.error);
    assert.deepEqual(
      [r.linked.part_of.name, r.linked.part_of.doc, r.linked.part_of.label],
      ["paper.docx", w.docOf("paper.docx"), "SI"],
    );
    const supp = await w.file("SI.docx");
    const main = await w.file("paper.docx");
    assert.equal(supp.part_of.manuscript, "paper.docx", "relative to its folder");
    assert.deepEqual(
      main.parts.map((p) => [p.doc, p.label]),
      [["SI.docx", "SI"]],
    );
    assert.equal(main.parts[0].joined_at, supp.part_of.joined_at);
    assert.deepEqual(
      supp.comments.map((c) => c.ref),
      ["R1.1", "R1.2", "R2.1"],
    );
    assert.deepEqual(supp.comments[0].requirements, main.comments[0].requirements);
    assert.deepEqual(
      supp.ledger.map((x) => [x.id, x.text]),
      [["L1", NOVEL.text]],
    );
    // The manuscript's pane is told what is linked.
    const mainLoad = await w.ask("paper.docx", "review_load");
    assert.deepEqual(
      mainLoad.linked.parts.map((p) => [p.name, p.doc]),
      [["SI.docx", w.docOf("SI.docx")]],
    );
    // Typed: the review file, a relative name, a quoted name.
    for (const typed of ["paper.revision.json", "paper.docx", '"paper.docx"']) {
      const again = await w.ask("SI.docx", "review_link_join", { main: typed });
      assert.equal(again.ok, true, `${typed}: ${again.error}`);
    }
  } finally {
    await w.cleanup();
  }
});

test("join is idempotent: a second join leaves both files byte for byte", async () => {
  const w = await world();
  try {
    assert.equal((await join1(w)).ok, true);
    const before = [await w.raw("SI.docx"), await w.raw("paper.docx")];
    await sleep(5);
    const again = await join1(w);
    assert.equal(again.ok, true, again.error);
    assert.equal(again.linked.part_of.name, "paper.docx");
    assert.deepEqual([await w.raw("SI.docx"), await w.raw("paper.docx")], before);
    // And a load of the supplement that finds everything in line writes nothing.
    assert.equal((await w.ask("SI.docx", "review_load")).ok, true);
    assert.deepEqual([await w.raw("SI.docx"), await w.raw("paper.docx")], before);
  } finally {
    await w.cleanup();
  }
});

test("join is refused with a plain sentence, and nothing is written", async () => {
  const w = await world();
  try {
    const refused = async (key, main, pattern) => {
      const r = await w.ask(key, "review_link_join", { main });
      assert.equal(r.ok, false, `${key} -> ${main} was not refused`);
      assert.match(r.error, pattern);
    };
    // The manuscript has no comments (no review file, or an empty one).
    await refused(
      "SI.docx",
      w.docOf("nothing.docx"),
      /nothing\.docx has no saved reviewer comments/,
    );
    await w.ask("empty.docx", "review_load");
    await refused("SI.docx", w.docOf("empty.docx"), /empty\.docx has no saved reviewer comments/);
    await refused("SI.docx", "", /Type the path of the manuscript/);
    await refused("SI.docx", w.docOf("SI.docx"), /the file you have open/);
    await assert.rejects(readFile(w.pathOf("SI.docx")), "no file was made for the supplement");
    // This file has changes under a comment the manuscript does not list.
    await w.ask("S2.docx", "review_save_comments", {
      comments: [{ ref: "R9.1", reviewer: 9, text: "Another paper." }],
    });
    await withRevision(
      w.pathOf("S2.docx"),
      w.docOf("S2.docx"),
      (rev) =>
        addChange(rev, "R9.1", {
          ...CHANGE,
          rationale: { ...CHANGE.rationale, requirement_ids: [] },
        }).revision,
    );
    await refused(
      "S2.docx",
      w.docOf("paper.docx"),
      /This file has changes for R9\.1, which paper\.docx does not list/,
    );
    assert.equal("part_of" in (await w.file("S2.docx")), false);
    assert.equal("parts" in (await w.file("paper.docx")), false);
    // The manuscript is itself a supplement; this file has supplements.
    assert.equal((await join1(w)).ok, true);
    await refused(
      "S3.docx",
      w.docOf("SI.docx"),
      /SI\.docx is itself part of the revision of paper\.docx/,
    );
    await w.ask("other.docx", "review_save_comments", { comments: COMMENTS.slice(0, 1) });
    await refused(
      "paper.docx",
      w.docOf("other.docx"),
      /Other files are linked to this file \(SI\.docx\)/,
    );
    // It belongs to another revision already.
    await refused(
      "SI.docx",
      w.docOf("other.docx"),
      /already belongs to the revision of paper\.docx\. Leave that revision first/,
    );
    assert.equal((await w.file("SI.docx")).part_of.manuscript, "paper.docx");
  } finally {
    await w.cleanup();
  }
});

test("the copy follows the manuscript on every load of the supplement", async () => {
  const w = await world();
  try {
    assert.equal((await join1(w)).ok, true);
    const more = [...COMMENTS, { ref: "R2.2", reviewer: 2, text: "Add a scale bar." }];
    assert.equal((await w.ask("paper.docx", "review_save_comments", { comments: more })).ok, true);
    assert.equal(
      (
        await w.ask("paper.docx", "review_set_requirements", {
          ref: "R2.2",
          requirements: ["a", "b"],
        })
      ).ok,
      true,
    );
    const loaded = await w.ask("SI.docx", "review_load");
    assert.deepEqual(
      loaded.revision.comments.map((c) => c.ref),
      ["R1.1", "R1.2", "R2.1", "R2.2"],
    );
    assert.deepEqual(
      loaded.revision.comments[3].requirements,
      (await w.file("paper.docx")).comments[3].requirements,
    );
    assert.equal(loaded.linked.note, null);
  } finally {
    await w.cleanup();
  }
});

test("a lost end of the link is put back on the next load of the supplement", async () => {
  const w = await world();
  try {
    assert.equal((await join1(w)).ok, true);
    const joined = (await w.file("SI.docx")).part_of;
    // The manuscript was saved by an earlier daemon: parts is gone.
    await saveAsEarlier(w, "paper.docx");
    assert.equal("parts" in (await w.file("paper.docx")), false);
    assert.equal((await w.ask("SI.docx", "review_load")).ok, true);
    assert.deepEqual((await w.file("paper.docx")).parts, [
      { doc: "SI.docx", label: joined.label, joined_at: joined.joined_at },
    ]);
    // The supplement was: part_of is gone, the manuscript still lists it.
    await saveAsEarlier(w, "SI.docx");
    assert.equal("part_of" in (await w.file("SI.docx")), false);
    const healed = await w.ask("SI.docx", "review_load");
    assert.equal(healed.linked.part_of.name, "paper.docx");
    assert.deepEqual((await w.file("SI.docx")).part_of, joined);
    // Both lost: nothing links by itself, and joining again works.
    await saveAsEarlier(w, "paper.docx");
    await saveAsEarlier(w, "SI.docx");
    const lost = await w.ask("SI.docx", "review_load");
    assert.equal(lost.linked.part_of, null);
    assert.equal(lost.revision.comments.length, 3, "the copy is still there");
    assert.equal((await join1(w)).ok, true);
    assert.equal((await w.file("paper.docx")).parts.length, 1);
  } finally {
    await w.cleanup();
  }
});

test("a manuscript whose review file is gone does not empty the supplement", async () => {
  const w = await world();
  try {
    assert.equal((await join1(w)).ok, true);
    await rm(w.pathOf("paper.docx"));
    const r = await w.ask("SI.docx", "review_load");
    assert.equal(r.revision.comments.length, 3);
    assert.match(r.linked.note, /review file of paper\.docx was not found/);
    await assert.rejects(readFile(w.pathOf("paper.docx")), "no file is made for the manuscript");
  } finally {
    await w.cleanup();
  }
});

test("leave: the link goes at both ends, with the copy", async () => {
  const w = await world();
  try {
    assert.equal((await w.ask("SI.docx", "review_link_leave")).ok, false);
    assert.equal((await join1(w)).ok, true);
    const r = await w.ask("SI.docx", "review_link_leave");
    assert.equal(r.ok, true, r.error);
    assert.deepEqual([r.linked.part_of, r.revision.comments, r.revision.ledger], [null, [], []]);
    assert.equal("part_of" in (await w.file("SI.docx")), false);
    assert.equal("parts" in (await w.file("paper.docx")), false);
    // Not put back by the next load.
    assert.equal((await w.ask("SI.docx", "review_load")).linked.part_of, null);
    assert.equal((await w.file("paper.docx")).comments.length, 3);
  } finally {
    await w.cleanup();
  }
});

// The write of the manuscript's file is held back. A join or a leave that
// ran it inside the update of the supplement's file could not have saved the
// supplement by then.
test("join and leave save the supplement first, and not inside the save of the manuscript", async () => {
  const w = await world();
  try {
    const hold = () => {
      let release;
      const gate = new Promise((r) => (release = r));
      const held = withRevision(w.pathOf("paper.docx"), w.docOf("paper.docx"), async () => {
        await gate;
        return null;
      });
      return { release, held };
    };
    let h = hold();
    let done = false;
    const joining = join1(w).then((r) => ((done = true), r));
    await until(async () => {
      try {
        return "part_of" in (await w.file("SI.docx"));
      } catch {
        return false;
      }
    }, "the supplement saved while the manuscript is held");
    assert.equal(done, false, "the join still waits for the manuscript");
    assert.equal("parts" in (await w.file("paper.docx")), false);
    h.release();
    await h.held;
    assert.equal((await joining).ok, true);
    assert.equal((await w.file("paper.docx")).parts.length, 1);

    h = hold();
    done = false;
    const leaving = w.ask("SI.docx", "review_link_leave").then((r) => ((done = true), r));
    await until(
      async () => !("part_of" in (await w.file("SI.docx"))),
      "the supplement left while the manuscript is held",
    );
    assert.equal(done, false);
    assert.equal((await w.file("paper.docx")).parts.length, 1);
    h.release();
    await h.held;
    assert.equal((await leaving).ok, true);
    assert.equal("parts" in (await w.file("paper.docx")), false);
  } finally {
    await w.cleanup();
  }
});

test("in a supplement, what the manuscript owns is refused with the name of the manuscript", async () => {
  const w = await world();
  try {
    const requests = {
      review_split: { text: "Reviewer #1\n\n1. Please state the unit." },
      review_save_comments: { comments: [] },
      review_set_requirements: { ref: "R1.1", requirements: ["x"] },
      ledger_add: { item: NOVEL },
      ledger_update: { id: "L1", patch: { text: "y" } },
      ledger_remove: { id: "L1" },
      ledger_import: { doc: { items: [NOVEL] } },
      ledger_suggest: { text: "Do not write novel." },
      profile_set: { overrides: { response_lang: "de" } },
      response_import_previous: { previous: { "R1.1": "old" } },
      response_save: { ref: "R1.1", text: "We did." },
      response_draft: { ref: "R1.1", snapshot: { paragraphs: [] } },
      response_translate: { ref: "R1.1" },
      response_check: { snapshot: { paragraphs: [] } },
      response_export: { snapshot: { paragraphs: [] } },
      profile_rules_apply: {
        text: "Do not write novel.",
        candidates: [{ ...NOVEL, source: { kind: "guideline", sentence: "Do not write novel." } }],
      },
    };
    // Not refused before the join (the request may fail for its own reasons).
    for (const type of ["review_set_requirements", "ledger_add", "response_save"]) {
      const r = await w.ask("SI.docx", type, requests[type]);
      assert.doesNotMatch(String(r.error ?? ""), /Edit this in/, type);
    }
    await rm(w.pathOf("SI.docx"), { force: true });
    assert.equal((await join1(w)).ok, true);
    const before = await w.raw("SI.docx");
    for (const [type, payload] of Object.entries(requests)) {
      const r = await w.ask("SI.docx", type, payload);
      assert.equal(r.ok, false, `${type} was not refused`);
      assert.equal(r.error, "Edit this in paper.docx.", type);
    }
    assert.equal(w.pushed.length, 0);
    assert.equal(await w.raw("SI.docx"), before, "nothing was written");
    // What is the supplement's own still works.
    const status = await w.ask("SI.docx", "review_set_status", { ref: "R1.1", status: "done" });
    assert.equal(status.ok, true, status.error);
    // The same requests work in the manuscript.
    const inMain = await w.ask(
      "paper.docx",
      "review_set_requirements",
      requests.review_set_requirements,
    );
    assert.equal(inMain.ok, true, inMain.error);
  } finally {
    await w.cleanup();
  }
});

test("the manuscript keeps a comment that has changes in a supplement", async () => {
  const w = await world();
  try {
    assert.equal((await join1(w)).ok, true);
    await withRevision(
      w.pathOf("SI.docx"),
      w.docOf("SI.docx"),
      (rev) =>
        addChange(rev, "R1.2", {
          ...CHANGE,
          rationale: { ...CHANGE.rationale, requirement_ids: [] },
        }).revision,
    );
    const dropped = await w.ask("paper.docx", "review_save_comments", {
      comments: COMMENTS.filter((c) => c.ref !== "R1.2"),
    });
    assert.equal(dropped.ok, false);
    assert.match(dropped.error, /Comment R1\.2 has recorded change\(s\) in SI\.docx; keep it/);
    assert.equal((await w.file("paper.docx")).comments.length, 3);
    // A comment without changes anywhere can go.
    const ok = await w.ask("paper.docx", "review_save_comments", {
      comments: COMMENTS.filter((c) => c.ref !== "R2.1"),
    });
    assert.equal(ok.ok, true, ok.error);
    assert.deepEqual(
      (await w.file("paper.docx")).comments.map((c) => c.ref),
      ["R1.1", "R1.2"],
    );
  } finally {
    await w.cleanup();
  }
});

test("the other open panes of the revision are told; the pane that asked is not", async () => {
  const w = await world();
  try {
    w.open.add("other.docx");
    const told = () =>
      w.events.filter((e) => e.event === "linked_changed").map((e) => [e.key, e.doc]);
    assert.equal((await join1(w)).ok, true);
    assert.deepEqual(told(), [["paper.docx", "SI.docx"]]);
    w.events.length = 0;
    // A change of what the supplement copies, made in the manuscript.
    assert.equal((await w.ask("paper.docx", "ledger_remove", { id: "L1" })).ok, true);
    assert.deepEqual(told(), [["SI.docx", "paper.docx"]]);
    w.events.length = 0;
    assert.equal(
      (await w.ask("paper.docx", "review_set_status", { ref: "R1.1", status: "done" })).ok,
      true,
    );
    assert.deepEqual(told(), [], "the status of a comment is each file's own");
    assert.equal((await w.ask("SI.docx", "review_link_leave")).ok, true);
    assert.deepEqual(told(), [["paper.docx", "SI.docx"]]);
  } finally {
    await w.cleanup();
  }
});

test("a supplement in a folder below the manuscript: relative one way, absolute the other", async () => {
  const w = await world();
  try {
    await mkdir(join(w.dir, "suppl"));
    const r = await w.ask("suppl/S1.docx", "review_link_join", { main: "../paper.docx" });
    assert.equal(r.ok, true, r.error);
    assert.equal((await w.file("suppl/S1.docx")).part_of.manuscript, w.docOf("paper.docx"));
    assert.deepEqual(
      (await w.file("paper.docx")).parts.map((p) => p.doc),
      ["suppl/S1.docx"],
    );
    const loaded = await w.ask("suppl/S1.docx", "review_load");
    assert.equal(loaded.linked.part_of.name, "paper.docx");
    assert.equal(loaded.revision.comments.length, 3);
  } finally {
    await w.cleanup();
  }
});

// ---- the session hint of the files of a revision (daemon/sessions.mjs) ----------

test("a manuscript with a supplement and the supplement each get a session hint of their own; any other document keeps the folder's", async () => {
  const { sessionDocKey } = await import("../daemon/revision-parts.mjs");
  const { emptyRevision, saveRevision, revisionPathFor } =
    await import("../daemon/revision-store.mjs");
  const { mkdtemp, rm, writeFile } = await import("node:fs/promises");
  const { tmpdir } = await import("node:os");
  const { join } = await import("node:path");
  const dir = await mkdtemp(join(tmpdir(), "session-key-"));
  try {
    const doc = (name) => join(dir, name);
    const at = "2026-03-01T10:00:00.000Z";
    const put = (name, extra) =>
      saveRevision(revisionPathFor(doc(name)), { ...emptyRevision(doc(name)), ...extra });
    await put("paper.docx", { parts: [{ doc: "SI.docx", label: "", joined_at: at }] });
    await put("SI.docx", { part_of: { manuscript: "paper.docx", label: "", joined_at: at } });
    await put("other.docx", {});
    await put("left.docx", { parts: [] });
    await writeFile(revisionPathFor(doc("broken.docx")), "{ not json");
    assert.equal(await sessionDocKey(doc("paper.docx")), doc("paper.docx"));
    assert.equal(await sessionDocKey(doc("SI.docx")), doc("SI.docx"));
    // A review of its own and no link; a manuscript whose supplement left;
    // no review file; a review file that cannot be read; no document.
    for (const name of ["other.docx", "left.docx", "none.docx", "broken.docx"])
      assert.equal(await sessionDocKey(doc(name)), null, name);
    assert.equal(await sessionDocKey(null), null);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
