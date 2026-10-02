// A supplementary Word document in the revision of a manuscript, step A3:
// the response letter of the manuscript counts and cites the accepted
// changes of the supplementary file. The manuscript reads that file from its
// review file only (what its own pane wrote at its last status refresh), and
// only while that record can be confirmed: the refresh was taken on a saved
// document and the file on disk is still the file that was read.
// Invented files only.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, utimes, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { RevisionSchema, withRevision, applyStates } from "../daemon/revision-store.mjs";
import {
  citedInMain,
  linkedSummary,
  partUnconfirmed,
  partsForLetter,
} from "../daemon/revision-parts.mjs";
import { markStale, partCitations, partLabel, staleReasons } from "../daemon/response-quotes.mjs";
import { checkResponses, partStateText } from "../daemon/response-check.mjs";
import { buildLetterModel, letterPlainText } from "../daemon/response-model.mjs";
import { buildResponsePrompt } from "../daemon/response-prompt.mjs";
import { fileLine, lineRequests } from "../daemon/response-lines.mjs";
import { withDefaults } from "../daemon/journal-profile.mjs";
import { createReviewService } from "../daemon/review-service.mjs";
import { createResponseService, refreshed } from "../daemon/response-service.mjs";
import { createAuxService } from "../daemon/aux-service.mjs";
import { submissionRows } from "../daemon/submission-check.mjs";
import { plainPane } from "./helpers/plain-pane.mjs";

const ENV = { CC_OFFICE_WSL_PATHS: "0", CC_OFFICE_CONFIG_DIR: "/nonexistent-config" };
const profile = withDefaults({});
const L = profile.response_letter;

// ---- the record of a supplementary file ------------------------------------------

const JOINED = "2026-10-02T08:00:00.000Z";
const FILE = { mtime: "2026-10-02T09:00:00.000Z", bytes: 14000 };
const state = (over = {}) => ({
  at: "2026-10-02T09:00:01.000Z",
  text_hash: "h",
  tracked_paragraphs: 0,
  comments: 0,
  tool_comments: 0,
  track_changes_on: true,
  saved: true,
  file: FILE,
  ...over,
});
const QUOTE = "Cultures were incubated at 30 °C.";
const SI_TEXT = ["Supplementary Methods", `${QUOTE} They were shaken.`];
const change = (id, st, ids = ["q1"]) => ({
  id,
  kind: "replace",
  find: "incubated",
  text: "incubated at 30 °C",
  state: st,
  paragraph_index: 1,
  rationale: { what: "temperature", why: "asked", addresses: "R1.1", requirement_ids: ids },
});
const cite = (ids, over = {}) => ({
  change_ids: ids,
  kind: "revised",
  section: "Supplementary Methods",
  paragraph_index: 1,
  quote: QUOTE,
  lines: null,
  lines_pdf: null,
  problem: null,
  ...over,
});
// A part as partsForLetter gives it. changes, citations: of comment R1.1.
const part = ({
  changes = [change("R1.1-1", "accepted")],
  citations,
  why = null,
  ...over
} = {}) => ({
  doc: "/w/SI.docx",
  name: "SI.docx",
  label: "Supplementary Information",
  joined_at: JOINED,
  paragraphs: SI_TEXT,
  confirmed: why === null,
  why,
  revision: {
    comments: [
      {
        ref: "R1.1",
        changes,
        response: {
          citations:
            citations ?? changes.filter((ch) => ch.state === "accepted").map((ch) => cite([ch.id])),
        },
      },
    ],
    snapshot: { paragraphs: SI_TEXT },
    doc_state: state(),
  },
  ...over,
});

test("the record of a supplement is confirmed only when it was refreshed after the join, on a saved document, and the file is still that file", () => {
  const p = (doc_state) => ({ revision: { doc_state }, joined_at: JOINED });
  assert.equal(partUnconfirmed(p(state()), FILE), null);
  // Never refreshed, no review file, refreshed before it joined.
  assert.equal(partUnconfirmed(p(null), FILE), "no_record");
  assert.equal(partUnconfirmed({ revision: null, joined_at: JOINED }, FILE), "no_record");
  assert.equal(partUnconfirmed(p(state({ at: "2026-10-02T07:59:59.000Z" })), FILE), "no_record");
  // Word held unsaved changes, or did not say.
  assert.equal(partUnconfirmed(p(state({ saved: false })), FILE), "unsaved");
  assert.equal(partUnconfirmed(p(state({ saved: null })), FILE), "unsaved");
  // The file was saved again since (another time, another size), cannot be
  // read now, or was not read then.
  assert.equal(
    partUnconfirmed(p(state()), { ...FILE, mtime: "2026-10-02T09:30:00.000Z" }),
    "changed",
  );
  assert.equal(partUnconfirmed(p(state()), { ...FILE, bytes: 14001 }), "changed");
  assert.equal(partUnconfirmed(p(state()), null), "changed");
  assert.equal(partUnconfirmed(p(state({ file: null })), FILE), "changed");
  // What the author reads, each with the way to put it right.
  const text = (why) => partStateText({ name: "SI.docx", why });
  assert.equal(
    text("changed"),
    "SI.docx changed since Sidewrite last read it. Open it, save, press Refresh.",
  );
  assert.equal(
    text("unsaved"),
    "SI.docx had changes that were not saved when Sidewrite last read it. Open it, save, press Refresh.",
  );
  assert.equal(
    text("no_record"),
    "Sidewrite has not read SI.docx since it joined this revision. Open it, save, press Refresh.",
  );
  for (const why of ["changed", "unsaved", "no_record"])
    assert.doesNotMatch(text(why), /—|[0-9a-f]{8}/);
});

test("only accepted changes of a supplement are cited, with its file name and no line numbers", () => {
  // Accepted and revised by hand are cited; pending and rejected are not.
  const p = part({
    changes: [
      change("R1.1-1", "accepted"),
      change("R1.1-2", "pending"),
      change("R1.1-3", "rejected"),
      change("R1.1-4", "user_edited"),
    ],
    citations: [cite(["R1.1-1", "R1.1-4"]), cite(["R1.1-2"]), cite(["R1.1-3"])],
  });
  assert.deepEqual(partCitations(p, "R1.1"), [{ ...cite(["R1.1-1", "R1.1-4"]), doc: "SI.docx" }]);
  assert.deepEqual(partCitations(p, "R9.9"), []);
  // A stored citation that names a change which is not accepted is not used
  // for any of its changes: the accepted one then has no citation, and says so.
  const mixed = partCitations(
    part({
      changes: [change("R1.1-1", "accepted"), change("R1.1-2", "pending")],
      citations: [cite(["R1.1-1", "R1.1-2"])],
    }),
    "R1.1",
  );
  assert.deepEqual(
    mixed.map((c) => [c.change_ids, c.quote, c.problem, c.doc]),
    [[["R1.1-1"], "", "no citation on record; open the file and press Refresh", "SI.docx"]],
  );
  // The stored quote is not in the stored text of its paragraph: a problem, no quote.
  const moved = partCitations(
    part({ paragraphs: ["Supplementary Methods", "Other words."] }),
    "R1.1",
  );
  assert.deepEqual(
    [moved[0].quote, moved[0].problem],
    ["", "the quote is not in the file as it was last read"],
  );
  // Line numbers of the part's own review file are not taken over.
  const lined = partCitations(
    part({ citations: [cite(["R1.1-1"], { lines: { start: 3, end: 3 }, lines_pdf: "x" })] }),
    "R1.1",
  );
  assert.deepEqual([lined[0].lines, lined[0].lines_pdf], [null, null]);
  // A removal has no quote to check.
  const removed = partCitations(
    part({ citations: [cite(["R1.1-1"], { kind: "removed", quote: "" })] }),
    "R1.1",
  );
  assert.deepEqual([removed[0].kind, removed[0].problem], ["removed", null]);
  assert.equal(partLabel("", "SI_v2.docx"), "SI_v2");
  assert.equal(partLabel(" Supplementary Information ", "SI.docx"), "Supplementary Information");
});

// ---- the checks of the manuscript ------------------------------------------------

const SNAP = {
  paragraphs: [
    { index: 0, text: "Methods", w_ins: 0, w_del: 0 },
    { index: 1, text: "Cells were grown in flasks.", w_ins: 0, w_del: 0 },
  ],
  comments: [],
};
const HEADINGS = [{ index: 0, level: 1, text: "Methods" }];
const mainRev = (parts, over = {}) => ({
  comments: [
    {
      ref: "R1.1",
      reviewer: 1,
      text: "State the incubation temperature in the Supplementary Methods.",
      requirements: [{ id: "q1", text: "state the temperature in the Supplementary Methods" }],
      changes: [],
      figure_changes: [],
      declined: null,
      blocked: [],
      response: {
        text: "We now state the temperature.",
        gloss: "",
        previous: null,
        declined_requirements: [],
        stale: [],
        stale_at: {},
        figure_ids: [],
        citations: parts.flatMap((p) => partCitations(p, "R1.1")),
      },
      ...over,
    },
  ],
  ledger: [],
  response_render: null,
  parts: parts.map((p) => ({ doc: p.name, label: p.label, joined_at: JOINED })),
});
const check = (parts, over) =>
  checkResponses({
    revision: mainRev(parts, over),
    snapshot: SNAP,
    headings: HEADINGS,
    profile,
    parts,
  });
const fails = (rows) => rows.filter((r) => r.level === "fail").map((r) => `${r.id}: ${r.detail}`);

test("a request is answered by an accepted change in any file; a pending one says so with the file name; a rejected one does not count", () => {
  // Accepted in the supplement only: nothing fails. No line numbers are asked for it.
  assert.deepEqual(fails(check([part()])), []);
  // Pending there.
  const pending = check([part({ changes: [change("R1.1-1", "pending")] })]);
  assert.deepEqual(fails(pending), [
    "requirements: q1 has only pending changes in SI.docx: state the temperature in the Supplementary Methods",
  ]);
  assert.deepEqual(
    pending.filter((r) => r.id === "undecided_changes").map((r) => [r.level, r.ref, r.detail]),
    [
      [
        "warn",
        "R1.1",
        "1 pending in SI.docx: decide them in Word there, save and refresh; only accepted changes are cited",
      ],
    ],
  );
  // Rejected there: the request is open.
  assert.deepEqual(fails(check([part({ changes: [change("R1.1-1", "rejected")] })])), [
    "requirements: q1 is neither cited nor declined: state the temperature in the Supplementary Methods",
  ]);
  // An accepted change that names another request answers that one only.
  assert.equal(
    fails(check([part({ changes: [change("R1.1-1", "accepted", ["q9"])] })]))[0].startsWith(
      "requirements: q1 is neither cited nor declined",
    ),
    true,
  );
  // Without the parts (a manuscript on its own) the same request is open.
  assert.equal(
    fails(checkResponses({ revision: mainRev([]), snapshot: SNAP, headings: HEADINGS, profile }))
      .length,
    1,
  );
});

test("part_state fails while the record of a supplement cannot be confirmed, and passes when it can", () => {
  assert.deepEqual(
    check([part()]).filter((r) => r.id === "part_state"),
    [],
  );
  for (const why of ["changed", "unsaved", "no_record"]) {
    const rows = check([part({ why })]).filter((r) => r.id === "part_state");
    assert.deepEqual(
      rows.map((r) => [r.level, r.ref, r.detail]),
      [["fail", null, partStateText({ name: "SI.docx", why })]],
    );
  }
});

test("a citation of a supplement is checked against the text of that file; a problem fails and names the file", () => {
  // The citation stored in the manuscript's file no longer fits the part.
  const stale = mainRev([part()]);
  const now = [part({ paragraphs: ["Supplementary Methods", "Rewritten."] })];
  const rows = checkResponses({
    revision: stale,
    snapshot: SNAP,
    headings: HEADINGS,
    profile,
    parts: now,
  });
  assert.deepEqual(
    fails(rows),
    [
      "stale: quote_changed: a quoted sentence changed after this answer was written; read the answer, then save or draft it again",
      `quote_exact: quote is not in SI.docx as it was last read: ${QUOTE}`,
    ].slice(1),
  );
  // An accepted change the pane of the part wrote no citation for.
  const bare = check([part({ citations: [] })]);
  assert.deepEqual(fails(bare), [
    "quote_exact: SI.docx: no citation on record; open the file and press Refresh",
  ]);
  // The heading of the part is not a heading of the manuscript: no section failure.
  assert.deepEqual(
    check([part()]).filter((r) => r.id === "section" || r.id === "lines"),
    [],
  );
});

test("an answer written before the accepted changes of a supplement changed is marked change_state, a changed quote quote_changed", () => {
  const c = mainRev([part()]).comments[0];
  // As written.
  assert.deepEqual(staleReasons(c, SNAP, null, [part()]), []);
  // The change was rejected in the supplement.
  assert.deepEqual(
    staleReasons(c, SNAP, null, [part({ changes: [change("R1.1-1", "rejected")] })]),
    ["change_state"],
  );
  // A second change was accepted there.
  assert.deepEqual(
    staleReasons(c, SNAP, null, [
      part({ changes: [change("R1.1-1", "accepted"), change("R1.1-2", "accepted")] }),
    ]),
    ["change_state"],
  );
  // The quoted sentence of the supplement was edited.
  assert.deepEqual(
    staleReasons(c, SNAP, null, [part({ paragraphs: ["Supplementary Methods", "Edited."] })]),
    ["quote_changed"],
  );
  // The file left the revision.
  assert.deepEqual(staleReasons(c, SNAP, null, []), ["quote_changed", "change_state"]);
  // An answer from before the part had an accepted change.
  const before = mainRev([part({ changes: [change("R1.1-1", "pending")] })]).comments[0];
  assert.deepEqual(before.response.citations, []);
  assert.deepEqual(staleReasons(before, SNAP, null, [part()]), ["change_state"]);
  // A caller that does not know the parts judges nothing of them.
  assert.deepEqual(staleReasons(c, SNAP, null, null), []);
  // markStale keeps the mark until the answer is written again.
  const marked = markStale(mainRev([part()]), SNAP, [
    part({ changes: [change("R1.1-1", "rejected")] }),
  ]);
  assert.deepEqual(marked.comments[0].response.stale, ["change_state"]);
  assert.deepEqual(markStale(marked, SNAP, [part()]).comments[0].response.stale, ["change_state"]);
});

test("refreshed keeps the line numbers of the manuscript and puts the citations of the parts after them", () => {
  const LINES = { start: 9, end: 9, start_page: 1, end_page: 1, match_count: 1 };
  const own = {
    id: "R1.1-1",
    kind: "replace",
    find: "grown",
    text: "grown at 30 °C",
    state: "accepted",
    paragraph_index: 1,
    before: "Cells were grown in flasks.",
    after: "Cells were grown at 30 °C in flasks.",
    rationale: { what: "w", why: "y", addresses: "a", requirement_ids: [] },
    context: null,
  };
  const snap = {
    paragraphs: [SNAP.paragraphs[0], { ...SNAP.paragraphs[1], text: own.after }],
    comments: [],
  };
  const stored = RevisionSchema.parse({
    version: 1,
    manuscript: "paper.docx",
    updated_at: JOINED,
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "t",
        changes: [{ ...own, created_at: JOINED, word_comment_id: null }],
        response: {
          text: "a",
          citations: [
            {
              change_ids: ["R1.1-1"],
              kind: "revised",
              section: "Methods",
              paragraph_index: 1,
              quote: own.after,
              lines: LINES,
              lines_pdf: "pdf",
            },
            // What an earlier refresh took from a part that has since changed.
            { ...cite(["R1.1-9"], { quote: "Old words." }), doc: "SI.docx" },
          ],
        },
      },
    ],
  });
  const next = refreshed(stored, snap, HEADINGS, profile, [part()]);
  const cits = next.comments[0].response.citations;
  assert.deepEqual(
    cits.map((c) => [c.doc ?? null, c.change_ids, c.quote, c.lines, c.lines_pdf]),
    [
      [null, ["R1.1-1"], own.after, LINES, "pdf"],
      ["SI.docx", ["R1.1-1"], QUOTE, null, null],
    ],
  );
  // No line request for a citation of a part, whose text is not in the PDF.
  assert.deepEqual(
    lineRequests(next.comments, snap).map((r) => r.id),
    ["R1.1#0"],
  );
  // Without parts the citations of a part go.
  assert.deepEqual(
    refreshed(stored, snap, HEADINGS, profile).comments[0].response.citations.map((c) => c.doc),
    [undefined],
  );
});

// ---- the letter -------------------------------------------------------------------

test("with a part every location line names its file; with none the letter is as it was", () => {
  const LINES = { start: 211, end: 213, start_page: 4, end_page: 4, match_count: 1 };
  const rev = mainRev([part()]);
  rev.comments[0].response.citations.unshift({
    change_ids: ["R1.1-1"],
    kind: "revised",
    section: "Results",
    paragraph_index: 1,
    quote: "The plateau lasted two days.",
    lines: LINES,
    lines_pdf: "pdf",
    problem: null,
  });
  const lines = (r) =>
    buildLetterModel(r, profile)
      .paragraphs.filter((p) => p.kind === "location" || p.kind === "quote")
      .map((p) => p.text);
  assert.deepEqual(lines(rev), [
    "Manuscript [Revised in Results, lines 211-213]",
    "The plateau lasted two days.",
    "Supplementary Information [Revised in Supplementary Methods]",
    QUOTE,
  ]);
  // The label of the link; without one, the file name without its extension.
  rev.parts[0].label = "";
  assert.equal(lines(rev)[2], "SI [Revised in Supplementary Methods]");
  // The frames are those of the profile.
  const framed = withDefaults({
    response_letter: { file_line: "{line} ({file})", main_label: "Main text" },
  });
  assert.equal(
    fileLine(rev.comments[0].response.citations[0], framed.response_letter),
    "[Revised in Results, lines 211-213] (Main text)",
  );
  // A removal in a part, and a citation of a part without a section.
  assert.equal(
    fileLine({ ...cite(["x"], { kind: "removed" }), doc: "SI.docx" }, L, { "SI.docx": "SI" }),
    "SI [Removed from Supplementary Methods]",
  );
  assert.equal(
    fileLine({ ...cite(["x"], { section: null }), doc: "SI.docx" }, L, { "SI.docx": "SI" }),
    "SI [Revised]",
  );
  // No parts: the same record without the link prints the lines as before.
  const alone = structuredClone(rev);
  delete alone.parts;
  alone.comments[0].response.citations.pop();
  assert.deepEqual(lines(alone), [
    "[Revised in Results, lines 211-213]",
    "The plateau lasted two days.",
  ]);
  const empty = structuredClone(alone);
  empty.parts = [];
  assert.equal(
    JSON.stringify(buildLetterModel(empty, profile)),
    JSON.stringify(buildLetterModel(alone, profile)),
  );
  // The plain text of a reviewer carries the same lines.
  assert.match(
    letterPlainText(buildLetterModel(rev, profile), 1),
    /\nSI \[Revised in Supplementary Methods\]\nCultures were incubated at 30 °C\.$/,
  );
});

test("the draft prompt lists the changes of each supplement after those of the manuscript, with the same rows", () => {
  const comment = mainRev([]).comments[0];
  const p = (changes) => [{ name: "SI.docx", label: "Supplementary Information", changes }];
  const built = buildResponsePrompt({
    comment,
    profile,
    parts: p([change("R1.1-1", "accepted"), change("R1.1-2", "pending")]),
  });
  const text = built.prompt;
  const at = (s) => text.indexOf(s);
  assert.ok(at("Changes in the manuscript for this comment") > 0);
  assert.ok(
    at(
      "Changes in Supplementary Information (SI.docx) for this comment (accepted, or revised by the author; you may report them):",
    ) > at("Changes in the manuscript for this comment"),
  );
  assert.match(
    text,
    /\n- \[R1\.1-1\] replaced <<<incubated>>> with <<<incubated at 30 °C>>> \(answers q1\)\n {2}What: temperature Why: asked\n/,
  );
  assert.match(
    text,
    /\nNot in Supplementary Information \(do not report any of these\):\n- \[R1\.1-2\] not decided in Word yet; do not claim: replaced <<<incubated>>>/,
  );
  // The rules allow a change in a supplementary file to be reported, and
  // keep file names out of the prose.
  assert.match(
    text,
    /\n3\. Never claim a change that is not listed as in the manuscript or as in a supplementary file\.\n/,
  );
  assert.match(
    text,
    /\n4\. If no change is listed as in the manuscript or as in a supplementary file, do not say that anything was revised\.\n/,
  );
  assert.match(
    text,
    /Do not write file names\. The tool adds the line that names the file after your text\./,
  );
  assert.deepEqual(built.partChangeIds, ["SI.docx#R1.1-1"]);
  // No parts, or a part without a change for this comment: the prompt of before.
  const plain = buildResponsePrompt({ comment, profile }).prompt;
  assert.equal(buildResponsePrompt({ comment, profile, parts: [] }).prompt, plain);
  assert.equal(buildResponsePrompt({ comment, profile, parts: p([]) }).prompt, plain);
  assert.doesNotMatch(plain, /supplementary file|file names/i);
});

// ---- the stored form ---------------------------------------------------------------

test("a citation from before the file name saves as it was; one with it keeps it", () => {
  const old = {
    change_ids: ["R1.1-1"],
    kind: "revised",
    section: "Methods",
    paragraph_index: 1,
    quote: "q",
    lines: null,
    lines_pdf: null,
    problem: null,
  };
  const parse = (citations) =>
    RevisionSchema.parse({
      version: 1,
      manuscript: "paper.docx",
      updated_at: JOINED,
      comments: [{ ref: "R1.1", reviewer: 1, text: "t", response: { text: "a", citations } }],
    }).comments[0].response.citations;
  assert.equal(JSON.stringify(parse([old])), JSON.stringify([old]));
  assert.equal(parse([{ ...old, doc: "SI.docx" }])[0].doc, "SI.docx");
});

// ---- the submission checklist -------------------------------------------------------

test("the checklist gets one line per linked file for tracked changes, tool comments and a clean file, from its confirmed record", () => {
  const live = { snapshot: SNAP, headings: HEADINGS, format: null, saved: true };
  const rows = (parts) =>
    submissionRows({
      revision: mainRev(parts),
      profile,
      live,
      files: {},
      file: "paper.docx",
      parts,
    })
      .filter((r) => r.file === "SI.docx")
      .map((r) => [r.id, r.state, r.text, r.go]);
  // Confirmed and clean.
  assert.deepEqual(rows([part()]), [
    ["changes_decided", "ready", "No tracked change is waiting for a decision in SI.docx.", null],
    ["tool_comments", "ready", "The tool's comments are removed from SI.docx.", null],
    ["file_clean", "ready", "SI.docx is clean: no tracked changes and no comments are left.", null],
  ]);
  // Confirmed, with work left in it.
  const busy = part();
  busy.revision.doc_state = state({ tracked_paragraphs: 2, comments: 3, tool_comments: 1 });
  assert.deepEqual(rows([busy]), [
    [
      "changes_decided",
      "not_ready",
      "2 paragraphs of SI.docx still hold tracked changes. Accept or reject them in Word, save and refresh there.",
      null,
    ],
    [
      "tool_comments",
      "not_ready",
      "1 comment of the tool is still in SI.docx. Remove it there, save and refresh.",
      null,
    ],
    [
      "file_clean",
      "not_ready",
      "SI.docx is not clean yet. Left in it: 2 paragraphs with tracked changes, 1 comment of the tool, 2 comments of yours or the reviewers.",
      null,
    ],
  ]);
  // Not confirmed: never guessed, each line says how to put it right.
  const open =
    "Open SI.docx and refresh: Sidewrite cannot confirm what it last read there (save the file first).";
  assert.deepEqual(rows([part({ why: "changed" })]), [
    ["changes_decided", "not_ready", open, null],
    ["tool_comments", "not_ready", open, null],
    ["file_clean", "not_ready", open, null],
  ]);
  // The rows of the open file are as they were, and stand before those of the part.
  const all = submissionRows({
    revision: mainRev([part()]),
    profile,
    live,
    files: {},
    file: "paper.docx",
    parts: [part()],
  });
  assert.deepEqual(
    all.map((r) => `${r.id}:${r.file}`),
    [
      "answers:paper.docx",
      "requests:paper.docx",
      "changes_decided:paper.docx",
      "tool_comments:paper.docx",
      "file_clean:paper.docx",
      "changes_decided:SI.docx",
      "tool_comments:SI.docx",
      "file_clean:SI.docx",
      "format_counts:paper.docx",
      "letter:paper.docx",
      "marked_copy:paper.docx",
      "anonymized_copy:paper.docx",
      "figures:paper.docx",
    ],
  );
  // An unconfirmed part is not counted as a failing answer: its own lines say it.
  const unconfirmed = submissionRows({
    revision: mainRev([part()]),
    profile,
    live,
    files: {},
    file: "paper.docx",
    parts: [part({ why: "unsaved" })],
  });
  assert.equal(unconfirmed.find((r) => r.id === "answers").state, "ready");
  // Without parts: the ten rows of before.
  assert.equal(
    submissionRows({ revision: mainRev([]), profile, live, files: {}, file: "paper.docx" }).length,
    10,
  );
});

// ---- two files, the services ---------------------------------------------------------

const COMMENTS = [
  {
    ref: "R1.1",
    reviewer: 1,
    text: "State the incubation temperature in the Supplementary Methods.",
    requirements: ["state the temperature in the Supplementary Methods"],
  },
];
const CHANGE = {
  word_comment_id: "9",
  paragraph_index: 1,
  kind: "replace",
  find: "incubated",
  text: "incubated at 30 °C",
  before: "Cultures were incubated. They were shaken.",
  after: `${QUOTE} They were shaken.`,
  rationale: { what: "temperature", why: "asked", addresses: "R1.1", requirement_ids: ["q1"] },
};
const snapOf = (...texts) => ({
  track_changes_on: true,
  has_fields: false,
  paragraphs: texts.map((text, index) => ({ index, id: null, text, w_ins: 0, w_del: 0 })),
  comments: [],
});
const PAPER = snapOf("Methods", "Cells were grown in flasks.");
const SI_BEFORE = snapOf("Supplementary Methods", CHANGE.before);
const SI_AFTER = snapOf("Supplementary Methods", CHANGE.after);
const SI_HEADINGS = [{ index: 0, level: 1, text: "Supplementary Methods" }];

async function world() {
  const dir = await mkdtemp(join(tmpdir(), "parts-letter-"));
  const docOf = (name) => join(dir, name);
  const events = [];
  const waiters = new Map();
  const done = new Map();
  const sendEvent = (key, e) => {
    events.push({ key, ...e });
    if (e.event !== "job_done") return;
    if (waiters.has(e.job_id)) waiters.get(e.job_id)(e.result);
    else done.set(e.job_id, e.result);
  };
  const deps = {
    getContext: (key) => ({ activeDoc: docOf(key), caps: ["target_hash"] }),
    sendEvent,
    log: () => {},
    env: ENV,
  };
  const review = createReviewService({
    ...deps,
    callTool: async () => ({}),
    pushTurn: () => {},
    listPanes: () =>
      ["paper.docx", "SI.docx"].map((key) => ({ key, host: "word", activeDoc: docOf(key) })),
    afterRefresh: (rev, snap, parts) => markStale(rev, snap, parts),
  });
  const handlers = {
    ...review.handlers,
    // The real Python tools: the letter is written and verified.
    ...createResponseService({ ...deps, oneshotJob: async () => ({ run: null, model: null }) })
      .handlers,
    ...createAuxService(deps).handlers,
  };
  const ask = (key, type, payload = {}) =>
    new Promise((resolve) => {
      const reply = (r) => {
        if (!r.accepted) return resolve(r);
        if (done.has(r.job_id)) resolve(done.get(r.job_id));
        else waiters.set(r.job_id, resolve);
      };
      handlers[type]({ type, request_id: "q", ...plainPane(payload) }, reply, key, "word");
    });
  const pathOf = (name) => join(dir, name.replace(/\.docx$/, ".revision.json"));
  const file = async (name) => JSON.parse(await readFile(pathOf(name), "utf8"));
  const w = { dir, docOf, pathOf, ask, file, events, review };
  w.cleanup = () => rm(dir, { recursive: true, force: true });
  await writeFile(docOf("paper.docx"), "the manuscript as saved");
  await writeFile(docOf("SI.docx"), "the supplement as saved");
  const ok = (r) => (assert.equal(r.ok, true, r.error), r);
  ok(await ask("paper.docx", "review_save_comments", { comments: COMMENTS }));
  ok(
    await ask("SI.docx", "review_link_join", {
      main: docOf("paper.docx"),
      label: "Supplementary Information",
    }),
  );
  // The pane of the supplement: a status refresh, then the citations of this
  // file (what review-pane.js does after every refresh in a part).
  w.refreshSI = async (snapshot, saved = true) => {
    ok(await ask("SI.docx", "review_refresh", { snapshot, saved }));
    return ok(await ask("SI.docx", "response_state", { snapshot, headings: SI_HEADINGS }));
  };
  w.main = (type, more = {}) =>
    ask("paper.docx", type, {
      snapshot: PAPER,
      headings: [{ index: 0, level: 1, text: "Methods" }],
      ...more,
    });
  // The change of the supplement is recorded (as a tracked edit Word applied).
  await w.refreshSI(SI_BEFORE);
  await review.trackedEditDeps("SI.docx").onEdit({ ref: "R1.1", change: CHANGE });
  return w;
}
const failing = (r) =>
  r.checks.filter((x) => x.level === "fail").map((x) => `${x.id}: ${x.detail}`);

test("two files: the accepted change of the supplement answers the request, is cited with its file, and the exported letter verifies", async () => {
  const w = await world();
  try {
    // Still tracked in the supplement: the manuscript's check says where.
    await w.refreshSI(snapOf("Supplementary Methods", CHANGE.before));
    let r = await w.main("response_save", { ref: "R1.1", text: "We now state the temperature." });
    assert.equal(r.ok, true, r.error);
    // (The refresh above read the old text: the change is not in the file.)
    r = await w.main("response_check");
    assert.equal(
      failing(r).some((x) => x.startsWith("requirements: q1")),
      true,
    );

    // Accepted in Word, saved, refreshed in the pane of the supplement.
    await w.refreshSI(SI_AFTER);
    const si = await w.file("SI.docx");
    assert.equal(si.comments[0].changes[0].state, "accepted");
    assert.deepEqual(
      si.comments[0].response.citations.map((c) => [c.section, c.quote, c.lines]),
      [["Supplementary Methods", QUOTE, null]],
    );
    assert.equal(si.doc_state.saved, true);

    // The manuscript: the answer was written before that change was accepted.
    r = await w.main("response_check");
    assert.deepEqual(failing(r), [
      "stale: change_state: the accepted changes or adopted figure changes are not the ones this answer was written for; read the answer, then save or draft it again",
    ]);
    assert.deepEqual(
      r.revision.comments[0].response.citations.map((c) => [c.doc, c.section, c.quote, c.lines]),
      [["SI.docx", "Supplementary Methods", QUOTE, null]],
    );
    assert.deepEqual(
      r.linked_changes.map((x) => [x.file, x.ref, x.state]),
      [["SI.docx", "R1.1", "accepted"]],
    );
    assert.deepEqual(r.parts_unconfirmed, []);
    // Saved again against the citation the card shows now.
    r = await w.main("response_save", {
      ref: "R1.1",
      text: "We now state the temperature.",
      shown: r.revision.comments[0].response.citations,
    });
    assert.deepEqual(r.warnings, []);
    r = await w.main("response_check");
    assert.deepEqual(failing(r), []);

    // Export, not forced: the letter names the file, and the verifier read it back.
    const x = await w.main("response_export");
    assert.equal(x.ok, true, x.error);
    assert.equal(x.exported, true, JSON.stringify(x.summary));
    assert.equal(x.verify.ok, true);
    const model = JSON.parse(await readFile(x.paths.model, "utf8")).paragraphs;
    const at = model.findIndex((p) => p.kind === "location");
    assert.deepEqual(
      [model[at].text, model[at + 1].kind, model[at + 1].text],
      ["Supplementary Information [Revised in Supplementary Methods]", "quote", QUOTE],
    );
    // The paragraphs the quotes were checked against: the manuscript's, then the part's.
    assert.deepEqual(JSON.parse(await readFile(x.paths.manuscript, "utf8")), [
      "Methods",
      "Cells were grown in flasks.",
      "Supplementary Methods",
      CHANGE.after,
    ]);
    // The pane of the supplement is told which of its changes the letter cites.
    const linked = (await w.ask("SI.docx", "review_load")).linked;
    assert.deepEqual(linked.cited_in_main, [
      {
        ref: "R1.1",
        kind: "revised",
        section: "Supplementary Methods",
        quote: QUOTE,
        change_ids: ["R1.1-1"],
      },
    ]);
    assert.equal(linked.saved_at_refresh, true);
    assert.equal(linked.part_of.label, "Supplementary Information");
    // The checklist of the manuscript has the lines of the supplement.
    const rows = (await w.main("submission_check", { saved: true })).rows;
    assert.deepEqual(
      rows.filter((row) => row.file === "SI.docx").map((row) => [row.id, row.state]),
      [
        ["changes_decided", "ready"],
        ["tool_comments", "ready"],
        ["file_clean", "ready"],
      ],
    );
    assert.equal(
      rows.find((row) => row.id === "letter").state,
      "ready",
      rows.find((row) => row.id === "letter").text,
    );
  } finally {
    await w.cleanup();
  }
});

test("two files: a supplement that changed on disk, or was read with unsaved changes, fails part_state and holds the export", async () => {
  const w = await world();
  try {
    await w.refreshSI(SI_AFTER);
    let r = await w.main("response_save", { ref: "R1.1", text: "We now state the temperature." });
    r = await w.main("response_check");
    assert.deepEqual(failing(r), []);

    // The author rejects the change in Word and saves, and does not refresh:
    // the file on disk is no longer the file Sidewrite read.
    await writeFile(w.docOf("SI.docx"), "the supplement, saved again with the change rejected");
    r = await w.main("response_check");
    assert.deepEqual(failing(r), [
      "part_state: SI.docx changed since Sidewrite last read it. Open it, save, press Refresh.",
    ]);
    assert.deepEqual(r.parts_unconfirmed, ["SI.docx"]);
    let x = await w.main("response_export");
    assert.equal(x.exported, false, "an unforced export is refused");
    assert.equal(x.summary.fail, 1);
    // The same bytes with another time on the disk are another file too.
    const again = new Date(Date.now() + 5000);
    await writeFile(w.docOf("SI.docx"), "the supplement as saved");
    await utimes(w.docOf("SI.docx"), again, again);
    assert.equal(failing(await w.main("response_check")).length, 1);
    // The checklist does not guess: the lines of the file say to open it.
    const rows = (await w.main("submission_check", { saved: true })).rows.filter(
      (row) => row.file === "SI.docx",
    );
    assert.deepEqual([...new Set(rows.map((row) => row.state))], ["not_ready"]);
    assert.match(rows[0].text, /^Open SI\.docx and refresh/);

    // Refreshed while Word held unsaved changes: still not confirmed.
    await w.refreshSI(SI_AFTER, false);
    assert.deepEqual(failing(await w.main("response_check")), [
      "part_state: SI.docx had changes that were not saved when Sidewrite last read it. Open it, save, press Refresh.",
    ]);
    // Word did not say: the same.
    await w.refreshSI(SI_AFTER, null);
    assert.equal(failing(await w.main("response_check"))[0].startsWith("part_state:"), true);

    // Saved and refreshed, with the change rejected: the request is open
    // again and the answer is marked as written for another set of changes.
    await w.refreshSI(SI_BEFORE, true);
    r = await w.main("response_check");
    // (The sentence it quoted is gone from the supplement as well.)
    assert.deepEqual(failing(r), [
      "requirements: q1 is neither cited nor declined: state the temperature in the Supplementary Methods",
      "stale: quote_changed: a quoted sentence changed after this answer was written; read the answer, then save or draft it again",
      "stale: change_state: the accepted changes or adopted figure changes are not the ones this answer was written for; read the answer, then save or draft it again",
    ]);
    assert.deepEqual(r.revision.comments[0].response.stale, ["quote_changed", "change_state"]);
    assert.deepEqual(r.revision.comments[0].response.citations, []);
    x = await w.main("response_export");
    assert.equal(x.exported, false);
  } finally {
    await w.cleanup();
  }
});

test("two files: a supplement that was never refreshed after it joined is not confirmed", async () => {
  const w = await world();
  try {
    // doc_state from before the join does not count: put the time of the join after it.
    await withRevision(w.pathOf("paper.docx"), w.docOf("paper.docx"), (rev) => ({
      ...rev,
      parts: rev.parts.map((p) => ({ ...p, joined_at: "2999-01-01T00:00:00.000Z" })),
    }));
    const parts = await partsForLetter(w.docOf("paper.docx"), await w.file("paper.docx"));
    assert.deepEqual(
      parts.map((p) => [p.name, p.label, p.confirmed, p.why]),
      [["SI.docx", "Supplementary Information", false, "no_record"]],
    );
    const r = await w.main("response_check");
    assert.equal(
      failing(r)[0],
      "part_state: Sidewrite has not read SI.docx since it joined this revision. Open it, save, press Refresh.",
    );
  } finally {
    await w.cleanup();
  }
});

test("two files: a refresh of the manuscript's Review tab judges the answers against the supplement too", async () => {
  const w = await world();
  try {
    await w.refreshSI(SI_AFTER);
    await w.main("response_save", { ref: "R1.1", text: "We now state the temperature." });
    let r = await w.main("response_check");
    assert.deepEqual(failing(r), []);
    // A refresh of the Review tab with nothing changed marks nothing.
    await w.ask("paper.docx", "review_refresh", { snapshot: PAPER, saved: true });
    assert.deepEqual((await w.file("paper.docx")).comments[0].response.stale, []);
    // The change is rejected in the supplement and its pane refreshes.
    await w.refreshSI(SI_BEFORE);
    await w.ask("paper.docx", "review_refresh", { snapshot: PAPER, saved: true });
    assert.deepEqual((await w.file("paper.docx")).comments[0].response.stale, [
      "quote_changed",
      "change_state",
    ]);
  } finally {
    await w.cleanup();
  }
});

test("the name of a supplement in the letter is set from its own pane, at both ends; without one it is the file name", async () => {
  const w = await world();
  try {
    const labels = async () => [
      (await w.file("SI.docx")).part_of.label,
      (await w.file("paper.docx")).parts[0].label,
    ];
    assert.deepEqual(await labels(), ["Supplementary Information", "Supplementary Information"]);
    let r = await w.ask("SI.docx", "review_link_label", {
      label: "  Supplementary Methods and Tables ",
    });
    assert.equal(r.ok, true, r.error);
    assert.equal(r.linked.part_of.label, "Supplementary Methods and Tables");
    assert.deepEqual(await labels(), [
      "Supplementary Methods and Tables",
      "Supplementary Methods and Tables",
    ]);
    // The time of the join stays.
    const joined = (await w.file("paper.docx")).parts[0].joined_at;
    assert.equal((await w.file("SI.docx")).part_of.joined_at, joined);
    // The manuscript's pane is told.
    assert.equal(
      w.events.some((e) => e.key === "paper.docx" && e.event === "linked_changed"),
      true,
    );
    // Emptied: the letter falls back to the file name without its extension.
    r = await w.ask("SI.docx", "review_link_label", { label: "" });
    assert.equal(r.linked.part_of.label, "SI");
    const summary = linkedSummary(w.docOf("paper.docx"), await w.file("paper.docx"));
    assert.equal(summary.parts[0].label, "SI");
    // Refused: too long, and in a file that is no part.
    r = await w.ask("SI.docx", "review_link_label", { label: "x".repeat(81) });
    assert.match(r.error, /too long/);
    r = await w.ask("paper.docx", "review_link_label", { label: "Main" });
    assert.match(r.error, /not part of another revision/);
    assert.deepEqual(citedInMain(null, "SI.docx"), []);
  } finally {
    await w.cleanup();
  }
});

// A row applied with applyStates directly, for the judges that do not go
// through a refresh (kept so the import is used where a test needs it).
test("applyStates is what a refresh of the supplement writes: the manuscript reads the state it stores", async () => {
  const w = await world();
  try {
    await withRevision(w.pathOf("SI.docx"), w.docOf("SI.docx"), (rev) =>
      applyStates(rev, [
        { ref: "R1.1", change_id: "R1.1-1", state: "pending", paragraph_index: 1 },
      ]),
    );
    const parts = await partsForLetter(w.docOf("paper.docx"), await w.file("paper.docx"));
    assert.deepEqual(partCitations(parts[0], "R1.1"), []);
  } finally {
    await w.cleanup();
  }
});
