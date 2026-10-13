// Fail closed when the fields of a paragraph could not be read. The daemon
// guards a field (a citation, a reference list over several paragraphs)
// through the reviewed text of the last status refresh, which shows it as
// U+0013 instruction U+0014 result U+0015. A paragraph the task pane gave no
// reviewed text for shows nothing, so an edit or a rewrite there is refused
// unless the document is known to hold no field. And a rewrite never reaches
// a paragraph of a field that runs over several paragraphs, where the task
// pane alone sees no marker. All texts are invented.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { applyTrackedEdit, applyTrackedRewrite, FIELD_CAP } from "../daemon/review-tools.mjs";
import { snapshotRead } from "../daemon/doc-diff.mjs";
import { createReviewService } from "../daemon/review-service.mjs";
import { RevisionSchema } from "../daemon/revision-store.mjs";
import { FIELDS_UNREAD, fieldsUnread, xmlHasField } from "../taskpane/shared/field-text.js";
import { rewritesFromLog, refusalsFromLog } from "../scripts/e2e/replay-lib.mjs";

const B = "\u0013";
const S = "\u0014";
const E = "\u0015";
// A reference list that is one field over three paragraphs: the begin marker
// in the first, the end marker in the third, no marker in the middle one.
const LIST = [
  `${B} ADDIN ZOTERO_BIBL {"uncited":[]} CSL_BIBLIOGRAPHY ${S}Arden, Q. (2019). Shaken flasks.`,
  "Belloc, M. (2020). A lag in slim vessels.",
  `Corvane, I. (2021). The medium, not the flask.${E}`,
];
const DOC = ["Title", "Cells grew slowly in narrow flasks.", ...LIST, "A closing sentence."];
const IDS = DOC.map((_, k) => `U${k}`);
const TURN = { ref: "R1.2" };
const R = { what: "w", why: "y", addresses: "a" };

function deps(ctx = {}) {
  const d = {
    calls: [],
    edits: [],
    attempts: 0,
    lines: [],
    log: (s) => d.lines.push(s),
    getLedger: async () => [],
    getReviewContext: async () => ({ paragraphs: DOC, paragraph_ids: IDS, ...ctx }),
    paneChecksTarget: () => true,
    paneCan: (cap) =>
      ["target_hash", "author_insertion_guard", "tracked_rewrite", FIELD_CAP].includes(cap),
    onAttempt: async () => ((d.attempts += 1), { id: "R1.2-b1" }),
    call: async (name, args) => {
      d.calls.push(args);
      return {
        ok: true,
        word_comment_id: args.no_comment ? null : 7,
        paragraph_index: args.expected_paragraph.index,
        before: DOC[args.expected_paragraph.index],
        after: "after",
        target_checked: true,
      };
    },
    onEdit: async ({ change }) => (d.edits.push(change), { id: `R1.2-${d.edits.length}` }),
    onNotApplied: async () => {},
    onStaleTarget: async () => {},
  };
  return d;
}
const edit = (find, replace) => ({ comment_ref: "R1.2", find, replace, rationale: R });
const rewrite = (id, text) => ({
  comment_ref: "R1.2",
  paragraph_id: id,
  text,
  what: "w",
  why: "y",
  addresses: "a",
});
const nothingSent = (d) =>
  assert.deepEqual([d.calls.length, d.attempts, d.edits.length], [0, 0, 0]);

test("the rule: a paragraph is unread when its reviewed text, or that of one before it, is missing and the document may hold a field", () => {
  // No reviewed text for paragraph 3; whether the document has fields is not known.
  const read = { unreviewed: [3], fields: null, unknown: false };
  assert.deepEqual(
    [0, 2, 3, 4, 9].map((k) => fieldsUnread(read, k)),
    [false, false, true, true, true],
  );
  assert.equal(fieldsUnread({ ...read, fields: true }, 3), true);
  // Known to hold no field anywhere: nothing to cut through.
  assert.equal(fieldsUnread({ ...read, fields: false }, 3), false);
  // Every reviewed text was read.
  assert.equal(fieldsUnread({ unreviewed: [], fields: true, unknown: false }, 3), false);
  // A snapshot that says nothing about what was read.
  assert.equal(fieldsUnread({ unreviewed: null, fields: null, unknown: true }, 0), true);
  assert.equal(fieldsUnread({ unreviewed: null, fields: false, unknown: true }, 0), false);
  // A caller that hands in nothing (a test of something else) is not refused.
  assert.equal(fieldsUnread(undefined, 0), false);
});

test("what a snapshot says about the fields: which paragraphs have no reviewed text, and whether a field is anywhere", () => {
  const p = (index, text, reviewed_text) => ({ index, text, reviewed_text });
  // The pane read every reviewed text; one shows a field.
  assert.deepEqual(
    snapshotRead({ paragraphs: [p(0, "T", "T"), p(1, "x", `${B} REF a ${S}x${E}`)] }),
    { unreviewed: [], fields: true },
  );
  // No reviewed text at all (Word gave none): what the pane found in the XML decides.
  const none = [p(1, "b", null), p(0, "a", null)];
  assert.deepEqual(snapshotRead({ paragraphs: none, has_fields: true }), {
    unreviewed: [0, 1],
    fields: true,
  });
  assert.deepEqual(snapshotRead({ paragraphs: none, has_fields: false }), {
    unreviewed: [0, 1],
    fields: false,
  });
  // A pane that does not say: not known. A paragraph without the key counts as unread.
  assert.deepEqual(snapshotRead({ paragraphs: [{ index: 0, text: "a" }, p(1, "b", "b")] }), {
    unreviewed: [0],
    fields: null,
  });
  // A field seen in a reviewed text wins over a pane that found none in the XML.
  assert.equal(
    snapshotRead({ paragraphs: [p(0, "x", `${B}PAGE${S}1${E}`)], has_fields: false }).fields,
    true,
  );
});

test("the XML of a paragraph: complex fields, simple fields and hyperlinks are fields; plain runs are not", () => {
  const pkg = (body) =>
    `<pkg:package><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body><w:p>${body}</w:p></w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/styles.xml"><pkg:xmlData><w:fldSimple w:instr="PAGE"/></pkg:xmlData></pkg:part></pkg:package>`;
  assert.equal(xmlHasField(pkg("<w:r><w:t>plain</w:t></w:r>")), false);
  assert.equal(xmlHasField(pkg('<w:r><w:fldChar w:fldCharType="begin"/></w:r>')), true);
  assert.equal(xmlHasField(pkg('<w:r><w:fldChar w:fldCharType="end"/></w:r>')), true);
  assert.equal(xmlHasField(pkg('<w:fldSimple w:instr="REF a"><w:r/></w:fldSimple>')), true);
  assert.equal(xmlHasField(pkg('<w:hyperlink r:id="rId5"><w:r/></w:hyperlink>')), true);
  assert.equal(
    xmlHasField(pkg("<w:del><w:r><w:delInstrText>REF a</w:delInstrText></w:r></w:del>")),
    true,
  );
  assert.equal(xmlHasField(""), false);
});

test("an edit in a paragraph whose reviewed text is missing is refused before Word, with a sentence of its own", async () => {
  for (const fields of [true, null]) {
    const d = deps({ fields_read: { unreviewed: [1], fields, unknown: false } });
    const r = await applyTrackedEdit(d, edit("slowly", "at a low rate"), TURN);
    assert.deepEqual(
      [r.ok, r.code, r.reason, r.paragraph_index, r.applied],
      [false, "FieldsUnread", "unread", 1, undefined],
    );
    assert.equal(r.message, FIELDS_UNREAD);
    assert.match(r.message, /Tell the author to press Refresh status/);
    assert.match(r.message, /Nothing was changed/);
    nothingSent(d);
    assert.deepEqual(d.lines, ["[review] fields unread R1.2 paragraph=1 by=daemon"]);
    // The acceptance run counts it as a refusal of its own kind.
    assert.deepEqual(refusalsFromLog(d.lines.join("\n"), "R1.2"), { FieldsUnread: 1 });
  }
  // A paragraph before the missing one was read: it is edited as before.
  const before = deps({ fields_read: { unreviewed: [5], fields: true, unknown: false } });
  assert.equal((await applyTrackedEdit(before, edit("slowly", "at a low rate"), TURN)).ok, true);
  assert.equal(before.calls.length, 1);
  // A document known to hold no field: the edit goes to Word as before.
  const plain = deps({ fields_read: { unreviewed: [1], fields: false, unknown: false } });
  assert.equal((await applyTrackedEdit(plain, edit("slowly", "at a low rate"), TURN)).ok, true);
  assert.equal(plain.calls.length, 1);
  // A snapshot from before the record of what was read: refused until a refresh.
  const old = deps({ fields_read: { unreviewed: null, fields: null, unknown: true } });
  assert.equal(
    (await applyTrackedEdit(old, edit("slowly", "at a low rate"), TURN)).code,
    "FieldsUnread",
  );
  nothingSent(old);
});

test("a rewrite of a paragraph whose reviewed text is missing is refused the same way", async () => {
  const d = deps({ fields_read: { unreviewed: [1], fields: null, unknown: false } });
  const r = await applyTrackedRewrite(
    d,
    rewrite("U1", "Cells grew at a low rate in slim flasks."),
    TURN,
  );
  assert.deepEqual(
    [r.ok, r.code, r.applied_parts, r.paragraph_index],
    [false, "FieldsUnread", 0, 1],
  );
  assert.equal(r.message, FIELDS_UNREAD);
  nothingSent(d);
  assert.deepEqual(d.lines, [
    "[review] rewrite R1.2 paragraph=1 parts=0 applied=0 stopped=FieldsUnread",
  ]);
  assert.deepEqual(rewritesFromLog(d.lines.join("\n"), "R1.2"), [
    { paragraph: 1, parts: 0, applied: 0, stopped: "FieldsUnread" },
  ]);
  const plain = deps({ fields_read: { unreviewed: [1], fields: false, unknown: false } });
  const ok = await applyTrackedRewrite(
    plain,
    rewrite("U1", "Cells grew at a low rate in slim flasks."),
    TURN,
  );
  assert.equal(ok.ok, true);
  assert.ok(plain.calls.length >= 1);
});

test("a field over three paragraphs: the middle one is refused for a tracked edit and for a rewrite, with its reviewed text present", async () => {
  // The edit: the carry from the first paragraph puts the whole middle one inside the field.
  const e = deps();
  const r = await applyTrackedEdit(e, edit("A lag in slim vessels", "A lag"), TURN);
  assert.deepEqual(
    [r.ok, r.code, r.reason, r.paragraph_index],
    [false, "FieldNotEditable", "inside", 3],
  );
  nothingSent(e);
  // The rewrite: the task pane would see no field in the middle paragraph
  // (no marker in its XML), so the daemon must not send it one part.
  for (const [id, index, text] of [
    ["U3", 3, "Belloc, M. (2020). A delay in slim vessels, measured again."],
    ["U2", 2, "Arden, Q. (2019). Shaken flasks, a second look at them."],
    ["U4", 4, "Corvane, I. (2021). The broth, not the vessel, decides."],
  ]) {
    const d = deps();
    const w = await applyTrackedRewrite(d, rewrite(id, text), TURN);
    assert.deepEqual(
      [w.ok, w.code, w.applied_parts, w.paragraph_index],
      [false, "ParagraphNotPlain", 0, index],
      id,
    );
    assert.match(w.message, /part of a field that runs over several paragraphs/);
    assert.match(w.message, /is not rewritten as a whole/);
    assert.match(w.message, /Nothing was changed/);
    nothingSent(d);
    assert.deepEqual(d.lines, [
      `[review] rewrite R1.2 paragraph=${index} parts=0 applied=0 stopped=ParagraphNotPlain`,
    ]);
    assert.deepEqual(refusalsFromLog(d.lines.join("\n"), "R1.2"), { ParagraphNotPlain: 1 });
  }
  // The paragraph after the list is plain again.
  const after = deps();
  const ok = await applyTrackedRewrite(
    after,
    rewrite("U5", "A closing sentence, now a little longer than before."),
    TURN,
  );
  assert.equal(ok.ok, true);
});

// ---- through the service: a refresh records what was read ----------------------

async function world() {
  const dir = await mkdtemp(join(tmpdir(), "fields-unread-"));
  const doc = join(dir, "paper.docx");
  const calls = [];
  const callTool = async (name, args) => {
    calls.push(args);
    return {
      ok: true,
      word_comment_id: 9,
      paragraph_index: args.expected_paragraph.index,
      before: "Cells grew slowly in narrow flasks.",
      after: "Cells grew at a low rate in narrow flasks.",
      target_checked: true,
    };
  };
  const svc = createReviewService({
    getContext: () => ({
      activeDoc: doc,
      caps: ["target_hash", "author_insertion_guard", "tracked_rewrite", FIELD_CAP],
    }),
    callTool,
    pushTurn: () => {},
    sendEvent: () => {},
    log: () => {},
    env: { CC_OFFICE_WSL_PATHS: "0" },
  });
  const ask = (type, payload = {}) =>
    new Promise((resolve) =>
      svc.handlers[type]({ type, request_id: "q", ...payload }, resolve, "k", "word"),
    );
  await ask("review_save_comments", { comments: [{ ref: "R1.2", reviewer: 1, text: "Reword." }] });
  const apply = () =>
    applyTrackedEdit(
      { ...svc.trackedEditDeps("k"), call: callTool },
      edit("slowly", "at a low rate"),
      TURN,
    );
  const file = async () => JSON.parse(await readFile(join(dir, "paper.revision.json"), "utf8"));
  return { ask, apply, calls, file, cleanup: () => rm(dir, { recursive: true, force: true }) };
}
const snapshotOf = (reviewed, extra = {}) => ({
  paragraphs: ["Title", "Cells grew slowly in narrow flasks."].map((text, index) => ({
    index,
    id: `U${index}`,
    text,
    reviewed_text: reviewed ? text : null,
    w_ins: 0,
    w_del: 0,
  })),
  comments: [],
  ...extra,
});

test("a refresh without reviewed text: the edit is refused while fields are there or not known, and goes in once the pane says there are none", async () => {
  const w = await world();
  try {
    // The pane could not read the reviewed text and found a field in the XML.
    await w.ask("review_refresh", { snapshot: snapshotOf(false, { has_fields: true }) });
    assert.deepEqual((await w.file()).snapshot.unreviewed, [0, 1]);
    assert.equal((await w.file()).snapshot.fields, true);
    assert.equal((await w.apply()).code, "FieldsUnread");
    // An older pane that does not say: not known, refused as well.
    await w.ask("review_refresh", { snapshot: snapshotOf(false) });
    assert.equal((await w.file()).snapshot.fields, null);
    assert.equal((await w.apply()).code, "FieldsUnread");
    assert.equal(w.calls.length, 0);
    // No field in the XML of any paragraph: nothing to cut through.
    await w.ask("review_refresh", { snapshot: snapshotOf(false, { has_fields: false }) });
    assert.equal((await w.apply()).ok, true);
    assert.equal(w.calls.length, 1);
    // The usual case: the reviewed text was read.
    await w.ask("review_refresh", { snapshot: snapshotOf(true, { has_fields: true }) });
    const stored = (await w.file()).snapshot;
    assert.deepEqual([stored.unreviewed, stored.fields], [[], true]);
  } finally {
    await w.cleanup();
  }
});

test("a review file from before the record of what was read loads and saves as it was", () => {
  const before = {
    taken_at: "2026-03-01T10:00:00.000Z",
    paragraphs: ["Title", "Cells grew slowly in narrow flasks."],
    ids: ["U0", "U1"],
  };
  const parsed = RevisionSchema.parse({
    version: 1,
    manuscript: "paper.docx",
    updated_at: "2026-03-01T10:00:00.000Z",
    snapshot: before,
  });
  // The two members get no default: the snapshot is written back as it was.
  assert.deepEqual(parsed.snapshot, before);
  assert.equal(JSON.stringify(parsed.snapshot), JSON.stringify(before));
});
