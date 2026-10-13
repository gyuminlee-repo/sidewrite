// A tracked edit next to a citation field (the defect of the cite acceptance
// run of 2026-10-02: the visible text of an EndNote citation was deleted on
// its own, the field stayed with an empty result). The rule: an edit leaves
// every field as it was or removes one as a whole. Tested here on both
// sides, without Word:
//   the daemon (review-tools.mjs applyTrackedEdit) refuses from the text of
//     the last status refresh, which shows a field as
//     U+0013 instruction U+0014 result U+0015;
//   the task pane (tools-word.js toolTrackedEdit) refuses from its own
//     reviewed text, and after an edit reads the fields of the paragraph
//     from its XML and takes the edit back when one is not whole.
// The Word here is a stub that does what each test tells it to: what the
// real Word does with Range.delete() over a field is not known from here.
import { test } from "node:test";
import assert from "node:assert/strict";
import {
  applyTrackedEdit,
  applyTrackedRewrite,
  currentTexts,
  editTarget,
  FIELD_CAP,
} from "../daemon/review-tools.mjs";
import { paragraphHash } from "../daemon/doc-diff.mjs";
import {
  fieldView,
  FIELD_REFUSAL,
  undoOutcome,
  UNDO_TOOK_MORE,
} from "../taskpane/shared/field-text.js";
import { refusalsFromLog } from "../scripts/e2e/replay-lib.mjs";

const B = "\u0013";
const S = "\u0014";
const E = "\u0015";
const zotero = (shown, id) =>
  `${B} ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"${id}","properties":{"plainCitation":"${shown}"}} ${S}${shown}${E}`;
const endnote = (shown) =>
  `${B} ADDIN EN.CITE ${B} ADDIN EN.CITE.DATA \u0001${E}\u0001${S}${shown}${E}`;

// The paragraphs as a status refresh gives them (reviewed text).
const BELLOC = `Earlier invented work reported slow growth ${zotero("(Belloc, 2020)", "c2")} in narrow flasks at low shaking speed.`;
const HALLIVANE = `Optical density was read every hour. The reader had been calibrated against dry weight in an earlier invented study ${endnote("(Hallivane, 2016)")}. Samples were kept on ice.`;
const BIBL = [
  `${B} ADDIN ZOTERO_BIBL {"uncited":[]} CSL_BIBLIOGRAPHY ${S}Arden, Q. (2019). Shaken flasks.`,
  "Belloc, M. (2020). A lag in slim vessels.",
  `Corvane, I. (2021). The medium, not the flask.${E}`,
];
const DOC = ["Title", BELLOC, HALLIVANE, "A plain paragraph.", ...BIBL];
const IDS = DOC.map((_, k) => `U${k}`);
const SENTENCE =
  " The reader had been calibrated against dry weight in an earlier invented study (Hallivane, 2016).";

// ---- the daemon ---------------------------------------------------------------

const R = { what: "w", why: "y", addresses: "a" };
const edit = (find, replace, extra = {}) => ({
  comment_ref: "R1.2",
  find,
  replace,
  rationale: R,
  ...extra,
});
const TURN = { ref: "R1.2" };
function deps(
  result,
  { caps = ["target_hash", "author_insertion_guard", "tracked_rewrite", FIELD_CAP], ctx = {} } = {},
) {
  const d = {
    calls: [],
    edits: [],
    attempts: 0,
    notApplied: [],
    lines: [],
    log: (s) => d.lines.push(s),
    getLedger: async () => [],
    getReviewContext: async () => ({ paragraphs: DOC, paragraph_ids: IDS, ...ctx }),
    paneChecksTarget: () => caps.includes("target_hash"),
    paneCan: (cap) => caps.includes(cap),
    onAttempt: async () => ((d.attempts += 1), { id: "R1.2-b1" }),
    call: async (name, args) => {
      d.calls.push(args);
      return typeof result === "function" ? result(args) : result;
    },
    onEdit: async ({ change }) => (d.edits.push(change), { id: "R1.2-1" }),
    onNotApplied: async (x) => d.notApplied.push(x.attempt_id),
    onStaleTarget: async () => {},
  };
  return d;
}
const applied = (i) => ({
  ok: true,
  word_comment_id: 7,
  paragraph_index: i,
  before: fieldView(DOC[i]).text,
  after: "after",
  target_checked: true,
});

test("daemon: the paragraph of a find is told in the text as Word shows it, across a field", () => {
  const texts = currentTexts(DOC, []);
  // The sentence with its citation: not in the text with the instruction in it.
  assert.equal(texts[2].includes(SENTENCE), false);
  assert.deepEqual(editTarget(texts, IDS, SENTENCE, undefined), { index: 2, code: null });
  assert.deepEqual(editTarget(texts, IDS, "(Belloc, 2020) in narrow", undefined), {
    index: 1,
    code: null,
  });
  // Words of an instruction are not text of the document.
  assert.equal(editTarget(texts, IDS, "ADDIN ZOTERO_ITEM", undefined).code, "NotFound");
  assert.equal(editTarget(texts, IDS, "citationID", undefined).code, "NotFound");
  // Without a field nothing changes.
  assert.deepEqual(editTarget(texts, IDS, "plain paragraph", undefined), { index: 3, code: null });
});

test("daemon: an edit in the text of a field, across its edge, or a replacement over it is refused before Word is called", async () => {
  const cases = [
    // The three edits of the run: the second one hollowed the field.
    ["(Hallivane, 2016)", "", "inside", 2],
    ["(Belloc, 2020)", "(Belloc and others, 2020)", "inside", 1],
    ["growth (Belloc,", "rise (Belloc;", "partial", 1],
    [
      "slow growth (Belloc, 2020) in narrow flasks",
      "reduced growth (Belloc, 2020) in slim flasks",
      "replace-across",
      1,
    ],
    // In a list that is one field over several paragraphs.
    ["A lag in slim vessels", "A lag", "inside", 5],
    ["The medium, not the flask", "The medium", "inside", 6],
  ];
  for (const [find, replace, reason, index] of cases) {
    const d = deps(applied(index));
    const r = await applyTrackedEdit(d, edit(find, replace), TURN);
    assert.deepEqual(
      [r.ok, r.code, r.reason, r.paragraph_index, r.applied],
      [false, "FieldNotEditable", reason, index, undefined],
      find,
    );
    assert.ok(r.message.startsWith(FIELD_REFUSAL), find);
    assert.match(r.message, /Nothing was changed\.$/);
    assert.deepEqual([d.calls.length, d.attempts, d.edits.length], [0, 0, 0], find);
    assert.deepEqual(d.lines, [
      `[review] field edit refused R1.2 why=${reason} by=daemon paragraph=${index}`,
    ]);
    assert.deepEqual(refusalsFromLog(d.lines.join("\n"), "R1.2"), { FieldNotEditable: 1 });
  }
  // An insertion right after the text of a citation.
  const d = deps(applied(1));
  const r = await applyTrackedEdit(
    d,
    { comment_ref: "R1.2", find: "(Belloc, 2020)", insert_after: " and others", rationale: R },
    TURN,
  );
  assert.deepEqual([r.code, r.reason, d.calls.length], ["FieldNotEditable", "inside", 0]);
});

test("daemon: the words around a field go to Word as before, with the hash of the paragraph as the pane reads it", async () => {
  const cases = [
    ["slow growth", "reduced growth", 1],
    [" in narrow flasks", " in slim flasks", 1],
    ["Samples were kept on ice", "Samples were held on ice", 2],
    // The find holds the citation, the words that differ do not: narrowed.
    ["slow growth (Belloc, 2020)", "reduced growth (Belloc, 2020)", 1],
    ["A plain paragraph", "A plain one", 3],
  ];
  for (const [find, replace, index] of cases) {
    const d = deps(applied(index));
    const r = await applyTrackedEdit(d, edit(find, replace), TURN);
    assert.equal(r.ok, true, `${find}: ${r.message}`);
    assert.equal(d.calls.length, 1);
    assert.equal(d.calls[0].find, find);
    // The task pane hashes getReviewedText of the paragraph, instruction and all.
    assert.equal(d.calls[0].expected_before_hash, paragraphHash(DOC[index]), find);
    assert.deepEqual(
      d.lines.filter((l) => /field edit/.test(l)),
      [],
    );
  }
  const d = deps(applied(1));
  await applyTrackedEdit(
    d,
    edit("slow growth (Belloc, 2020)", "reduced growth (Belloc, 2020)"),
    TURN,
  );
  assert.deepEqual(d.calls[0].narrow, { find: "slow", replace: "reduced" });
});

test("daemon: a whole sentence with its citation, deleted in one edit, is sent to a task pane that guards fields", async () => {
  let d = deps(applied(2));
  let r = await applyTrackedEdit(d, edit(SENTENCE, ""), TURN);
  assert.equal(r.ok, true, r.message);
  assert.deepEqual([d.calls.length, d.calls[0].find, d.calls[0].replace], [1, SENTENCE, ""]);
  assert.equal(d.calls[0].expected_before_hash, paragraphHash(HALLIVANE));
  assert.deepEqual(d.calls[0].expected_paragraph, { index: 2, id: "U2" });
  assert.equal(d.edits[0].kind, "delete");
  // A task pane loaded before the daemon was updated checks no fields: it
  // is not sent a deletion over one.
  d = deps(applied(2), { caps: ["target_hash", "author_insertion_guard", "tracked_rewrite"] });
  r = await applyTrackedEdit(d, edit(SENTENCE, ""), TURN);
  assert.deepEqual(
    [r.ok, r.code, r.reason, d.calls.length],
    [false, "FieldNotEditable", "pane-outdated", 0],
  );
  assert.match(r.message, /reload the task pane/);
  // ... while the words around a field still go to it.
  r = await applyTrackedEdit(d, edit("slow growth", "reduced growth"), TURN);
  assert.equal(r.ok, true);
});

test("daemon: after such a deletion is recorded the paragraph reads without the field, and the next edit is judged on that", async () => {
  const gone = { kind: "delete", find: SENTENCE, text: "", paragraph_index: 2 };
  const texts = currentTexts(DOC, [gone]);
  assert.equal(texts[2], "Optical density was read every hour. Samples were kept on ice.");
  // A change next to a field leaves its instruction where it is.
  const next = currentTexts(DOC, [
    { kind: "replace", find: "slow growth", text: "reduced growth", paragraph_index: 1 },
  ]);
  assert.equal(next[1], BELLOC.replace("slow growth", "reduced growth"));
  // A narrowed change is put in by the find and replace the model asked for.
  const narrowed = currentTexts(DOC, [
    {
      kind: "replace",
      find: "slow",
      text: "reduced",
      paragraph_index: 1,
      narrowed_from: {
        find: "slow growth (Belloc, 2020)",
        text: "reduced growth (Belloc, 2020)",
      },
    },
  ]);
  assert.equal(fieldView(narrowed[1]).text.includes("reduced growth (Belloc, 2020) in"), true);
  const d = deps(applied(2), { ctx: { recent_changes: [gone] } });
  const r = await applyTrackedEdit(d, edit("kept on ice", "held on ice"), TURN);
  assert.equal(r.ok, true, r.message);
  assert.equal(d.calls[0].expected_before_hash, paragraphHash(texts[2]));
});

test("daemon: what the task pane answers about a field is logged; an edit it could not take back is recorded and the model is told to leave it", async () => {
  // Refused by the pane after the edit, taken back.
  let d = deps({
    ok: false,
    code: "FieldNotEditable",
    reason: "after-edit",
    detail: "hollowed",
    undone: true,
    paragraph_index: 2,
    message: `${FIELD_REFUSAL} Word did not keep the field whole, so the edit was taken back. (hollowed) Nothing was changed.`,
  });
  let r = await applyTrackedEdit(d, edit(SENTENCE, ""), TURN);
  assert.deepEqual(
    [r.ok, r.code, r.reason, r.applied],
    [false, "FieldNotEditable", "after-edit", undefined],
  );
  assert.deepEqual([d.edits.length, d.notApplied], [0, ["R1.2-b1"]]);
  assert.deepEqual(d.lines, [
    "[review] field edit refused R1.2 why=after-edit by=pane paragraph=2 found=hollowed undone=1",
  ]);
  assert.deepEqual(refusalsFromLog(d.lines.join("\n"), "R1.2"), { FieldNotEditable: 1 });
  // Not taken back: the change is in the document, so it is on record.
  d = deps({
    ok: false,
    applied: true,
    code: "FieldEditNotUndone",
    detail: "count 2->0",
    message: "The edit is in the document ...",
    paragraph_index: 2,
    before: fieldView(HALLIVANE).text,
    after: "Optical density was read every hour. Samples were kept on ice.",
    target_checked: true,
  });
  r = await applyTrackedEdit(d, edit(SENTENCE, ""), TURN);
  assert.deepEqual([r.ok, r.applied, r.comment_missing], [true, true, "FieldEditNotUndone"]);
  assert.deepEqual([d.edits.length, d.edits[0].comment_missing], [1, "FieldEditNotUndone"]);
  assert.match(r.warning, /did not leave a citation or another field .* whole \(count 2->0\)/);
  assert.match(r.warning, /Do not repeat it and do not try to repair it/);
  assert.match(r.warning, /reject change R1\.2-1 in Word/);
  assert.ok(
    d.lines.includes("[review] field edit not undone R1.2-1 paragraph=2 found=count_2->0"),
    d.lines.join("\n"),
  );
});

test("daemon: a paragraph with fields is still not rewritten as a whole", async () => {
  const d = deps((args) =>
    args.plain_only
      ? { ok: false, code: "ParagraphNotPlain", message: "The paragraph holds a field." }
      : applied(1),
  );
  const r = await applyTrackedRewrite(
    d,
    {
      comment_ref: "R1.2",
      paragraph_id: "U1",
      text: BELLOC.replace("slow growth", "reduced growth").replace("narrow", "slim"),
      what: "w",
      why: "y",
      addresses: "a",
    },
    TURN,
  );
  assert.deepEqual([r.ok, r.code, r.applied_parts], [false, "ParagraphNotPlain", 0]);
  assert.equal(d.calls.length, 1);
  assert.equal(d.calls[0].plain_only, true);
  assert.equal(d.edits.length, 0);
});

// ---- the task pane ------------------------------------------------------------

const run = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const fld = (kind) => `<w:r><w:fldChar w:fldCharType="${kind}"/></w:r>`;
const instr = (t) => `<w:r><w:instrText xml:space="preserve">${t}</w:instrText></w:r>`;
const del = (xml) =>
  `<w:del w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z">${xml
    .replaceAll("w:instrText", "w:delInstrText")
    .replaceAll("<w:t ", "<w:delText ")
    .replaceAll("</w:t>", "</w:delText>")}</w:del>`;
const pkg = (...parts) =>
  `<pkg:package><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body><w:p>${parts.join("")}</w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
const CODE = [
  fld("begin"),
  instr(" ADDIN EN.CITE "),
  fld("begin"),
  instr(" ADDIN EN.CITE.DATA "),
  fld("end"),
  fld("separate"),
].join("");
const HEAD = run("Optical density was read every hour.");
const WORDS = run(SENTENCE.slice(0, SENTENCE.indexOf("(Hallivane")));
const SHOWN = run("(Hallivane, 2016)");
const TAIL = run(" Samples were kept on ice.");
// The paragraph of the run in its XML: as it was, with the field whole in
// one deletion, and with its result alone in a deletion.
const XML = {
  live: pkg(HEAD, WORDS, CODE, SHOWN, fld("end"), run("."), TAIL),
  whole: pkg(HEAD, del(WORDS + CODE + SHOWN + fld("end") + run(".")), TAIL),
  hollow: pkg(HEAD, WORDS, CODE, del(SHOWN), fld("end"), run("."), TAIL),
  split: pkg(HEAD, del(WORDS + CODE), SHOWN, fld("end"), run("."), TAIL),
  nextTo: pkg(HEAD, WORDS, CODE, SHOWN, fld("end"), run("."), del(TAIL), run(" Samples on ice.")),
  // Word deleted the field outright, without a tracked change.
  gone: pkg(HEAD, TAIL),
};

// A Word with the one paragraph of the run after a title.
//   word(op)   what Word makes of an edit: the key of XML the paragraph has
//              after it, and (reviewed) its reviewed text
//   api16      Office says WordApi 1.6 is there (tracked changes of a range)
//   rejects    true: rejecting this edit restores the paragraph; false:
//              it changes nothing; { xml, reviewed }: what the paragraph is
//              after it (Word took back less, or more, than the edit)
//   reviewed   what getReviewedText gives before the edit
//   ooxml      false: Paragraph.getOoxml throws
function fakeWord({
  word = () => ({ xml: "live" }),
  api16 = true,
  rejects = true,
  reviewed = HALLIVANE,
  ooxml = true,
  pending = [],
  afterChanges,
  trackedChanges = true,
} = {}) {
  const shown = fieldView(HALLIVANE).text;
  const s = {
    ops: [],
    comments: [],
    rejected: 0,
    rejectAllCalls: 0,
    rejectedIds: [],
    mode: "Off",
    xml: "live",
    reviewed,
    changes: pending.map((c, i) => ({ id: `prior-${i}`, ...c })),
  };
  const restore = () => {
    if (rejects === true) Object.assign(s, { xml: "live", reviewed });
    else if (rejects) Object.assign(s, rejects);
  };
  const coll = (items) => ({ items, load() {} });
  const hits = (text, q) => {
    const out = [];
    for (let i = text.indexOf(q); i !== -1; i = text.indexOf(q, i + q.length)) out.push(i);
    return out;
  };
  const did = (op) => {
    s.ops.push(op);
    const own = [
      ...(op.where === "After" ? [] : [{ id: "own-delete", type: "Deleted", text: op.on }]),
      ...(op.text ? [{ id: "own-insert", type: "Added", text: op.text }] : []),
    ];
    s.changes.push(...(afterChanges ? afterChanges(own) : own));
    const next = word(op);
    s.xml = next.xml;
    if (next.reviewed !== undefined) s.reviewed = next.reviewed;
  };
  let para;
  const range = (text) => ({
    text,
    load() {},
    paragraphs: { getFirst: () => para },
    search: (q) => coll(hits(text, q).map(() => range(q))),
    insertText: (t, where) => (did({ op: "insertText", on: text, text: t, where }), range(t)),
    delete: () => did({ op: "delete", on: text }),
    insertComment: (c) => (s.comments.push(c), { id: `C${s.comments.length}`, load() {} }),
    compareLocationWith: () => ({ value: "Before" }),
    getReviewedText: () => ({ value: s.reviewed }),
    getTrackedChanges() {
      if (!trackedChanges) throw new Error("Tracked changes unavailable");
      return {
        load() {},
        items: s.changes.map((c) => ({
          // WordApi 1.6 has type/text/getRange/reject, no stable change ID.
          type: c.type,
          text: c.text,
          getRange: () => ({
            compareLocationWith: (r) => ({
              value: c.location ?? (r.text === c.text ? "Equal" : "Before"),
            }),
          }),
          reject() {
            s.rejected += 1;
            s.rejectedIds.push(c.id);
            if (rejects !== false) {
              s.changes = s.changes.filter((x) => x !== c);
              if (!s.changes.some((x) => x.id.startsWith("own-"))) restore();
            }
          },
        })),
        // Model the old unsafe operation too: it also rejects preexisting
        // changes in this range. The regression fails against that code.
        rejectAll() {
          s.rejectAllCalls += 1;
          s.rejected += 1;
          s.rejectedIds.push(...s.changes.map((c) => c.id));
          if (rejects !== false) s.changes = [];
          restore();
        },
      };
    },
  });
  para = {
    ...range(shown),
    text: shown,
    style: "Normal",
    uniqueLocalId: "U1",
    getRange: () => range(shown),
    getOoxml() {
      if (!ooxml) throw new Error("GeneralException: getOoxml");
      return { value: XML[s.xml] };
    },
  };
  const title = { text: "Title", style: "Title", uniqueLocalId: "U0", load() {} };
  title.search = () => coll([]);
  const document = {
    load() {},
    get changeTrackingMode() {
      return s.mode;
    },
    set changeTrackingMode(v) {
      s.mode = v;
    },
    getSelection: () => ({ select() {} }),
    body: { paragraphs: coll([title, para]), search: (q) => para.search(q) },
  };
  globalThis.Word = {
    run: async (fn) => fn({ document, sync: async () => {} }),
    InsertLocation: { replace: "Replace", after: "After" },
    ChangeTrackingMode: { trackAll: "TrackAll" },
    ChangeTrackingVersion: { current: "Current" },
  };
  globalThis.Office = {
    context: {
      requirements: { isSetSupported: (set, v) => set === "WordApi" && v === "1.6" && api16 },
    },
  };
  return s;
}

const { toolTrackedEdit, PANE_CAPS } = await import("../taskpane/shared/tools-word.js");
const call = (find, replace, extra = {}) =>
  toolTrackedEdit({
    find,
    replace,
    comment_text: "[R1.2] w / y / a",
    expected_before_hash: paragraphHash(HALLIVANE),
    ...extra,
  });
const untouched = (s) => [s.ops.length, s.comments.length, s.rejected, s.mode];

test("task pane: it names the capability the daemon asks for before a deletion over a field", () => {
  assert.ok(PANE_CAPS.includes(FIELD_CAP));
});

test("task pane: the visible text of a citation alone is not deleted, replaced or cut into", async () => {
  const cases = [
    ["(Hallivane, 2016)", "", "inside"],
    ["(Hallivane, 2016)", "(Hallivane and others, 2016)", "inside"],
    ["Hallivane", "Halivane", "inside"],
    ["study (Hallivane,", "work (Hallivane;", "partial"],
    ["2016). Samples", "2017). Samples", "partial"],
    [SENTENCE, " The reader was calibrated earlier.", "replace-across"],
  ];
  for (const [find, replace, reason] of cases) {
    const s = fakeWord();
    const r = await call(find, replace);
    assert.deepEqual(
      [r.ok, r.code, r.reason, r.paragraph_index, r.applied],
      [false, "FieldNotEditable", reason, 1, undefined],
      find,
    );
    assert.ok(r.message.startsWith(FIELD_REFUSAL));
    assert.deepEqual(
      untouched(s),
      [0, 0, 0, "Off"],
      `${find}: nothing changed, track mode as it was`,
    );
  }
  // An insertion at the end of the text of a citation.
  const s = fakeWord();
  const r = await toolTrackedEdit({
    find: "(Hallivane, 2016)",
    insert_after: " and others",
    comment_text: "[R1.2] w / y / a",
    expected_before_hash: paragraphHash(HALLIVANE),
  });
  assert.deepEqual([r.code, r.reason, s.ops.length], ["FieldNotEditable", "inside", 0]);
});

test("task pane: a whole sentence with its citation is deleted as one range, and the field lies whole in the deletion", async () => {
  const s = fakeWord({ word: () => ({ xml: "whole" }) });
  const r = await call(SENTENCE, "");
  assert.equal(r.ok, true, r.message);
  // Range.delete() on the range the search found, not a replacement by "".
  assert.deepEqual(s.ops, [{ op: "delete", on: SENTENCE }]);
  assert.deepEqual([s.rejected, s.comments.length, s.mode], [0, 1, "Off"]);
  assert.deepEqual([r.paragraph_index, r.target_checked, r.word_comment_id], [1, true, "C1"]);
  assert.equal(r.after, "Optical density was read every hour. Samples were kept on ice.");
});

test("task pane: when Word does not put the field into the deletion as a whole, the edit is taken back and refused", async () => {
  for (const [xml, found] of [
    ["hollow", "hollowed"],
    ["split", "split"],
    ["gone", "count 2->0"],
  ]) {
    const s = fakeWord({ word: () => ({ xml }) });
    const r = await call(SENTENCE, "");
    assert.deepEqual(
      [r.ok, r.code, r.reason, r.detail, r.undone, r.applied],
      [false, "FieldNotEditable", "after-edit", found, true, undefined],
      xml,
    );
    assert.match(r.message, new RegExp(`taken back\\. \\(${found.replace(">", "\\>")}\\) Nothing`));
    assert.deepEqual(
      [s.ops.length, s.rejected, s.xml, s.comments.length, s.mode],
      [1, 1, "live", 0, "Off"],
    );
  }
});

test("task pane: an edit that broke a field and cannot be taken back is reported as applied, with what was found", async () => {
  const s = fakeWord({ word: () => ({ xml: "hollow" }), rejects: false });
  const r = await call(SENTENCE, "");
  assert.deepEqual(
    [r.ok, r.applied, r.code, r.detail],
    [false, true, "FieldEditNotUndone", "hollowed"],
  );
  assert.match(r.message, /could not be taken back\. Reject this change in Word\./);
  // What the daemon records a change from.
  assert.deepEqual([r.paragraph_index, r.paragraph_id, r.target_checked], [1, "U1", true]);
  assert.equal(r.before, fieldView(HALLIVANE).text);
  assert.deepEqual([s.rejected, s.comments.length, s.mode], [1, 0, "Off"]);
});

// ---- the undo: clean, still there, or more than the edit ------------------------

// The paragraph in its reviewed text at three moments. Before the edit it
// holds a pending change of an earlier comment ("every hour", counted as
// accepted); the edit hollows the field; an undo that takes the neighbour
// with it leaves the text the paragraph had before that earlier change.
const HOLLOWED = HALLIVANE.replace(`${S}(Hallivane, 2016)${E}`, `${S}${E}`);
const NEIGHBOUR_GONE = HALLIVANE.replace("every hour", "each hour");
const hollowing = () => ({ xml: "hollow", reviewed: HOLLOWED });

test("the outcome of an undo is told from the fields and three texts: before the edit, after it, after the undo", () => {
  const o = (fieldsOk, before, after, undone) => undoOutcome({ fieldsOk, before, after, undone });
  // The paragraph reads as before the edit and its fields are as before.
  assert.equal(o(true, "B", "A", "B"), "clean");
  // It still reads as after the edit, or a field is still not as it was.
  assert.equal(o(true, "B", "A", "A"), "kept");
  assert.equal(o(false, "B", "A", "B"), "kept");
  assert.equal(o(false, "B", "A", "X"), "kept");
  // The fields are as before and it reads as neither: more went than the edit.
  assert.equal(o(true, "B", "A", "X"), "more");
  // A text that could not be read: never "gone".
  assert.equal(o(true, null, "A", "X"), "kept");
  assert.equal(o(true, "B", "A", null), "kept");
  assert.equal(o(true, "B", null, "X"), "kept");
  assert.equal(o(true, "B", null, "B"), "clean");
});

test("task pane: a clean undo is refused as before, with nothing left of the edit", async () => {
  const s = fakeWord({ word: hollowing });
  const r = await call(SENTENCE, "");
  assert.deepEqual(
    [r.ok, r.code, r.reason, r.detail, r.undone, r.took_more, r.applied],
    [false, "FieldNotEditable", "after-edit", "hollowed", true, undefined, undefined],
  );
  assert.match(r.message, /Nothing was changed\.$/);
  assert.deepEqual([s.rejected, s.xml, s.reviewed, s.comments.length], [1, "live", HALLIVANE, 0]);
});

test("task pane: an edit that is still in the document after the undo is reported as applied", async () => {
  // Word rejected nothing: the field is still hollow, the text still the edited one.
  let s = fakeWord({ word: hollowing, rejects: false });
  let r = await call(SENTENCE, "");
  assert.deepEqual(
    [r.ok, r.applied, r.code, r.detail, r.took_more],
    [false, true, "FieldEditNotUndone", "hollowed", undefined],
  );
  assert.equal(s.reviewed, HOLLOWED);
  // The fields read whole again, the paragraph still reads as after the edit.
  s = fakeWord({ word: hollowing, rejects: { xml: "live" } });
  r = await call(SENTENCE, "");
  assert.deepEqual([r.ok, r.applied, r.code], [false, true, "FieldEditNotUndone"]);
  assert.match(r.message, /could not be taken back\. Reject this change in Word\./);
});

test("task pane: unexpected host over-undo still uses the defensive took_more postcondition", async () => {
  const s = fakeWord({ word: hollowing, rejects: { xml: "live", reviewed: NEIGHBOUR_GONE } });
  const r = await call(SENTENCE, "");
  assert.deepEqual(
    [r.ok, r.code, r.reason, r.detail, r.undone, r.took_more, r.applied],
    [false, "FieldNotEditable", "after-edit", "hollowed", true, true, undefined],
  );
  assert.deepEqual([r.paragraph_index, r.paragraph_id], [1, "U1"]);
  assert.ok(r.message.startsWith(FIELD_REFUSAL));
  assert.ok(r.message.endsWith(UNDO_TOOK_MORE));
  assert.match(r.message, /also took back an earlier pending change in the same paragraph/);
  // "Nothing was changed" would be wrong here.
  assert.doesNotMatch(r.message, /Nothing was changed/);
  assert.deepEqual(
    [s.rejected, s.xml, s.reviewed, s.comments.length, s.mode],
    [1, "live", NEIGHBOUR_GONE, 0, "Off"],
  );
});

test("daemon: an undo that took more is not recorded as applied; the attempt goes, a refresh is asked for, the model is told", async () => {
  const took = [];
  const d = deps({
    ok: false,
    code: "FieldNotEditable",
    reason: "after-edit",
    detail: "hollowed",
    undone: true,
    took_more: true,
    paragraph_index: 2,
    message: `${FIELD_REFUSAL} Word did not keep the field whole (hollowed), so the edit was taken back. ${UNDO_TOOK_MORE}`,
  });
  d.onUndoTookMore = async (x) => took.push([x.ref, x.paragraph_index]);
  const r = await applyTrackedEdit(d, edit(SENTENCE, ""), TURN);
  assert.deepEqual(
    [r.ok, r.code, r.took_more, r.applied],
    [false, "FieldNotEditable", true, undefined],
  );
  assert.match(r.message, /do not repeat this edit/);
  assert.deepEqual([d.edits.length, d.notApplied], [0, ["R1.2-b1"]]);
  assert.deepEqual(took, [["R1.2", 2]]);
  assert.deepEqual(d.lines, [
    "[review] field edit refused R1.2 why=after-edit by=pane paragraph=2 found=hollowed undone=1 took_more=1",
  ]);
  // A clean undo asks for no refresh.
  const clean = deps({
    ok: false,
    code: "FieldNotEditable",
    reason: "after-edit",
    detail: "hollowed",
    undone: true,
    paragraph_index: 2,
    message: "m",
  });
  clean.onUndoTookMore = async (x) => took.push(x.ref);
  await applyTrackedEdit(clean, edit(SENTENCE, ""), TURN);
  assert.equal(took.length, 1);
});

test("task pane: without a way to take the edit back, or to read the paragraph, a deletion over a field is not tried", async () => {
  for (const world of [{ api16: false }, { ooxml: false }]) {
    const s = fakeWord({ word: () => ({ xml: "whole" }), ...world });
    const r = await call(SENTENCE, "");
    assert.deepEqual([r.ok, r.code, r.reason], [false, "FieldNotEditable", "no-undo"]);
    assert.deepEqual(untouched(s), [0, 0, 0, "Off"]);
    // The words around the field still go in.
    const t = fakeWord({ word: () => ({ xml: "nextTo" }), ...world });
    const ok = await call("Samples were kept on ice", "Samples on ice");
    assert.equal(ok.ok, true);
    assert.equal(t.ops.length, 1);
  }
  // An Office object that is not there at all.
  const s = fakeWord({ word: () => ({ xml: "whole" }) });
  delete globalThis.Office;
  const r = await call(SENTENCE, "");
  assert.deepEqual([r.code, r.reason, s.ops.length], ["FieldNotEditable", "no-undo", 0]);
});

test("task pane: an edit of the words next to a field goes in as before, and the fields are read again after it", async () => {
  const s = fakeWord({ word: () => ({ xml: "nextTo" }) });
  const r = await call("Samples were kept on ice", "Samples on ice");
  assert.equal(r.ok, true, r.message);
  assert.deepEqual(s.ops, [
    { op: "insertText", on: "Samples were kept on ice", text: "Samples on ice", where: "Replace" },
  ]);
  assert.deepEqual([s.rejected, s.comments.length], [0, 1]);
  // Narrowed: the words that differ, inside the hit.
  const n = fakeWord({ word: () => ({ xml: "nextTo" }) });
  const narrowed = await call(
    "an earlier invented study (Hallivane, 2016)",
    "a previous invented study (Hallivane, 2016)",
    {
      narrow: { find: "an earlier", replace: "a previous" },
    },
  );
  assert.deepEqual([narrowed.ok, narrowed.narrowed], [true, true]);
  assert.deepEqual(
    n.ops.map((o) => [o.op, o.on, o.text]),
    [["insertText", "an earlier", "a previous"]],
  );
});

test("task pane: a Word whose reviewed text does not show the fields is still held to the rule by the check after the edit", async () => {
  // getReviewedText without the marks: the place of the edit among the
  // fields cannot be told, so the edit of the run goes to Word ...
  const plain = fieldView(HALLIVANE).text;
  const s = fakeWord({ reviewed: plain, word: () => ({ xml: "hollow" }) });
  const r = await call("(Hallivane, 2016)", "", { expected_before_hash: paragraphHash(plain) });
  // ... and comes back: the XML of the paragraph shows the hollowed field.
  assert.deepEqual(
    [r.ok, r.code, r.reason, r.detail, r.undone],
    [false, "FieldNotEditable", "after-edit", "hollowed", true],
  );
  // One deletion, no insertion for the empty replacement; reject once.
  assert.deepEqual([s.ops.length, s.rejected, s.xml, s.comments.length], [1, 1, "live", 0]);
  // A replacement that turned a field into words (the field whole in the
  // deletion, its text typed again) is not a removal the edit named.
  const f = fakeWord({ reviewed: plain, word: () => ({ xml: "whole" }) });
  const flat = await call(SENTENCE, " The reader was calibrated earlier (Hallivane, 2016).", {
    expected_before_hash: paragraphHash(plain),
  });
  assert.deepEqual(
    [flat.code, flat.reason, flat.detail, f.rejected],
    ["FieldNotEditable", "after-edit", "removed 1", 2],
  );
});

// These are synthetic Office.js contract tests, not Word acceptance tests.
test("task pane: a pending change anywhere in the field paragraph prevents a deletion needing undo", async () => {
  for (const type of ["Added", "Deleted", "Formatted"]) {
    const s = fakeWord({ pending: [{ type, text: "every hour" }], word: hollowing });
    const prior = [...s.changes];
    const r = await call(SENTENCE, "");
    assert.deepEqual([r.code, r.reason, r.applied], ["FieldNotEditable", "no-undo", undefined]);
    assert.deepEqual(untouched(s), [0, 0, 0, "Off"]);
    assert.deepEqual(s.changes, prior);
  }
});

test("task pane: a field-breaking replacement cannot reject a preexisting overlapping change", async () => {
  const s = fakeWord({
    pending: [{ type: "Added", text: "Samples were kept on ice" }],
    word: hollowing,
  });
  // The old words are that whole pending insertion: an earlier change of
  // the tool (own_insertions), else the edit is refused before Word is
  // touched (WholeAuthorInsertion, comment-anchor-dropped.test.mjs).
  const r = await call("Samples were kept on ice", "Samples on ice", {
    own_insertions: ["Samples were kept on ice"],
  });
  assert.deepEqual([r.applied, r.code, s.rejected], [true, "FieldEditNotUndone", 0]);
  assert.ok(s.changes.some((c) => c.id === "prior-0"));
  assert.equal(s.mode, "Off");
});

test("task pane: an unavailable revision collection cannot authorize field-deletion undo", async () => {
  const s = fakeWord({ trackedChanges: false, word: hollowing });
  const r = await call(SENTENCE, "");
  assert.deepEqual(
    [r.code, r.reason, ...untouched(s)],
    ["FieldNotEditable", "no-undo", 0, 0, 0, "Off"],
  );
});

test("task pane: every undo candidate must match this edit's type, text and range before any rejection", async () => {
  const uncertain = [
    (own) => [...own, { id: "concurrent", type: "Added", text: "user words" }],
    (own) => own.map((c) => ({ ...c, text: c.text + " merged user words" })),
    (own) => own.map((c) => ({ ...c, location: "Contains" })),
    (own) => own.map((c) => ({ ...c, location: "OverlapsBefore" })),
    (own) => own.map((c) => ({ ...c, type: "Formatted" })),
    (own) => own.map((c) => ({ ...c, location: "Before" })),
    () => [],
  ];
  for (const afterChanges of uncertain) {
    const s = fakeWord({ word: hollowing, afterChanges });
    const r = await call(SENTENCE, "");
    assert.deepEqual([r.applied, r.code, s.rejected], [true, "FieldEditNotUndone", 0]);
    assert.equal(s.mode, "Off");
  }
});

test("task pane: an isolated rollback rejects each owned change once and never calls rejectAll", async () => {
  const s = fakeWord({ word: hollowing });
  const r = await call("Samples were kept on ice", "Samples on ice");
  assert.equal(r.undone, true);
  assert.deepEqual(s.rejectedIds, ["own-delete", "own-insert"]);
  assert.equal(s.rejected, 2);
  assert.equal(s.rejectAllCalls, 0);
  assert.deepEqual(s.changes, []);
});

test("task pane: an isolated insertion rollback rejects only the inserted change", async () => {
  const s = fakeWord({ word: hollowing });
  const r = await toolTrackedEdit({
    find: "Samples were kept on ice",
    insert_after: " for transport",
    comment_text: "[R1.2] w / y / a",
    expected_before_hash: paragraphHash(HALLIVANE),
  });
  assert.equal(r.undone, true);
  assert.deepEqual(s.rejectedIds, ["own-insert"]);
  assert.equal(s.rejectAllCalls, 0);
});
