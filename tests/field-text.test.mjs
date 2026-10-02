// Fields in the text and the XML of a paragraph (taskpane/shared/field-text.js).
// The task pane and the daemon decide with these functions whether a tracked
// edit may be made next to a citation, and whether the fields of a paragraph
// are still whole after it. The inputs are written by hand in the form Word
// gave in the cite acceptance run of 2026-10-02: getReviewedText shows a
// field as U+0013 instruction U+0014 result U+0015, a hollowed one as
// U+0014 U+0015 with nothing between.
import { test } from "node:test";
import assert from "node:assert/strict";
import {
  FIELD_REFUSAL,
  fieldRefusal,
  fieldVerdict,
  fieldView,
  fieldViews,
  fieldsKept,
  hasFieldMarks,
  paragraphFields,
  spliceShown,
} from "../taskpane/shared/field-text.js";
import { normalizeParaText } from "../daemon/revision-classify.mjs";
import { refusalCode } from "../scripts/e2e/judge.mjs";

const B = "\u0013";
const S = "\u0014";
const E = "\u0015";
const zotero = (shown, id = "c1") =>
  `${B} ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"${id}","properties":{"plainCitation":"${shown}"}} ${S}${shown}${E}`;
// The EndNote citation with its record in w:fldData: a nested field without
// a result inside the instruction, as the run showed it.
const endnote = (shown) =>
  `${B} ADDIN EN.CITE ${B} ADDIN EN.CITE.DATA \u0001${E}\u0001${S}${shown}${E}`;

const BELLOC = `Earlier invented work reported slow growth ${zotero("(Belloc, 2020)")} in narrow flasks at low shaking speed.`;
const HALLIVANE = `Optical density was read every hour. The reader had been calibrated against dry weight in an earlier invented study ${endnote("(Hallivane, 2016)")}. Samples were kept on ice.`;
const TWO = `Oxygen transfer falls as the fill volume rises ${zotero("(Dellmar, 2018; Eskridge, 2022)", "c4")}. The invented strain grows fastest when a tenth of the flask is filled ${zotero("(Arden, 2019)", "c5")}.`;

test("the shown text of a paragraph leaves out the instruction and the marks, normalized as normalizeParaText", () => {
  const v = fieldView(BELLOC);
  assert.equal(
    v.text,
    "Earlier invented work reported slow growth (Belloc, 2020) in narrow flasks at low shaking speed.",
  );
  const at = v.text.indexOf("(Belloc, 2020)");
  assert.deepEqual(v.spans, [{ start: at, end: at + "(Belloc, 2020)".length }]);
  assert.equal(v.carry, 0);
  // Every shown character knows where it stands in the text it came from.
  assert.equal(v.raw.length, v.text.length);
  for (let i = 0; i < v.text.length; i++) {
    assert.ok(v.text[i] === " " || BELLOC[v.raw[i]] === v.text[i], String(i));
  }
  // The nested field of the EndNote form has no result: nothing of it shows.
  assert.equal(
    fieldView(HALLIVANE).text,
    "Optical density was read every hour. The reader had been calibrated against dry weight in an earlier invented study (Hallivane, 2016). Samples were kept on ice.",
  );
  // Without a field the view is normalizeParaText, whatever the spacing.
  for (const s of ["plain", "  a\u0005 b\t\tc \n", "", null, " nbsp  x ", "한글  문단"]) {
    assert.equal(fieldView(s).text, normalizeParaText(s), JSON.stringify(s));
    assert.deepEqual(fieldView(s).spans, []);
    assert.equal(hasFieldMarks(s), false);
  }
  assert.equal(hasFieldMarks(BELLOC), true);
  // Normalizing first changes nothing: the daemon holds normalized texts.
  assert.deepEqual(fieldView(normalizeParaText(HALLIVANE)).text, fieldView(HALLIVANE).text);
  // A hollowed field: a place, no text.
  const hollow = fieldView(`before ${B} ADDIN EN.CITE ${S}${E}. After`);
  assert.equal(hollow.text, "before . After");
  assert.deepEqual(hollow.spans, [{ start: 7, end: 7 }]);
});

test("a field over several paragraphs is carried from one paragraph to the next", () => {
  const texts = [
    "References",
    `${B} ADDIN ZOTERO_BIBL {"uncited":[]} CSL_BIBLIOGRAPHY ${S}Arden, Q. (2019). Shaken flasks.`,
    "Belloc, M. (2020). Slow growth.",
    `Corvane, I. (2021). The medium.${E}`,
    "After the list.",
  ];
  const views = fieldViews(texts);
  assert.deepEqual(
    views.map((v) => v.carry),
    [0, 1, 1, 0, 0],
  );
  assert.deepEqual(views[1].spans, [{ start: 0, end: Infinity }]);
  assert.deepEqual(views[2].spans, [{ start: 0, end: Infinity }]);
  assert.deepEqual(views[3].spans, [{ start: 0, end: views[3].text.length }]);
  assert.deepEqual(views[4].spans, []);
  // Every edit inside the list is inside a field.
  for (const [k, find] of [
    [1, "Shaken flasks"],
    [2, "Slow growth"],
    [3, "The medium"],
  ]) {
    assert.deepEqual(fieldVerdict(views[k], find, { replace: "x" }), {
      ok: false,
      reason: "inside",
    });
  }
  assert.deepEqual(fieldVerdict(views[4], "the list", { replace: "it" }), { ok: true, remove: 0 });
  // The last paragraph alone, as the task pane reads it: an end without a begin.
  assert.deepEqual(fieldVerdict(fieldView(texts[3]), "The medium", { replace: "x" }), {
    ok: false,
    reason: "inside",
  });
});

// Each row: [name, paragraph, find, edit, narrow, expected]. The expected
// answers were written before the function was run on them.
const NO = (reason) => ({ ok: false, reason });
const YES = (remove = 0) => ({ ok: true, remove });
const VERDICTS = [
  // ---- refused ----
  [
    "the text of a citation alone, deleted",
    HALLIVANE,
    "(Hallivane, 2016)",
    { replace: "" },
    null,
    NO("inside"),
  ],
  [
    "the text of a citation alone, replaced",
    BELLOC,
    "(Belloc, 2020)",
    { replace: "(Belloc and others, 2020)" },
    null,
    NO("inside"),
  ],
  [
    "a part of the text of a citation",
    BELLOC,
    "Belloc",
    { replace: "Bellock" },
    null,
    NO("inside"),
  ],
  [
    "it starts before the citation and ends inside it",
    BELLOC,
    "growth (Belloc,",
    { replace: "rise (Belloc;" },
    null,
    NO("partial"),
  ],
  [
    "it starts inside the citation and ends after it",
    BELLOC,
    "2020) in narrow",
    { replace: "2021) in slim" },
    null,
    NO("partial"),
  ],
  [
    "it starts at the first character of the citation",
    BELLOC,
    "(Belloc, 2020) in",
    { replace: "" },
    null,
    NO("partial"),
  ],
  [
    "it ends at the last character of the citation",
    BELLOC,
    "slow growth (Belloc, 2020)",
    { replace: "" },
    null,
    NO("partial"),
  ],
  [
    "a replacement over a whole citation",
    BELLOC,
    "slow growth (Belloc, 2020) in narrow flasks",
    { replace: "reduced growth in slim flasks" },
    null,
    NO("replace-across"),
  ],
  [
    "a replacement over a whole citation whose narrow form still holds it",
    BELLOC,
    "slow growth (Belloc, 2020) in narrow flasks",
    { replace: "reduced growth (Belloc, 2020) in slim flasks" },
    {
      find: "slow growth (Belloc, 2020) in narrow",
      replace: "reduced growth (Belloc, 2020) in slim",
    },
    NO("replace-across"),
  ],
  [
    "an insertion right after the text of a citation",
    BELLOC,
    "(Belloc, 2020)",
    { insert_after: " and others" },
    null,
    NO("inside"),
  ],
  [
    "an insertion in the middle of a citation",
    BELLOC,
    "(Belloc,",
    { insert_after: " M.," },
    null,
    NO("inside"),
  ],
  // ---- allowed ----
  ["words before the citation", BELLOC, "slow growth", { replace: "reduced growth" }, null, YES()],
  [
    "words right after the citation",
    BELLOC,
    " in narrow flasks",
    { replace: " in slim flasks" },
    null,
    YES(),
  ],
  [
    "the space and words before it, up to its first character",
    BELLOC,
    "slow growth ",
    { replace: "a lag " },
    null,
    YES(),
  ],
  ["the full stop after a citation", HALLIVANE, ". Samples", { replace: "; samples" }, null, YES()],
  [
    "a whole sentence with its citation, deleted",
    HALLIVANE,
    " The reader had been calibrated against dry weight in an earlier invented study (Hallivane, 2016).",
    { replace: "" },
    null,
    YES(1),
  ],
  [
    "a deletion over two citations",
    TWO,
    "rises (Dellmar, 2018; Eskridge, 2022). The invented strain grows fastest when a tenth of the flask is filled (Arden, 2019).",
    { replace: "" },
    null,
    YES(2),
  ],
  [
    "a replacement whose narrow form lies before the citation",
    BELLOC,
    "slow growth (Belloc, 2020)",
    { replace: "reduced growth (Belloc, 2020)" },
    { find: "slow", replace: "reduced" },
    YES(),
  ],
  [
    "an insertion after words that end before the citation",
    BELLOC,
    "slow growth",
    { insert_after: " and yield" },
    null,
    YES(),
  ],
  [
    "an insertion after an anchor that holds the whole citation",
    BELLOC,
    "growth (Belloc, 2020) in narrow flasks",
    { insert_after: " of glass" },
    null,
    YES(),
  ],
  [
    "a paragraph without a field",
    "Plain (Belloc, 2020) text.",
    "(Belloc, 2020)",
    { replace: "" },
    null,
    YES(),
  ],
];

test("where an edit lies among the fields: refused inside, across an edge and as a replacement over one; a deletion may take a field whole", () => {
  for (const [name, para, find, edit, narrow, want] of VERDICTS) {
    assert.deepEqual(fieldVerdict(fieldView(para), find, edit, narrow), want, name);
  }
  assert.ok(VERDICTS.filter((r) => !r[5].ok).length >= 5);
  assert.ok(VERDICTS.filter((r) => r[5].ok).length >= 5);
  // A hollowed field between the words of a deletion goes with it; at the
  // edge of the range it is left alone.
  const hollow = fieldView(`before ${B} ADDIN EN.CITE ${S}${E}. After`);
  assert.deepEqual(fieldVerdict(hollow, "before . After", { replace: "" }), YES(1));
  assert.deepEqual(fieldVerdict(hollow, ". After", { replace: "" }), YES());
  // A find that is not in the shown text once: nothing can be told here.
  for (const find of ["not there", "e"]) {
    assert.deepEqual(fieldVerdict(fieldView(BELLOC), find, { replace: "" }), {
      ok: true,
      remove: 0,
      unknown: true,
    });
  }
  // A narrow form that is in the find twice is not the one applied.
  assert.deepEqual(
    fieldVerdict(
      fieldView(BELLOC),
      "slow growth (Belloc, 2020) in narrow",
      { replace: "x" },
      { find: "o", replace: "0" },
    ),
    NO("replace-across"),
  );
});

test("the refusal names the field, says what to do, and is read as FieldNotEditable", () => {
  for (const reason of [
    "inside",
    "partial",
    "replace-across",
    "no-undo",
    "pane-outdated",
    "after-edit",
  ]) {
    const m = fieldRefusal(reason);
    assert.ok(m.startsWith(FIELD_REFUSAL), reason);
    assert.match(m, /Nothing was changed\.$/, reason);
    assert.ok(m.length > FIELD_REFUSAL.length + 30, `${reason} has its own sentence`);
    assert.equal(refusalCode(m), "FieldNotEditable", reason);
    assert.equal(refusalCode(`Not applied: ${m}`), "FieldNotEditable", reason);
  }
  assert.match(FIELD_REFUSAL, /delete the whole sentence in one edit/);
  assert.match(fieldRefusal("after-edit", "(hollowed)"), /taken back\. \(hollowed\) Nothing/);
});

test("a recorded change is put into a paragraph with fields in its shown text, and a deletion takes its field out whole", () => {
  // Next to the field: the instruction stays as it is.
  const reworded = spliceShown(BELLOC, "slow growth", "reduced growth");
  assert.equal(reworded, BELLOC.replace("slow growth", "reduced growth"));
  // A sentence with its citation: the field goes, marks and instruction.
  const find =
    " The reader had been calibrated against dry weight in an earlier invented study (Hallivane, 2016).";
  const cut = spliceShown(HALLIVANE, find, "");
  assert.equal(cut, "Optical density was read every hour. Samples were kept on ice.");
  assert.equal(hasFieldMarks(cut), false);
  // An insertion after an anchor before the field.
  assert.equal(
    spliceShown(BELLOC, "slow growth", " and yield", true),
    BELLOC.replace("slow growth", "slow growth and yield"),
  );
  // Not there, or there twice: null.
  assert.equal(spliceShown(BELLOC, "not there", "x"), null);
  assert.equal(spliceShown(BELLOC, "o", "0"), null);
});

// ---- the XML of a paragraph ----

const run = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const fld = (kind) => `<w:r><w:fldChar w:fldCharType="${kind}"/></w:r>`;
const instr = (t) => `<w:r><w:instrText xml:space="preserve">${t}</w:instrText></w:r>`;
const del = (xml) =>
  `<w:del w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z">${xml
    .replaceAll("w:instrText", "w:delInstrText")
    .replaceAll("<w:t ", "<w:delText ")
    .replaceAll("</w:t>", "</w:delText>")}</w:del>`;
const p = (...parts) =>
  `<w:p><w:pPr><w:rPr><w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr>${parts.join("")}</w:p>`;
// The EndNote citation of the run: [begin]{EN.CITE}[begin]{EN.CITE.DATA}[end][separate]result[end].
const CODE = [
  `<w:r><w:fldChar w:fldCharType="begin"><w:fldData xml:space="preserve">QUJD</w:fldData></w:fldChar></w:r>`,
  instr(" ADDIN EN.CITE "),
  fld("begin"),
  instr(" ADDIN EN.CITE.DATA "),
  fld("end"),
  fld("separate"),
].join("");
const SHOWN = run("(Hallivane, 2016)");
const LIVE = p(run("an earlier invented study "), CODE, SHOWN, fld("end"), run(". Samples."));
// The three deletions of the run: the words before, the result alone, the full stop.
const HOLLOW = p(
  del(run("an earlier invented study")),
  run(" "),
  CODE,
  del(SHOWN),
  fld("end"),
  del(run(". ")),
  run("Samples."),
);
const HOLLOW_ACCEPTED = p(run(" "), CODE, fld("end"), run("Samples."));
const WHOLE = p(
  del(run("an earlier invented study ") + CODE + SHOWN + fld("end") + run(". ")),
  run("Samples."),
);
const SPLIT = p(run("study "), del(CODE), SHOWN, fld("end"), run(". Samples."));
const PART = p(run("study "), CODE, run("(Hallivane, "), del(run("2016")), run(")"), fld("end"));
const NEXT_TO = p(
  del(run("an earlier")),
  `<w:ins w:id="8" w:author="A" w:date="2026-01-01T00:00:00Z">${run("a previous")}</w:ins>`,
  run(" invented study "),
  CODE,
  SHOWN,
  fld("end"),
  run(". Samples."),
);

test("the fields of a paragraph from its XML: live, deleted whole, split, hollowed", () => {
  const states = (xml) => paragraphFields(xml).map((f) => [f.code.trim(), f.state, f.text, f.all]);
  assert.deepEqual(
    [LIVE, WHOLE, SPLIT].map((xml) => paragraphFields(xml)[0].marks),
    ["-", "del", "-,del"],
  );
  assert.deepEqual(states(LIVE), [
    ["ADDIN EN.CITE", "live", "(Hallivane, 2016)", "(Hallivane, 2016)"],
    ["ADDIN EN.CITE.DATA", "live", "", ""],
  ]);
  // The mark of a deleted paragraph end in w:rPr (an empty w:del) opens nothing.
  assert.equal(paragraphFields(LIVE).stray, "");
  assert.deepEqual(states(HOLLOW)[0], ["ADDIN EN.CITE", "live", "", "(Hallivane, 2016)"]);
  assert.deepEqual(states(HOLLOW_ACCEPTED)[0], ["ADDIN EN.CITE", "live", "", ""]);
  assert.deepEqual(
    states(WHOLE).map((f) => f[1]),
    ["deleted", "deleted"],
  );
  assert.deepEqual(states(WHOLE)[0], ["ADDIN EN.CITE", "deleted", "", "(Hallivane, 2016)"]);
  assert.deepEqual(
    states(SPLIT).map((f) => f[1]),
    ["split", "deleted"],
  );
  assert.deepEqual(states(PART)[0], [
    "ADDIN EN.CITE",
    "live",
    "(Hallivane, )",
    "(Hallivane, 2016)",
  ]);
  // Inside the package Paragraph.getOoxml returns: only the document part is read.
  const pkg =
    `<pkg:package><pkg:part pkg:name="/word/styles.xml"><pkg:xmlData>${fld("begin")}</pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body>${LIVE}</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
  assert.deepEqual(states(pkg), states(LIVE));
  // A paragraph that starts inside a field (the last one of a bibliography)
  // and one that ends inside one.
  const tail = paragraphFields(p(run("Corvane, I. (2021)."), fld("end")));
  assert.deepEqual([tail.length, tail.stray], [0, "-"]);
  const head = paragraphFields(
    p(fld("begin"), instr(" ADDIN ZOTERO_BIBL "), fld("separate"), run("Arden")),
  );
  assert.deepEqual([head.length, head[0].state, head[0].text], [1, "open", "Arden"]);
  // Entities of an instruction are read as text.
  const amp = paragraphFields(
    p(fld("begin"), instr(" REF a&amp;b "), fld("separate"), run("x"), fld("end")),
  );
  assert.equal(amp[0].code, " REF a&b ");
  assert.deepEqual(paragraphFields(""), Object.assign([], { stray: "" }));
});

test("after an edit: every field as it was, or gone as a whole with a deletion; nothing in between", () => {
  const f = paragraphFields;
  // Unchanged, and an edit of the words next to the field.
  assert.equal(fieldsKept(f(LIVE), f(LIVE)), null);
  assert.equal(fieldsKept(f(LIVE), f(NEXT_TO)), null);
  // The defect of the run, pending and accepted.
  assert.equal(fieldsKept(f(LIVE), f(HOLLOW)), "hollowed");
  assert.equal(fieldsKept(f(LIVE), f(HOLLOW_ACCEPTED)), "hollowed");
  assert.equal(
    fieldsKept(f(LIVE), f(HOLLOW), 1),
    "hollowed",
    "a hollowed field is not a removed one",
  );
  // The field whole in the deletion: only when the edit was to remove it.
  // The nested field goes with its parent and is not counted on its own.
  assert.equal(fieldsKept(f(LIVE), f(WHOLE), 1), null);
  assert.equal(fieldsKept(f(LIVE), f(WHOLE), Infinity), null);
  assert.equal(fieldsKept(f(LIVE), f(WHOLE), 0), "removed 1");
  assert.deepEqual(
    f(LIVE).map((x) => x.nested),
    [false, true],
  );
  // Half of the field in a deletion, a result partly deleted.
  assert.equal(fieldsKept(f(LIVE), f(SPLIT), 2), "split");
  assert.equal(fieldsKept(f(LIVE), f(PART)), "result");
  // A field gone without a tracked change, another instruction.
  assert.equal(fieldsKept(f(LIVE), f(p(run("Samples.")))), "count 2->0");
  assert.equal(fieldsKept(f(LIVE), f(LIVE.replace("EN.CITE ", "EN.CITED "))), "code");
  // What was hollowed or deleted before the edit and is the same after it passes.
  assert.equal(fieldsKept(f(HOLLOW), f(HOLLOW)), null);
  assert.equal(fieldsKept(f(WHOLE), f(WHOLE)), null);
  // The end marker of a field that began before the paragraph, and the
  // begin of one that ends after it: in a deletion or gone, the field is half.
  const tail = p(run("Corvane."), fld("end"));
  assert.equal(fieldsKept(f(tail), f(tail)), null);
  assert.equal(fieldsKept(f(tail), f(p(del(run("Corvane.") + fld("end"))))), "split");
  assert.equal(fieldsKept(f(tail), f(p(run("Corvane.")))), "split");
  const head = [fld("begin"), instr(" ADDIN ZOTERO_BIBL "), fld("separate"), run("Arden")];
  assert.equal(fieldsKept(f(p(...head)), f(p(...head))), null);
  assert.equal(fieldsKept(f(p(...head)), f(p(del(head.join(""))))), "hollowed");
});
