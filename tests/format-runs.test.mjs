// What the pane reads for the format checks (taskpane/shared/format-runs.js).
import { test } from "node:test";
import assert from "node:assert/strict";
import { runsFromOoxml, formatParagraphs } from "../taskpane/shared/format-runs.js";
import { findFormatIssues } from "../daemon/format-check.mjs";

const pkg = (body) =>
  `<pkg:package><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body><w:p>${body}</w:p></w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/styles.xml"><w:r><w:t>not the document</w:t></w:r></pkg:part></pkg:package>`;
const run = (text, rPr = "") =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const plain = (text, more = {}) => ({ text, italic: false, sub: false, sup: false, ...more });

test("runs carry italic, subscript and superscript; equal neighbours are joined", () => {
  const xml = pkg(
    run("The value ") +
      run("p", "<w:i/>") +
      run(" &lt; 0.05 for CO") +
      run("2", '<w:vertAlign w:val="subscript"/>') +
      run(" and m") +
      run("2", '<w:vertAlign w:val="superscript"/>') +
      run(" in ") +
      run("vi", "<w:b/><w:i/>") +
      run("vo", '<w:i w:val="true"/>'),
  );
  assert.deepEqual(runsFromOoxml(xml), [
    plain("The value "),
    plain("p", { italic: true }),
    plain(" < 0.05 for CO"),
    plain("2", { sub: true }),
    plain(" and m"),
    plain("2", { sup: true }),
    plain(" in "),
    plain("vivo", { italic: true }),
  ]);
});

test("w:i switched off, w:iCs and the formatting a tracked change replaced are not italic", () => {
  const xml = pkg(
    run("a", '<w:i w:val="0"/>') +
      run("b", '<w:i w:val="false"/>') +
      run("c", "<w:iCs/>") +
      run("d", '<w:rPrChange w:id="1" w:author="x"><w:rPr><w:i/></w:rPr></w:rPrChange>'),
  );
  assert.deepEqual(runsFromOoxml(xml), [plain("abcd")]);
});

test("deleted and moved-away text is left out, inserted text is kept, a marked paragraph end eats nothing", () => {
  const xml = pkg(
    '<w:pPr><w:rPr><w:del w:id="9" w:author="x"/></w:rPr></w:pPr>' +
      run("kept ") +
      `<w:del w:id="1" w:author="x"><w:r><w:delText>gone</w:delText></w:r></w:del>` +
      `<w:moveFrom w:id="2" w:author="x">${run("moved")}</w:moveFrom>` +
      `<w:ins w:id="3" w:author="x">${run("new", "<w:i/>")}</w:ins>` +
      run(" end"),
  );
  assert.deepEqual(runsFromOoxml(xml), [
    plain("kept "),
    plain("new", { italic: true }),
    plain(" end"),
  ]);
});

test("tabs and breaks are the characters Word gives; a text box and a field code are not text", () => {
  const xml = pkg(
    "<w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/></w:r>" +
      `<w:r><w:drawing><wp:inline><w:txbxContent><w:p>${run("in a box")}</w:p></w:txbxContent></wp:inline></w:drawing></w:r>` +
      '<w:r><w:instrText xml:space="preserve"> REF x </w:instrText></w:r>' +
      run("c"),
  );
  assert.deepEqual(runsFromOoxml(xml), [plain("a\tb\u000bc")]);
  assert.deepEqual(runsFromOoxml(""), []);
  assert.deepEqual(runsFromOoxml(null), []);
});

const item = (text, more = {}) => ({
  ooxml: pkg(run(text)),
  first_line: 0,
  left: 0,
  is_list: false,
  table_level: 0,
  ...more,
});

test("the paragraph list: index, indent in points, list and table flags, a table on its first paragraph", () => {
  const { paragraphs, tables_read } = formatParagraphs(
    [
      item("Body.", { first_line: 36 }),
      item("Item.", { is_list: true, first_line: -18, left: 36 }),
      item("", { table_level: 1 }),
      item("", { table_level: 1 }),
      item("After.", { first_line: Number.NaN }),
      item("x", { table_level: 1 }),
    ],
    [
      { level: 1, rows: 1, values: [["", ""]] },
      { level: 1, rows: 1, values: [["x"]] },
    ],
  );
  assert.equal(tables_read, true);
  assert.deepEqual(
    paragraphs.map((p) => [p.index, p.indent.first_line, p.in_table, p.is_list]),
    [
      [0, 36, false, false],
      [1, -18, false, true],
      [2, 0, true, false],
      [3, 0, true, false],
      [4, null, false, false],
      [5, 0, true, false],
    ],
  );
  assert.deepEqual(paragraphs[2].table, { rows: 1, cols: 2, cells: ["", ""] });
  assert.equal(paragraphs[3].table, undefined);
  assert.deepEqual(paragraphs[5].table, { rows: 1, cols: 1, cells: ["x"] });
  // What the daemon makes of it: the empty table is found, the full one is not.
  assert.deepEqual(
    findFormatIssues(paragraphs, {}).map((f) => [f.paragraph_index, f.rule]),
    [[2, "empty_table"]],
  );
});

test("tables that cannot be matched to the paragraphs are left out and said so", () => {
  const r = formatParagraphs(
    [item("a"), item("", { table_level: 1 })],
    [
      { level: 1, rows: 1, values: [[""]] },
      { level: 1, rows: 1, values: [[""]] },
    ],
  );
  assert.equal(r.tables_read, false);
  assert.ok(r.paragraphs.every((p) => p.table === undefined));
  // A nested table is not a table of the body's first level.
  const nested = formatParagraphs(
    [item("", { table_level: 1 }), item("", { table_level: 2 })],
    [
      { level: 1, rows: 1, values: [[""]] },
      { level: 2, rows: 1, values: [[""]] },
    ],
  );
  assert.equal(nested.tables_read, true);
});
