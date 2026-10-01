// What the format checks read from Word (daemon/format-check.mjs), as pure
// functions: the runs of a paragraph with their italic and sub/superscript
// marks from the paragraph's OOXML, and the paragraph list the daemon gets.
//
// Only formatting set on the run itself is read (w:i, w:vertAlign in its
// w:rPr). Italics that come from a character or paragraph style are not
// seen, so text in such a style is reported as not italic.

import { RUN_CHAR, xmlText } from "./review-view.js";

const DOC_PART = '<pkg:part pkg:name="/word/document.xml"';
const OFF = /\bw:val="(?:0|false|off)"/;

// Is the toggle property <w:tag/> on in this w:rPr?
function on(rPr, tag) {
  const m = new RegExp(`<w:${tag}(?=[\\s/>])([^>]*)>`).exec(rPr);
  return !!m && !OFF.test(m[1]);
}

const RUN = /<w:r(?=[\s>])[^>]*>([\s\S]*?)<\/w:r>/g;
const RUN_TEXT =
  /<w:t(?=[\s>])[^>]*>([\s\S]*?)<\/w:t>|<w:(noBreakHyphen|softHyphen|tab|ptab|br|cr|sym)(?=[\s/>])[^>]*>/g;

// [{ text, italic, sub, sup }] of one paragraph, as the text reads with its
// tracked changes accepted: deleted and moved-away runs are left out, and so
// is what is not in the line of text (drawings, text boxes, embedded objects)
// and the formatting a tracked change replaced.
export function runsFromOoxml(xml) {
  let s = String(xml ?? "");
  const at = s.indexOf(DOC_PART);
  if (at !== -1) {
    const end = s.indexOf("</pkg:part>", at);
    s = s.slice(at, end === -1 ? undefined : end);
  }
  s = s
    .replace(/<mc:AlternateContent(?=[\s>])[\s\S]*?<\/mc:AlternateContent>/g, "")
    .replace(/<w:(drawing|pict|object)(?=[\s>])[\s\S]*?<\/w:\1>/g, "")
    .replace(/<w:rPrChange(?=[\s>])[\s\S]*?<\/w:rPrChange>/g, "")
    // A mark on a paragraph end or a table row is an empty element.
    .replace(/<w:(del|ins|moveFrom|moveTo)(?=[\s/>])[^>]*\/>/g, "")
    .replace(/<w:(del|moveFrom)(?=[\s>])[^>]*>[\s\S]*?<\/w:\1>/g, "");
  const runs = [];
  for (const m of s.matchAll(RUN)) {
    const rPr = /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(m[1])?.[1] ?? "";
    const body = m[1].replace(/<w:rPr>[\s\S]*?<\/w:rPr>/, "");
    let text = "";
    for (const t of body.matchAll(RUN_TEXT)) text += t[2] ? RUN_CHAR[t[2]] : xmlText(t[1]);
    if (!text) continue;
    const align = /<w:vertAlign(?=[\s/>])[^>]*\bw:val="(\w+)"/.exec(rPr)?.[1];
    const run = {
      text,
      italic: on(rPr, "i"),
      sub: align === "subscript",
      sup: align === "superscript",
    };
    const last = runs.at(-1);
    if (last && last.italic === run.italic && last.sub === run.sub && last.sup === run.sup)
      last.text += text;
    else runs.push(run);
  }
  return runs;
}

const points = (n) => (typeof n === "number" && Number.isFinite(n) ? n : null);

// The paragraph list of a format check.
// items: one per paragraph of the body, in order: { ooxml, first_line, left,
//   is_list, table_level } (first_line and left in points, as Word reports
//   them with the style applied; table_level 0 outside a table).
// tables: the tables of the body in order: { level, rows, values }.
// A table is put on its first paragraph. Two tables never touch in Word, so
// the n-th table of level 1 starts where the paragraphs enter a table for
// the n-th time. When the two counts differ the tables are left out and
// tables_read is false.
export function formatParagraphs(items, tables = []) {
  const paragraphs = items.map((p, index) => ({
    index,
    runs: runsFromOoxml(p.ooxml),
    indent: { first_line: points(p.first_line), left: points(p.left), source: "direct" },
    in_table: p.table_level > 0,
    is_list: !!p.is_list,
  }));
  const starts = [];
  items.forEach((p, i) => {
    if (p.table_level > 0 && !(items[i - 1]?.table_level > 0)) starts.push(i);
  });
  const top = tables.filter((t) => t.level === 1);
  const tables_read = starts.length === top.length;
  if (tables_read) {
    top.forEach((t, k) => {
      const values = Array.isArray(t.values) ? t.values : [];
      paragraphs[starts[k]].table = {
        rows: t.rows,
        cols: Math.max(0, ...values.map((row) => row.length)),
        cells: values.flat().map(String),
      };
    });
  }
  return { paragraphs, tables_read };
}
