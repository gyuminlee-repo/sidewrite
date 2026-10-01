// Format rules for a manuscript (P11): italics of statistics symbols, of
// species and genus names and of gene names, subscripts in formulas,
// first-line indents and empty tables. Pure functions over what the task
// pane reads: no model, no Word, no file.
//
// Paragraph: {
//   index,
//   runs: [{ text, italic, sub, sup }],
//   indent?: { first_line, left, source },   first_line in points, or
//                                            { chars: n }, or null (unread);
//                                            source "direct" | "style"
//   in_table?, is_list?,
//   table?: { rows, cols, empty? | cells?: string[] }   on the paragraph
//                                            that stands for a table
// }
// The journal profile is the only source of style. No style rule runs unless
// the profile turns it on; with an empty profile the only finding left is an
// empty table, which is a defect of content, not a matter of style.
//   italics.statistics: true           statistics symbols in italics
//   italics.genera, italics.epithets   species and genus names in italics
//   italics.genes                      gene names in italics
//   subscripts                         digits of these formulas in subscript
//   first_line_indent                  "indented" | "none" | "consistent"
//                                      (outliers of the document's own clear
//                                      majority), or { expect, min_pt,
//                                      skip_after_heading }
// Also read: ref_section_head, ref_section_end_heads, legend_heads (where
// the reference list and the captions are; they turn nothing on).
//
// Finding: { paragraph_index, rule, find?, start?, spans?, want, found?,
// context?, locate? }. spans are the characters to change (paragraph
// offsets). locate.text occurs once in the paragraph and locate.offset is
// where `find` starts inside it, so a single letter can be found again.
//
// Called by the Tools tab (aux-service.mjs format_check) with the pane's
// format snapshot and the heading list. The findings are reported only.

import { referenceSection, paraText, wordsIn } from "./ref-section.mjs";

// A statistics symbol: one letter P, p, Q, q before a comparison, "value(s)"
// or "adjusted"; t before "("; n before = or a comparison. Not after a
// letter, a digit or a full stop, so author initials ("P.E.") and strain
// names ("Pm02") do not match.
export const STAT =
  /(?<![\p{L}\p{N}.])(?:[pqPQ](?=\s*(?:[<>=≤≥]|-?values?\b|-?adjusted))|t(?=\()|n(?=\s*[=≤≥]))/gu;

const esc = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
const list = (x) => (Array.isArray(x) ? x.filter((s) => typeof s === "string" && s) : []);
const alt = (names) =>
  [...names]
    .sort((a, b) => b.length - a.length)
    .map(esc)
    .join("|");
// A listed name must not be part of a longer word ("lacZY", "sCO2").
const whole = (body) => new RegExp(`(?<![A-Za-z0-9])(?:${body})(?![A-Za-z0-9])`, "gu");

// "Genus epithet", "G. epithet" and a genus on its own. An epithet on its
// own is not reported ("coli" is too easily another word).
function speciesPattern(genera, epithets) {
  if (!genera.length) return null;
  const genus = alt(genera);
  if (!epithets.length) return whole(genus);
  const initials = [...new Set(genera.map((g) => g[0]))].map(esc).join("|");
  const epithet = alt(epithets);
  return whole(`(?:${genus})(?:\\s+(?:${epithet}))?|(?:${initials})\\.\\s?(?:${epithet})`);
}

function flatten(p) {
  let text = "";
  const italic = [];
  const sub = [];
  for (const r of p.runs ?? []) {
    const t = String(r.text ?? "");
    text += t;
    for (let i = 0; i < t.length; i++) {
      italic.push(!!r.italic);
      sub.push(!!r.sub);
    }
  }
  return { text, italic, sub };
}

// The shortest stretch of text around [start, start + length) that occurs
// once in the paragraph.
function locate(text, start, length) {
  let a = start;
  let b = start + length;
  for (;;) {
    const piece = text.slice(a, b);
    if (text.indexOf(piece) === a && text.indexOf(piece, a + 1) === -1) break;
    if (a === 0 && b === text.length) break;
    a = Math.max(0, a - 1);
    b = Math.min(text.length, b + 1);
  }
  return { text: text.slice(a, b), offset: start - a };
}

function textFindings(p, rules) {
  const { text, italic, sub } = flatten(p);
  const out = [];
  for (const { rule, re, flag, want, digitsOnly } of rules) {
    for (const m of text.matchAll(re)) {
      const start = m.index;
      const spans = [];
      if (digitsOnly) {
        for (const d of m[0].matchAll(/\d+/g)) {
          spans.push({ start: start + d.index, length: d[0].length });
        }
      } else {
        spans.push({ start, length: m[0].length });
      }
      const marks = flag === "italic" ? italic : sub;
      const done = spans.every((s) => {
        for (let i = s.start; i < s.start + s.length; i++) {
          if (!marks[i] && /\S/.test(text[i])) return false;
        }
        return true;
      });
      if (done || !spans.length) continue;
      out.push({
        paragraph_index: p.index,
        rule,
        find: m[0],
        start,
        spans,
        want,
        context: text.slice(Math.max(0, start - 20), start + m[0].length + 20),
        locate: locate(text, start, m[0].length),
      });
    }
  }
  return out.sort((a, b) => a.start - b.start);
}

const plainText = paraText;
const words = wordsIn;

// Indexes of the reference list: its heading line and every paragraph in it.
function referenceIndexes(paragraphs, profile) {
  const { heads, entries } = referenceSection(paragraphs, profile);
  return new Set([...heads, ...entries.map((p) => p.index)]);
}

const CAPTION = /^(?:Fig(?:ure|\.)?|Table)\s*S?\d+/i;

// The paragraphs the indent rule judges: running text with indent data. Not
// headings, captions, list items, table cells, reference entries, lines under
// 20 words or lines that do not end a sentence.
export function bodyParagraphs(paragraphs, profile = {}, { headings = [] } = {}) {
  const heads = new Set(headings.map((h) => h.index));
  const refs = referenceIndexes(paragraphs, profile);
  const captions = [
    CAPTION,
    ...Object.values(profile.legend_heads ?? {})
      .filter((s) => typeof s === "string" && s)
      .map((s) => new RegExp(s, "i")),
  ];
  return paragraphs.filter((p) => {
    if (!p.indent || p.table || p.in_table || p.is_list) return false;
    if (heads.has(p.index) || refs.has(p.index)) return false;
    const t = plainText(p);
    if (words(t) < 20 || !/[.!?]["'”’)\]]*$/.test(t)) return false;
    return !captions.some((re) => re.test(t));
  });
}

// +1 indented, 0 flush or hanging, null unread.
function side(first) {
  if (first === null || first === undefined) return null;
  const n = typeof first === "object" ? first.chars : first;
  return typeof n === "number" && Number.isFinite(n) ? (n > 0 ? 1 : 0) : null;
}

function indentFindings(paragraphs, profile, { headings = [], styleIndentsKnown = true } = {}) {
  const setting = profile.first_line_indent ?? null;
  const opt = typeof setting === "object" && setting !== null ? setting : { expect: setting };
  const heads = new Set(headings.map((h) => h.index));
  const after = new Set(
    paragraphs.filter((_, i) => i > 0 && heads.has(paragraphs[i - 1].index)).map((p) => p.index),
  );
  const judged = bodyParagraphs(paragraphs, profile, { headings }).filter(
    (p) => side(p.indent.first_line) !== null && (styleIndentsKnown || p.indent.source !== "style"),
  );
  let expect = opt.expect ?? null;
  if (expect === "consistent") {
    expect = null;
    const indented = judged.filter((p) => side(p.indent.first_line) === 1).length;
    if (judged.length && indented / judged.length >= 0.8) expect = "indented";
    else if (judged.length && (judged.length - indented) / judged.length >= 0.8) expect = "none";
  }
  if (expect !== "indented" && expect !== "none") return new Map();
  const minPt = typeof opt.min_pt === "number" ? opt.min_pt : null;
  const seen = new Map();
  for (const p of judged) {
    const v = p.indent.first_line;
    if (typeof v === "number" && v > 0) seen.set(v, (seen.get(v) ?? 0) + 1);
  }
  const common = [...seen].sort((a, b) => b[1] - a[1] || a[0] - b[0])[0]?.[0] ?? null;
  const wanted = expect === "none" ? 0 : (minPt ?? common);
  const out = new Map();
  for (const p of judged) {
    const v = p.indent.first_line;
    const s = side(v);
    if (expect === "indented" && opt.skip_after_heading && after.has(p.index)) continue;
    const small = minPt !== null && typeof v === "number" && v < minPt;
    const bad = expect === "indented" ? s === 0 || small : s === 1;
    if (!bad) continue;
    out.set(p.index, {
      paragraph_index: p.index,
      rule: "indent",
      want: { first_line_indent: wanted },
      found: v,
      context: plainText(p).slice(0, 60),
    });
  }
  return out;
}

function emptyTable(p) {
  const t = p.table;
  if (!t) return null;
  const empty = Array.isArray(t.cells)
    ? t.cells.every((c) => !String(c ?? "").replace(/[\s\u0007]/g, ""))
    : t.empty === true;
  if (!empty) return null;
  return {
    paragraph_index: p.index,
    rule: "empty_table",
    want: { table: "filled or removed" },
    found: { rows: t.rows, cols: t.cols },
  };
}

export function findFormatIssues(paragraphs, profile = {}, extra = {}) {
  const italics = profile.italics ?? {};
  const species = speciesPattern(list(italics.genera), list(italics.epithets));
  const genes = list(italics.genes);
  const formulas = list(profile.subscripts);
  const rules = [
    italics.statistics === true && {
      rule: "stat",
      re: STAT,
      flag: "italic",
      want: { italic: true },
    },
    species && { rule: "species", re: species, flag: "italic", want: { italic: true } },
    genes.length && { rule: "gene", re: whole(alt(genes)), flag: "italic", want: { italic: true } },
    formulas.length && {
      rule: "subscript",
      re: whole(alt(formulas)),
      flag: "sub",
      want: { sub: true },
      digitsOnly: true,
    },
  ].filter(Boolean);
  const refs = referenceIndexes(paragraphs, profile);
  const indents = indentFindings(paragraphs, profile, extra);
  const out = [];
  for (const p of paragraphs) {
    if (!refs.has(p.index)) out.push(...textFindings(p, rules));
    if (indents.has(p.index)) out.push(indents.get(p.index));
    const table = emptyTable(p);
    if (table) out.push(table);
  }
  return out;
}
