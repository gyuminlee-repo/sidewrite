// Pure parts of marked-diff-word-check.mjs: which pairs of the corpus are
// checked in Word, how two structures that Word reported are compared, and
// how one pair is judged. No Word, no files.
//
// A structure is what marked-diff-word.ps1 reads from an open document:
// { paragraphs: [text], inline: [n], fields: [n], codes: [text],
//   tables: [{ rows, columns, uniform, cells: [text], at: ["row,column"] }],
//   bookmarks: [{ name, text }], comments, footnotes, endnotes }. inline
// and fields are, for every paragraph, the number of its inline shapes
// (pictures, objects) and of its fields; codes is, for every paragraph, the
// code and the result of each of its fields (code, U+001F, result, U+001E).
// comments, footnotes and endnotes are counts of the whole document. uniform is Word's Table.Uniform (false when cells are merged), at
// the row and column index of every cell: a cell that was merged into the
// one above it is not there. The texts are Word's own (Range.Text), so
// a paragraph ends in "\r", a cell in "\r\u0007", and breaks are the control
// characters U+000B (line), U+000C (page) and U+000E (column). Both sides of
// every comparison come from Word, so these conventions cancel out.

// The kinds of change the selection must cover.
export const KINDS = [
  "word",
  "paragraph-added",
  "paragraph-deleted",
  "paragraph-split",
  "paragraph-merged",
  "paragraph-moved",
  "empty-paragraph",
  "last-paragraph",
  "before-table",
  "after-table",
  "table-row",
  "table-column",
  "table-added",
  "table-deleted",
  "table-moved",
  "cell",
  "break-line",
  "break-page",
  "break-column",
  "break-clear",
  "break-deleted",
  "tab",
  "link",
  "bookmark",
  "pending-changes",
  "footnote-header",
  "cell-merge",
  "picture",
  "field",
  "citation",
  "field-result",
  "bibliography",
  "sign",
  "picture-deleted",
  "footnote",
  "equation-kept",
  "comment",
  "content-control",
  "simple-field",
];

// The pairs, by their name in tests/marked_diff_corpus.py, with the kinds each
// one covers. The list is fixed: the same pairs every run.
export const PAIRS = [
  ["a word changed", ["word"]],
  ["a paragraph added at the start", ["paragraph-added"]],
  ["a paragraph added in the middle", ["paragraph-added"]],
  ["a paragraph added at the end", ["paragraph-added", "last-paragraph"]],
  ["a paragraph added before a table", ["paragraph-added", "before-table"]],
  ["a paragraph added after a table", ["paragraph-added", "after-table"]],
  ["a paragraph deleted at the start", ["paragraph-deleted"]],
  ["a paragraph deleted in the middle", ["paragraph-deleted"]],
  ["a paragraph deleted at the end", ["paragraph-deleted", "last-paragraph"]],
  ["a paragraph deleted before a table", ["paragraph-deleted", "before-table"]],
  ["a paragraph deleted after a table", ["paragraph-deleted", "after-table"]],
  ["a paragraph split in two", ["paragraph-split"]],
  ["the last paragraph split in two", ["paragraph-split", "last-paragraph"]],
  ["two paragraphs merged", ["paragraph-merged"]],
  ["the last two paragraphs merged", ["paragraph-merged", "last-paragraph"]],
  ["a sentence moved to the next paragraph", ["paragraph-merged", "paragraph-split"]],
  ["an empty paragraph added", ["empty-paragraph"]],
  ["an empty paragraph removed", ["empty-paragraph"]],
  ["the last paragraph replaced", ["last-paragraph"]],
  ["a paragraph moved", ["paragraph-moved"]],
  ["a table row added", ["table-row"]],
  ["a table row removed", ["table-row"]],
  ["a table column added", ["table-column"]],
  ["a table added", ["table-added"]],
  ["the only table removed", ["table-deleted"]],
  ["a table moved between paragraphs", ["table-moved"]],
  ["a word changed in a cell", ["cell"]],
  ["a paragraph added in a cell", ["cell", "paragraph-added"]],
  ["the last paragraph of a cell deleted", ["cell", "last-paragraph"]],
  ["a line break in a changed paragraph", ["break-line"]],
  ["a page break in a changed paragraph", ["break-page"]],
  ["a column break in a changed paragraph", ["break-column"]],
  ["a break that clears floats in a changed paragraph", ["break-clear"]],
  ["a break written with clear none", ["break-clear"]],
  ["a page break removed", ["break-deleted", "break-page"]],
  ["a tab in a changed paragraph", ["tab"]],
  ["a link in a changed paragraph", ["link"]],
  ["a bookmark in a changed paragraph", ["bookmark"]],
  ["a pending insertion in the clean file", ["pending-changes"]],
  ["a footnote and a header, change elsewhere", ["footnote-header"]],
  ["a cell merged across two columns", ["cell-merge"]],
  ["a cell merged down a column", ["cell-merge"]],
  ["a merged cell split again", ["cell-merge"]],
  ["a paragraph with a picture added", ["picture", "paragraph-added"]],
  ["a paragraph with a field added", ["field", "paragraph-added"]],
  ["a word changed next to a citation", ["citation"]],
  ["words changed on both sides of a citation", ["citation"]],
  ["a citation added", ["citation"]],
  ["a citation deleted", ["citation"]],
  ["the shown text of a citation changed", ["citation", "field-result"]],
  ["a paragraph with citations deleted", ["citation", "paragraph-deleted"]],
  ["a paragraph with a citation added", ["citation", "paragraph-added"]],
  ["a word changed in a caption with a number and a cross-reference", ["field"]],
  ["the result of a cross-reference changed", ["field-result"]],
  ["a word changed in a paragraph with a hard hyphen, a symbol and a return", ["sign"]],
  ["a symbol taken out", ["sign"]],
  ["a word changed in a paragraph with a picture", ["picture"]],
  ["a picture swapped for another", ["picture-deleted"]],
  ["a paragraph with a picture deleted", ["picture-deleted", "paragraph-deleted"]],
  ["a table with a picture that lost a row", ["picture-deleted", "table-row"]],
  ["a footnote in the changed paragraph", ["footnote"]],
  ["a footnote in a deleted paragraph", ["footnote", "paragraph-deleted"]],
  ["a footnote swapped for another", ["footnote"]],
  ["a word changed in a paragraph with an equation", ["equation-kept"]],
  ["a word changed in a paragraph with a comment", ["comment"]],
  ["a comment added in a changed paragraph", ["comment"]],
  ["a comment only in the baseline, the paragraph changed", ["comment"]],
  ["the first entry of a bibliography reworded", ["bibliography"]],
  ["the first entry of a bibliography replaced", ["bibliography"]],
  ["an entry added to a bibliography", ["bibliography"]],
  ["a word changed in a paragraph with a simple field", ["simple-field"]],
  ["a simple field added to a paragraph", ["simple-field"]],
  ["a word changed in a paragraph with a content control", ["content-control"]],
].map(([name, covers]) => ({ name, covers }));

// Kinds of KINDS that no pair covers.
export const uncovered = (pairs = PAIRS) =>
  KINDS.filter((k) => !pairs.some((p) => p.covers.includes(k)));

// The first 80 characters of a text, with control characters made visible.
export const show = (text) => JSON.stringify(String(text ?? "").slice(0, 80));

function firstTextDifference(what, got, want) {
  const n = Math.min(got.length, want.length);
  for (let i = 0; i < n; i++) {
    if (got[i] !== want[i]) {
      return `${what} ${i + 1}: expected ${show(want[i])}, got ${show(got[i])}`;
    }
  }
  if (got.length !== want.length) {
    const more = got.length > want.length;
    const extra = more ? got[n] : want[n];
    return `${what} count: expected ${want.length}, got ${got.length}; ${what} ${n + 1} is ${more ? "extra" : "missing"}: ${show(extra)}`;
  }
  return null;
}

// The first paragraph whose count of pictures or fields differs.
function firstCountDifference(what, got, want) {
  for (let i = 0; i < want.length; i++) {
    if (got[i] !== want[i]) {
      return `paragraph ${i + 1} ${what}: expected ${want[i]}, got ${got[i]}`;
    }
  }
  return null;
}

// What differs between two structures, as one line, or null. `bookmarks`
// and `comments`: compare them too (the accept side: the output carries the
// clean document's).
export function compareStructure(got, want, { bookmarks = true, comments = true } = {}) {
  const paragraphs = firstTextDifference("paragraph", got.paragraphs ?? [], want.paragraphs ?? []);
  if (paragraphs) return paragraphs;
  for (const [key, what] of [
    ["inline", "pictures"],
    ["fields", "fields"],
  ]) {
    const d = firstCountDifference(what, got[key] ?? [], want[key] ?? []);
    if (d) return d;
  }
  const codes = firstTextDifference("paragraph fields", got.codes ?? [], want.codes ?? []);
  if (codes) return codes;
  for (const key of ["footnotes", "endnotes", ...(comments ? ["comments"] : [])]) {
    if ((got[key] ?? 0) !== (want[key] ?? 0)) {
      return `${key}: expected ${want[key] ?? 0}, got ${got[key] ?? 0}`;
    }
  }
  const [gt, wt] = [got.tables ?? [], want.tables ?? []];
  if (gt.length !== wt.length) return `table count: expected ${wt.length}, got ${gt.length}`;
  for (let i = 0; i < wt.length; i++) {
    for (const key of ["rows", "columns"]) {
      if (gt[i][key] !== wt[i][key]) {
        return `table ${i + 1} ${key}: expected ${wt[i][key]}, got ${gt[i][key]}`;
      }
    }
    if (gt[i].uniform !== wt[i].uniform) {
      return `table ${i + 1} uniform (no merged cells): expected ${wt[i].uniform}, got ${gt[i].uniform}`;
    }
    const at = firstTextDifference(`table ${i + 1} cell position`, gt[i].at ?? [], wt[i].at ?? []);
    if (at) return at;
    const cell = firstTextDifference(`table ${i + 1} cell`, gt[i].cells ?? [], wt[i].cells ?? []);
    if (cell) return cell;
  }
  if (!bookmarks) return null;
  const byName = (list) => new Map((list ?? []).map((b) => [b.name, b.text]));
  const [gb, wb] = [byName(got.bookmarks), byName(want.bookmarks)];
  for (const name of [...new Set([...wb.keys(), ...gb.keys()])].sort()) {
    if (!wb.has(name)) return `bookmark ${name}: not expected, got ${show(gb.get(name))}`;
    if (!gb.has(name)) return `bookmark ${name}: expected ${show(wb.get(name))}, missing`;
    if (gb.get(name) !== wb.get(name)) {
      return `bookmark ${name}: expected ${show(wb.get(name))}, got ${show(gb.get(name))}`;
    }
  }
  return null;
}

const counts = (shape) => `${shape.paragraphs} paragraphs, ${shape.tables} tables`;

// One pair. tool: exit codes of the tool { track, color }. accept, reject:
// the track-changes output after Revisions.AcceptAll / RejectAll. clean,
// base: the two inputs as Word reads them. color: the colour output, opened
// only. Each is what the ps1 reports for one job, or undefined when the job
// did not run. Returns { pass, failures: [line], first }.
export function judgePair({ tool, accept, reject, clean, base, color }) {
  const failures = [];
  const opened = (job, what) => {
    if (job?.ok) return true;
    failures.push(`Word does not open ${what}: ${job?.error ?? "the job did not run"}`);
    return false;
  };
  if (tool?.track !== 0)
    failures.push(`the tool did not write the track-changes file (exit ${tool?.track})`);
  if (tool?.color !== 0)
    failures.push(`the tool did not write the colour file (exit ${tool?.color})`);
  // A clean or baseline file Word cannot read says nothing about the tool.
  const inputs =
    opened(clean, "the clean input (a fault of the test pair)") &&
    opened(base, "the baseline input (a fault of the test pair)");
  if (tool?.track === 0) {
    const a = opened(accept, "the track-changes output");
    const r = a && opened(reject, "the track-changes output (second time)");
    if (a && accept.dirty_on_open && !clean?.dirty_on_open) {
      failures.push("Word changed the track-changes output while opening it (a repair?)");
    }
    if (a && accept.revisions_left !== 0) {
      failures.push(`${accept.revisions_left} revision(s) left after AcceptAll`);
    }
    if (r && reject.revisions_left !== 0) {
      failures.push(`${reject.revisions_left} revision(s) left after RejectAll`);
    }
    if (a && inputs) {
      const d = compareStructure(accept.shape, clean.shape);
      if (d) failures.push(`accept-all is not the clean document: ${d}`);
    }
    if (r && inputs) {
      const d = compareStructure(reject.shape, base.shape, { bookmarks: false, comments: false });
      if (d) failures.push(`reject-all is not the baseline: ${d}`);
      const [left, kept] = [reject.shape.comments ?? 0, clean.shape.comments ?? 0];
      if (left !== kept) {
        failures.push(
          `reject-all has ${left} comment(s), the clean document ${kept} (a comment is no tracked change and stays)`,
        );
      }
    }
  }
  if (tool?.color === 0 && opened(color, "the colour output")) {
    if (color.dirty_on_open && !clean?.dirty_on_open) {
      failures.push("Word changed the colour output while opening it (a repair?)");
    }
    // Both outputs lay out the same paragraphs and tables; only the marks differ.
    if (accept?.ok && counts(color.raw) !== counts(accept.raw)) {
      failures.push(
        `the colour output has ${counts(color.raw)}, the track-changes output ${counts(accept.raw)}`,
      );
    }
  }
  return { pass: failures.length === 0, failures, first: failures[0] ?? "" };
}

// The Word jobs of one pair. files: { track, color, clean, base } as Windows
// paths, or null for a file the tool did not write.
export function jobsFor(index, files) {
  const job = (role, path, action) => (path ? [{ id: `${index}:${role}`, path, action }] : []);
  return [
    ...job("accept", files.track, "accept"),
    ...job("reject", files.track, "reject"),
    ...job("clean", files.clean, "none"),
    ...job("base", files.base, "none"),
    ...job("color", files.color, "none"),
  ];
}

// Exit code of the run: 0 only when every pair passed and Word was closed
// the way it should be.
export function exitCode({ rows, wordError, cleanupFail }) {
  if (wordError || cleanupFail) return 1;
  return rows.length > 0 && rows.every((r) => r.result === "PASS") ? 0 : 1;
}

// The table of the run: one line per pair, the first failure in full (it
// holds up to 80 characters of the expected and of the actual text).
export function renderRows(rows) {
  const w = Math.max(4, ...rows.map((r) => r.step.length));
  const line = (a, b, c) => `${a.padEnd(w)}  ${b.padEnd(6)}  ${c}`.trimEnd();
  return [
    line("pair", "result", "first difference"),
    ...rows.map((r) => line(r.step, r.result, r.got ?? "")),
  ].join("\n");
}
