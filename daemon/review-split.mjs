// Split a pasted review into reviewer comments with refs R{reviewer}.{n}.
// Regex first (no model call); a single LLM one-shot only when the regex
// finds at most one comment. The user edits the result in the task pane.
// The one-shot recovers document structure: it labels each numbered line of
// the review and returns no text, and the code builds the comments from the
// source lines, so no line is lost or rewritten.
// L5: every comment carries requirements [{ id, text }], the separate things
// the reviewer asks for. Both splits leave them empty; extractRequirements
// fills them afterwards the same way: the code splits the comment sentences
// and marks where a sentence could be cut (clauseCuts), and one one-shot
// labels each sentence and each mark, so a requirement is always a sentence
// of the comment, or a part of one, as written.

import { numberRequirements } from "./revision-store.mjs";

// Head patterns, one regex source per kind; a profile may replace any of
// them. The first capture group of reviewerHead is the reviewer (a number, or
// one capital letter: A is 1), of commentHead and listItem the comment number,
// of sectionHead the section label; combinedHead captures the reviewer, then
// the comment number. Defaults are stricter than the plain
// design patterns so body lines do not start a new comment: "Reviewer 2
// raised the same point" (reviewer word at line start, followed by prose),
// "Q10 values" (temperature coefficient), "3.5 mM" (decimal at line start),
// "2019. The" (a year), "Line 5 of the abstract" (no mark after the number),
// "-20 °C" and "*p < 0.05" (no space after the sign).
// The layouts covered are listed with their expected split in
// tests/fixtures/review_formats.
export const DEFAULT_PATTERNS = {
  reviewerHead: String.raw`^\s*(?:(?:Comments?|Reports?|Remarks?)\s+(?:from|of|by)\s+(?:the\s+)?)?(?:Reviewer|Referee|심사\s*위원|심사자|리뷰어)\s*(?:#|No\.?|:)?\s*(\d+|[A-Z])(?:\s*\([^)]*\))?(?:\s*(?:Report|Comments?|Remarks?|의\s*의견|의견))?\s*(?:[:.)\-–].*)?$`,
  combinedHead: String.raw`^\s*(?:(?:Reviewer|Referee|심사\s*위원|심사자|리뷰어)\s*(?:#|No\.?|:)?\s*(\d+|[A-Z])\s*[,;:.\-\u2013]?\s*(?:(?:Comment|Point|Question|Remark|Issue)\s*(?:#|No\.?)?\s*|[QC]\s*#?\s*|의견\s*)(\d+)|R\s*(\d+)\s*C\s*(\d+))(?!\d)`,
  editorHead: String.raw`^\s*(?:(?:(?:Comments?|Remarks?)\s+(?:from|of|by)\s+(?:the\s+)?)?(?:(?:Associate|Academic|Handling|Guest|Section|Managing|Senior|Deputy|Chief)\s+)?Editors?(?:-in-Chief|\s+in\s+Chief)?(?:['’]s?)?(?:\s+and\s+(?:the\s+)?(?:Reviewers?|Referees?)['’]?)?(?:\s+(?:Comments?|Remarks?|Decision|Report|Letter))?(?:\s+(?:to|for)\s+(?:the\s+)?Authors?)?|편집\s*위원장?(?:\s*의?\s*의견)?|편집장(?:\s*의?\s*의견)?)\s*(?=:|$)`,
  sectionHead: String.raw`^\s*(?:(?:\d+|[IVX]+|[A-Z])[.)]\s+)?((?:Major|Minor|Specific|General|Other|Additional|Detailed|Technical|Main|Editorial|Further)\s+(?:comments?|points?|issues?|concerns?|remarks?|suggestions?|questions?|corrections?)|Major|Minor|주요\s*(?:의견|지적\s*사항|사항)|세부\s*(?:의견|사항)|일반\s*(?:의견|사항)|기타\s*(?:의견|사항)|총평|종합\s*의견|추가\s*의견)\s*(?:[:.]\s*$|:|$)`,
  commentHead: String.raw`^\s*(?:(?:Comment|Point)\s*#?\s*(\d+)\b|(?:의견|질의|지적\s*사항)\s*(\d+)\s*(?:[:.)]|$)|[QC]\s*#?\s*(\d+)\s*[:.)]|R\s*\d+\s*[.\-]\s*(\d+)\s*(?:[:.)]|$))`,
  listItem: String.raw`^\s*\(?(\d{1,3})[.)](?:\s|$)`,
  lineHead: String.raw`^\s*(?:(?:Page|pp?\.)\s*\d+\s*[,;]?\s*)?(?:Lines?|ll?)\.?\s*\d+(?:\s*[-–~]\s*\d+)?\s*(?:[:,.)\-–]|$)|^\s*(?:Page|pp?\.)\s*\d+\s*[:,)\-–]`,
  bullet: String.raw`^\s*[-*•–●◦·▪]\s+(?=\S)`,
};

function compile(patterns) {
  const p = { ...DEFAULT_PATTERNS, ...patterns };
  return {
    reviewer: new RegExp(p.reviewerHead, "id"),
    combined: new RegExp(p.combinedHead, "i"),
    editor: new RegExp(p.editorHead, "i"),
    section: new RegExp(p.sectionHead, "i"),
    comment: new RegExp(p.commentHead, "i"),
    list: new RegExp(p.listItem),
    line: new RegExp(p.lineHead, "i"),
    bullet: new RegExp(p.bullet),
  };
}

function firstNumber(m) {
  for (let i = 1; i < m.length; i++) if (m[i] !== undefined) return Number(m[i]);
  return null;
}

const HEAD_TRIM = /^[\s:.)\-]+/;

// A journal's own label on or under a head line, not the reviewer's words.
const STOCK_LABEL =
  /^(?:(?:review(?:er)?\s+)?(?:reports?|comments?|remarks?)(?:\s+and\s+suggestions)?(?:\s+(?:to|for)\s+(?:the\s+)?(?:authors?|editors?))?|의\s*의견|의견)[^\w가-힣]*$/i;

// "Reviewer(s)' Comments to Author:" (ScholarOne): reviewers start, no number.
const ANON_REVIEWER =
  /^\s*(?:Reviewer|Referee)s?(?:\(s\))?['’]?s?\s+(?:Comments?|Reports?)(?:\s+(?:to|for)\s+(?:the\s+)?Authors?)?\s*:?\s*$/i;

// A comment marker "R2.3": its reviewer, used when the review has no reviewer head.
const R_MARK = /^\s*R\s*(\d+)\s*[.\-]\s*\d+/i;

// A letter closing on its own line, and the first words of a portal footer.
const CLOSING =
  /^\s*(?:(?:yours\s+)?(?:sincerely|faithfully|truly)(?:\s+yours)?|(?:with\s+)?(?:best|kind|warm)(?:est)?\s+(?:regards|wishes)|regards)\s*[,.]?\s*$/i;
const FOOTER =
  /^\s*(?:In compliance with data protection regulations|This (?:e-?mail|message) (?:and any|is intended|was sent|has been sent)|To ensure this e-?mail|#AU_)/i;

// The line before ends inside a sentence that goes on with a number
// ("... in Table" / "2. The legend"), so the number starts no comment. A
// period after a whole word ends the sentence ("... in the table."); after
// an abbreviation it does not ("... of Fig.").
const WRAP_END =
  /(?:\b(?:(?:Figs?|Eqs?|Refs?|Secs?|approx)\.?|Figures?|Tables?|Equations?|Sections?|Supplementary|lines?|pages?|steps?|and|or|of|in|to|at|from|by|than|about|approximately|for|with|the)|[A-Za-z]-)\s*$/i;

// What may follow a combined head ("Reviewer 1, Comment 2") on its line: the
// end, a mark, or a new sentence. "Reviewer 1, Comment 2 was answered" is prose.
const COMBINED_REST = /^(?:\s*$|\s*[:.)\-\u2013]|\s+["'(\u201C]?[A-Z0-9])/;

// A rule line between the question blocks of a questionnaire letter.
const SEPARATOR = /^\s*[*_=\-]{5,}\s*$/;

// The label of a review round, alone on a line or as the text of a reviewer
// head line ("Reviewer 1: Second round").
const ROUND_LABEL =
  /^(?:(?:first|second|third|fourth|fifth|1st|2nd|3rd|\d+th)\s+(?:round|review|revision)|(?:review\s+)?round\s*#?\s*\d+|revision\s*#?\s*\d+|R\d+\s+comments?|\d+\s*\uCC28\s*(?:\uC2EC\uC0AC|\uC218\uC815|\uB9AC\uBDF0|\uAC80\uD1A0)(?:\s*\uC758\uACAC)?)$/i;

// A short label line of the reviewer's own ("Abstract:", "Methods:").
const SHORT_LABEL = /^\s*([^\s:.!?,;][^:.!?,;]{0,38}):\s*$/;
// The name of a manuscript part: a label also with no blank line before it.
const PART_NAME =
  /^(?:Title|Abstract|Summary|Keywords|Introduction|Background|(?:Materials\s+and\s+)?Methods|Experimental(?:\s+section)?|Results(?:\s+and\s+Discussion)?|Discussion|Conclusions?|References|Figures?|Tables?|Supplementary(?:\s+(?:materials?|information))?|\uC81C\uBAA9|\uCD08\uB85D|\uC11C\uB860|\uBC29\uBC95|\uACB0\uACFC|\uACE0\uCC30|\uACB0\uB860|\uCC38\uACE0\uBB38\uD5CC)$/i;

const QUOTE = /^\s*(?:>\s?)+/;

// Full-width forms and odd spaces as ASCII, one UTF-16 unit for one, so an
// index into the folded line is an index into the line as pasted. Only used
// to match; the comment text keeps the pasted characters.
const fold = (line) =>
  line
    .replace(/[！-～]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0))
    .replace(/[   　]/g, " ");

const trimHead = (raw) => raw.slice(HEAD_TRIM.exec(fold(raw))?.[0].length ?? 0);

// Lines of the review. An email quote mark is taken off when every line has one.
function sourceLines(text) {
  const lines = String(text ?? "")
    .replace(/\r\n?/g, "\n")
    .split("\n");
  const full = lines.filter((l) => l.trim());
  return full.length && full.every((l) => QUOTE.test(l))
    ? lines.map((l) => l.replace(QUOTE, ""))
    : lines;
}

// Text after a head label on its line; a bracket ("(Anonymous)") and a stock
// label give "". One mark is taken off, so "Reviewer 3: - Add a control."
// leaves the bullet.
function labelRest(raw) {
  const s = raw.replace(/^\s*\([^)]*\)/, "");
  const rest = s.slice(/^\s*(?:[:.)]+|[-\u2013])?\s*/.exec(fold(s))[0].length).trim();
  return STOCK_LABEL.test(fold(rest)) ? "" : rest;
}

// Text on a reviewer head line after the reviewer number ("Reviewer 1: Please
// report n." gives "Please report n."; a stock label gives ""). The regex
// needs the d flag.
function reviewerRest(m, line) {
  const end = m.indices?.[1]?.[1];
  return end == null ? "" : labelRest(line.slice(end));
}

// 3 for "3", 2 for "B"; null for anything else ("a" under the i flag).
const reviewerNumber = (s) =>
  /^\d+$/.test(s ?? "") ? Number(s) : /^[A-Z]$/.test(s ?? "") ? s.charCodeAt(0) - 64 : null;

const indent = (s) => /^\s*/.exec(s)[0].replace(/\t/g, "    ").length;

export function splitReviewRegex(text, patterns = DEFAULT_PATTERNS) {
  return splitBlocks(text, patterns).comments;
}

// The regex split. Returns the comments, the refs of reviewer blocks kept
// whole (unsplit), the kept preambles (preserved), the refs of editor text
// after a reviewer (late), the refs split only at blank lines (weak), the
// refs of form text (form), the lines dropped as boilerplate (dropped: a
// journal label line under a reviewer head, an unnumbered reviewer label
// before a numbered head, a rule line of a questionnaire letter), the
// section labels taken off their lines (labels) and
// `over`: every line of the review without the head labels the split took
// off, for the check of text in no comment.
//
// Blocks. A reviewer head opens a reviewer block, an editor head an editor
// block. Text after a head on its line is the block's first body line. When
// the review has a head, the text before the first one (the editor letter)
// is one editor comment. A letter closing ("Sincerely,") or a portal footer
// inside a reviewer block opens an editor block too, so a signature is not
// part of the last comment. Until the next blank line a closing also keeps
// the lines under it from being read as heads ("Editor-in-Chief").
// A combined head ("Reviewer 1, Comment 2:", "R2C3:") opens or goes on
// with the block of its reviewer and starts comment R1.2.
//
// Questionnaire letters. When a reviewer has more than one block and one of
// them is only a short answer on its head line ("Reviewer #1: Partly", at
// most 40 characters, no text under it, not starting with a comment head
// that goes on with text: "1. Add units." is a comment, "3 - Partly" and "3."
// are answers), the longest other block of that
// reviewer is the review and the answers are form text. A block with text
// under its head line is a review, whatever its head-line text ("Reviewer 1:
// First round"), so two rounds of a reviewer stay reviewer comments. Form
// text: each run of form text (question, answers, the label between them)
// is one editor comment with the answer lines as pasted. In such a letter a
// rule line ("**********") ends a reviewer block and is dropped. Text under
// an unnumbered reviewer label before the first numbered head is form text
// too.
//
// Comments of a block. Section heads ("Major comments", "Minor comments", a
// short label line such as "Abstract:" before a list) cut the block into
// parts, and each comment keeps the label of its part as `section`. In each
// part the first kind of head that occurs splits it: "Comment n" heads, then
// numbered items, then line or page heads, then bullets. Heads of a later
// kind are body text, and so are lettered items ("a)", "i."), "1a.", an
// item indented deeper than the first one, a numbered list after a line
// that ends with a colon, a number lower than the comment before it, and a
// number after a line that ends inside a sentence ("in Table"). A list after
// a colon ends where its numbers stop counting up. When its next number is
// also the next comment number, the item is the next comment only if that
// number comes again below or the item is written unlike the first
// sub-item (capital first letter, final period).
//
// Refs. E.{k} for editor comments, k counting up in reading order.
// R{reviewer}.{n} for reviewer comments: n is the number of the head.
// When the numbers start again (a new section, or "1." after "3." with no
// section head), the part goes on from the highest number so far, so
// Major 1-3 and Minor 1-2 are R1.1 to R1.5. A head with no number (a line
// head, a bullet, a paragraph) takes the next number. Numbers inside one
// run are kept as written, so a doubled "Comment 1" gives a doubled ref and
// a warning. A later block of the same reviewer (a second round, or the
// reviewer's head a second time) goes on from the reviewer's highest number:
// round 2 after R1.1 and R1.2 starts at R1.3. Its comments carry the round
// label as `section` ("Round 2" on a line of its own after the first
// reviewer head, or as the text of the head line: "Reviewer 1: Second
// round"), or "Block 2" where the letter gives none; a section inside the
// block is joined to it ("Round 2 / Minor comments"). Where the ref differs
// from the number as written, the written number stays in the comment text.
// Text before the first comment of a later block is a comment of its own
// with the next number. One warning per reviewer names the blocks. A later
// block whose comments all keep their written numbers gets no "Block n"
// label and no warning (one comment per head line: "Reviewer 1: 2. ...").
// Blocks opened by combined heads carry their own numbers and are left as
// written.
// R{reviewer}.0 is the text before the first comment of the first block of a
// reviewer with a head; without any reviewer head that text is dropped (and
// named by the check). A reviewer block with no head of any kind is split
// at blank lines when it has more than one paragraph (a weak split: it does
// not stop the model split), else kept whole as R{reviewer}.1.
//
// Limits, left to the model split or the user:
// - A sub-list after a colon whose last item has the next comment number and
//   is written like the sub-items stays in its comment, also when it was
//   meant as the next comment.
// - A "Round 1" line before the first reviewer head is part of the letter,
//   so the first round carries no round label unless the head line has one.
// - The same combined head twice ("Reviewer 1, Comment 1:" in two rounds)
//   gives a doubled ref and a warning.
// - In a questionnaire letter pasted without its rule lines, a question
//   after the reviews ends up in the last comment of the last reviewer.
// - A closing or footer after the last comment of a review without any
//   reviewer head stays in that comment.
// - Lettered items ("a)", "b)") are never comment heads; a reviewer who
//   only letters the comments is kept whole or split at blank lines.
function splitBlocks(text, patterns = DEFAULT_PATTERNS) {
  const re = compile(patterns);
  const lines = sourceLines(text);
  let over;
  let blocks;
  let forced; // line index -> comment number of a combined head
  let gone; // line indexes dropped as boilerplate
  let rounds; // round labels taken off their lines
  const txt = (i) => fold(over[i]);
  const hasText = (idx) => idx.some((i) => over[i].trim());

  // cut: rule lines end a reviewer block (questionnaire letter).
  const build = (cut) => {
    over = lines.slice();
    blocks = [];
    forced = new Map();
    gone = [];
    rounds = [];
    let round = null;
    let cur = { reviewer: 1, idx: [], headed: false };
    let top = 0;
    let reviewers = false;
    let closed = false;
    // An unnumbered reviewer label right before a numbered head is no block;
    // text under it is form text, not the words of a reviewer.
    const open = (b) => {
      if (cur.anon && b.numbered) {
        Object.assign(cur, { editor: true, form: true, reviewer: 0 });
        gone.push(cur.head);
      }
      if (!(cur.anon && !hasText(cur.idx) && b.numbered)) blocks.push(cur);
      cur = b;
      if (!b.editor) {
        if (round) b.round = round;
        reviewers = true;
        top = Math.max(top, b.reviewer);
      }
    };
    lines.forEach((line, i) => {
      const s = fold(line);
      if (!s.trim()) {
        closed = false;
        cur.idx.push(i);
        return;
      }
      if (cut && SEPARATOR.test(s)) {
        over[i] = "";
        gone.push(i);
        if (cur.headed && !cur.editor)
          open({ reviewer: 0, editor: true, form: true, first: true, idx: [i] });
        else cur.idx.push(i);
        return;
      }
      // A round line after the first reviewer: the label of the blocks below.
      const mark = line.trim().replace(/\s*[:\uFF1A]$/, "");
      if (reviewers && !closed && ROUND_LABEL.test(fold(mark))) {
        over[i] = "";
        rounds.push(mark);
        round = mark;
        if (cur.headed && !cur.editor && !hasText(cur.idx)) cur.round = mark;
        cur.idx.push(i);
        return;
      }
      const c = closed ? null : re.combined.exec(s);
      const both = c && [...c].slice(1).filter((x) => x !== undefined);
      const who = c && reviewerNumber(both[0]);
      if (who && /^\d+$/.test(both[1] ?? "") && COMBINED_REST.test(s.slice(c[0].length))) {
        over[i] = trimHead(line.slice(c[0].length)).trim();
        forced.set(i, Number(both[1]));
        if (cur.headed && !cur.editor && cur.reviewer === who) cur.idx.push(i);
        else
          open({ reviewer: who, headed: true, numbered: true, combined: true, head: i, idx: [i] });
        return;
      }
      const m = closed ? null : re.reviewer.exec(s);
      const number = m && reviewerNumber(m[1]);
      const e = closed || number ? null : re.editor.exec(s);
      const closing = CLOSING.test(s);
      if (number) {
        over[i] = reviewerRest(m, line);
        open({ reviewer: number, headed: true, numbered: true, head: i, idx: [i] });
      } else if (!closed && ANON_REVIEWER.test(s)) {
        over[i] = "";
        open({ reviewer: top + 1, headed: true, anon: true, head: i, idx: [i] });
      } else if (e) {
        over[i] = labelRest(line.slice(e[0].length));
        open({ reviewer: 0, editor: true, headed: true, late: reviewers, head: i, idx: [i] });
      } else if (cur.headed && !cur.editor && (closing || FOOTER.test(s))) {
        open({ reviewer: 0, editor: true, letter: true, late: true, idx: [i] });
        closed = closing;
      } else {
        if (closing) closed = true;
        else if (cur.headed && !cur.editor && !hasText(cur.idx) && STOCK_LABEL.test(s.trim())) {
          over[i] = "";
          gone.push(i);
        }
        cur.idx.push(i);
      }
    });
    blocks.push(cur);
  };

  // Blocks that are only a short answer on the head line, of a reviewer
  // whose review is another block. `pure`: nothing but the head-line text.
  // `tail`: text under the head line after a blank line. Such a block is a
  // review ("Reviewer 1: First round" / "1. Add units.") unless the block
  // right before it is the pure answer of another reviewer: then it is the
  // last answer to a question, and the text under it is the next question
  // of a letter pasted without rule lines.
  // A head-line text that starts with a comment head and goes on with text
  // ("1. Add units.", "(2) State n.", "- Add a control.", "Line 12: ...") is a
  // comment of the reviewer, never an answer. A digit alone, or with anything
  // but "." or ")" after it, is no comment head, so "3 - Partly", "1 (poor)",
  // "4" and "3." are answers.
  const comment = (b) =>
    ["comment", "list", "line", "bullet"].some((k) => {
      const m = re[k].exec(txt(b.head));
      return m && txt(b.head).slice(m[0].length).trim() !== "";
    });
  const short = (b) =>
    over[b.head].trim() !== "" && over[b.head].trim().length <= 40 && !comment(b);
  const pure = (b) => short(b) && !b.idx.some((i) => i !== b.head && over[i].trim());
  const tail = (b) => short(b) && !pure(b) && !over[b.idx[1]].trim();
  // The review among the blocks of one reviewer: the longest block that is
  // not an answer, else the answer with the longest head-line text.
  const size = (b) => b.idx.reduce((n, i) => n + over[i].trim().length, 0);
  const longest = (list, len) => list.reduce((a, b) => (len(b) >= len(a) ? b : a));
  const answersBy = (answer) => {
    const by = new Map();
    for (const b of blocks) {
      if (b.numbered && !b.combined) by.set(b.reviewer, [...(by.get(b.reviewer) ?? []), b]);
    }
    return [...by.values()].flatMap((list) => {
      if (list.length < 2) return [];
      const real = list.filter((b) => !answer(b));
      const main = real.length
        ? longest(real, size)
        : longest(list, (b) => over[b.head].trim().length);
      return list.filter((b) => b !== main && answer(b));
    });
  };
  const answers = () => {
    const first = new Set(answersBy(pure));
    const last = new Set(
      blocks.filter((b, k) => b.numbered && !b.combined && tail(b) && first.has(blocks[k - 1])),
    );
    return last.size ? answersBy((b) => pure(b) || last.has(b)) : [...first];
  };
  build(false);
  if (answers().length) build(true);
  for (const b of answers()) {
    over[b.head] = lines[b.head];
    Object.assign(b, { editor: true, form: true, reviewer: 0 });
  }
  // A run of form text is one editor comment; a rule line starts a new run.
  const count = blocks.length;
  blocks = blocks.reduce((list, b) => {
    if (b.form && !b.first && list.at(-1)?.form) list.at(-1).idx.push(...b.idx);
    else list.push(b);
    return list;
  }, []);

  // A short label line: after a blank line or the block head, before a head.
  // The name of a manuscript part needs no blank line, unless the line
  // before it ends inside a sentence ("... are missing in the").
  const isHead = (i) => ["comment", "list", "line", "bullet"].some((k) => re[k].test(txt(i)));
  const shortLabel = (b, k) => {
    const i = b.idx[k];
    const m = i === b.head ? null : SHORT_LABEL.exec(txt(i));
    if (!m || m[1].trim().split(/\s+/).length > 3 || isHead(i)) return null;
    const before = k > 0 ? txt(b.idx[k - 1]) : "";
    if (before.trim() && !(PART_NAME.test(m[1].trim()) && !WRAP_END.test(before))) return null;
    const next = b.idx.slice(k + 1).find((j) => over[j].trim());
    return next !== undefined && isHead(next) ? m[1].trim() : null;
  };

  // How a list item is written: capital first letter, final period.
  const shape = (s, m) => {
    const t = s.slice(m[0].length).trim();
    return `${/^[A-Z]/.test(t)},${/[.!?]$/.test(t)}`;
  };

  // Heads of one part. Returns the lines before the first head and the items.
  const partItems = (idx, mode) => {
    const pre = [];
    const items = [];
    let item = null;
    let prev = 0;
    let sub = null;
    let subShape = null;
    let topIndent = null;
    let last = null;
    let gap = true;
    idx.forEach((i, k) => {
      const s = txt(i);
      if (forced.has(i)) {
        item = { n: forced.get(i), restart: false, idx: [i] };
        items.push(item);
        last = i;
        gap = false;
        return;
      }
      if (!s.trim()) {
        (item ? item.idx : pre).push(i);
        gap = true;
        return;
      }
      const m = re[mode].exec(s);
      const n = m && (mode === "comment" || mode === "list") ? firstNumber(m) : null;
      let body = !m;
      let restart = false;
      if (!body && (mode === "list" || mode === "bullet")) {
        if (topIndent === null) topIndent = indent(s);
        else if (item && indent(s) >= topIndent + 2) body = true;
      }
      if (!body && mode === "list" && n !== null) {
        const before = last === null ? "" : txt(last);
        // The same number comes again below: this one is the sub-item.
        const later = () => idx.slice(k + 1).some((j) => Number(re.list.exec(txt(j))?.[1]) === n);
        if (!gap && WRAP_END.test(before)) body = true;
        else if (
          item &&
          sub !== null &&
          n === sub + 1 &&
          (n !== prev + 1 || later() || shape(s, m) === subShape)
        ) {
          sub = n;
          body = true;
        } else if (item && n === 1 && /:\s*$/.test(before)) {
          sub = 1;
          subShape = shape(s, m);
          body = true;
        } else if (item && n !== 1 && n <= prev) body = true;
        else {
          restart = !!item && n === 1;
          sub = null;
          prev = n;
        }
      }
      if (!body) {
        item = { n, restart, idx: [i], raw: over[i] };
        items.push(item);
        if (mode !== "line") over[i] = trimHead(over[i].slice(m[0].length));
      } else {
        (item ? item.idx : pre).push(i);
      }
      last = i;
      gap = false;
    });
    return { pre, items };
  };

  const paragraphs = (idx) => {
    const out = [];
    let p = null;
    for (const i of idx) {
      if (!over[i].trim()) p = null;
      else if (p) p.push(i);
      else out.push((p = [i]));
    }
    return out;
  };

  const labels = [...rounds]; // section and round labels taken off their lines
  const seen = new Map(); // reviewer -> blocks so far
  const tops = new Map(); // reviewer -> highest comment number so far
  const blockRows = (b) => {
    if (!b.editor && !hasText(b.idx)) return [];
    const base = b.editor ? { editor: true, late: !!b.late } : { reviewer: b.reviewer };
    // A later block of a reviewer goes on from the reviewer's highest number.
    // Its comments carry the round label, or "Block n" where the letter gives
    // none. Combined heads carry their own numbers and make no later block.
    const nth = b.editor ? 1 : (seen.get(b.reviewer) ?? 0) + 1;
    if (!b.editor) seen.set(b.reviewer, nth);
    const later = nth > 1 && !b.combined;
    let round = b.round ?? null;
    if (!b.editor && !b.combined && b.head !== undefined) {
      const rest = over[b.head].trim();
      if (ROUND_LABEL.test(fold(rest))) {
        round = rest;
        labels.push(rest);
        over[b.head] = "";
      }
    }
    // "Block n" is only a stand-in label: it is taken back (in `done`) when
    // every comment of the block has its number as written, so one comment
    // per head line ("Reviewer 1: 2. State n.") is not a block of its own.
    const standIn = later && round === null;
    if (later) {
      round ??= `Block ${nth}`;
      base.block = nth;
    }
    const sec = (label) => [round, label].filter(Boolean).join(" / ") || undefined;
    const done = (rows) => {
      if (standIn && rows.filter((r) => hasText(r.idx)).every((r) => r.written)) {
        for (const r of rows) {
          r.block = undefined;
          r.section = r.section === round ? undefined : r.section?.replace(`${round} / `, "");
        }
      }
      if (!b.editor) {
        tops.set(b.reviewer, Math.max(tops.get(b.reviewer) ?? 0, ...rows.map((r) => r.number)));
      }
      return rows;
    };
    let max = later ? (tops.get(b.reviewer) ?? 0) : 0;
    const parts = [{ idx: [] }];
    b.idx.forEach((i, k) => {
      const m = i === b.head ? null : re.section.exec(txt(i));
      const label = m ? (m[1] ?? m[0]).replace(/\s+/g, " ").trim() : shortLabel(b, k);
      if (label === null) return parts.at(-1).idx.push(i);
      over[i] = m ? trimHead(over[i].slice(m[0].length)).trim() : "";
      labels.push(label);
      parts.push({ section: label, idx: [i] });
    });
    const modeOf = (p) =>
      p.idx.some((i) => forced.has(i))
        ? "comment"
        : ["comment", "list", "line", "bullet"].find((k) => p.idx.some((i) => re[k].test(txt(i))));
    if (!parts.some(modeOf)) {
      // An editor block without heads: one comment per section.
      if (b.editor) return parts.map((p) => ({ ...base, kept: true, ...p }));
      const paras = parts.flatMap((p) =>
        paragraphs(p.idx).map((idx) => ({ idx, section: sec(p.section) })),
      );
      if (paras.length > 1)
        return done(paras.map((p) => ({ ...base, number: ++max, weak: true, ...p })));
      if (!b.headed) return [];
      return done([
        { ...base, number: ++max, whole: true, idx: b.idx, section: paras[0]?.section ?? sec() },
      ]);
    }
    const out = [];
    const preamble = { ...base, number: 0, kept: true, idx: [] };
    for (const p of parts) {
      const mode = modeOf(p);
      const { pre, items } = mode ? partItems(p.idx, mode) : { pre: p.idx, items: [] };
      if (hasText(pre)) {
        if (out.length) out.push({ ...base, number: ++max, idx: pre, section: sec(p.section) });
        else {
          preamble.idx.push(...pre);
          preamble.section ??= p.section;
        }
      }
      let off = null;
      for (const it of items) {
        // In a later block the text before the first comment is a comment of
        // its own, with the next number.
        if (later && !out.length && !preamble.number && hasText(preamble.idx)) {
          preamble.number = ++max;
        }
        let number;
        if (it.n === null) number = ++max;
        else {
          if (off === null || it.restart) off = it.n <= max ? max - it.n + 1 : 0;
          number = it.n + off;
          max = Math.max(max, number);
          // The number as written stays in the text where the ref differs.
          if (later && number !== it.n && it.raw !== undefined) over[it.idx[0]] = it.raw;
        }
        const written = it.n !== null && number === it.n;
        const r = !b.headed && mode === "comment" ? R_MARK.exec(fold(lines[it.idx[0]])) : null;
        out.push({
          ...base,
          number,
          idx: it.idx,
          written,
          section: sec(p.section),
          ...(r && { reviewer: Number(r[1]) }),
        });
      }
    }
    preamble.section = sec(preamble.section);
    return done(b.headed ? [preamble, ...out] : out);
  };

  let editors = 0;
  const rows = blocks
    .flatMap((b) =>
      (!b.headed && count > 1) || b.letter || b.form
        ? [{ editor: true, kept: !b.form, form: !!b.form, late: !!b.late, idx: b.idx }]
        : blockRows(b),
    )
    .map((r) => ({ ...r, text: trimHead(r.idx.map((i) => over[i]).join("\n")).trim() }))
    .filter((r) => r.text)
    .map((r) => ({
      ...r,
      comment: {
        ref: r.editor ? `E.${++editors}` : `R${r.reviewer}.${r.number}`,
        reviewer: r.editor ? 0 : r.reviewer,
        text: r.text,
        requirements: [],
        ...(r.section && { section: r.section }),
      },
    }));
  const refs = (key) => rows.filter((x) => x[key]).map((x) => x.comment.ref);
  return {
    comments: rows.map((x) => x.comment),
    unsplit: refs("whole"),
    preserved: rows.filter((x) => x.kept).map((x) => x.comment),
    late: refs("late"),
    weak: refs("weak"),
    form: refs("form"),
    over,
    labels,
    dropped: gone.sort((a, b) => a - b).map((i) => lines[i].trim()),
    // Reviewers with more than one block: [{ reviewer, blocks, later: [{ block, refs }] }].
    multi: [...seen]
      .filter(([reviewer]) => rows.some((x) => x.block && x.reviewer === reviewer))
      .map(([reviewer, blocks]) => ({
        reviewer,
        blocks,
        later: [
          ...new Set(rows.filter((x) => x.block && x.reviewer === reviewer).map((x) => x.block)),
        ].map((block) => ({
          block,
          refs: rows
            .filter((x) => x.block === block && x.reviewer === reviewer)
            .map((x) => x.comment.ref),
        })),
      })),
  };
}

// late: refs of editor text that comes after a reviewer (label split only).
const preservedWarnings = (comments, late = []) =>
  comments.map((c) =>
    late.includes(c.ref)
      ? `Editor text after a reviewer is kept as ${c.ref}. Delete it if it asks for nothing.`
      : c.reviewer === 0
        ? `Text before the first reviewer (usually the editor letter) is kept as ${c.ref}. Delete it if it asks for nothing.`
        : `Reviewer ${c.reviewer} text before the first numbered comment is kept as ${c.ref}. Delete it if it asks for nothing.`,
  );

// refs: comments of a reviewer block with no head, split at blank lines.
function weakWarnings(refs) {
  const by = new Map();
  for (const ref of refs) {
    const n = ref.slice(1, ref.indexOf("."));
    by.set(n, [...(by.get(n) ?? []), ref]);
  }
  return [...by].map(
    ([n, list]) =>
      `Reviewer ${n} text has no numbered comments; it was split at blank lines into ${list[0]} to ${list.at(-1)}. Join or split them by hand where a comment spans paragraphs.`,
  );
}

const formWarnings = (refs) =>
  refs.length
    ? [
        `Form text and short answers between the reviews are kept as ${refs.join(", ")}. Delete them if they ask for nothing.`,
      ]
    : [];

// multi: reviewers with more than one block (splitBlocks).
const blockWarnings = (multi) =>
  multi.map(
    ({ reviewer, blocks, later }) =>
      `Reviewer ${reviewer} has ${blocks} blocks; numbering was continued: ${later
        .map(
          ({ block, refs }) =>
            `block ${block} is ${refs.length > 1 ? `${refs[0]} to ${refs.at(-1)}` : refs[0]}`,
        )
        .join(", ")}.`,
  );

const droppedWarnings = (lines) =>
  lines.length
    ? [
        `${lines.length} boilerplate line(s) were dropped: ${quoteUnits(lines, 40)}. Add them back if they ask for something.`,
      ]
    : [];

const unsplitWarnings = (refs) =>
  refs.map(
    (ref) =>
      `Reviewer ${ref.slice(1, ref.indexOf("."))} text was not split into numbered comments; it is kept whole as ${ref}. Split it by hand if it holds several requests.`,
  );

function duplicateWarnings(comments) {
  const seen = new Set();
  const w = [];
  for (const c of comments) {
    if (seen.has(c.ref)) w.push(`duplicate ref ${c.ref}: renumber it before saving`);
    seen.add(c.ref);
  }
  return w;
}

const squash = (s) =>
  String(s ?? "")
    .replace(/\s+/g, " ")
    .trim();

// Sentence units for the coverage check: split at line breaks and sentence
// ends, drop a leading list marker (a list head loses it in its comment) and keep
// units with a letter. An abbreviation ("e.g.", "Fig. 2") only splits a
// sentence into shorter pieces, each still inside a comment that holds it.
const LIST_MARK = /^(?:\(?\d+[.)]|\(?[a-z][.)]|[-*\u2022])\s+/i;
function sentenceUnits(text) {
  return String(text ?? "")
    .split(/\n|(?<=[.!?])\s+/)
    .map((u) => squash(u).replace(LIST_MARK, ""))
    .filter((u) => /\p{L}/u.test(u));
}

const quoteUnits = (units, max = 80) =>
  units
    .slice(0, 3)
    .map((u) => `"${u.length > max ? `${u.slice(0, max - 3)}...` : u}"`)
    .join(", ") + (units.length > 3 ? ` and ${units.length - 3} more` : "");

// Sentences of the review, without the head labels the regex split takes
// off (the text after a label on its line stays), that are inside none of
// the comments, and the section labels that are the `section` of none.
// Lines dropped as boilerplate are not named here: the split result lists
// them as `dropped`.
export function uncoveredText(text, comments, patterns) {
  const { over, labels } = splitBlocks(text, patterns);
  const texts = comments.map((c) => squash(c.text));
  const sections = comments.map((c) => squash(c.section));
  return [
    ...sentenceUnits(over.join("\n")).filter((u) => !texts.some((t) => t.includes(u))),
    ...labels.filter((l) => !sections.some((x) => x.split(" / ").includes(squash(l)))),
  ];
}

export const SPLIT_SYSTEM =
  "You recover the structure of a peer review. Give every numbered line exactly one type. Do not return any text of the review. Reply with JSON only.";

export const LINE_TYPES = ["editor", "reviewer_head", "comment_start", "continue", "drop"];

// lines: the non-empty lines of the review, numbered from 1 in the prompt.
export function splitPrompt(lines) {
  return [
    'Return {"lines":[{"n":<line number>,"type":"<type>"}]} with every line number exactly once.',
    'Types: "editor" (the editor letter), "reviewer_head" (a new reviewer starts; add "reviewer":<int>), "comment_start" (a new comment starts), "continue" (same comment as the line before), "drop" (only a page header, page footer or rule line).',
    "Lines:",
    ...lines.map((l, i) => `${i + 1}: ${l}`),
  ].join("\n");
}

// The model reply as one label per line, or null unless every line number
// 1..count is there exactly once with a known type, and every reviewer_head
// has a positive integer reviewer.
function validLabels(json, count) {
  const items = json?.lines;
  if (!Array.isArray(items) || items.length !== count) return null;
  const out = new Array(count);
  for (const it of items) {
    const n = it?.n;
    if (!Number.isInteger(n) || n < 1 || n > count || out[n - 1]) return null;
    if (!LINE_TYPES.includes(it.type)) return null;
    if (it.type === "reviewer_head" && !(Number.isInteger(it.reviewer) && it.reviewer > 0)) {
      return null;
    }
    out[n - 1] = it;
  }
  return out;
}

// Comments from the source lines and one label per non-empty line. A blank
// line joins the block before it. Only recognized heads lose text: the
// reviewer label of a reviewer_head line (regex reviewerHead, as in the regex
// split) and the comment marker of a comment_start line (commentHead or
// listItem, whose number becomes the ref). Every other line not dropped is in
// exactly one comment. Refs follow the regex split: E.{k} for each editor run,
// R{n}.0 for reviewer text before its first comment, R{n}.1 for a reviewer
// with no comment start, otherwise the marker number or the next free one.
function commentsFromLabels(src, labels, patterns) {
  const { reviewer: reReviewer, comment: reComment, list: reList } = compile(patterns);
  const counts = Object.fromEntries(LINE_TYPES.map((t) => [t, 0]));
  const groups = [];
  const dropped = [];
  let cur = null;
  let reviewer = 1;
  let editors = 0;
  let n = 0;
  const open = (g) => groups.push((cur = g));
  for (const line of src) {
    if (!line.trim()) {
      cur?.lines.push(line);
      continue;
    }
    const label = labels[n++];
    counts[label.type] += 1;
    const type = label.type === "continue" && !cur ? "comment_start" : label.type;
    if (type === "drop") {
      dropped.push(line.trim());
    } else if (type === "continue" || (type === "editor" && cur?.editor)) {
      cur.lines.push(line);
    } else if (type === "editor") {
      const late = groups.some((g) => !g.editor);
      open({ reviewer: 0, editor: true, late, number: ++editors, lines: [line] });
    } else if (type === "reviewer_head") {
      reviewer = label.reviewer;
      const m = reReviewer.exec(fold(line));
      open({ reviewer, number: 0, pre: true, lines: [m ? reviewerRest(m, line) : line] });
    } else {
      const m = reComment.exec(fold(line)) ?? reList.exec(fold(line));
      const used = groups.filter((g) => !g.editor && g.reviewer === reviewer).map((g) => g.number);
      const number = (m && firstNumber(m)) ?? Math.max(0, ...used) + 1;
      open({
        reviewer,
        number,
        lines: [m ? trimHead(line.slice(m[0].length)) : line],
      });
    }
  }
  const rows = groups
    .map((g) => {
      const whole =
        !!g.pre && !groups.some((c) => !c.pre && !c.editor && c.reviewer === g.reviewer);
      const number = whole ? 1 : g.number;
      return {
        whole,
        kept: !!g.editor || (!!g.pre && !whole),
        late: !!g.late,
        comment: {
          ref: g.editor ? `E.${number}` : `R${g.reviewer}.${number}`,
          reviewer: g.reviewer,
          text: g.lines.join("\n").trim(),
          requirements: [],
        },
      };
    })
    .filter((x) => x.comment.text)
    .sort((a, b) => a.comment.reviewer - b.comment.reviewer);
  const comments = rows.map((x) => x.comment);
  const warnings = [
    ...preservedWarnings(
      rows.filter((x) => x.kept).map((x) => x.comment),
      rows.filter((x) => x.late).map((x) => x.comment.ref),
    ),
    ...unsplitWarnings(rows.filter((x) => x.whole).map((x) => x.comment.ref)),
  ];
  if (dropped.length) {
    warnings.push(
      `${dropped.length} line(s) were dropped as page headers, footers or rules: ${quoteUnits(dropped, 40)}. Add them back if they ask for something.`,
    );
  }
  if (dropped.length * 2 > labels.length) {
    warnings.push(
      `The model dropped more than half of the lines (${dropped.length} of ${labels.length}); check the split against the review.`,
    );
  }
  return {
    comments,
    warnings: [...warnings, ...duplicateWarnings(comments)],
    labels: counts,
    dropped,
  };
}

// oneshot: (req) => Promise<{ ok, json, error }> (daemon/oneshot.mjs).
// The regex result stands when its comment heads found more than one
// comment; blocks kept whole, kept preambles, form text and comments split
// only at blank lines do not count, so they cannot hide a review the regex barely
// split. Kept-whole blocks and kept preambles
// always come back as warnings.
// A regex result ends with a check of the review text that is in no comment
// (without a reviewer head the regex drops the text before the first comment
// head); it is named, not added back. A result built from line labels needs
// no check: only dropped lines are left out, and a warning names them.
export async function splitReview(text, { patterns, oneshot } = {}) {
  const r = await splitReviewInner(text, { patterns, oneshot });
  if (r.labels) return r;
  const lost = uncoveredText(text, r.comments, patterns);
  if (!lost.length) return r;
  return {
    ...r,
    warnings: [
      ...r.warnings,
      `${lost.length} sentence(s) of the review are in no comment: ${quoteUnits(lost)}. Add them if they ask for something.`,
    ],
  };
}

async function splitReviewInner(text, { patterns, oneshot } = {}) {
  const {
    comments: regexOut,
    unsplit,
    preserved,
    late,
    weak,
    form,
    dropped,
    multi,
  } = splitBlocks(text, patterns);
  const whole = [
    ...preservedWarnings(preserved, late),
    ...blockWarnings(multi),
    ...formWarnings(form),
    ...unsplitWarnings(unsplit),
    ...weakWarnings(weak),
    ...droppedWarnings(dropped),
  ];
  const regex = (method, warnings) => ({ method, comments: regexOut, warnings, dropped });
  const soft = unsplit.length + preserved.length + weak.length + form.length;
  if (regexOut.length - soft > 1) {
    return regex("regex", [...whole, ...duplicateWarnings(regexOut)]);
  }
  if (!oneshot) {
    return regex("regex", [
      "regex found at most one comment; no LLM fallback configured",
      ...whole,
    ]);
  }
  // The same lines as the regex split: an email quote mark on every line is
  // taken off before the model sees them.
  const src = sourceLines(text);
  const numbered = src.filter((l) => l.trim());
  if (!numbered.length) return regex("regex", whole);
  const r = await oneshot({ system: SPLIT_SYSTEM, prompt: splitPrompt(numbered) });
  if (!r.ok) {
    return regex("llm", [`LLM split failed: ${r.error}`, ...whole]);
  }
  // A reply that does not label every line once (empty, another JSON shape,
  // a line missing, doubled or unknown): the regex split stands, with a warning.
  const labels = validLabels(r.json, numbered.length);
  if (!labels) {
    return regex("regex", [
      "model split returned no usable line labels; kept the regex split",
      ...whole,
      ...duplicateWarnings(regexOut),
    ]);
  }
  return { method: "llm", ...commentsFromLabels(src, labels, patterns) };
}

// ---- L5 requirement extraction ----------------------------------------------

// A period after one of these does not end a sentence ("e.g.", "Fig. 2",
// "et al. Nature"). A lowercase letter after the space does not either.
const ABBREV =
  /(?:^|[\s(])(?:e\.g|i\.e|et al|vs|cf|Figs?|Eqs?|Refs?|Tabs?|Nos?|Suppl|approx|resp|ca|Dr|Prof|Sec)\.$/i;
// A bare list number at the start of a unit or after a colon ("points: 1.").
const LIST_NUMBER = /(?:^|[:;]\s*)\(?\d{1,2}[.)]$/;
const LIST_LINE = /^\s*(?:\(?\d+[.)]|\(?[a-z][.)]|[-*\u2022])\s+/i;
const BOUNDARY = /[.!?]+["')\]]*\s+/g;

// Sentences of a comment, whitespace squashed, a leading list marker left
// off, only units with a letter. A line break is a space (pasted text wraps
// mid-sentence) except before a list line and at a blank line. Every
// sentence is a substring of the comment text with whitespace squashed.
export function splitSentences(text) {
  const chunks = [];
  let cur = [];
  for (const line of String(text ?? "")
    .replace(/\r\n?/g, "\n")
    .split("\n")) {
    if (!line.trim() || LIST_LINE.test(line)) {
      if (cur.length) chunks.push(cur.join(" "));
      cur = line.trim() ? [line] : [];
    } else {
      cur.push(line);
    }
  }
  if (cur.length) chunks.push(cur.join(" "));
  const out = [];
  for (const chunk of chunks) {
    const t = squash(chunk);
    let start = 0;
    for (const m of t.matchAll(BOUNDARY)) {
      const end = m.index + m[0].length;
      const head = t.slice(start, end).trimEnd();
      const next = t[end] ?? "";
      if (/\p{Ll}/u.test(next) || ABBREV.test(head) || LIST_NUMBER.test(head)) continue;
      out.push(head);
      start = end;
    }
    if (start < t.length) out.push(t.slice(start));
  }
  return out.map((u) => u.replace(LIST_MARK, "").trim()).filter((u) => /\p{L}/u.test(u));
}

// S2-30: one sentence may ask for two things ("please state how growth was
// followed, and describe the carbon source as novel"). The code marks the
// places where it could be cut; the model only says split or keep at each
// mark and gives each part a type, so a requirement is still text of the
// comment as written. A mark is a coordinator outside brackets: "; ",
// ", and ", " and also ", or " and " before an imperative verb or "please".
const IMPERATIVE =
  "add|address|avoid|change|check|cite|clarify|comment|compare|consider|correct|define|delete|describe|discuss|elaborate|explain|expand|give|include|indicate|justify|list|mention|move|name|note|perform|present|provide|quantify|remove|rename|rephrase|replace|report|revise|rewrite|show|specify|state|summarise|summarize|test|update|use";
const CUT = new RegExp(
  `;\\s+|,\\s+and\\s+(?:also\\s+)?|\\s+and\\s+also\\s+|\\s+and\\s+(?=(?:please\\s+)?(?:${IMPERATIVE})\\b)`,
  "gi",
);

// Cut points of one sentence: [{ end, start }], the part before ends at
// `end` and the part after starts at `start`; the coordinator between them
// belongs to no part. Only at bracket depth 0, and only where both sides
// hold a letter.
export function clauseCuts(sentence) {
  const t = String(sentence ?? "");
  const depth = [];
  let d = 0;
  // By UTF-16 index, as matchAll reports it.
  for (let i = 0; i < t.length; i++) {
    const ch = t[i];
    if (ch === "(" || ch === "[") d += 1;
    depth.push(d);
    if ((ch === ")" || ch === "]") && d > 0) d -= 1;
  }
  const cuts = [];
  let from = 0;
  for (const m of t.matchAll(CUT)) {
    if (depth[m.index]) continue;
    const end = m.index;
    const start = m.index + m[0].length;
    if (!/\p{L}/u.test(t.slice(from, end)) || !/\p{L}/u.test(t.slice(start))) continue;
    cuts.push({ end, start });
    from = start;
  }
  return cuts;
}

// Parts of a sentence for one "split"/"keep" per cut point: a kept cut joins
// its neighbours with the coordinator as written. Each part is a substring
// of the sentence (trailing ",;" and spaces left off).
export function requirementParts(sentence, decisions) {
  const t = String(sentence ?? "");
  const cuts = clauseCuts(t);
  const parts = [];
  let start = 0;
  cuts.forEach((c, i) => {
    if (decisions[i] !== "split") return;
    parts.push(t.slice(start, c.end));
    start = c.start;
  });
  parts.push(t.slice(start));
  return parts.map((x) => x.replace(/[\s,;]+$/u, "").trim());
}

export const REQUIREMENT_SYSTEM =
  "You sort the sentences of peer-review comments. Label every sentence exactly once: one type, or cuts and types where cut marks are shown. Do not return any text of the review. Reply with JSON only.";

export const REQUIREMENT_TYPES = ["request", "other"];

export const REQUIREMENT_WARNING = "model could not label requirement sentences; add them by hand";

// A sentence with its cut points marked <c1>, <c2>, ... where each part
// before the coordinator ends.
function markCuts(sentence) {
  let out = "";
  let at = 0;
  clauseCuts(sentence).forEach((c, i) => {
    out += `${sentence.slice(at, c.end)}<c${i + 1}>`;
    at = c.end;
  });
  return out + sentence.slice(at);
}

// items: [{ ref, sentences, section? }], sentences numbered from 1 per ref.
// The section of a comment (taken off its line by the split) goes in as a
// context line, so "this section" can be read; it is not a sentence to label.
export function requirementPrompt(items) {
  return [
    'Return {"sentences":[{"ref":"R1.2","s":1,"type":"request"}, ...]} with every ref and sentence number exactly once.',
    'Types: "request" (asks the authors to do, change, add, explain or answer something; questions too), "other" (praise, summary, background, greeting).',
    'A sentence with marks <c1>, <c2>, ... may ask for more than one thing. For it give "cuts" instead of "type": one "split" or "keep" per mark in order, and "types": one type per part after the splits, e.g. {"ref":"R1.3","s":1,"cuts":["split"],"types":["request","request"]}. Split only where the parts ask for different things; keep when they are one request.',
    ...(items.some((it) => it.section)
      ? [
          'A "[ref section]" line names the part of the review the comment is under. It is context, not a sentence: give it no label.',
        ]
      : []),
    "Sentences:",
    ...items.flatMap((it) => [
      ...(it.section ? [`[${it.ref} section] ${it.section}`] : []),
      ...it.sentences.map((t, i) => `[${it.ref} s${i + 1}] ${markCuts(t)}`),
    ]),
  ].join("\n");
}

// One label of a sentence: a type for a sentence without cut points; for
// one with k cut points, k decisions and one type per part. null otherwise.
function sentenceLabel(x, sentence) {
  const k = clauseCuts(sentence).length;
  if (!k) return REQUIREMENT_TYPES.includes(x.type) && !("cuts" in x) ? [[], [x.type]] : null;
  const { cuts, types } = x;
  if (!Array.isArray(cuts) || cuts.length !== k) return null;
  if (!cuts.every((c) => c === "split" || c === "keep")) return null;
  const parts = 1 + cuts.filter((c) => c === "split").length;
  if (!Array.isArray(types) || types.length !== parts) return null;
  if (!types.every((t) => REQUIREMENT_TYPES.includes(t))) return null;
  return [cuts, types];
}

// The request texts per ref as { ref: [texts] }, or null unless every asked
// sentence is there exactly once with a valid label and nothing else is.
function validSentenceLabels(json, items) {
  const list = json?.sentences;
  const total = items.reduce((n, it) => n + it.sentences.length, 0);
  if (!Array.isArray(list) || list.length !== total) return null;
  const labels = Object.fromEntries(items.map((it) => [it.ref, new Array(it.sentences.length)]));
  const sentences = Object.fromEntries(items.map((it) => [it.ref, it.sentences]));
  for (const x of list) {
    const slots = Object.hasOwn(labels, x?.ref) ? labels[x.ref] : null;
    const s = x?.s;
    if (!slots || !Number.isInteger(s) || s < 1 || s > slots.length || slots[s - 1]) return null;
    const label = sentenceLabel(x, sentences[x.ref][s - 1]);
    if (!label) return null;
    slots[s - 1] = label;
  }
  return Object.fromEntries(
    items.map((it) => [
      it.ref,
      it.sentences.flatMap((sentence, i) => {
        const [cuts, types] = labels[it.ref][i];
        return requirementParts(sentence, cuts).filter((_, n) => types[n] === "request");
      }),
    ]),
  );
}

// comments: the split result. A comment that already has requirements, has
// no sentence or shares its ref is left as it is. One one-shot labels the
// sentences of the rest; a request sentence, or a request part of a sentence
// the model split at a cut point, becomes a requirement as written. A failed
// call or a reply that does not label every sentence once leaves the
// requirements empty, with a warning; the comments are otherwise unchanged.
// Returns { comments, warnings, asked, input, requests }: input is the prompt
// (for a fingerprint), requests the count of requirements per ref.
export async function extractRequirements(comments, { oneshot } = {}) {
  const none = { comments, warnings: [], asked: false, input: null, requests: null };
  if (!oneshot) return none;
  // A doubled ref (named in a split warning) is left for the user to renumber.
  const once = (ref) => comments.filter((c) => c.ref === ref).length === 1;
  const items = comments
    .filter((c) => !(c.requirements ?? []).length && once(c.ref))
    .map((c) => ({
      ref: c.ref,
      sentences: splitSentences(c.text),
      ...(c.section && { section: c.section }),
    }))
    .filter((it) => it.sentences.length);
  if (!items.length) return none;
  const input = requirementPrompt(items);
  const r = await oneshot({ system: REQUIREMENT_SYSTEM, prompt: input });
  const failed = (warning) => ({ ...none, warnings: [warning], asked: true, input });
  if (!r.ok) return failed(`requirement extraction failed: ${r.error}; add them by hand`);
  const texts = validSentenceLabels(r.json, items);
  if (!texts) return failed(REQUIREMENT_WARNING);
  return {
    comments: comments.map((c) =>
      texts[c.ref] ? { ...c, requirements: numberRequirements([], texts[c.ref]) } : c,
    ),
    warnings: [],
    asked: true,
    input,
    requests: Object.fromEntries(items.map((it) => [it.ref, texts[it.ref].length])),
  };
}
