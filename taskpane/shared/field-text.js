// Fields (citations of a reference manager, cross-references, hyperlinks)
// in the text and the XML of a paragraph. Pure: no Office.js, no DOM, no
// Node API. The task pane (tools-word.js toolTrackedEdit) and the daemon
// (review-tools.mjs) both use it, so they decide the same way.
//
// The rule both hold a tracked edit to: an edit either leaves every field
// exactly as it was, or removes a field as a whole (begin through end
// inside the tracked deletion, so accepting removes the field and rejecting
// restores it). Nothing in between: a field whose visible text alone is
// deleted stays in the document with an empty result, and the reference
// manager writes the text again at its next refresh.
//
// Word evidence (the cite acceptance run of 2026-10-02, Word desktop):
//   - Range.getReviewedText gives a field as
//       U+0013 instruction U+0014 result U+0015
//     (a nested field inside the instruction the same way, U+0001 for its
//     data). Paragraph.text gives the result only. A field over several
//     paragraphs has its U+0013 ... U+0014 in the first and its U+0015 in
//     the last.
//   - A field whose result was deleted (tracked) reads U+0014 U+0015 with
//     nothing between them.
//   - A search matches the visible text across a field: the find
//     "... before use (Garrowby, 2017)." was found in its paragraph.
// Not seen, assumed: how a field that lies whole in a tracked deletion
// reads there (taken to be absent, marks and all).

const DOC_PART = 'pkg:name="/word/document.xml"';
const BEGIN = "\u0013";
const SEPARATE = "\u0014";
const END = "\u0015";
const MARKS = /[\u0013\u0014\u0015]/;

// What a refused edit says, to the model and, in the task pane, to the
// author. scripts/e2e/judge.mjs REFUSAL_TEXTS reads the first sentence.
export const FIELD_REFUSAL =
  "This edit would change the text of a citation or another field. Fields are written by the reference manager, so they are left as they are. Edit the words around it; to remove a citation, delete the whole sentence in one edit.";
const DETAILS = {
  inside: "The text to change lies inside a field.",
  partial: "The text to change starts or ends inside a field.",
  "replace-across":
    "The text to change holds a whole field; a replacement would turn it into plain words. Change the words before it and the words after it in separate edits.",
  "no-undo":
    "Deleting a sentence with its field needs an isolated edit that Word can safely take back if the field does not go as a whole; that cannot be established here.",
  "pane-outdated":
    "The task pane is older than the daemon and does not check fields. Tell the author to reload the task pane (close and reopen it).",
  "after-edit": "Word did not keep the field whole, so the edit was taken back.",
};
// What an undo of a tracked edit came to (tools-word.js takeBack), told from
// the fields of the paragraph and from three reviewed texts of it: before
// the edit, after the edit, and after the undo.
//   "clean"  the fields are as before the edit and the paragraph reads as
//            before the edit: the edit is gone and nothing else is.
//   "kept"   the edit is still in the document: a field is still not as it
//            was, or the paragraph still reads as it did after the edit.
//            Also the answer whenever a text could not be read: an edit
//            nobody can vouch for as gone is reported as in the document.
//   "more"   the fields are as before, the paragraph reads neither as before
//            the edit nor as after it: the edit is gone, and so is something
//            else (Word had merged the edit with a pending change of the
//            same author next to it, and rejected both).
// fieldsOk: fieldsKept(before, now, 0) === null. The texts are compared as
// the caller hands them in (the task pane: their paragraphHash).
export function undoOutcome({ fieldsOk, before, after, undone }) {
  if (!fieldsOk || before == null || undone == null) return "kept";
  if (undone === before) return "clean";
  if (after == null || undone === after) return "kept";
  return "more";
}
// What the model is told when Word took more back than the edit.
export const UNDO_TOOK_MORE =
  "While the edit was taken back, Word also took back an earlier pending change in the same paragraph. The edit is not in the document. The task pane is refreshing the status: read the paragraph again before you edit it, and do not repeat this edit.";

// A paragraph whose fields could not be read at the last status refresh: the
// edit is refused although nothing says that it touches a field.
export const FIELDS_UNREAD =
  "Not applied: the task pane could not read this paragraph as Word shows it with its fields (citations, a reference list), so an edit here could cut through one unseen. Nothing was changed. Tell the author to press Refresh status in the Review tab; if the paragraph still cannot be read, it has to be changed by hand in Word.";
export const fieldRefusal = (reason, extra = "") =>
  `${FIELD_REFUSAL} ${DETAILS[reason] ?? ""}${extra ? ` ${extra}` : ""} Nothing was changed.`.replace(
    /\s+/g,
    " ",
  );

// The paragraph as it is shown, normalized like normalizeParaText (U+0005
// dropped, white space collapsed, trimmed), with the instructions and the
// marks of its fields left out.
//   text    the shown text
//   raw     for each character of text its index in s
//   spans   the results of the outermost fields, { start, end } in text;
//           start === end for a field with no result. end is Infinity for a
//           field that does not end in this paragraph, start is 0 for one
//           that began in an earlier paragraph (carry).
//   carry   how many fields are open at the end (for the next paragraph)
// s may be normalized already. carry: fields open at the start, as the
// paragraph before left them.
export function fieldView(s, carry = 0) {
  const str = String(s ?? "");
  let text = "";
  const raw = [];
  const spans = [];
  // One entry per open field: is its result reached? Fields carried in are
  // in their result (a field instruction does not run over paragraphs).
  const open = Array.from({ length: carry }, () => true);
  let top = carry > 0 ? { start: 0, end: Infinity } : null;
  if (top) spans.push(top);
  for (let i = 0; i < str.length; i++) {
    const ch = str[i];
    if (ch === BEGIN) {
      if (open.length === 0) {
        top = { start: text.length, end: Infinity };
        spans.push(top);
      }
      open.push(false);
    } else if (ch === SEPARATE) {
      if (open.length) open[open.length - 1] = true;
      if (open.length === 1 && top) top.start = text.length;
    } else if (ch === END) {
      if (open.length === 0) {
        // The end of a field that began before this text and was not
        // carried in: everything so far was its result.
        spans.unshift({ start: 0, end: text.length });
        continue;
      }
      const hadResult = open.pop();
      if (open.length === 0 && top) {
        if (!hadResult) top.start = text.length;
        top.end = text.length;
        top = null;
      }
    } else if (ch === "\u0005" || !open.every(Boolean)) {
      // A comment anchor, or a character of an instruction.
    } else if (/\s/.test(ch)) {
      if (text && text[text.length - 1] !== " ") {
        text += " ";
        raw.push(i);
      }
    } else {
      text += ch;
      raw.push(i);
    }
  }
  if (text.endsWith(" ")) {
    text = text.slice(0, -1);
    raw.pop();
  }
  for (const sp of spans) {
    sp.start = Math.min(sp.start, text.length);
    if (sp.end !== Infinity) sp.end = Math.min(sp.end, text.length);
  }
  return { text, raw, spans, carry: open.length };
}

// The views of the paragraphs of a document in order: a field that runs
// over several paragraphs (a bibliography) is carried from one to the next.
export function fieldViews(texts) {
  let carry = 0;
  return (texts ?? []).map((t) => {
    const v = fieldView(t, carry);
    carry = v.carry;
    return v;
  });
}

// Does this text (getReviewedText) show a field?
export const hasFieldMarks = (s) => MARKS.test(String(s ?? ""));

// Is a paragraph one whose fields nobody has seen? Yes, when its reviewed
// text (or that of a paragraph before it: a field that runs over several
// paragraphs is carried from there) was not read at the last status refresh,
// unless the document is known to hold no field at all.
//   read.unreviewed  indexes of the paragraphs without reviewed text
//   read.fields      does the document hold a field anywhere: true, false,
//                    or null when that is not known
//   read.unknown     true: the snapshot says nothing about what was read
// No `read` at all (a caller that knows nothing of this): not refused.
export function fieldsUnread(read, index) {
  if (!read || read.fields === false) return false;
  if (read.unknown) return true;
  return Array.isArray(read.unreviewed) && read.unreviewed.some((k) => k <= index);
}

// Does the XML of a paragraph (Paragraph.getOoxml) hold a field of any kind:
// a complex field, a simple field or a hyperlink?
export function xmlHasField(xml) {
  const s = String(xml ?? "");
  const at = s.indexOf(DOC_PART);
  const end = at === -1 ? -1 : s.indexOf("</pkg:part>", at);
  const doc = at === -1 ? s : s.slice(at, end === -1 ? undefined : end);
  return /<w:(?:fldChar|fldSimple|instrText|delInstrText|hyperlink)\b/.test(doc);
}

const collapse = (s) =>
  String(s ?? "")
    .replace(/\u0005/g, "")
    .replace(/\s+/g, " ");

function occurrences(text, f) {
  const out = [];
  if (!f) return out;
  for (let i = text.indexOf(f); i !== -1; i = text.indexOf(f, i + 1)) out.push(i);
  return out;
}

// Where [ms, me) of the shown text lies with respect to the fields:
// { where: "none" | "whole" | "inside" | "partial", fields } (fields: how
// many it holds whole). "whole" needs a character of the range before the
// field and one after it: only then does the range Word finds for the text
// hold the field from its begin marker to its end marker.
function place(spans, ms, me) {
  let whole = 0;
  let where = "none";
  for (const { start, end } of spans) {
    const apart = start === end ? !(ms < start && me > start) : me <= start || ms >= end;
    if (apart) continue;
    if (ms < start && me > end) whole += 1;
    else if (ms >= start && me <= end) return { where: "inside", fields: 0 };
    else where = "partial";
  }
  if (where === "partial") return { where, fields: 0 };
  return { where: whole ? "whole" : "none", fields: whole };
}

// May this edit be made in the paragraph `view` (fieldView)?
//   find     the text of the call
//   edit     { replace } or { insert_after }, as the tool takes them
//   narrow   { find, replace }: the narrower form, when it is the one
//            applied (the task pane: it knows; the daemon: when narrow.find
//            is in `find` exactly once, which is when the pane narrows)
// Returns
//   { ok: true, remove: n }   n fields go as a whole with a deletion
//   { ok: false, reason }     inside, partial, replace-across
//   { ok: true, remove: 0, unknown: true }   `find` is not in the shown text
//       exactly once: nothing can be told here. The check of the paragraph
//       after the edit (fieldsKept) still holds.
export function fieldVerdict(view, find, edit = {}, narrow = null) {
  if (!view || !view.spans.length) return { ok: true, remove: 0 };
  const f = collapse(find);
  const at = occurrences(view.text, f);
  if (at.length !== 1) return { ok: true, remove: 0, unknown: true };
  let ms = at[0];
  let me = ms + f.length;
  if (typeof edit.insert_after === "string") {
    // Only the place right after the anchor matters: inside a result, or at
    // its end (before the end marker or after it cannot be told), is refused.
    const hit = view.spans.some((sp) => sp.start < me && me <= sp.end);
    return hit ? { ok: false, reason: "inside" } : { ok: true, remove: 0 };
  }
  let replace = edit.replace;
  if (narrow && typeof narrow.find === "string" && narrow.find) {
    const nf = collapse(narrow.find);
    const sub = occurrences(f, nf);
    if (sub.length === 1) {
      ms += sub[0];
      me = ms + nf.length;
      replace = narrow.replace;
    }
  }
  const p = place(view.spans, ms, me);
  if (p.where === "none") return { ok: true, remove: 0 };
  if (p.where !== "whole") return { ok: false, reason: p.where };
  if (replace !== "") return { ok: false, reason: "replace-across" };
  return { ok: true, remove: p.fields };
}

// `text` (a paragraph as getReviewedText gives it, normalized or not) with
// the one place of `find` in its shown text replaced by `to`, or with `to`
// put in after it (insert). A range that holds a field whole takes the
// field with it, instruction and marks. null when `find` is not in the
// shown text exactly once.
export function spliceShown(text, find, to, insert = false) {
  const str = String(text ?? "");
  const view = fieldView(str);
  const f = collapse(find);
  const at = occurrences(view.text, f);
  if (at.length !== 1) return null;
  const a = view.raw[at[0]];
  const b = view.raw[at[0] + f.length - 1] + 1;
  return str.slice(0, insert ? b : a) + to + str.slice(b);
}

// ---- the fields of a paragraph, from its XML --------------------------------

const SIDE = { ins: "ins", moveTo: "ins", del: "del", moveFrom: "del" };
const XML_TOKEN = new RegExp(
  [
    // 1 2: a tracked element opens (2: "/" when it is empty, as in w:rPr)
    "<w:(ins|del|moveTo|moveFrom)\\b[^>]*?(/?)>",
    // 3: it closes
    "</w:(ins|del|moveTo|moveFrom)>",
    // 4: a field marker
    '<w:fldChar\\b[^>]*?w:fldCharType="(begin|separate|end)"',
    // an empty text element
    "<w:(?:t|delText|instrText|delInstrText)\\b[^>]*?/>",
    // 5 6: text of a run, 5 says which kind
    "<w:(t|delText|instrText|delInstrText)\\b[^>]*>([\\s\\S]*?)</w:\\5>",
  ].join("|"),
  "g",
);
const ENTITY = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };
const unescape = (s) =>
  s.replace(/&(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);/g, (m, e) =>
    e[0] === "#"
      ? String.fromCodePoint(e[1] === "x" ? parseInt(e.slice(2), 16) : Number(e.slice(1)))
      : ENTITY[e],
  );

// The complex fields of a paragraph (Paragraph.getOoxml, or the XML of a
// w:p), in document order, each
//   { code, state, marks, nested, text, all }
// state: "live" (its markers and its instruction outside every tracked
// change), "deleted" / "inserted" (all of them inside deletions /
// insertions), "split" (not in one kind of place), "open" (no end marker in
// this paragraph). marks: the kinds of place its markers lie in ("-",
// "del", "-,del"), which tells an open field whose begin went into a
// deletion from one that is as it was. nested: it lies inside another
// field of the paragraph. text: its result as it reads with
// the tracked changes accepted; all: its result with the deleted text too.
// The same rules as tools/citation_fields_check.py, which judges the saved
// file. stray (a property of the list): the end and separate markers
// without a begin in this paragraph (it starts inside a field), each with
// the kind of place it lies in.
export function paragraphFields(xml) {
  const s = String(xml ?? "");
  const at = s.indexOf(DOC_PART);
  const end = at === -1 ? -1 : s.indexOf("</pkg:part>", at);
  const doc = at === -1 ? s : s.slice(at, end === -1 ? undefined : end);
  const tracked = [];
  const open = [];
  const fields = [];
  const stray = [];
  for (const m of doc.matchAll(XML_TOKEN)) {
    if (m[1]) {
      if (!m[2]) tracked.push(SIDE[m[1]]);
      continue;
    }
    if (m[3]) {
      const k = tracked.lastIndexOf(SIDE[m[3]]);
      if (k !== -1) tracked.splice(k, 1);
      continue;
    }
    const deleted = tracked.includes("del") || m[5] === "delText" || m[5] === "delInstrText";
    const ctx = deleted ? "del" : tracked.includes("ins") ? "ins" : "-";
    const top = open[open.length - 1];
    if (m[4] === "begin") {
      const f = {
        code: "",
        places: new Set([ctx]),
        phase: "code",
        text: "",
        all: "",
        state: "open",
        nested: open.length > 0,
      };
      fields.push(f);
      open.push(f);
    } else if (m[4]) {
      if (!top) {
        stray.push(ctx);
        continue;
      }
      top.places.add(ctx);
      if (m[4] === "separate") top.phase = "result";
      else {
        open.pop();
        const [only] = top.places;
        top.state =
          top.places.size !== 1 ? "split" : { "-": "live", del: "deleted", ins: "inserted" }[only];
      }
    } else if (m[5] === "instrText" || m[5] === "delInstrText") {
      if (top && top.phase === "code") {
        top.code += unescape(m[6]);
        top.places.add(ctx);
      }
    } else if (m[5] && top && top.phase === "result") {
      const t = unescape(m[6]);
      top.all += t;
      if (ctx !== "del") top.text += t;
    }
  }
  const list = fields.map(({ code, state, places, nested, text, all }) => ({
    code,
    state,
    marks: [...places].sort().join(","),
    nested,
    text,
    all,
  }));
  list.stray = stray.join(",");
  return list;
}

// Did an edit keep the fields of its paragraph? before, after:
// paragraphFields of the paragraph before and after the edit. remove: how
// many fields the edit was to take away as a whole (fieldVerdict). null
// when it did, else what is wrong:
//   "count a->b"   a field is gone or new without a tracked change
//   "code"         an instruction differs
//   "split"        the markers of a field are not in one kind of place
//   "hollowed"     a field is still there and nothing of its result is left
//   "result"       a field is still there with another result
//   "removed n"    more fields went into the deletion than the edit named
//                  (a field inside another one goes with it and is not counted)
// A field that is as it was passes; so does one that went into a tracked
// deletion with everything it has, when the edit was to remove fields.
export function fieldsKept(before, after, remove = 0) {
  if (before.length !== after.length) return `count ${before.length}->${after.length}`;
  if ((before.stray ?? "") !== (after.stray ?? "")) return "split";
  let removed = 0;
  for (let i = 0; i < before.length; i++) {
    const b = before[i];
    const a = after[i];
    if (a.code !== b.code) return "code";
    if (a.state === b.state && a.marks === b.marks && a.text === b.text && a.all === b.all)
      continue;
    if (a.state === "split") return "split";
    if (b.state === "live" && a.state === "deleted" && a.text === "" && a.all === b.all) {
      if (!a.nested) removed += 1;
      continue;
    }
    return a.text.trim() === "" && b.text.trim() !== "" ? "hollowed" : "result";
  }
  return removed > remove ? `removed ${removed}` : null;
}
