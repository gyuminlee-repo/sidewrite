// Decide what happened to each tracked edit the review flow made, from a
// snapshot of the document as it is now. Pure: no Office.js, no I/O.
//
// Word evidence (2026-09-30, spike session f2748a6f, spike tool outputs):
//   - A pending tracked replace shows BOTH the inserted and the deleted text
//     in paragraph.text, separated by U+0005 comment-anchor characters, and
//     the paragraph OOXML has w:ins/w:del elements.
//   - After accept or reject every w:ins/w:del in the paragraph is gone and
//     the text (minus U+0005) equals the "after" or the "before" sentence.
//   - A comment on the deleted range survives accept (empty range) and
//     reject. A comment on the inserted range is removed by reject.
//
// Snapshot shape (built by the task pane):
//   { paragraphs: [{ index, id, text, w_ins, w_del }],
//     comments:   [{ id, paragraph_index, paragraph_id }] }

export const CHANGE_STATES = ["pending", "accepted", "rejected", "user_edited", "unlocated"];

export function normalizeParaText(s) {
  return String(s ?? "")
    .replace(/\u0005/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function byIndex(snapshot) {
  const m = new Map();
  for (const p of snapshot.paragraphs ?? []) m.set(p.index, p);
  return m;
}

// Where does this change live now? Tried in order:
//   1. paragraph_id (Word uniqueLocalId recorded at edit time)
//   2. the paragraph that holds the rationale comment of the change
//   3. the only paragraph whose text equals the after or the before sentence
//   4. the recorded paragraph_index, if that paragraph still matches
// Returns { index, method }; index is null when nothing fits or the text
// match is ambiguous (identical paragraphs), so a caller never guesses.
export function locateParagraph(change, snapshot) {
  const paras = snapshot.paragraphs ?? [];
  const idx = byIndex(snapshot);
  if (change.paragraph_id) {
    const p = paras.find((x) => x.id && x.id === change.paragraph_id);
    if (p) return { index: p.index, method: "paragraph_id" };
  }
  if (change.word_comment_id != null) {
    const c = (snapshot.comments ?? []).find(
      (x) => String(x.id) === String(change.word_comment_id),
    );
    if (c && c.paragraph_index != null && idx.has(c.paragraph_index)) {
      return { index: c.paragraph_index, method: "comment" };
    }
  }
  const before = normalizeParaText(change.before);
  const after = normalizeParaText(change.after);
  const same = paras.filter((p) => {
    const t = normalizeParaText(p.text);
    return t === after || t === before;
  });
  if (same.length === 1) return { index: same[0].index, method: "text" };
  if (same.length > 1) return { index: null, method: "ambiguous" };
  const rec = idx.get(change.paragraph_index);
  if (rec && (rec.w_ins ?? 0) + (rec.w_del ?? 0) > 0) {
    return { index: rec.index, method: "recorded_index" };
  }
  return { index: null, method: "not_found" };
}

// ---- per-change context -----------------------------------------------------
//
// One paragraph can hold several tool changes, so the whole paragraph text
// says little about one of them. At record time the daemon keeps `context`:
// up to CONTEXT_WORDS words right before and right after the range the task
// pane changed (for an insertion, around the insertion point). A change is
// then judged by which form of its own range is in the paragraph now
// (L = left, R = right, D = change.find, I = change.text):
//   replace  pending L+I+D+R (Word shows inserted then deleted text),
//            accepted L+I+R, rejected L+D+R
//   insert   L+I+R present (pending while revision marks remain, else
//            accepted), L+R present: rejected
//   delete   L+D+R present (pending while marks remain, else rejected),
//            L+R present: accepted
// A sibling change can alter the context words, so the check is repeated
// with fewer context words (3, 2, 1, 0) until one form alone matches, once
// and on word boundaries (boundedMatches). For a
// pure insertion or deletion the pending text equals the accepted or the
// rejected text, so the paragraph's revision marks still decide pending;
// such a change next to a pending sibling reads pending.

export const CONTEXT_WORDS = 3;

const collapse = (s) =>
  String(s ?? "")
    .replace(/\u0005/g, "")
    .replace(/\s+/g, " ");

function tailWords(s, k) {
  return k > 0 ? (s.match(/\S+\s*/g) ?? []).slice(-k).join("") : "";
}

function headWords(s, k) {
  return k > 0 ? (s.match(/\s*\S+/g) ?? []).slice(0, k).join("") : "";
}

// before: the paragraph text the task pane saw; find: the find of the tool
// call (unique in that paragraph); narrowFind: the part of find the task
// pane actually replaced (L3), if it narrowed. Returns { left, right } or
// null when find is not exactly once in before.
export function changeContext(before, find, { kind = "replace", narrowFind = null } = {}) {
  const b = normalizeParaText(before);
  const f = collapse(find);
  const i = f.trim() ? b.indexOf(f) : -1;
  if (i === -1 || b.indexOf(f, i + 1) !== -1) return null;
  let start = i;
  let end = i + f.length;
  if (kind === "insert") {
    start = end;
  } else if (narrowFind) {
    const j = f.indexOf(collapse(narrowFind));
    if (j === -1) return null;
    start = i + j;
    end = start + collapse(narrowFind).length;
  }
  return {
    left: tailWords(b.slice(0, start), CONTEXT_WORDS),
    right: headWords(b.slice(end), CONTEXT_WORDS),
  };
}

// Start indexes of the occurrences of `form` in `text` that do not cut a
// word or a number: the character before and after a match may not be a
// letter or a digit when the form's own first or last character is one.
// Without context words a bare "7" would otherwise match inside "70" (Codex
// review 3). Shared with doc-diff.mjs.
const WORD_CHAR = /[\p{L}\p{N}]/u;
export function boundedMatches(text, form) {
  const out = [];
  if (!form) return out;
  const wordStart = WORD_CHAR.test(form[0]);
  const wordEnd = WORD_CHAR.test(form[form.length - 1]);
  for (let i = text.indexOf(form); i !== -1; i = text.indexOf(form, i + 1)) {
    const end = i + form.length;
    if (wordStart && i > 0 && WORD_CHAR.test(text[i - 1])) continue;
    if (wordEnd && end < text.length && WORD_CHAR.test(text[end])) continue;
    out.push(i);
  }
  return out;
}

function classifyByContext(change, t, revisions) {
  const { left, right } = change.context;
  const I = collapse(change.text);
  const D = collapse(change.find);
  // Exactly one bounded match: a form found several times says nothing
  // about which occurrence is this change.
  const has = (form) => boundedMatches(t, normalizeParaText(form)).length === 1;
  for (let k = CONTEXT_WORDS; k >= 0; k--) {
    const L = tailWords(left, k);
    const R = headWords(right, k);
    if (change.kind === "replace") {
      // With no context words left, a form that contains the other ("source"
      // inside "source, followed by ...") is found inside an edited version
      // of the other and says nothing: an accepted then edited change read
      // as rejected.
      if (k === 0 && (I.includes(D) || D.includes(I))) continue;
      if (has(L + I + D + R)) return "pending";
      const a = has(L + I + R);
      const b = has(L + D + R);
      if (a && !b) return "accepted";
      if (b && !a) return "rejected";
    } else if (change.kind === "insert") {
      const a = has(L + I + R);
      const b = has(L + R);
      if (a && !b) return revisions > 0 ? "pending" : "accepted";
      if (b && !a) return "rejected";
    } else {
      const a = has(L + R);
      const b = has(L + D + R);
      if (b && !a) return revisions > 0 ? "pending" : "rejected";
      if (a && !b) return "accepted";
    }
  }
  return null;
}

export function classifyChange(change, snapshot) {
  const loc = locateParagraph(change, snapshot);
  const commentIds = new Set((snapshot.comments ?? []).map((c) => String(c.id)));
  // Word removes the comment on a rejected inserted range, but the comment
  // cleanup (comment_cleaned_at) and the author remove comments too. So a
  // gone comment backs a reject only where the text already reads the old
  // form, never on its own (Codex review 7).
  const commentGone =
    change.word_comment_id != null &&
    !change.comment_cleaned_at &&
    !commentIds.has(String(change.word_comment_id));
  if (loc.index === null) {
    // A rejected pure insertion whose paragraph is no longer unique or was
    // edited since: rejected when the comment Word removes is gone and the
    // text reads the old form, either every paragraph matching a sentence
    // reads the old one or the recorded paragraph holds the old form of the
    // change's own range (context).
    if (change.kind === "insert" && commentGone) {
      const before = normalizeParaText(change.before);
      const after = normalizeParaText(change.after);
      const texts = (snapshot.paragraphs ?? []).map((p) => normalizeParaText(p.text));
      const rec = byIndex(snapshot).get(change.paragraph_index);
      const oldForm =
        loc.method === "ambiguous"
          ? texts.includes(before) && !texts.includes(after)
          : !!change.context &&
            !!rec &&
            classifyByContext(change, normalizeParaText(rec.text), 0) === "rejected";
      if (oldForm) {
        return { state: "rejected", paragraph_index: null, method: `${loc.method}+comment_gone` };
      }
    }
    return { state: "unlocated", paragraph_index: null, method: loc.method };
  }
  const p = byIndex(snapshot).get(loc.index);
  const t = normalizeParaText(p.text);
  const before = normalizeParaText(change.before);
  const after = normalizeParaText(change.after);
  const out = (state) => ({ state, paragraph_index: loc.index, method: loc.method });
  const revisions = (p.w_ins ?? 0) + (p.w_del ?? 0);
  // Changes recorded with context are judged on their own range; changes
  // recorded before context existed, or no single form matched, fall
  // through to the whole-paragraph rule below.
  if (change.context) {
    const state = classifyByContext(change, t, revisions);
    if (state) return out(state);
  }
  if (revisions > 0) {
    // A pending replace reads as inserted+deleted text, i.e. neither side.
    // A pending pure insertion reads as "after" and a pending pure deletion
    // as "before", so for those kinds any revision mark means pending.
    if (change.kind !== "replace" || (t !== before && t !== after)) return out("pending");
  }
  if (t === after) return out("accepted");
  if (t === before) return out("rejected");
  return out("user_edited");
}

// Classify every change in a revision. Returns rows the store can apply.
export function classifyRevision(revision, snapshot) {
  const rows = [];
  for (const c of revision.comments ?? []) {
    for (const ch of c.changes ?? []) {
      rows.push({ ref: c.ref, change_id: ch.id, ...classifyChange(ch, snapshot) });
    }
  }
  return rows;
}

export function summarizeStates(rows) {
  const counts = Object.fromEntries(CHANGE_STATES.map((s) => [s, 0]));
  for (const r of rows) counts[r.state] = (counts[r.state] ?? 0) + 1;
  return counts;
}
