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
//      (a change whose paragraph had an identical twin when it was made,
//      duplicate_before: the old text counts only at the recorded index,
//      elsewhere it may be the twin the tool never touched)
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
  const twin =
    same.length === 1 &&
    !!change.duplicate_before &&
    normalizeParaText(same[0].text) !== after &&
    same[0].index !== change.paragraph_index;
  if (same.length === 1 && !twin) return { index: same[0].index, method: "text" };
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
//   insert   L+I+R present (pending while I is a tracked insertion, else
//            accepted), L+R present: rejected
//   delete   L+D+R present (pending while D is a tracked deletion, else
//            rejected), L+R present: accepted
// "Tracked" is read from the paragraph's tracked runs when the snapshot
// has them (placedRuns); a snapshot without them falls back to the
// paragraph's revision mark total.
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

// The tracked runs of a snapshot paragraph (p.tracked, from the task pane:
// [{ kind: "ins" | "del", text, start, marks }]) as ranges of its normalized
// text t: { runs: [{ kind, start, end }], short: { ins, del } }. A run is
// placed where its text occurs nearest to the start it reports (Word's
// paragraph text and the OOXML text can differ by a character or two).
// short[kind] is true when the list cannot be trusted to hold every tracked
// run of that kind: a run whose text is not in t (or that has no text), or
// whose text is in t twice equally near its start, was left out, or the paragraph has more revision marks of that kind (w_ins,
// w_del, counted on their own) than the placed runs stand for. A change of
// that kind is then never read as "not tracked" (Codex review of the
// follow-up: a run that lost a no-break hyphen was dropped and its pending
// insertion read accepted).
// null: the snapshot has no per-run list (an older task pane).
export function placedRuns(p, t) {
  if (!Array.isArray(p?.tracked)) return null;
  const runs = [];
  const marks = { ins: 0, del: 0 };
  const short = { ins: false, del: false };
  for (const run of p.tracked) {
    const whole = collapse(run.text);
    const text = whole.trim();
    // run.start is the length of the collapsed, start-trimmed text before the
    // run (review-view.js:109-113 place()), the coordinates of t. The space a
    // run opens with is one character of t, unless the text before it ends
    // with a space (the two are one) or the run opens the paragraph (trimmed).
    const lead = whole.startsWith(" ") && run.start > 0 && t[run.start - 1] !== " " ? 1 : 0;
    const want = run.start + lead;
    let best = -1;
    let tie = false;
    if (text) {
      for (let i = t.indexOf(text); i !== -1; i = t.indexOf(text, i + 1)) {
        const d = Math.abs(i - want) - Math.abs(best - want);
        if (best === -1 || d < 0) {
          best = i;
          tie = false;
        } else if (d === 0) tie = true;
      }
    }
    if (best === -1 || tie) {
      if (run.kind in short) short[run.kind] = true;
      continue;
    }
    runs.push({ kind: run.kind, start: best, end: best + text.length });
    marks[run.kind] += run.marks ?? 1;
  }
  if ((p.w_ins ?? 0) > marks.ins) short.ins = true;
  if ((p.w_del ?? 0) > marks.del) short.del = true;
  return { runs, short };
}

// placed: placedRuns of the paragraph, or null. With it, a pure insertion or
// deletion is pending while its OWN text is tracked, or while the list is
// short for its kind (then nothing proves it is not); without it the
// paragraph's revision mark total decides (revisions), which reads an
// accepted change pending while any sibling is (Codex review 16).
function classifyByContext(change, t, revisions, placed = null) {
  const { left, right } = change.context;
  const I = collapse(change.text);
  const D = collapse(change.find);
  // Exactly one bounded match: a form found several times says nothing
  // about which occurrence is this change.
  const has = (form) => boundedMatches(t, normalizeParaText(form)).length === 1;
  // Is the middle part of the form L + own + R (found once) still tracked?
  // With context words (L or R) the match is the place of the change. With
  // none, the form is the change's own text alone and its one bounded match
  // is not known to be the change: "s" inserted after "control" and still
  // tracked, the words around it rewritten ("Our controls were assessed over
  // 5 s."), matches only the unit "s" at the end, where nothing is tracked,
  // and read accepted. So without context a tracked run of the same kind
  // that holds the change's text anywhere in the paragraph counts as its
  // own: the two cannot be told apart, and the change stays pending. A run
  // of other text, or of the other kind, does not hold it back.
  const ownTracked = (kind, L, own, R) => {
    if (!placed) return revisions > 0;
    if (placed.short[kind]) return true;
    const form = normalizeParaText(L + own + R);
    const at = boundedMatches(t, form)[0];
    const a = at + collapse(L).trimStart().length;
    const b = at + form.length - collapse(R).trimEnd().length;
    const noContext = !collapse(L).trim() && !collapse(R).trim();
    return placed.runs.some(
      (r) =>
        r.kind === kind &&
        ((r.start < b && r.end > a) || (noContext && t.slice(r.start, r.end).includes(form))),
    );
  };
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
      // A later change of the same turn inside this one's inserted text (the
      // find of a sibling falls in it) leaves no deleted text to show: Word
      // drops deleted text of a pending insertion outright. The form then
      // reads accepted while its text is still a tracked insertion.
      if (a && !b) return placed && ownTracked("ins", L, I, R) ? "pending" : "accepted";
      if (b && !a) return "rejected";
    } else if (change.kind === "insert") {
      const a = has(L + I + R);
      const b = has(L + R);
      if (a && !b) return ownTracked("ins", L, I, R) ? "pending" : "accepted";
      // L + R is also found where a sibling rewrote the right context and
      // its new text opens with the same words: the insertion still stands,
      // tracked, before the sibling. Its own text in a tracked insertion is
      // no reject; fewer context words decide.
      if (b && !a) {
        const own = boundedMatches(t, normalizeParaText(I)).length > 0;
        if (!(placed && own && ownTracked("ins", "", I, ""))) return "rejected";
      }
    } else {
      const a = has(L + R);
      const b = has(L + D + R);
      if (b && !a) return ownTracked("del", L, D, R) ? "pending" : "rejected";
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
    const state = classifyByContext(change, t, revisions, placedRuns(p, t));
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
