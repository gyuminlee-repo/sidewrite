// The other places of a comment: paragraphs its changes did not touch that
// are likely to state the same thing as the paragraphs they did.
//
// A turn changes the sentence a comment points at. Replays against a real
// revision showed what it then leaves as it was: the section heading above,
// the other paragraphs of that section, the legends of the figures the
// changed paragraphs cite, and the methods paragraph on the same analysis.
// A second turn of the comment ("Check other places" on its card) is given
// these places by name (daemon/paragraph-check.mjs followNote) and may
// change no other paragraph (review-tools.mjs NotAFollowPlace). This file
// finds them; it judges nothing.
//
// One implementation for both sides, like named-terms.js: the daemon lists
// the places in the prompt, the task pane counts them for the button. No
// imports but the other shared files, nothing of the browser or of node.
//
// Where the places come from, in the order they are listed:
//   heading  the nearest heading above each changed paragraph
//   legend   the captions of the figures, tables and schemes named in the
//            comment or in the text a change took out or put in, then those
//            only the rest of a changed paragraph names, as long as fewer
//            than LEGEND_PLACES legend paragraphs are listed: a paragraph
//            that cites ten figures must not fill the list with legends.
//            ("Fig. 5", "Figs. 2, 3, and 4", "Figure 5C", "Table S5";
//            "Supplementary Figure 2" and "Supp. Fig. 2" are S2; a panel
//            letter is not read.) With a caption go the paragraphs between
//            it and the next caption or heading when they are at most
//            CONTINUATION: a legend often goes on there; a longer run is
//            body text that follows a caption.
//   methods  the methods paragraphs that hold a term the comment names
//            (named-terms.js), a removed word that the comment and the
//            reason of its change both name (residue.js claimTerms), or one
//            of at most METHOD_WORDS other removed words: six letters or
//            more, or a technical token (a digit, a hyphen or a capital
//            inside), not a common word of reports (COMMON), and in at most
//            2% of the body and methods paragraphs (never fewer than 2,
//            never more than 5). Those in the fewest paragraphs first, and
//            at most METHOD_PLACES paragraphs for one word. A judged run
//            listed methods paragraphs for "across" and "significant".
//   section  the other body and methods paragraphs under the same heading
// A paragraph is listed once, under the first of these. Left out: the
// paragraphs the comment changed, those where the author rejected a change
// of this comment (the author has decided there), empty ones, and the
// reference list (from a heading "References" or "Bibliography" to the next
// caption or heading; a short entry without a full stop at its end reads as
// a heading, so there only a line of at most four words without a digit or
// a full stop ends it).
// A paragraph without two letters in a row ("/", "0.05", a table cell) is
// no place, no heading and no end of a section: paragraphKinds takes such a
// line for a heading, here it is passed over.
// At most MAX_PLACES paragraphs and MAX_CHARS characters of text. A
// paragraph that does not fit into what is left of the characters is passed
// over; when none fits, the first is listed alone. A comment without a
// change that is in the document (not rejected, not unlocated) has no places.
// Assumed: English; paragraphKinds tells headings, captions and methods.

import { paragraphKinds, termPlaces } from "./named-terms.js";
import { claimTerms, removedTerms } from "./residue.js";

export const MAX_PLACES = 20;
export const MAX_CHARS = 40_000;
const CONTINUATION = 3;
const LEGEND_PLACES = 8;
const METHOD_WORDS = 3;
const METHOD_PLACES = 3;
// The longest range of figure numbers that is read as one ("Figs. 2-4").
const RANGE = 10;

const live = (ch) => ch.state !== "rejected" && ch.state !== "unlocated";

// The changes of a comment that are in the document.
export const liveChanges = (comment) => (comment?.changes ?? []).filter(live);

// The indexes of the paragraphs some changes are in: the paragraph the task
// pane id of a change names, else the one at its index. ids: the task pane
// ids of the paragraphs, or null; count: how many paragraphs there are.
function indexesOf(changes, ids, count) {
  const out = new Set();
  for (const ch of changes) {
    const byId = ch.paragraph_id == null ? -1 : (ids ?? []).indexOf(ch.paragraph_id);
    const i = byId !== -1 ? byId : ch.paragraph_index;
    if (Number.isInteger(i) && i >= 0 && i < count) out.add(i);
  }
  return [...out].sort((a, b) => a - b);
}

// The paragraphs the live changes of a comment are in.
export const changedIndexes = (comment, ids, count) => indexesOf(liveChanges(comment), ids, count);

const LABEL = String.raw`S?\d+[A-Za-z]?(?![\p{L}\p{N}])`;
const SUPPLEMENT = String.raw`supp(?:l(?:ement(?:ary|al)?)?)?\.?\s+`;
const POINTER = new RegExp(
  String.raw`\b(${SUPPLEMENT})?(figs?|figures?|tables?|schemes?)\.?\s*(${LABEL}(?:\s*(?:,\s*(?:and|&)|,|and|&|to|–|\u2014|-)\s*${LABEL})*)`,
  "giu",
);
const CAPTION_LABEL = /^(supplementary\s+)?(fig(?:ure|\.)?|table|scheme)\s*\.?\s*(S?)(\d+)/i;
const NAMES = { fig: "Fig.", tab: "Table", sch: "Scheme" };
const key = (word, s, n) => `${word.toLowerCase().slice(0, 3)}:${s ? "S" : ""}${Number(n)}`;
const keyName = (k) => `${NAMES[k.slice(0, 3)]} ${k.slice(4)}`;

// The figures, tables and schemes a text points at, as keys ("fig:5",
// "fig:S7", "tab:S5"). "Figs. 2-4" is 2, 3 and 4.
export function figurePointers(text) {
  const out = new Set();
  for (const m of String(text ?? "").matchAll(POINTER)) {
    let last = null;
    for (const p of m[3].matchAll(/(to|–|\u2014|-)?\s*(S?)(\d+)/giu)) {
      const s = !!m[1] || !!p[2];
      const n = Number(p[3]);
      const from = p[1] && last?.s === s && n > last.n && n - last.n <= RANGE ? last.n : n;
      for (let i = from; i <= n; i++) out.add(key(m[2], s, i));
      last = { s, n };
    }
  }
  return [...out];
}

// The key of the figure or table a caption paragraph is the legend of.
function captionKey(text) {
  const m = CAPTION_LABEL.exec(String(text ?? "").trim());
  return m ? key(m[2], !!m[1] || !!m[3], m[4]) : null;
}

const REFERENCE_LIST = /^(?:\d+\.?\s*)?(?:references|bibliography|literature cited|works cited)\b/i;
const escapeRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
const wordRe = (w) => new RegExp(`(?<![\\p{L}\\p{N}])${escapeRe(w)}(?![\\p{L}\\p{N}])`, "iu");

// Words any report has: a change that drops one says nothing about the
// methods. On top of the function words residue.js leaves out.
const COMMON = new Set(
  (
    "across significant significantly biological different several observed increased decreased higher lower " +
    "similar respectively however therefore further between within during results result showed shown " +
    "indicate indicates indicated suggest suggests suggested various overall compared relative general " +
    "generally important importantly likely possible potential potentially previous previously recent " +
    "recently present presented obtained performed consistent demonstrated revealed reported found " +
    "strongly clearly highly largely mainly primarily robust robustly notably markedly substantially"
  ).split(" "),
);
// A removed word as the change had it: a technical token has a digit, a
// hyphen or a capital after its first letter ("GPS-2", "McAdam").
function technical(word, changes) {
  if (/[\d-]/.test(word)) return true;
  const re = wordRe(word);
  return changes.some((ch) => {
    const form = re.exec(String(ch.narrowed_from?.find ?? ch.find ?? ""))?.[0] ?? "";
    return /\p{Lu}/u.test(form.slice(1));
  });
}

// [{ index, id, kind, why }]: the other places of a comment, in the order
// they are listed. paragraphs: the texts as they read now; ids: their task
// pane ids, or null.
export function followPlaces(
  comment,
  paragraphs,
  ids = null,
  { maxPlaces = MAX_PLACES, maxChars = MAX_CHARS } = {},
) {
  const texts = (paragraphs ?? []).map((t) => String(t ?? ""));
  const changed = changedIndexes(comment, ids, texts.length);
  if (!changed.length) return [];
  const kinds = paragraphKinds(texts);
  const changes = liveChanges(comment);
  const words = texts.map((t) => /\p{L}{2}/u.test(t));
  const heading = (i) => kinds[i] === "heading" && words[i];

  const skip = new Set([
    ...changed,
    ...indexesOf(
      (comment?.changes ?? []).filter((ch) => ch.state === "rejected"),
      ids,
      texts.length,
    ),
  ]);
  let inList = false;
  texts.forEach((t, i) => {
    if (kinds[i] === "caption") inList = false;
    else if (heading(i) && REFERENCE_LIST.test(t.trim())) inList = true;
    else if (heading(i) && !/[\d.]/.test(t) && t.trim().split(/\s+/).length <= 4) inList = false;
    if (inList || !words[i]) skip.add(i);
  });

  const found = new Map();
  const add = (index, why) => {
    if (skip.has(index) || found.has(index)) return false;
    found.set(index, why);
    return true;
  };

  // A changed heading is the heading of its own section.
  const headOf = (i) => {
    for (let h = i; h >= 0; h--) if (heading(h)) return h;
    return -1;
  };
  const heads = [...new Set(changed.map(headOf))].filter((h) => h !== -1);
  for (const h of heads) add(h, "heading above a changed paragraph");

  const named = new Set(
    [
      comment?.text,
      ...changes.flatMap((ch) => [
        ch.narrowed_from?.find ?? ch.find,
        ch.narrowed_from?.text ?? ch.text,
      ]),
    ].flatMap(figurePointers),
  );
  const cited = new Set(
    changed.flatMap((i) => figurePointers(texts[i])).filter((k) => !named.has(k)),
  );
  let legends = 0;
  for (const [keys, most] of [
    [named, Infinity],
    [cited, LEGEND_PLACES],
  ]) {
    texts.forEach((t, i) => {
      const k = kinds[i] === "caption" ? captionKey(t) : null;
      if (!k || !keys.has(k) || legends >= most) return;
      if (add(i, `legend of ${keyName(k)}`)) legends++;
      const run = [];
      for (let j = i + 1; j < texts.length && kinds[j] !== "caption" && !heading(j); j++)
        if (words[j]) run.push(j);
      if (run.length > CONTINUATION) return;
      for (const j of run) if (add(j, `continues the legend of ${keyName(k)}`)) legends++;
    });
  }

  const inMethods = (re) =>
    texts.flatMap((t, i) => (kinds[i] === "methods" && re.test(t) ? [i] : []));
  const also = (i, word) => add(i, `methods; also has "${word}"`);
  for (const row of termPlaces(comment, texts, ids))
    for (const p of row.places) if (p.kind === "methods") also(p.index, row.term);
  const claimed = new Set(changes.flatMap((ch) => claimTerms(ch, comment?.text)));
  for (const word of claimed) for (const i of inMethods(wordRe(word))) also(i, word);
  const prose = texts.flatMap((t, i) =>
    words[i] && (kinds[i] === "body" || kinds[i] === "methods") ? [t] : [],
  );
  const most = Math.min(5, Math.max(2, Math.ceil(prose.length * 0.02)));
  [...new Set(changes.flatMap(removedTerms))]
    .filter((w) => !claimed.has(w) && !COMMON.has(w))
    .filter((w) => w.length >= 6 || technical(w, changes))
    .map((word) => {
      const re = wordRe(word);
      return { word, n: prose.filter((t) => re.test(t)).length, at: inMethods(re) };
    })
    .filter((r) => r.at.length && r.n <= most)
    .sort((a, b) => a.n - b.n)
    .slice(0, METHOD_WORDS)
    .forEach((r) => r.at.slice(0, METHOD_PLACES).forEach((i) => also(i, r.word)));

  for (const h of heads)
    for (let i = h + 1; i < texts.length && !heading(i); i++)
      if (kinds[i] !== "caption") add(i, "same section as a changed paragraph");

  const out = [];
  let chars = 0;
  const place = ([index, why]) => ({
    index,
    id: ids?.[index] ?? `p${index}`,
    kind: kinds[index],
    why,
  });
  for (const row of found) {
    if (out.length >= maxPlaces) break;
    if (chars + texts[row[0]].length > maxChars) continue;
    chars += texts[row[0]].length;
    out.push(place(row));
  }
  if (!out.length && found.size) out.push(place([...found][0]));
  return out;
}

// What the card of a comment shows of this. snapshot: revision.snapshot.
// { offer, line }: offer is the number of places for the button "Check other
// places (N)", or null when there is nothing to check or the check ran and
// the comment got no change since (comment.follow_through.seen_change_ids);
// line is what the last check did, or null when none ran.
export function followState(comment, snapshot) {
  const ran = comment?.follow_through;
  const seen = new Set(ran?.seen_change_ids ?? []);
  const fresh = !ran || liveChanges(comment).some((ch) => !seen.has(ch.id));
  const n = fresh ? followPlaces(comment, snapshot?.paragraphs, snapshot?.ids).length : 0;
  const count = (k, word) => `${k} ${word}${k === 1 ? "" : "s"}`;
  const did = !ran
    ? ""
    : ran.change_ids.length || ran.questions
      ? `${count(ran.change_ids.length, "change")}, ${count(ran.questions, "question")}`
      : "no change needed";
  return {
    offer: n || null,
    line: ran ? `Other places checked: ${ran.candidate_ids.length} (${did}).` : null,
  };
}
