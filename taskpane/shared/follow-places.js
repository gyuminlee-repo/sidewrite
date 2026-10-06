// The other places of a comment: paragraphs its changes did not touch that
// are likely to state the same thing as the paragraphs they did.
//
// A turn changes the sentence a comment points at. Replays against a real
// revision showed what it then leaves as it was: the section heading above,
// the other paragraphs of that section, the legends of the figures the
// changed paragraphs cite, and the methods paragraph on the same analysis.
// A second turn of the comment ("Check other places" on its card) is given
// these places by name (daemon/paragraph-check.mjs followNote). This file
// finds them; it judges nothing.
//
// One implementation for both sides, like named-terms.js: the daemon lists
// the places in the prompt, the task pane counts them for the button. No
// imports but the other shared files, nothing of the browser or of node.
//
// Where the places come from, in the order they are listed:
//   heading  the nearest heading above each changed paragraph
//   legend   the captions of the figures, tables and schemes named in a
//            changed paragraph, in the text a change took out or put in, or
//            in the comment ("Fig. 5", "Figs. 2 and 3", "Figure 5C", "Table
//            S5"; "Supplementary Figure 2" is S2; a panel letter is not
//            read), and the paragraphs right after such a caption up to the
//            next caption or heading: a legend often goes on there. At most
//            CONTINUATION of those: a caption that stands between body
//            paragraphs must not bring the rest of its section as legend.
//   methods  the methods paragraphs that hold a term the comment names
//            (named-terms.js) or a word one of its changes took out
//            (residue.js removedTerms; a word that stands in more than a
//            fifth of the paragraphs is vocabulary of the paper, as there)
//   section  the other body and methods paragraphs under the same heading
// A paragraph is listed once, under the first of these. Left out: the
// paragraphs the comment changed, empty ones and the reference list (from a
// heading "References" or "Bibliography" to the next heading or caption).
// At most MAX_PLACES paragraphs and MAX_CHARS characters of text; what is
// over is dropped from the end. A comment without a change that is in the
// document (not rejected, not unlocated) has no places.
// Assumed: English; paragraphKinds tells headings, captions and methods.

import { paragraphKinds, termPlaces } from "./named-terms.js";
import { headingAbove } from "./paragraph-label.js";
import { removedTerms } from "./residue.js";

export const MAX_PLACES = 20;
export const MAX_CHARS = 40_000;
const CONTINUATION = 3;

const live = (ch) => ch.state !== "rejected" && ch.state !== "unlocated";

// The changes of a comment that are in the document.
export const liveChanges = (comment) => (comment?.changes ?? []).filter(live);

// The indexes of the paragraphs those changes are in: the paragraph the task
// pane id of a change names, else the one at its index. ids: the task pane
// ids of the paragraphs, or null; count: how many paragraphs there are.
export function changedIndexes(comment, ids, count) {
  const out = new Set();
  for (const ch of liveChanges(comment)) {
    const byId = ch.paragraph_id == null ? -1 : (ids ?? []).indexOf(ch.paragraph_id);
    const i = byId !== -1 ? byId : ch.paragraph_index;
    if (Number.isInteger(i) && i >= 0 && i < count) out.add(i);
  }
  return [...out].sort((a, b) => a - b);
}

const LABEL = String.raw`S?\d+[A-Za-z]?(?![\p{L}\p{N}])`;
const POINTER = new RegExp(
  String.raw`\b(supplementary\s+)?(figs?|figures?|tables?|schemes?)\.?\s*(${LABEL}(?:\s*(?:,|and|&|to|–|-)\s*${LABEL})*)`,
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
    for (const p of m[3].matchAll(/(to|–|-)?\s*(S?)(\d+)/giu)) {
      const s = !!m[1] || !!p[2];
      const n = Number(p[3]);
      const from = p[1] && last?.s === s && n > last.n && n - last.n <= MAX_PLACES ? last.n : n;
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

  const skip = new Set(changed);
  let inList = false;
  texts.forEach((t, i) => {
    if (kinds[i] === "heading" || kinds[i] === "caption")
      inList = kinds[i] === "heading" && REFERENCE_LIST.test(t.trim());
    if (inList || !t.trim()) skip.add(i);
  });

  const found = new Map();
  const add = (index, why) => {
    if (!skip.has(index) && !found.has(index)) found.set(index, why);
  };

  // A changed heading is the heading of its own section.
  const heads = [
    ...new Set(changed.map((i) => (kinds[i] === "heading" ? i : headingAbove(kinds, i)))),
  ].filter((h) => h !== -1);
  for (const h of heads) add(h, "heading above a changed paragraph");

  const cited = new Set(
    [
      comment?.text,
      ...changed.map((i) => texts[i]),
      ...changes.flatMap((ch) => [
        ch.narrowed_from?.find ?? ch.find,
        ch.narrowed_from?.text ?? ch.text,
      ]),
    ].flatMap(figurePointers),
  );
  texts.forEach((t, i) => {
    const k = kinds[i] === "caption" ? captionKey(t) : null;
    if (!k || !cited.has(k)) return;
    add(i, `legend of ${keyName(k)}`);
    for (let j = i + 1, n = 0; j < texts.length && n < CONTINUATION; j++) {
      if (kinds[j] === "caption" || kinds[j] === "heading") break;
      if (!texts[j].trim()) continue;
      add(j, `continues the legend of ${keyName(k)}`);
      n++;
    }
  });

  for (const row of termPlaces(comment, texts, ids))
    for (const p of row.places)
      if (p.kind === "methods") add(p.index, `methods; holds "${row.term}"`);
  const filled = texts.filter((t) => t.trim()).length;
  const most = Math.min(15, Math.max(3, Math.ceil(filled / 5)));
  for (const word of new Set(changes.flatMap(removedTerms))) {
    const re = new RegExp(`(?<![\\p{L}\\p{N}])${escapeRe(word)}(?![\\p{L}\\p{N}])`, "iu");
    const at = texts.flatMap((t, i) => (re.test(t) ? [i] : []));
    if (at.length > most) continue;
    for (const i of at)
      if (kinds[i] === "methods") add(i, `methods; holds "${word}", which a change took out`);
  }

  for (const h of heads)
    for (let i = h + 1; i < texts.length && kinds[i] !== "heading"; i++)
      if (kinds[i] !== "caption") add(i, "same section as a changed paragraph");

  const out = [];
  let chars = 0;
  for (const [index, why] of found) {
    chars += texts[index].length;
    if (out.length >= maxPlaces || chars > maxChars) break;
    out.push({ index, id: ids?.[index] ?? `p${index}`, kind: kinds[index], why });
  }
  return out;
}

// What the card of a comment shows of this. snapshot: revision.snapshot.
// { offer, line }: offer is the number of places for the button "Check other
// places (N)", or null when there is nothing to check or the check ran and
// the comment got no change since (comment.follow_through.seen_change_ids);
// line is what the last check did, or null when none ran.
export function followState(comment, snapshot) {
  const ran = comment?.follow_through ?? null;
  const seen = new Set(ran?.seen_change_ids ?? []);
  const fresh = !ran || liveChanges(comment).some((ch) => !seen.has(ch.id));
  const n = fresh ? followPlaces(comment, snapshot?.paragraphs, snapshot?.ids).length : 0;
  const count = (k, word) => `${k} ${word}${k === 1 ? "" : "s"}`;
  return {
    offer: n || null,
    line: ran
      ? `Other places checked: ${ran.candidate_ids.length} (${count(ran.change_ids.length, "change")}, ${count(ran.questions, "question")}).`
      : null,
  };
}
