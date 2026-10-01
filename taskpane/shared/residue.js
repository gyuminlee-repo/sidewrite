// Terms a comment's changes took out of the manuscript that are still in
// other paragraphs. A reviewer who rejects a claim rejects it everywhere it
// is made; a turn tends to change the sentence the comment points at and
// leave the abstract, the legends and the discussion as they were. This is
// a report for the card of the comment: nothing is refused.
//
// A term: a word of five letters or more in the text a change replaced or
// deleted that the new text no longer has in any form (the first letters
// are compared, so "inhibits" stays when "inhibited" comes in) and that is
// not a common function word. Assumed: English prose. Numbers are not
// terms. A word found in more than a fifth of the paragraphs (never fewer
// than 3, never more than `maxPlaces`) is taken to be vocabulary of the
// paper and not the claim, and is left out: "growth" in a paper on growth.
// At most `maxTerms` terms are reported, those in the fewest places first.
//
// Which removed words are the claim: any rewording drops words ("suggests",
// "primary", "likely"), and a real revision of 20 comments gave 88 lines of
// those. The term of a claim is one the reviewer named and the model names
// again as its reason, so a removed word counts only when the comment text
// and the why or addresses of its change both hold it, in any form. (What a
// change says it did quotes the replaced text, so it names every word.)
// A comment that names no term ("the claim in line 203"), and a rationale
// in another language that does not quote the term, give no line here. In
// front of these lines goes what the turn itself said is left
// (comment.same_term_at, from its last line "Same term still at: ...").

import { namedTerms, namedTermLines } from "./named-terms.js";

const STOP = new Set(
  (
    "about above after again against along already also although among another around because been before being below " +
    "between beyond could during either every further having hence however itself might other others rather should " +
    "since still their theirs there thereby therefore these those though through throughout under until upon where " +
    "whereas whether which while whose within without would using based shown"
  ).split(" "),
);

const words = (text) =>
  String(text ?? "")
    .toLowerCase()
    .match(/\p{L}[\p{L}\p{N}'-]*/gu) ?? [];
const stem = (w) => w.slice(0, Math.max(5, w.length - 3));

// The terms one change took out, in the order they stood.
export function removedTerms(change) {
  if (!change || change.kind === "insert") return [];
  const before = words(change.narrowed_from?.find ?? change.find);
  const after = words(change.narrowed_from?.text ?? change.text);
  const out = [];
  for (const w of before) {
    if (w.length < 5 || STOP.has(w) || out.includes(w)) continue;
    if (after.some((a) => a.startsWith(stem(w)))) continue;
    out.push(w);
  }
  return out;
}

// The removed words of a change that the reason of the change and the
// reviewer comment both name.
export function claimTerms(change, commentText) {
  const r = change?.rationale;
  const names = (text) => {
    const said = words(text);
    return (w) => said.some((x) => x.startsWith(stem(w)));
  };
  const reason = names([r?.why, r?.addresses].filter(Boolean).join(" "));
  const reviewer = names(commentText);
  return removedTerms(change).filter((w) => reason(w) && reviewer(w));
}

const escapeRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

// [{ term, change_ids, places }]: places are paragraph indexes. Rejected
// changes took nothing out.
export function stillPresent(comment, paragraphs, { maxPlaces = 15, maxTerms = 3 } = {}) {
  const texts = (paragraphs ?? []).map((t) => String(t ?? ""));
  const filled = texts.filter((t) => t.trim()).length;
  const most = Math.min(maxPlaces, Math.max(3, Math.ceil(filled / 5)));
  const by = new Map();
  for (const ch of comment?.changes ?? []) {
    if (ch.state === "rejected") continue;
    for (const term of claimTerms(ch, comment.text)) by.set(term, [...(by.get(term) ?? []), ch.id]);
  }
  const rows = [];
  for (const [term, change_ids] of by) {
    const re = new RegExp(`(?<![\\p{L}\\p{N}])${escapeRe(term)}(?![\\p{L}\\p{N}])`, "iu");
    const places = texts.flatMap((t, i) => (re.test(t) ? [i] : []));
    if (places.length && places.length <= most) rows.push({ term, change_ids, places });
  }
  return rows.sort((a, b) => a.places.length - b.places.length).slice(0, maxTerms);
}

// The lines of the card. snapshot: revision.snapshot ({ paragraphs, ids }).
export function stillPresentLines(comment, snapshot) {
  const said = comment?.same_term_at ?? [];
  const head = said.length ? [`The turn says the same term is still at: ${said.join(", ")}.`] : [];
  // The terms the reviewer named, once a turn has made changes: where they
  // still stand. A comment that names terms gets these lines in place of the
  // removed-word lines (named-terms.js); one that names none keeps those.
  const turnRan = (comment?.changes ?? []).length > 0;
  if (namedTerms(comment).length) {
    return [...(turnRan ? namedTermLines(comment, snapshot) : []), ...head];
  }
  return [...head, ...computedLines(comment, snapshot)];
}

function computedLines(comment, snapshot) {
  return stillPresent(comment, snapshot?.paragraphs).map((r) => {
    const n = r.places.length;
    const at = r.places.map((i) => snapshot.ids?.[i] ?? `p${i}`).join(", ");
    return `"${r.term}" (taken out by ${r.change_ids.join(", ")}) is still present in ${n} other place${n === 1 ? "" : "s"}: ${at}.`;
  });
}
