// How the pane names a paragraph to the author: the section it is in and its
// first words, never the internal id. A name like that is shown as a link that
// selects the paragraph in Word (review-pane.js).
//
// snapshot: revision.snapshot ({ paragraphs }). The section is the nearest
// heading above the paragraph (named-terms.js paragraphKinds: a short line
// without a sentence end).

import { paragraphKinds } from "./named-terms.js";

const WORDS = 8;
const SECTION_CHARS = 32;

const clean = (t) =>
  String(t ?? "")
    .replace(/\s+/g, " ")
    .trim();

// The first words of a text, for a paragraph that is no longer in the snapshot.
export function firstWords(text) {
  const all = clean(text).split(" ").filter(Boolean);
  return `${all.slice(0, WORDS).join(" ")}${all.length > WORDS ? "..." : ""}`;
}

// The index of the nearest heading above the paragraph at `index`, or -1.
// kinds: paragraphKinds of the paragraphs.
export function headingAbove(kinds, index) {
  for (let i = index - 1; i >= 0; i--) if (kinds[i] === "heading") return i;
  return -1;
}

// { section, words, label } for the paragraph at `index`. A paragraph the
// snapshot does not hold is "a paragraph" (the text may have changed since).
export function paragraphLabel(snapshot, index) {
  const paragraphs = snapshot?.paragraphs ?? [];
  const text = clean(paragraphs[index]);
  if (!text) return { section: "", words: "", label: "a paragraph" };
  const kinds = paragraphKinds(paragraphs);
  const above = kinds[index] === "heading" ? -1 : headingAbove(kinds, index);
  let section = above === -1 ? "" : clean(paragraphs[above]);
  if (section.length > SECTION_CHARS)
    section = `${section.slice(0, SECTION_CHARS - 3).trimEnd()}...`;
  const words = firstWords(text);
  return { section, words, label: `${section ? `${section}: ` : ""}“${words}”` };
}
