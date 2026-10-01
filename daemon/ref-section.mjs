// Where the reference list of a manuscript is. Shared by the format rules
// (which skip it) and the Crossref check (which reads it). Pure.
//
// A paragraph is { index, text } (review snapshot, reviewed_text preferred)
// or { index, runs: [{ text }] } (format snapshot).
// Profile keys: ref_section_head (a pattern for the heading line) and
// ref_section_end_heads (patterns for the heading that ends the list).

export const REF_HEAD =
  "^(?:\\d[\\d.]*\\s+)?(?:references?|bibliography|literature cited|works cited)$";

export const paraText = (p) =>
  String(p.reviewed_text ?? p.text ?? (p.runs ?? []).map((r) => String(r.text ?? "")).join(""))
    .replace(/\s+/g, " ")
    .trim();

export const wordsIn = (t) => (t ? t.split(" ").length : 0);

// { heads: [index], entries: [paragraph] }. A heading is a short line (six
// words at most) that matches the pattern, so a sentence that only mentions
// "references" does not open the list. Empty paragraphs are not entries.
export function referenceSection(paragraphs, profile = {}) {
  const head = new RegExp(profile.ref_section_head || REF_HEAD, "i");
  const ends = (Array.isArray(profile.ref_section_end_heads) ? profile.ref_section_end_heads : [])
    .filter((s) => typeof s === "string" && s)
    .map((s) => new RegExp(s, "i"));
  const isHead = (t, re) => wordsIn(t) <= 6 && re.test(t);
  const heads = [];
  const entries = [];
  let inside = false;
  for (const p of paragraphs) {
    const t = paraText(p);
    if (inside && ends.some((re) => isHead(t, re))) inside = false;
    else if (inside) {
      if (t) entries.push(p);
    } else if (isHead(t, head)) {
      inside = true;
      heads.push(p.index);
    }
  }
  return { heads, entries };
}
