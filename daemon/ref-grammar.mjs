// Figure, table and supplementary references in a sentence, from the profile
// grammar: ref_head + ref_item + (ref_sep + (ref_item | ref_item_cont))*.
// References joined by a separator alone are one group ("Fig. 1, Table S2").
// A caption title at the start of the paragraph is no reference.
//
// The profile regexes are written for Python `re`. They are compiled here
// without the u flag, so \w is ASCII only; tests/ref-grammar.test.mjs runs
// the whole profile corpus to show where the two dialects would differ.

// null when the profile has no grammar or one that does not compile
// (journal-profile.mjs regexProblems names the broken key).
export function buildRefGrammar(profile) {
  const heads = profile?.ref_head;
  if (!Array.isArray(heads) || !heads.length || typeof profile.ref_item !== "string") return null;
  const item = profile.ref_item;
  const sep = profile.ref_sep ?? String.raw`(?:\s*,\s*|\s+and\s+)`;
  const cont = profile.ref_item_cont ? `(?:${item}|${profile.ref_item_cont})` : `(?:${item})`;
  try {
    return {
      token: new RegExp(`(?<![\\w])(?:${heads.join("|")})${item}(?:${sep}${cont})*`, "g"),
      sepOnly: new RegExp(`^(?:${sep})$`),
      caption: profile.ref_caption ? new RegExp(profile.ref_caption) : null,
    };
  } catch {
    return null;
  }
}

// [{ start, end, text }] in text order.
export function refGroups(text, grammar) {
  if (!grammar) return [];
  const s = String(text ?? "");
  const cap = grammar.caption ? grammar.caption.exec(s) : null;
  const capEnd = cap && cap.index === 0 ? cap[0].length : -1;
  const out = [];
  for (const m of s.matchAll(grammar.token)) {
    if (m.index < capEnd) continue;
    const last = out[out.length - 1];
    if (last && grammar.sepOnly.test(s.slice(last.end, m.index))) last.end = m.index + m[0].length;
    else out.push({ start: m.index, end: m.index + m[0].length });
  }
  return out.map((g) => ({ ...g, text: s.slice(g.start, g.end) }));
}
