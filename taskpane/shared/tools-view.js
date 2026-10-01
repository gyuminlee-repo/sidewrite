// Wording of the Tools tab (tools-pane.js). Pure functions: no DOM, no Word.

// What a selection request was about, as the head of its answer.
const QUOTE_CHARS = 80;
export function selectionHead(kind, text, r) {
  const flat = String(text).replace(/\s+/g, " ").trim();
  const quote = flat.length > QUOTE_CHARS ? `${flat.slice(0, QUOTE_CHARS)}...` : flat;
  const what = kind === "translate" ? `Translation into ${r.lang}` : "Answer";
  return `${what} (${r.model}) for: "${quote}"`;
}

// How a rule candidate will be checked, in words.
const COUNT_WHAT = {
  section_words: "words",
  keywords: "keywords",
  figures: "figures",
  tables: "tables",
};
export function candidateKind(c) {
  const k = c.check ?? {};
  if (c.profile) return `Profile setting: ${JSON.stringify(c.profile)}`;
  if (k.type !== "count") return "Rule the model checks";
  if (k.target === "section_order") return `Count rule: section order ${k.order.join(", ")}`;
  const limits = [
    k.min != null ? `at least ${k.min}` : null,
    k.max != null ? `at most ${k.max}` : null,
  ].filter(Boolean);
  const where = k.target === "section_words" ? " in the section" : "";
  return `Count rule: ${limits.join(" and ")} ${COUNT_WHAT[k.target]}${where}`;
}
export const candidateCite = (c) => `Guideline: "${c.source.sentence}"`;

// What came of adding the picked rules.
export function rulesAdded(r) {
  const plural = (n, one) => `${n} ${one}${n === 1 ? "" : "s"}`;
  const parts = [`${plural(r.added, "rule")} added to the ledger`];
  if (r.skipped) parts.push(`${r.skipped} already there`);
  if (r.profile_keys?.length) parts.push(`profile setting saved: ${r.profile_keys.join(", ")}`);
  return `${parts.join("; ")}.`;
}

// A file larger than this does not fit one frame to the daemon (base64).
const FRAME_LIMIT = 100 * 1024 * 1024 - 4096;
export const fileTooLarge = (bytes) => Math.ceil(bytes / 3) * 4 > FRAME_LIMIT;

// The lines under "Make marked copy": [{ text, bad? }].
const MODE_WORDS = { track_changes: "tracked changes", color: "colours" };
export function markedLines(r) {
  if (!r.ok)
    return [
      { text: "No marked copy was written.", bad: true },
      { text: r.error, bad: true },
    ];
  const p = r.report?.paragraphs ?? {};
  return [
    { text: `Written: ${r.name} (${MODE_WORDS[r.mode] ?? r.mode}).` },
    {
      text: `Paragraphs: ${p.n_mod ?? 0} changed, ${p.n_ins ?? 0} added, ${p.n_del ?? 0} removed, ${p.n_eq ?? 0} unchanged.`,
    },
    ...(r.report?.warnings ?? []).map((w) => ({ text: String(w), bad: true })),
    { text: r.notice },
  ];
}

// Which file the copies are made from.
export const docLine = (doc) =>
  doc
    ? `Saved file: ${doc.doc_name}. Profile: ${doc.profile_name}. Copies are made from the saved file, so save before you make one.`
    : "Save the document to a local folder to use the tools.";
