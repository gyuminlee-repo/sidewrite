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

// The message after "Extract rules": { text, bad }. No candidate has three
// different reasons, and the tab says which one it is.
export function extractMessage(r) {
  const n = r.candidates.length;
  if (n)
    return {
      text: `${n} rule candidates. Untick the ones you do not want, then press Add selected.`,
      bad: false,
    };
  if (r.reason === "reply_unreadable")
    return {
      text: `The model's reply could not be read (${r.problem}), also when asked again. No rule was taken; run it again or add the rules by hand.`,
      bad: true,
    };
  if (r.reason === "request_failed")
    return { text: `The model could not be asked: ${r.warnings[0]}`, bad: true };
  return { text: "No rule was found in the guideline text.", bad: false };
}

// What came of adding the picked rules.
export function rulesAdded(r) {
  const plural = (n, one) => `${n} ${one}${n === 1 ? "" : "s"}`;
  const parts = [`${plural(r.added, "rule")} added to the ledger`];
  if (r.skipped) parts.push(`${r.skipped} already there`);
  if (r.profile_keys?.length) parts.push(`profile setting saved: ${r.profile_keys.join(", ")}`);
  return `${parts.join("; ")}.`;
}

// One format finding (daemon/format-check.mjs), in words.
const FORMAT_WANT = {
  stat: (f) => `"${f.find}" should be italic`,
  species: (f) => `"${f.find}" should be italic`,
  gene: (f) => `"${f.find}" should be italic`,
  subscript: (f) => `the digits of "${f.find}" should be subscript`,
  indent: (f) => {
    const found = typeof f.found === "object" ? `${f.found?.chars} characters` : `${f.found} pt`;
    return `first-line indent is ${found}, the others have ${f.want.first_line_indent} pt`;
  },
  empty_table: (f) => `empty table (${f.found.rows} x ${f.found.cols})`,
};
export function formatFindingText(f) {
  const what = FORMAT_WANT[f.rule]?.(f) ?? f.rule;
  const near = f.context && f.rule !== "indent" ? ` in "${f.context.trim()}"` : "";
  return `Paragraph ${f.paragraph_index}: ${what}${near}`;
}

// The rows under "Run checks": [{ text, index?, bad? }]. A row with an index
// goes to that paragraph when it is clicked.
export function checkRows(r) {
  const rows = [];
  const f = r.format;
  rows.push({
    text: f.rules.length
      ? `Format rules on: ${f.rules.join(", ")}, empty tables.`
      : "No format rule is on in the profile; only empty tables are reported.",
  });
  if (f.error) rows.push({ text: f.error, bad: true });
  for (const x of f.findings)
    rows.push({ text: formatFindingText(x), index: x.paragraph_index, bad: true });
  if (!f.error && !f.findings.length) rows.push({ text: "Format: nothing found." });
  if (!r.counts.length) rows.push({ text: "No count rule is in the ledger." });
  for (const c of r.counts) {
    if (c.status === "error") {
      rows.push({ text: `${c.id} could not be checked: ${c.error}. Rule: ${c.text}`, bad: true });
    } else if (c.status === "ok") {
      const n = c.count == null ? "" : ` (${c.count})`;
      rows.push({ text: `${c.id} ok${n}: ${c.text}` });
    } else {
      for (const v of c.violations) {
        const limits = [
          v.min != null ? `at least ${v.min}` : null,
          v.max != null ? `at most ${v.max}` : null,
        ].filter(Boolean);
        const found =
          v.count == null ? "out of order" : `${v.count} found, ${limits.join(" and ")} allowed`;
        rows.push({
          text: `${c.id}: ${found}. Rule: ${c.text}`,
          index: v.paragraph_index,
          bad: true,
        });
      }
    }
  }
  return rows;
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

// The lines under "Make anonymized copy": [{ text, bad? }]. The verifier's
// report (scripts/verify-anonymized.py) is shown whether the copy passed or
// not.
const yes = (b) => (b ? "yes" : "NO");
export function anonymizeLines(r) {
  const out = [];
  if (r.ok) out.push({ text: `Written: ${r.name}.` });
  else {
    out.push({
      text: r.name
        ? `The copy did not pass its check and was set aside as ${r.name}. Do not submit it.`
        : "No anonymized copy was written.",
      bad: true,
    });
    out.push({ text: r.error, bad: true });
  }
  const rep = r.report;
  if (rep) {
    out.push({
      text: `Verifier: total ${rep.total} left; text unchanged: ${yes(rep.body_equal)}; tracked changes unchanged: ${yes(rep.markup_equal)}; structure intact: ${yes(rep.structure_ok)}.`,
      bad: !r.ok,
    });
    for (const [id, c] of Object.entries(rep.categories ?? {})) {
      if (c.status === "cleared") continue;
      out.push({
        text: `${c.status === "not cleared" ? "Still present" : "Kept"}: ${id} (${c.count}): ${c.what}`,
        bad: c.status === "not cleared",
      });
    }
    for (const [name, n] of Object.entries(rep.name_mentions ?? {}))
      out.push({ text: `The text still has the name "${name}" (${n} times).`, bad: true });
  }
  if (r.notice) out.push({ text: r.notice });
  return out;
}

// Which file the copies are made from.
export const docLine = (doc) =>
  doc
    ? `Saved file: ${doc.doc_name}. Profile: ${doc.profile_name}. Copies are made from the saved file, so save before you make one.`
    : "Save the document to a local folder to use the tools.";
