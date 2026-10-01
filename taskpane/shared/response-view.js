// Pure helpers of the Response tab: no DOM, no Office.js, so they run under
// node --test.

// The daemon's WebSocket server keeps the `ws` default of 100 MiB per frame
// (daemon/bridge.mjs does not change it). A PDF goes to the daemon as base64
// inside one JSON frame.
export const PDF_FRAME_LIMIT = 100 * 1024 * 1024;
const ENVELOPE = 4096; // room for the rest of the JSON message

// base64 of a Uint8Array. String.fromCharCode takes a limited number of
// arguments, so the bytes go in chunks.
export function bytesToBase64(bytes) {
  let bin = "";
  const CHUNK = 8192;
  for (let i = 0; i < bytes.length; i += CHUNK)
    bin += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK));
  return btoa(bin);
}

// Would a PDF of this base64 length not fit one frame? Checked before sending:
// a frame that is too large is dropped without an answer.
export const pdfTooLarge = (base64Length) => base64Length + ENVELOPE > PDF_FRAME_LIMIT;

// Why a saved response is out of date (daemon/response-quotes.mjs staleReasons).
export const STALE_LABELS = {
  quote_changed: "Quote changed in the manuscript",
  change_state: "A cited change is no longer accepted",
  lines_old: "Line numbers are older than the manuscript",
};
export const staleText = (stale) => (stale ?? []).map((s) => STALE_LABELS[s] ?? s).join("; ");

// What a card shows as the citations of its comment, as the pane tells the
// daemon with a saved text (response_save `shown`). The daemon judges the
// text by these: they are what the author read while typing, and another
// pane may have replaced the stored ones since.
export const shownCitations = (comment) =>
  (comment?.response?.citations ?? []).map((c) => ({
    kind: c.kind,
    quote: c.quote,
    paragraph_index: c.paragraph_index,
    change_ids: c.change_ids,
  }));

// The line of a card whose saved answer has no second-language paragraph
// although the letter writes one (the `translation` row of
// daemon/response-check.mjs), or "". letter: profile.response_letter.
export function translationMissing(comment, letter) {
  const r = comment?.response;
  if (!letter?.gloss_enabled || !letter.gloss_lang) return "";
  if (!r?.text?.trim() || r.gloss?.trim()) return "";
  return `No ${letter.gloss_lang} translation of this answer; the letter prints the answer without one.`;
}

// The same rules as fmtLines and renderFrame of daemon/response-lines.mjs,
// written again here because the pane cannot import a daemon module.
// tests/response-view.test.mjs compares the two on the same inputs.
const fmtLines = (lines, L) =>
  !lines
    ? null
    : lines.start === lines.end
      ? L.line_one.replace("{a}", () => lines.start)
      : L.line_range.replace("{a}", () => lines.start).replace("{b}", () => lines.end);
function renderFrame(frame, { section = "", lines = null } = {}) {
  const f =
    lines == null
      ? String(frame).replace(/,?\s*\{lines\}/, "")
      : String(frame).replace("{lines}", () => lines);
  return f.replace("{section}", () => section ?? "").replace("{n}", () => "");
}

// A citation as the letter will print it: the location line and the quote
// under it. letter: response_letter of the profile.
export function citationText(cit, letter) {
  if (cit.problem) return { line: `Not located: ${cit.problem}`, quote: null };
  const removed = cit.kind === "removed";
  // No section known: the section-less frame (daemon locationLine).
  const frame = cit.section
    ? removed
      ? letter.removed_line
      : letter.location_line
    : removed
      ? (letter.removed_line_no_section ?? "[Removed]")
      : (letter.location_line_no_section ?? "[Revised, {lines}]");
  return {
    line: renderFrame(frame, { section: cit.section, lines: fmtLines(cit.lines, letter) }),
    quote: removed ? null : cit.quote,
  };
}

// One row of the checks: "FAIL R1.1 quote_exact: <detail>".
export const checkLine = (row) =>
  `${row.level === "fail" ? "FAIL" : "WARN"} ${row.ref ? `${row.ref} ` : ""}${row.id}: ${row.detail}`;

const count = (n, one, many = `${one}s`) => `${n} ${n === 1 ? one : many}`;

export function checksSummary({ fail, warn }) {
  const head = fail ? `${fail} failed` : "All checks passed";
  return warn ? `${head}, ${count(warn, "warning")}` : head;
}

// The export button: summary is that of the last check, null before one.
export function canExport(summary) {
  if (!summary) return { ok: false, label: "Run checks first" };
  return summary.fail
    ? { ok: false, label: "Export anyway" }
    : { ok: true, label: "Export letter (.docx)" };
}

const isDict = (x) => !!x && typeof x === "object" && !Array.isArray(x);
const shown = (v) => (v == null || v === "" ? "(none)" : typeof v === "string" ? v : String(v));

// What a loaded letter would change in the profile of this paper: one row
// per key (dot path) whose proposed value differs from the current one.
export function proposalRows(proposal, profile) {
  const rows = [];
  const add = (key, current, proposed) => {
    if ((current ?? null) === (proposed ?? null)) return;
    const seen = key.startsWith("colors.") ? proposal.evidence?.colors?.[proposed] : null;
    rows.push({
      key,
      current: shown(current),
      proposed: shown(proposed),
      evidence: seen ? count(seen.paragraphs, "paragraph") : "",
    });
  };
  for (const [k, v] of Object.entries(proposal?.overrides ?? {})) {
    if (isDict(v))
      for (const [k2, v2] of Object.entries(v)) add(`${k}.${k2}`, profile?.[k]?.[k2], v2);
    else add(k, profile?.[k], v);
  }
  return rows;
}

// The language form as profile_set overrides. Empty text is null: the key is
// removed and the profile (then the default) applies again.
export function overridesFromForm({ lang, rationaleLang, gloss, glossLang }) {
  const text = (s) => String(s ?? "").trim() || null;
  return {
    rationale_lang: text(rationaleLang),
    response_letter: {
      lang: text(lang),
      gloss_enabled: !!gloss,
      ...(glossLang === undefined ? {} : { gloss_lang: text(glossLang) }),
    },
  };
}

const CITED = new Set(["accepted", "user_edited"]);

// Requests of a comment that no change in the manuscript answers and that
// are not declined. pending: a change for it still waits in Word.
export function openRequirements(comment) {
  if (comment?.declined) return [];
  const changes = comment?.changes ?? [];
  const idsOf = (pred) =>
    new Set(changes.filter(pred).flatMap((ch) => ch.rationale?.requirement_ids ?? []));
  const answered = idsOf((ch) => CITED.has(ch.state));
  const pending = idsOf((ch) => ch.state === "pending");
  const declined = new Set((comment?.response?.declined_requirements ?? []).map((d) => d.id));
  return (comment?.requirements ?? [])
    .filter((q) => !answered.has(q.id) && !declined.has(q.id))
    .map((q) => ({ id: q.id, text: q.text, pending: pending.has(q.id) }));
}

// Changes of a comment that are neither accepted nor rejected yet: still
// tracked in Word (pending) or not found in the document (unlocated). The
// letter does not cite them.
export function undecidedChanges(comment) {
  const n = (state) => (comment?.changes ?? []).filter((ch) => ch.state === state).length;
  return { pending: n("pending"), unlocated: n("unlocated") };
}
export function undecidedText(comment) {
  const { pending, unlocated } = undecidedChanges(comment);
  const parts = [pending && `${pending} pending`, unlocated && `${unlocated} not found`].filter(
    Boolean,
  );
  // Sent to Word and not answered (outcome_unknown): it may be in the
  // manuscript with no record, so the letter cannot cite it.
  const unknown = (comment?.blocked ?? []).filter((b) => b.state === "outcome_unknown").length;
  return [
    parts.length
      ? `Decision left in Word: ${parts.join(", ")}. Only accepted changes are cited.`
      : "",
    unknown
      ? `${count(unknown, "edit")} got no answer from Word: settle ${unknown === 1 ? "it" : "them"} on the Review tab.`
      : "",
  ]
    .filter(Boolean)
    .join(" ");
}

// The one thing to do next. revision: revision.json as the daemon sent it;
// summary: { fail, warn } of the last check, null before one or after a change.
export function nextStep(revision, summary) {
  const comments = revision?.comments ?? [];
  if (!comments.length)
    return { id: "split", text: "Next: split the review into comments on the Review tab." };
  const missing = comments.filter((c) => !c.declined && !c.response?.text?.trim()).length;
  if (missing)
    return {
      id: "draft",
      text: `Next: write or draft the ${count(missing, "missing response")} (or decline the comment).`,
    };
  // Written for a sentence or a set of changes that is not there any more.
  const outdated = comments.filter(
    (c) =>
      !c.declined &&
      (c.response?.stale ?? []).some((x) => x === "quote_changed" || x === "change_state"),
  ).length;
  if (outdated)
    return {
      id: "reread",
      text: `Next: read the ${count(outdated, "response")} marked out of date, then save or draft ${outdated === 1 ? "it" : "them"} again.`,
    };
  const noLines = comments.some(
    (c) =>
      (c.response?.stale ?? []).includes("lines_old") ||
      (c.response?.citations ?? []).some((x) => x.kind === "revised" && !x.problem && !x.lines),
  );
  if (noLines)
    return {
      id: "locate",
      text: "Next: update line numbers (Word renders the manuscript to PDF).",
    };
  if (!summary) return { id: "check", text: "Next: run checks." };
  if (summary.fail)
    return {
      id: "fix",
      text: `Next: fix the ${count(summary.fail, "failed check")} below, then run checks again.`,
    };
  return { id: "export", text: "Next: export the letter." };
}

// What an export did. r: the response_export result. wanted: the usual
// name(s); a path that differs means a file of that name was there (or open
// in Word) and was left as it is.
export function exportMessage(r) {
  if (!r.exported)
    return `Not exported: ${checksSummary(r.summary)}. Fix them, or press Export anyway.`;
  if (r.path) {
    return r.wanted && r.wanted !== r.path
      ? `Exported under a new name: ${r.path} (${r.wanted} was already there or open in Word and was kept).`
      : `Exported: ${r.path}`;
  }
  const paths = r.paths ?? [];
  const kept = (r.wanted ?? []).filter((w, i) => w !== paths[i]);
  return `Exported: ${paths.join(", ")}${
    kept.length
      ? ` (${kept.join(", ")} ${kept.length === 1 ? "was" : "were"} already there and ${kept.length === 1 ? "was" : "were"} kept).`
      : ""
  }`;
}
