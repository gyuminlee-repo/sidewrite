// Automatic checks of the response letter before it is written. Pure: no I/O.
//
// checkResponses({ revision, snapshot, headings, profile }) returns rows
//   { id, level: "fail" | "warn", ref, detail }
// A letter may be exported when no row is a fail.
//
//   answered        every comment has a response or is declined
//   requirements    every request of a comment is answered by an accepted
//                   change or declined
//   declined_reason every declined comment and declined request has a reason
//                   that is not blank
//   stale           the cited changes are still the accepted changes
//   quote_exact     every quote is in its manuscript paragraph, character for
//                   character, in the accepted view
//   section         every citation names a manuscript heading
//   lines           every quote has one line range, read from the PDF of the
//                   manuscript text as it is now
//   ledger          response-scope ledger rules (regex) hold in every answer
//   prose_position  the answer text holds no location line or line number of
//                   its own; the tool adds them
//   rule_*          profile rules: word cap, opening, no other reviewer
//   warn only       lost_facts, pending_changes, rule_quote_words
//
// The profile must have its defaults filled (journal-profile.mjs).
// LLM-checked ledger items are not run here; the service runs them.

import { checkRegexItem, inScope } from "./ledger-check.mjs";
import { lostFacts } from "./edit-facts.mjs";
import { CITABLE, viewText, snapshotHash, staleReasons } from "./response-quotes.mjs";
import { normalizeParaText } from "./revision-classify.mjs";

// Counted by code, never taken from the model.
export const wordCount = (s) => (String(s ?? "").match(/\S+/g) ?? []).length;

const blank = (s) => !String(s ?? "").trim();

// Reviewer numbers a text names: "Reviewer 2", "Referee #3", a list after the
// plural ("Reviewers 1, 3 and 4") and a comment id of this tool ("R2.4").
// A bare "R2" is not taken: in a manuscript letter it is mostly R squared or
// a strain name.
const XREF =
  /\b(?:(Reviewers|Referees)|Reviewer|Referee)\s*#?\s*(\d+)((?:\s*(?:,\s*(?:and|or)|,|&|and|or)\s*#?\s*\d+)*)|\bR(\d+)\.\d+\b/gi;
function reviewersNamed(text) {
  const out = [];
  for (const m of text.matchAll(XREF)) {
    if (m[4]) out.push(Number(m[4]));
    else out.push(Number(m[2]), ...(m[1] ? (m[3].match(/\d+/g) ?? []).map(Number) : []));
  }
  return out;
}
// A whole line in square brackets: the shape of a location line.
const BRACKET_LINE = /^\s*\[.+\]\s*$/;
const escapeRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

// "line {a}" -> a regex for "line 12" as a word of its own.
function lineWordRe(frame) {
  if (typeof frame !== "string" || !frame.includes("{a}")) return null;
  const body = escapeRe(frame).replace(/\\\{[ab]\\\}/g, String.raw`\d+`);
  return new RegExp(String.raw`(?<![\p{L}\p{N}])${body}(?![\p{L}\p{N}])`, "iu");
}

export function checkResponses({ revision, snapshot, headings = [], profile }) {
  const rows = [];
  const add = (id, level, ref, detail) => rows.push({ id, level, ref, detail });
  const L = profile.response_letter;
  const texts = new Map((snapshot.paragraphs ?? []).map((p) => [p.index, viewText(p)]));
  const headSet = new Set(headings.map((h) => normalizeParaText(h.text)));
  const render = revision.response_render ?? null;
  const fresh = !!render && render.snapshot_hash === snapshotHash(snapshot);
  const items = (revision.ledger ?? []).filter(
    (i) => inScope(i, "response") && i.check?.type === "regex",
  );
  const positionRes = [lineWordRe(L.line_one), lineWordRe(L.line_range)].filter(Boolean);
  let opening = null;
  let openingError = null;
  if (L.opening_regex) {
    try {
      opening = new RegExp(L.opening_regex);
    } catch (e) {
      openingError = e.message;
    }
  }

  for (const c of revision.comments ?? []) {
    const r = c.response;
    const text = String(r?.text ?? "").trim();
    if (!text && !c.declined) add("answered", "fail", c.ref, "no response and not declined");
    // Declining counts as an answer only with a reason the letter can print.
    if (c.declined && blank(c.declined.reason))
      add("declined_reason", "fail", c.ref, "the comment is declined without a reason");

    const declined = new Map((r?.declined_requirements ?? []).map((d) => [d.id, d.reason]));
    const accepted = (c.changes ?? []).filter((ch) => CITABLE.has(ch.state));
    const answered = new Set(accepted.flatMap((ch) => ch.rationale?.requirement_ids ?? []));
    for (const q of c.requirements ?? []) {
      if (declined.has(q.id) && blank(declined.get(q.id)))
        add("declined_reason", "fail", c.ref, `${q.id} is declined without a reason: ${q.text}`);
      if (c.declined || declined.has(q.id) || answered.has(q.id)) continue;
      const pending = (c.changes ?? []).some(
        (ch) => ch.state === "pending" && (ch.rationale?.requirement_ids ?? []).includes(q.id),
      );
      const why = pending ? "has only pending changes in Word" : "is neither cited nor declined";
      add("requirements", "fail", c.ref, `${q.id} ${why}: ${q.text}`);
    }

    // quote_changed is the quote_exact axis and lines_old the lines axis below.
    if (staleReasons(c, snapshot, render).includes("change_state"))
      add("stale", "fail", c.ref, "the cited changes are not the accepted changes; refresh");

    for (const cit of r?.citations ?? []) {
      if (cit.problem) {
        add("quote_exact", "fail", c.ref, cit.problem);
        continue;
      }
      if (cit.section == null || !headSet.has(cit.section))
        add("section", "fail", c.ref, `section "${cit.section}" is not a manuscript heading`);
      if (cit.kind !== "revised") continue;
      const quote = String(cit.quote ?? "");
      if (!quote || !(texts.get(cit.paragraph_index) ?? "").includes(quote))
        add(
          "quote_exact",
          "fail",
          c.ref,
          `quote is not in paragraph ${cit.paragraph_index}: ${quote.slice(0, 60)}`,
        );
      const ln = cit.lines;
      if (!ln) add("lines", "fail", c.ref, "no line range; update line numbers");
      else if (ln.match_count !== 1)
        add("lines", "fail", c.ref, `quote found ${ln.match_count} times in the PDF`);
      else if (ln.start > ln.end && ln.start_page === ln.end_page)
        add("lines", "fail", c.ref, `range ${ln.start}-${ln.end}`);
      else if (!fresh || cit.lines_pdf !== render.pdf_sha)
        add("lines", "fail", c.ref, "manuscript changed since the PDF; update line numbers");
      if (L.max_quote_words && wordCount(quote) > L.max_quote_words)
        add("rule_quote_words", "warn", c.ref, `${wordCount(quote)} words quoted`);
    }

    if (!text) continue;
    const paras = text.split(/\n+/).map((t, i) => ({ index: i, text: t }));
    for (const item of items) {
      const out = checkRegexItem(item, paras);
      if (out.status === "violations")
        add(
          "ledger",
          "fail",
          c.ref,
          `${item.id}: ${out.violations.map((v) => v.match ?? "missing").join(", ")}`,
        );
      if (out.status === "error") add("ledger", "warn", c.ref, `${item.id}: ${out.error}`);
    }
    const position =
      paras.find((p) => BRACKET_LINE.test(p.text))?.text.trim() ??
      positionRes.map((re) => re.exec(text)?.[0]).find(Boolean);
    if (position)
      add(
        "prose_position",
        "fail",
        c.ref,
        `the response text holds "${position.slice(0, 60)}"; locations and line numbers are added by the tool`,
      );
    if (L.max_words && wordCount(text) > L.max_words)
      add("rule_words", "fail", c.ref, `${wordCount(text)} words, limit ${L.max_words}`);
    if (openingError) add("rule_opening", "warn", c.ref, `opening_regex: ${openingError}`);
    else if (opening && !opening.test(text))
      add("rule_opening", "fail", c.ref, "first sentence does not match the profile opening");
    // The reviewer being answered may be named, any other may not. An answer
    // to the editor (reviewer 0) has no reviewer of its own: every number is
    // a cross reference there.
    const others = L.forbid_cross_reviewer
      ? [...new Set(reviewersNamed(text).filter((n) => n !== c.reviewer))]
      : [];
    if (others.length) add("rule_xref", "fail", c.ref, `mentions Reviewer ${others.join(", ")}`);
    const lost = r.previous ? lostFacts(r.previous, text) : [];
    if (lost.length) add("lost_facts", "warn", c.ref, `not in the new draft: ${lost.join("; ")}`);
  }

  const pending = (snapshot.paragraphs ?? []).filter(
    (p) => (p.w_ins ?? 0) + (p.w_del ?? 0) > 0,
  ).length;
  if (pending)
    add(
      "pending_changes",
      "warn",
      null,
      `${pending} paragraph(s) still hold tracked changes; line numbers read them as accepted`,
    );
  return rows;
}

export const summarize = (rows) => ({
  fail: rows.filter((r) => r.level === "fail").length,
  warn: rows.filter((r) => r.level === "warn").length,
});
