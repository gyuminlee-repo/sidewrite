// Check the instruction ledger against the document, and pre-check new text
// against the forbid list before a tracked edit is applied.
//
// Paragraph shape (task pane snapshot): { index, text, reviewed_text }.
// reviewed_text is Word getReviewedText(current), i.e. the paragraph as if
// every pending tracked change were accepted; text is the raw paragraph text,
// which still contains pending deletions. Checks read reviewed_text so a
// pending deletion already counts as deleted.

import { normalizeParaText } from "./revision-classify.mjs";
import { paragraphHash } from "./doc-diff.mjs";

const viewText = (p) => normalizeParaText(p.reviewed_text ?? p.text);
const rawText = (p) => normalizeParaText(p.text);

export function compileCheck(check) {
  const flags = String(check.flags ?? "").replace(/g/g, "");
  return new RegExp(check.pattern, `${flags}g`);
}

// An empty match advances by one code point in unicode mode (u or v flag):
// one UTF-16 unit would land inside a surrogate pair, where exec() matches
// the same place again. There are at most text.length + 1 matches, so the
// loop also stops there as a guard.
function findAll(re, text) {
  const out = [];
  const unicode = re.unicode || re.unicodeSets;
  re.lastIndex = 0;
  let m;
  while (out.length <= text.length && (m = re.exec(text)) !== null) {
    out.push({ start: m.index, match: m[0] });
    if (m[0] === "") {
      const cp = text.codePointAt(re.lastIndex);
      re.lastIndex += unicode && cp > 0xffff ? 2 : 1;
    }
  }
  return out;
}

export function checkRegexItem(item, paragraphs) {
  let re;
  try {
    re = compileCheck(item.check);
  } catch (e) {
    return { id: item.id, status: "error", error: e.message, violations: [] };
  }
  const expect = item.check.expect ?? "absent";
  if (expect === "absent") {
    const violations = [];
    for (const p of paragraphs) {
      for (const hit of findAll(re, viewText(p))) {
        violations.push({ paragraph_index: p.index, ...hit });
      }
    }
    return { id: item.id, status: violations.length ? "violations" : "ok", violations };
  }
  // expect "present": a "do not delete X" rule.
  if (paragraphs.some((p) => findAll(re, viewText(p)).length > 0)) {
    return { id: item.id, status: "ok", violations: [] };
  }
  // Point at pending deletions of X (still in the raw text) when there are any.
  const pending = paragraphs.filter((p) => findAll(re, rawText(p)).length > 0);
  const violations = pending.length
    ? pending.map((p) => ({ paragraph_index: p.index, start: null, match: null, missing: true }))
    : [{ paragraph_index: null, start: null, match: null, missing: true }];
  return { id: item.id, status: "violations", violations };
}

// Paragraph batches for LLM checks, each at most maxChars of text (a longer
// single paragraph becomes its own batch).
export function batchParagraphs(paragraphs, maxChars = 6000) {
  const batches = [];
  let cur = [];
  let size = 0;
  for (const p of paragraphs) {
    const n = viewText(p).length;
    if (cur.length && size + n > maxChars) {
      batches.push(cur);
      cur = [];
      size = 0;
    }
    cur.push(p);
    size += n;
  }
  if (cur.length) batches.push(cur);
  return batches;
}

export const LEDGER_SYSTEM =
  "You check a manuscript against one author instruction. Report only clear violations and quote the violating words exactly. If you are not sure whether a paragraph breaks the instruction, list it under unsure; do not guess that it has no violation. Reply with JSON only.";

export function ledgerPrompt(item, batch) {
  return [
    `Instruction [${item.id}]: ${item.text}`,
    'Return {"violations":[{"paragraph_index":<int>,"quote":"<exact words>","reason":"<short>"}],"unsure":[{"paragraph_index":<int>,"reason":"<short>"}]}.',
    'Put a paragraph under "unsure" when you cannot tell whether it breaks the instruction. Use empty lists when nothing applies.',
    "Paragraphs:",
    ...batch.map((p) => `[${p.index}] ${viewText(p)}`),
  ].join("\n");
}

// The words of `text` (a normalized paragraph) that the model quoted, or null.
// Exact first; then the same words in another case or with other spacing.
// Anything further off is not evidence of where the violation is.
export function locateQuote(text, quote) {
  const q = normalizeParaText(quote);
  if (!q) return null;
  if (text.includes(q)) return q;
  const words = q.split(" ").map((w) => w.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
  return new RegExp(words.join("\\s+"), "i").exec(text)?.[0] ?? null;
}

export const QUOTE_NOT_FOUND = "model reported a violation but its quote was not found";

const notInteger = (idx) =>
  typeof idx === "boolean" || idx === null || idx === "" || !Number.isInteger(Number(idx));

// Why a one-shot reply is not {"violations":[{paragraph_index, quote,
// reason}], "unsure"?:[{paragraph_index, reason}]}, or null when it is.
// "unsure" is optional (a reply without it has no unsure paragraph). A
// malformed reply must not read as "no violations": checkLlmItem turns it
// into a check error.
export function llmReplyProblem(json) {
  if (!Array.isArray(json?.violations)) return "reply has no violations array";
  for (const [i, v] of json.violations.entries()) {
    if (v === null || typeof v !== "object") return `violations[${i}] is not an object`;
    if (notInteger(v.paragraph_index)) return `violations[${i}].paragraph_index is not an integer`;
    if (typeof v.quote !== "string") return `violations[${i}].quote is not a string`;
    if (typeof v.reason !== "string") return `violations[${i}].reason is not a string`;
  }
  if (json.unsure === undefined) return null;
  if (!Array.isArray(json.unsure)) return "unsure is not an array";
  for (const [i, u] of json.unsure.entries()) {
    if (u === null || typeof u !== "object") return `unsure[${i}] is not an object`;
    if (notInteger(u.paragraph_index)) return `unsure[${i}].paragraph_index is not an integer`;
    if (typeof u.reason !== "string") return `unsure[${i}].reason is not a string`;
  }
  return null;
}

// onDecision (optional): called once per paragraph the model judged, with
// { item_id, paragraph_index, instruction, paragraph, verdict, error? };
// verdict is violation, ok, unsure, or error (failed or malformed reply).
// It carries the texts so the caller can fingerprint them (decision-log.mjs).
export async function checkLlmItem(
  item,
  paragraphs,
  { oneshot, maxChars = 6000, maxBatches = 20, onDecision = null } = {},
) {
  if (!oneshot)
    return { id: item.id, status: "error", error: "no one-shot runner", violations: [] };
  const batches = batchParagraphs(paragraphs, maxChars);
  if (batches.length > maxBatches) {
    return {
      id: item.id,
      status: "error",
      error: `document needs ${batches.length} LLM batches; the limit is ${maxBatches}. Use a regex check or check a section.`,
      violations: [],
    };
  }
  const violations = [];
  const unsure = [];
  const decide = (batch, verdictOf, error) => {
    for (const p of batch) {
      onDecision?.({
        item_id: item.id,
        paragraph_index: p.index,
        instruction: item.text,
        paragraph: viewText(p),
        verdict: verdictOf(p.index),
        ...(error ? { error } : {}),
      });
    }
  };
  for (const batch of batches) {
    const r = await oneshot({ system: LEDGER_SYSTEM, prompt: ledgerPrompt(item, batch) });
    if (!r.ok) {
      decide(batch, () => "error", r.error);
      return { id: item.id, status: "error", error: r.error, violations, unsure };
    }
    const problem = llmReplyProblem(r.json);
    if (problem) {
      const error = `malformed LLM reply: ${problem}`;
      decide(batch, () => "error", error);
      return { id: item.id, status: "error", error, violations, unsure };
    }
    const before = { violations: violations.length, unsure: unsure.length };
    // Paragraphs with a reported violation whose quote is not in them.
    const unlocated = new Map();
    // An entry that names a paragraph outside the batch and cannot be placed:
    // the reply is about text that was not asked, so no paragraph of this
    // batch counts as ok on the strength of it (Codex review 13).
    let stray = null;
    for (const v of r.json.violations) {
      let para = batch.find((p) => p.index === Number(v?.paragraph_index));
      if (!para) {
        // Its quote in exactly one paragraph of the batch says which one.
        const holders = batch.filter(
          (p) => locateQuote(viewText(p), String(v?.quote ?? "")) !== null,
        );
        if (holders.length !== 1) {
          stray ??= v.paragraph_index;
          continue;
        }
        para = holders[0];
      }
      // A quote that is not in the paragraph is no evidence and is dropped:
      // the model must point at text.
      const match = locateQuote(viewText(para), String(v?.quote ?? ""));
      if (match === null) {
        if (!unlocated.has(para.index)) unlocated.set(para.index, String(v?.reason ?? ""));
        continue;
      }
      violations.push({
        paragraph_index: para.index,
        start: null,
        match,
        reason: String(v?.reason ?? ""),
      });
    }
    for (const u of r.json.unsure ?? []) {
      const para = batch.find((p) => p.index === Number(u.paragraph_index));
      if (para) unsure.push({ paragraph_index: para.index, reason: u.reason });
      else stray ??= u.paragraph_index;
    }
    // A paragraph under both lists is unsure: no violation is kept for it,
    // and the incremental check reads it again (Codex review 9).
    const open = new Set(unsure.slice(before.unsure).map((u) => u.paragraph_index));
    for (let i = violations.length - 1; i >= before.violations; i--) {
      if (open.has(violations[i].paragraph_index)) violations.splice(i, 1);
    }
    const hit = new Set(violations.slice(before.violations).map((v) => v.paragraph_index));
    // The model saw a violation it could not point at: the paragraph is not
    // ok. Unsure keeps it with the author and out of checked_hashes, so the
    // next incremental check reads it again (Codex review 12).
    for (const [index, reason] of unlocated) {
      if (hit.has(index) || open.has(index)) continue;
      unsure.push({
        paragraph_index: index,
        reason: reason ? `${QUOTE_NOT_FOUND} (${reason})` : QUOTE_NOT_FOUND,
      });
      open.add(index);
    }
    if (stray !== null) {
      for (const p of batch) {
        if (hit.has(p.index) || open.has(p.index)) continue;
        unsure.push({
          paragraph_index: p.index,
          reason: `model reply named paragraph ${stray}, which was not in the request; this paragraph has no result of its own`,
        });
        open.add(p.index);
      }
    }
    decide(batch, (i) => (hit.has(i) ? "violation" : open.has(i) ? "unsure" : "ok"));
  }
  return { id: item.id, status: llmStatus(violations, unsure), violations, unsure };
}

// A violation outranks unsure in the status of the item (a paragraph under
// both lists is unsure only, checkLlmItem); unsure paragraphs still wait for
// the author.
const llmStatus = (violations, unsure) =>
  violations.length ? "violations" : unsure.length ? "unsure" : "ok";

export async function checkLedger(items, paragraphs, deps = {}) {
  const outcomes = [];
  for (const item of items) {
    outcomes.push(
      item.check.type === "regex"
        ? checkRegexItem(item, paragraphs)
        : await checkLlmItem(item, paragraphs, deps),
    );
  }
  return outcomes;
}

// Ledger items that apply to manuscript edits. An item without a scope is a
// manuscript item (the schema default). The review prompt (review-prompt.mjs)
// and the pre-check below use this one rule, so what the model is told and
// what is refused cannot drift apart.
export function inManuscriptScope(item) {
  return (item?.scope ?? "manuscript") === "manuscript";
}

// The rules a tracked edit is pre-checked against: manuscript-scope
// forbid-list regex items (expect "absent"). LLM-checked forbid items cannot
// be pre-checked. A must-stay rule (expect "present") is not pre-checked
// either: an edit that deletes its phrase is applied, and the ledger check
// reports it afterwards. A rule whose regex does not compile is not among
// the rules: it is listed as broken, and applyTrackedEdit (review-tools.mjs)
// refuses the edit instead of reading the missing rule as "no hit".
function precheckRules(ledger) {
  const rules = [];
  const broken = [];
  for (const item of ledger ?? []) {
    if (item.kind !== "forbid" || item.check?.type !== "regex") continue;
    if (!inManuscriptScope(item)) continue;
    if ((item.check.expect ?? "absent") !== "absent") continue;
    try {
      rules.push({ id: item.id, re: compileCheck(item.check) });
    } catch (e) {
      broken.push({ id: item.id, error: e.message });
    }
  }
  return { rules, broken };
}

const forbidRules = (ledger) => precheckRules(ledger).rules;

// [{ id, error }] of the pre-check rules that cannot be compiled. The input
// schema refuses such a rule (revision-store.mjs LedgerItemSchema); a file
// from before that check can still hold one.
export const brokenForbidRules = (ledger) => precheckRules(ledger).broken;

// Ids of the pre-check rules that match `text` (the new text on its own).
export function forbidHits(ledger, text) {
  return forbidRules(ledger)
    .filter(({ re }) => findAll(re, String(text ?? "")).length)
    .map(({ id }) => id);
}

const collapse = (s) =>
  String(s ?? "")
    .replace(/\u0005/g, "")
    .replace(/\s+/g, " ");

// Matches of `re` in `text` that overlap or touch the range [a, b].
// Touching counts: a deletion can end a word right at its edge ("novelty"
// without "ty").
const matchesAt = (re, text, a, b) =>
  findAll(re, text).filter((m) => m.match && m.start <= b && m.start + m.match.length >= a).length;

// Ids of the pre-check rules that the edit would newly break in the text
// around it: the forbidden phrase is formed by the new text together with
// the words next to it (forbid "novel substrate", "a novel material",
// material -> substrate), or by a deletion that joins two words.
// texts: the paragraphs as they read now; edit: { find, text, kind } with
// kind replace | delete | insert (text is then inserted after find).
// A rule counts only when it has more matches at the place of the edit after
// it than before: a violation that was already there, next to the edit or
// elsewhere in the paragraph, does not block, and removing one passes.
// Every occurrence of `find` is tried (a paragraph_hint may pick any of
// them). A `find` that is in none of the texts gives no hit.
export function forbidHitsAfterEdit(ledger, texts, { find, text, kind }) {
  const rules = forbidRules(ledger);
  const f = collapse(find);
  const n = collapse(text);
  if (!rules.length || !f.trim()) return [];
  const hits = new Set();
  for (const raw of texts ?? []) {
    const t = normalizeParaText(raw);
    for (let i = t.indexOf(f); i !== -1; i = t.indexOf(f, i + 1)) {
      const end = i + f.length;
      const start = kind === "insert" ? end : i;
      const after = t.slice(0, start) + n + t.slice(end);
      for (const { id, re } of rules) {
        if (matchesAt(re, after, start, start + n.length) > matchesAt(re, t, start, end))
          hits.add(id);
      }
    }
  }
  return rules.map(({ id }) => id).filter((id) => hits.has(id));
}

// ---- L10: incremental check ------------------------------------------------
//
// Default check: only paragraphs whose text changed since this item was last
// checked (hash not in item.checked_hashes). Violations of the last check
// stay for paragraphs that are still there unchanged and skipped (hash in
// checked_hashes); a paragraph read again keeps only its new result. A
// paragraph that occurs twice keeps its violation in both copies. A full
// run happens when asked (full), for an item never checked or last in
// error, and always for an expect:"present" rule (it is about the whole
// document). An item outside the manuscript scope (inManuscriptScope) is
// not run against the manuscript: it comes back status "skipped", skipped
// "scope", with no violations and checked_hashes null, so a later scope
// change to manuscript runs in full.
// Returns [{ id, status, violations, unsure, error?, mode, checked, total,
// checked_hashes }]; violations and unsure entries carry `h`, the paragraph
// hash. An unsure paragraph's hash is left out of checked_hashes, so the
// next incremental check reads it again; unsure entries are not carried
// over (the recheck makes them fresh).

function needsFull(item, full) {
  return (
    full ||
    !Array.isArray(item.checked_hashes) ||
    item.status === "error" ||
    item.status === "unchecked" ||
    (item.check.type === "regex" && (item.check.expect ?? "absent") === "present")
  );
}

export async function checkLedgerIncremental(items, paragraphs, { full = false, ...deps } = {}) {
  const hashes = paragraphs.map((p) => paragraphHash(viewText(p)));
  const indexOfHash = new Map();
  hashes.forEach((h, i) => {
    if (!indexOfHash.has(h)) indexOfHash.set(h, []);
    indexOfHash.get(h).push(paragraphs[i].index);
  });
  const hashOfIndex = new Map(paragraphs.map((p, i) => [p.index, hashes[i]]));
  const allHashes = [...new Set(hashes)];
  const outcomes = [];
  for (const item of items) {
    if (!inManuscriptScope(item)) {
      outcomes.push({
        id: item.id,
        status: "skipped",
        skipped: "scope",
        violations: [],
        unsure: [],
        mode: "skipped",
        checked: 0,
        total: paragraphs.length,
        checked_hashes: null,
      });
      continue;
    }
    const whole = needsFull(item, full);
    const seen = new Set(whole ? [] : item.checked_hashes);
    const subset = whole ? paragraphs : paragraphs.filter((p, i) => !seen.has(hashes[i]));
    const r = (await checkLedger([item], subset, deps))[0];
    const fresh = r.violations.map((v) => ({
      ...v,
      h: v.paragraph_index == null ? null : (hashOfIndex.get(v.paragraph_index) ?? null),
    }));
    let kept = [];
    if (!whole) {
      // Only paragraphs the cache skipped keep their old violations; one
      // read again has the new result alone (Codex review 9).
      const once = new Map();
      for (const v of item.violations ?? []) {
        if (v.h && seen.has(v.h) && indexOfHash.has(v.h))
          once.set(`${v.h}|${v.start}|${v.match}`, v);
      }
      for (const v of once.values()) {
        for (const idx of indexOfHash.get(v.h)) kept.push({ ...v, paragraph_index: idx });
      }
    }
    const violations = [...kept, ...fresh].sort(
      (a, b) => (a.paragraph_index ?? -1) - (b.paragraph_index ?? -1),
    );
    const unsure = (r.unsure ?? []).map((u) => ({
      ...u,
      h: hashOfIndex.get(u.paragraph_index) ?? null,
    }));
    const open = new Set(unsure.map((u) => u.h));
    const status = r.status === "error" ? "error" : llmStatus(violations, unsure);
    outcomes.push({
      id: item.id,
      status,
      ...(r.error ? { error: r.error } : {}),
      violations,
      unsure,
      mode: whole ? "full" : "incremental",
      checked: subset.length,
      total: paragraphs.length,
      checked_hashes: status === "error" ? null : allHashes.filter((h) => !open.has(h)),
    });
  }
  return outcomes;
}
