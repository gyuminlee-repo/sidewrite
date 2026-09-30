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

function findAll(re, text) {
  const out = [];
  re.lastIndex = 0;
  let m;
  while ((m = re.exec(text)) !== null) {
    out.push({ start: m.index, match: m[0] });
    if (m[0] === "") re.lastIndex += 1;
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
  "You check a manuscript against one author instruction. Report only clear violations and quote the violating words exactly. Reply with JSON only.";

export function ledgerPrompt(item, batch) {
  return [
    `Instruction [${item.id}]: ${item.text}`,
    'Return {"violations":[{"paragraph_index":<int>,"quote":"<exact words>","reason":"<short>"}]} or {"violations":[]}.',
    "Paragraphs:",
    ...batch.map((p) => `[${p.index}] ${viewText(p)}`),
  ].join("\n");
}

export async function checkLlmItem(
  item,
  paragraphs,
  { oneshot, maxChars = 6000, maxBatches = 20 } = {},
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
  for (const batch of batches) {
    const r = await oneshot({ system: LEDGER_SYSTEM, prompt: ledgerPrompt(item, batch) });
    if (!r.ok) return { id: item.id, status: "error", error: r.error, violations };
    for (const v of Array.isArray(r.json?.violations) ? r.json.violations : []) {
      const para = batch.find((p) => p.index === Number(v?.paragraph_index));
      if (!para) continue;
      const quote = String(v?.quote ?? "");
      // Drop quotes that are not in the paragraph: the model must point at text.
      if (!quote || !viewText(para).includes(normalizeParaText(quote))) continue;
      violations.push({
        paragraph_index: para.index,
        start: null,
        match: quote,
        reason: String(v?.reason ?? ""),
      });
    }
  }
  return { id: item.id, status: violations.length ? "violations" : "ok", violations };
}

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

// Ids of forbid-list regex items (expect "absent") that match `text`.
// LLM-checked forbid items cannot be pre-checked and are skipped here.
export function forbidHits(ledger, text) {
  const hits = [];
  for (const item of ledger ?? []) {
    if (item.kind !== "forbid" || item.check?.type !== "regex") continue;
    if ((item.check.expect ?? "absent") !== "absent") continue;
    let re;
    try {
      re = compileCheck(item.check);
    } catch {
      continue;
    }
    if (findAll(re, String(text ?? "")).length) hits.push(item.id);
  }
  return hits;
}

// ---- L10: incremental check ------------------------------------------------
//
// Default check: only paragraphs whose text changed since this item was last
// checked (hash not in item.checked_hashes). Violations of the last check
// stay for paragraphs that are still there unchanged; a paragraph that
// occurs twice keeps its violation in both copies. A full run happens when
// asked (full), for an item never checked or last in error, and always for
// an expect:"present" rule (it is about the whole document).
// Returns [{ id, status, violations, error?, mode, checked, total,
// checked_hashes }]; violations carry `h`, the paragraph hash.

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
      const once = new Map();
      for (const v of item.violations ?? []) {
        if (v.h && indexOfHash.has(v.h)) once.set(`${v.h}|${v.start}|${v.match}`, v);
      }
      for (const v of once.values()) {
        for (const idx of indexOfHash.get(v.h)) kept.push({ ...v, paragraph_index: idx });
      }
    }
    const violations = [...kept, ...fresh].sort(
      (a, b) => (a.paragraph_index ?? -1) - (b.paragraph_index ?? -1),
    );
    const status = r.status === "error" ? "error" : violations.length ? "violations" : "ok";
    outcomes.push({
      id: item.id,
      status,
      ...(r.error ? { error: r.error } : {}),
      violations,
      mode: whole ? "full" : "incremental",
      checked: subset.length,
      total: paragraphs.length,
      checked_hashes: status === "error" ? null : allHashes,
    });
  }
  return outcomes;
}
