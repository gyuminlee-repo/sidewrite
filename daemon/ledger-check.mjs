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

// ---- LLM check: one verdict for every paragraph ---------------------------------
//
// The model is asked for a verdict on EACH paragraph of a batch, under the
// batch id of the request:
//   {"batch_id":"b27","verdicts":[[0,"ok"],[1,"unsure"],[2,"violation"]],
//    "violations":[{"paragraph_id":2,"quote":"...","reason":"..."}],
//    "unsure":[{"paragraph_id":1,"reason":"..."}]}
// Paragraphs go out under short numbers of the batch (0, 1, 2 ...), not their
// document index, and the daemon makes the batch id: the model computes
// nothing. A reply with only a list of violations cannot be told from one
// that skipped paragraphs ("no violations" for text it did not read), and
// the skipped paragraphs were cached as checked. So the set of paragraph
// numbers with a verdict must be exactly the set asked about; anything else
// (one left out, one twice, one outside the batch, another batch id, a
// verdict and its evidence list that disagree) fails the batch: the item is
// an error and nothing of it is cached. A reply in the wrong form is asked
// for once more, with what was wrong, from the same runner; a model call
// that failed (limit, refusal, timeout) is not.
// Limit: this catches a reply that leaves paragraphs out. It does not prove
// that the model read a paragraph it called "ok".

export const LEDGER_SYSTEM =
  'You check paragraphs of a manuscript against one author instruction. Give every paragraph exactly one verdict: "ok" (the paragraph does not break the instruction), "violation" (it clearly breaks it; quote the violating words exactly) or "unsure" (you cannot tell). Do not leave a paragraph out and do not guess "ok" when you are not sure. Reply with JSON only.';

const VERDICTS = ["ok", "violation", "unsure"];

let lastBatch = 0;
const nextBatchId = () => `b${++lastBatch}`;

// batch: the paragraphs; they are numbered 0 .. batch.length - 1 in order.
export function ledgerPrompt(item, batch, batchId) {
  const last = batch.length - 1;
  return [
    `Instruction [${item.id}]: ${item.text}`,
    `Batch ${batchId}: ${batch.length} paragraph(s), numbered 0 to ${last}.`,
    "Return JSON of this form, with one verdict for every paragraph number:",
    `{"batch_id":"${batchId}","verdicts":[[<paragraph number>,"ok"|"violation"|"unsure"], ...],"violations":[{"paragraph_id":<paragraph number>,"quote":"<exact words>","reason":"<short>"}],"unsure":[{"paragraph_id":<paragraph number>,"reason":"<short>"}]}`,
    'Every paragraph with verdict "violation" needs an entry in "violations". "violations" and "unsure" hold only paragraphs with that verdict; use empty lists when there are none.',
    "Paragraphs:",
    ...batch.map((p, k) => `[${k}] ${viewText(p)}`),
  ].join("\n");
}

// The second request for a batch whose reply could not be used.
const retryPrompt = (prompt, problem, n) =>
  `${prompt}\n\nYour last reply could not be used: ${problem}. Reply again in the JSON form above, with one verdict for each of the ${n} paragraph(s).`;

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

// A reply to a batch of n paragraphs with id batchId, read strictly.
// Returns { problem } (why it cannot be used) or { verdicts: [verdict of
// paragraph 0, 1 ...], violations: [{ id, quote, reason }], unsure: Map(id ->
// reason) }. "violations" and "unsure" may be left out when empty; an
// "unsure" verdict needs no entry (its reason is then "").
export function readLlmReply(json, batchId, n) {
  const bad = (problem) => ({ problem });
  if (json === null || typeof json !== "object" || Array.isArray(json))
    return bad("reply is not a JSON object");
  if (json.batch_id !== batchId)
    return bad(`batch_id is ${JSON.stringify(json.batch_id ?? null)}, expected "${batchId}"`);
  if (!Array.isArray(json.verdicts)) return bad("reply has no verdicts array");
  const verdicts = new Array(n).fill(null);
  for (const [i, v] of json.verdicts.entries()) {
    if (!Array.isArray(v) || v.length !== 2 || notInteger(v[0]) || !VERDICTS.includes(v[1]))
      return bad(`verdicts[${i}] is not [paragraph number, "ok" | "violation" | "unsure"]`);
    const id = Number(v[0]);
    if (id < 0 || id >= n)
      return bad(`verdicts[${i}] names paragraph ${id}, which is not in the batch (0 to ${n - 1})`);
    if (verdicts[id] !== null) return bad(`paragraph ${id} has more than one verdict`);
    verdicts[id] = v[1];
  }
  const none = verdicts.flatMap((v, id) => (v === null ? [id] : []));
  if (none.length) return bad(`no verdict for paragraph(s) ${none.join(", ")}`);

  // An evidence list: entries for paragraphs of the batch whose verdict is `want`.
  const entries = (key, want, fields) => {
    const list = json[key] ?? [];
    if (!Array.isArray(list)) return bad(`${key} is not an array`);
    for (const [i, e] of list.entries()) {
      if (e === null || typeof e !== "object") return bad(`${key}[${i}] is not an object`);
      if (notInteger(e.paragraph_id)) return bad(`${key}[${i}].paragraph_id is not an integer`);
      for (const f of fields)
        if (typeof e[f] !== "string") return bad(`${key}[${i}].${f} is not a string`);
      const id = Number(e.paragraph_id);
      if (verdicts[id] !== want)
        return bad(
          `${key}[${i}] is for paragraph ${id}, whose verdict is ${JSON.stringify(verdicts[id] ?? "none")}`,
        );
    }
    return { list: list.map((e) => ({ ...e, id: Number(e.paragraph_id) })) };
  };
  const v = entries("violations", "violation", ["quote", "reason"]);
  if (v.problem) return v;
  const u = entries("unsure", "unsure", ["reason"]);
  if (u.problem) return u;
  const bare = verdicts.findIndex(
    (verdict, id) => verdict === "violation" && !v.list.some((e) => e.id === id),
  );
  if (bare !== -1)
    return bad(`paragraph ${bare} has verdict "violation" and no entry in violations`);
  return {
    verdicts,
    violations: v.list,
    unsure: new Map(u.list.map((e) => [e.id, e.reason])),
  };
}

// What a one-shot runner answers when the model wrote something that is no
// JSON (oneshot.mjs and every provider under providers/ use this text). A
// transport failure that mentions JSON ("the reply was not JSON from the
// server") is another error and is not asked again.
const NOT_JSON = "reply was not JSON";

// The usable reply to one batch, or { error }. A reply in the wrong form
// (also one that is no JSON) is asked for once more.
async function askBatch(item, batch, oneshot) {
  const batchId = nextBatchId();
  const prompt = ledgerPrompt(item, batch, batchId);
  let problem = null;
  for (const again of [false, true]) {
    const r = await oneshot({
      system: LEDGER_SYSTEM,
      prompt: again ? retryPrompt(prompt, problem, batch.length) : prompt,
    });
    if (r.ok) {
      const read = readLlmReply(r.json, batchId, batch.length);
      if (!read.problem) return read;
      problem = read.problem;
    } else if (r.error === NOT_JSON) {
      problem = NOT_JSON;
    } else {
      return { error: r.error };
    }
  }
  return { error: `malformed LLM reply: ${problem}` };
}

// onDecision (optional): called once per paragraph the model judged, with
// { item_id, paragraph_index, instruction, paragraph, verdict, error? };
// verdict is violation, ok, unsure, or error (failed or malformed reply).
// It carries the texts so the caller can fingerprint them (decision-log.mjs).
// A paragraph that is in neither `violations` nor `unsure` of the result was
// given the verdict "ok" by the model, in so many words: only those (and
// located violations) are cached by checkLedgerIncremental.
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
    for (const [k, p] of batch.entries()) {
      onDecision?.({
        item_id: item.id,
        paragraph_index: p.index,
        instruction: item.text,
        paragraph: viewText(p),
        verdict: verdictOf(k),
        ...(error ? { error } : {}),
      });
    }
  };
  for (const batch of batches) {
    const reply = await askBatch(item, batch, oneshot);
    if (reply.error) {
      decide(batch, () => "error", reply.error);
      return { id: item.id, status: "error", error: reply.error, violations, unsure };
    }
    // Final verdict per paragraph of the batch. A violation whose quote is
    // not in the paragraph is no evidence and is dropped: the model must
    // point at text. A paragraph left with no located quote is not ok
    // either: it turns unsure, stays with the author and out of the cache,
    // so the next incremental check reads it again (Codex review 12).
    const final = [...reply.verdicts];
    const lostReason = new Map();
    const hit = new Set();
    for (const v of reply.violations) {
      const para = batch[v.id];
      const match = locateQuote(viewText(para), v.quote);
      if (match === null) {
        if (!lostReason.has(v.id)) lostReason.set(v.id, v.reason);
        continue;
      }
      hit.add(v.id);
      violations.push({ paragraph_index: para.index, start: null, match, reason: v.reason });
    }
    for (const [k, p] of batch.entries()) {
      if (final[k] === "unsure") {
        unsure.push({ paragraph_index: p.index, reason: reply.unsure.get(k) ?? "" });
      } else if (final[k] === "violation" && !hit.has(k)) {
        final[k] = "unsure";
        const reason = lostReason.get(k);
        unsure.push({
          paragraph_index: p.index,
          reason: reason ? `${QUOTE_NOT_FOUND} (${reason})` : QUOTE_NOT_FOUND,
        });
      }
    }
    decide(batch, (k) => final[k]);
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

// Does this ledger item apply to `scope` (manuscript | response | supplement)?
// An item without a scope is a manuscript item (the schema default).
export const inScope = (item, scope) => (item?.scope ?? "manuscript") === scope;

// Ledger items that apply to manuscript edits. The review prompt
// (review-prompt.mjs) and the pre-check below use this one rule, so what the
// model is told and what is refused cannot drift apart.
export function inManuscriptScope(item) {
  return inScope(item, "manuscript");
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
    // A count rule needs the headings of the document, which this check
    // does not have: the Tools tab checks it (aux-service.mjs format_check).
    if (!inManuscriptScope(item) || item.check.type === "count") {
      outcomes.push({
        id: item.id,
        status: "skipped",
        skipped: item.check.type === "count" ? "count" : "scope",
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
