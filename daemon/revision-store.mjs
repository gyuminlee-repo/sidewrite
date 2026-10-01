// Per-manuscript review state: `<manuscript stem>.revision.json` next to the
// document. Holds the reviewer comments, every tracked edit the review flow
// made (with its rationale comment id and before/after sentences), edits
// that the forbid list refused, and the instruction ledger.
//
// Writes are atomic (temp file in the same folder, then rename) and
// serialized per file inside this process, so two tracked edits that land
// back to back cannot lose each other.

import { readFile, writeFile, rename, unlink } from "node:fs/promises";
import { randomUUID } from "node:crypto";
import { basename, dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { z } from "zod";
import { toLocalPath } from "./workspace.mjs";
import { compileCheck } from "./ledger-check.mjs";

export const REVISION_VERSION = 1;

const RegexCheck = z.object({
  type: z.literal("regex"),
  pattern: z.string().min(1),
  flags: z
    .string()
    .regex(/^[imsuy]*$/, "flags may use i, m, s, u, y (g is added by the checker)")
    .default(""),
  // absent: every match is a violation. present: the pattern must still
  // match somewhere (a "do not delete X" rule).
  expect: z.enum(["absent", "present"]).default("absent"),
});
const LlmCheck = z.object({ type: z.literal("llm") });
// A count the manuscript must keep (count-check.mjs): the words of a section,
// the number of keywords, figures or tables, the order of sections. Checked
// from the Tools tab, which has the headings; the review check skips it.
const CountCheck = z.object({
  type: z.literal("count"),
  target: z.enum(["section_words", "keywords", "figures", "tables", "section_order"]),
  section: z.string().min(1).optional(),
  min: z.number().int().nonnegative().nullable().optional(),
  max: z.number().int().nonnegative().nullable().optional(),
  order: z.array(z.string().min(1)).optional(),
});
// A regex check as a rule may be entered: pattern and flags must compile
// the way the checker compiles them (ledger-check.mjs compileCheck). Letters
// that are each allowed can still be refused together ("ii").
const CompiledRegexCheck = RegexCheck.superRefine((check, ctx) => {
  try {
    compileCheck(check);
  } catch (e) {
    ctx.addIssue({ code: "custom", path: ["pattern"], message: `invalid regex: ${e.message}` });
  }
});

const Rationale = z.object({
  what: z.string().min(1),
  why: z.string().min(1),
  addresses: z.string().min(1),
  // L5: ids of the comment requirements this change answers (q1, q2 ...).
  requirement_ids: z.array(z.string()).default([]),
});

export const ChangeSchema = z.object({
  id: z.string(),
  word_comment_id: z.string().nullable(),
  paragraph_index: z.number().int().nullable(),
  paragraph_id: z.string().nullable().default(null),
  kind: z.enum(["replace", "delete", "insert"]),
  find: z.string(),
  // Replacement text (replace), "" (delete), or the inserted text (insert).
  text: z.string(),
  before: z.string(),
  after: z.string(),
  rationale: Rationale,
  override_of: z.array(z.string()).default([]),
  // L1: facts in the deleted text that the new text does not keep.
  lost_facts: z.array(z.string()).default([]),
  // L2: new facts and words found neither in the manuscript nor in the comment.
  unsourced_terms: z.array(z.string()).default([]),
  // L3: the find/replace the model asked for, when the applied range is smaller.
  narrowed_from: z.object({ find: z.string(), text: z.string() }).nullable().default(null),
  // Up to 3 words before and after the changed range at edit time, so one
  // change can be judged apart from others in the same paragraph. null for
  // changes recorded before this field (whole-paragraph rule).
  context: z.object({ left: z.string(), right: z.string() }).nullable().default(null),
  // Another paragraph had the same text as this one when the change was
  // made: the old text alone then does not say which of them this is
  // (revision-classify.mjs locateParagraph).
  duplicate_before: z.boolean().default(false),
  // The edit was applied but its rationale comment was not anchored: the
  // task pane error code (CommentAnchorMissing or CommentInsertFailed).
  // word_comment_id is then null.
  comment_missing: z.string().nullable().default(null),
  // When the task pane's comment cleanup (tool comments or Delete ALL)
  // removed the rationale comment: its absence is then no sign of a reject
  // and no missing comment (Codex review 7).
  comment_cleaned_at: z.string().nullable().default(null),
  created_at: z.string(),
  state: z.enum(["pending", "accepted", "rejected", "user_edited", "unlocated"]).default("pending"),
  checked_at: z.string().nullable().default(null),
});

// A tracked edit on its way to Word, or one the task pane did not answer
// (OutcomeUnknown): the change it would record and the document form it
// expects, so a late answer or a status refresh can record it once. It is
// written before Word is called (addAttempt, setBlockedAttempt).
const Attempt = z.object({
  // The bridge call id, once a call got no answer; null before.
  tool_call_id: z.string().nullable(),
  // When the edit was sent to Word; rewritten when its answer was given up.
  unknown_at: z.string(),
  comment_text: z.string(),
  change: z.record(z.string(), z.unknown()),
});

const Blocked = z.object({
  id: z.string(),
  args: z.record(z.string(), z.unknown()),
  // Ledger ids of the forbid rules that refused the edit. Empty: a normal
  // edit (nothing refused it), kept here with its attempt from before Word
  // is called until its answer, a late answer or a refresh settles it
  // (addAttempt). Its id is the id of the attempt.
  blocked_by: z.array(z.string()),
  created_at: z.string(),
  // outcome_unknown: the edit was sent and not answered; it may be in the
  // document, so a plain second override is refused (review-service.mjs).
  // applying: the edit is with Word now; any second override is refused.
  // An applying item no running call owns (the daemon stopped, or the
  // record of an applied edit could not be saved) turns outcome_unknown
  // (markStaleApplying).
  state: z.enum(["blocked", "applying", "outcome_unknown"]).default("blocked"),
  attempt: Attempt.nullable().default(null),
});

// L5: one thing the reviewer asks for, e.g. { id: "q1", text: "state the unit" }.
const Requirement = z.object({
  id: z.string().regex(/^q\d+$/),
  text: z.string().min(1),
});

// ---- response letter (stage 3) ---------------------------------------------
//
// Only `text`, `gloss` and the decline reasons are prose a person or the model
// wrote. Citations, line ranges and stale marks are computed by the daemon
// (response-quotes.mjs, response-lines.mjs) and never taken from a request.

// Line range of a quote, read from the margin numbers of the rendered PDF.
const Lines = z.object({
  start: z.number().int(),
  end: z.number().int(),
  start_page: z.number().int().nullable().default(null),
  end_page: z.number().int().nullable().default(null),
  // How often the quote is in the PDF inside its paragraph; 1 is the only good value.
  match_count: z.number().int(),
});

const Citation = z.object({
  change_ids: z.array(z.string()).default([]),
  kind: z.enum(["revised", "removed"]).default("revised"),
  // Manuscript heading above the change, the paragraph and the quoted sentence.
  section: z.string().nullable().default(null),
  paragraph_index: z.number().int().nullable().default(null),
  quote: z.string().default(""),
  lines: Lines.nullable().default(null),
  // sha256 of the PDF `lines` was read from (response_render.pdf_sha).
  lines_pdf: z.string().nullable().default(null),
  // Why the change could not be placed ("unlocated", "sentence not found").
  problem: z.string().nullable().default(null),
});

const Response = z.object({
  text: z.string().default(""),
  // Translation of `text`, written only when the profile asks for it.
  gloss: z.string().default(""),
  source: z.enum(["model", "author", "imported"]).default("model"),
  drafted_at: z.string().nullable().default(null),
  // The text this one replaced (the last draft, or the letter of an earlier round).
  previous: z.string().nullable().default(null),
  citations: z.array(Citation).default([]),
  stale: z.array(z.enum(["quote_changed", "change_state", "lines_old"])).default([]),
  // For quote_changed and change_state: the manuscript text (snapshotHash)
  // in which the saved citations were last found not to fit, answer or no
  // answer. A draft compares it between its request and its save
  // (response-quotes.mjs markStale).
  stale_at: z.record(z.string(), z.string()).default({}),
  declined_requirements: z
    .array(z.object({ id: z.string(), reason: z.string().min(1) }))
    .default([]),
});

// A question the model put to the author instead of settling it with an
// edit (review-tools.mjs askAuthor): what only the author's records can say.
// answer: what the author wrote on the card ("" until then). An answer goes
// into the prompt of every later turn and counts as a source of numbers.
const Question = z.object({
  id: z.string().regex(/^a\d+$/),
  // The listed request the question holds up, when there is one.
  requirement_id: z.string().nullable().default(null),
  question: z.string().min(1),
  why: z.string().default(""),
  asked_at: z.string(),
  answer: z.string().default(""),
  answered_at: z.string().nullable().default(null),
});

const Comment = z.object({
  // R<reviewer>.<n>; E.<n> is the editor letter with reviewer 0.
  ref: z.string().regex(/^(?:R\d+|E)\.\d+$/),
  reviewer: z.number().int().nonnegative(),
  text: z.string(),
  // Label of the part of the review the comment is in ("Minor comments"),
  // from the split. Refs do not carry it: numbers go on across sections.
  section: z.string().optional(),
  requirements: z.array(Requirement).default([]),
  status: z.enum(["pending", "in_progress", "done"]).default("pending"),
  changes: z.array(ChangeSchema).default([]),
  blocked: z.array(Blocked).default([]),
  // Highest q-number and blocked-edit number ever given out under this
  // comment (see "id counters" below). null: a file from before the counters.
  requirement_seq: z.number().int().nonnegative().nullable().default(null),
  blocked_seq: z.number().int().nonnegative().nullable().default(null),
  questions: z.array(Question).default([]),
  // The answer in the response letter, and the author's decision not to
  // change anything for this comment.
  response: Response.nullable().default(null),
  declined: z
    .object({ reason: z.string().min(1) })
    .nullable()
    .default(null),
});

// A paragraph the LLM check could not judge (abstain). Not a violation; it
// waits for the author and is read again by the next incremental check.
const Unsure = z.object({
  paragraph_index: z.number().int(),
  h: z.string().nullable(),
  reason: z.string().default(""),
});

const Violation = z.object({
  paragraph_index: z.number().int().nullable(),
  h: z.string().nullable(),
  start: z.number().int().nullable().default(null),
  match: z.string().nullable().default(null),
  reason: z.string().optional(),
  missing: z.boolean().optional(),
});

const ledgerItem = (regexCheck) =>
  z.object({
    id: z.string().regex(/^L\d+$/),
    kind: z.enum(["instruction", "forbid"]),
    text: z.string().min(1),
    scope: z.enum(["manuscript", "response", "supplement"]).default("manuscript"),
    check: z.discriminatedUnion("type", [regexCheck, LlmCheck, CountCheck]),
    // Where the rule comes from when the author did not type it: the sentence
    // of the journal guideline it was taken from, word for word.
    source: z
      .object({ kind: z.literal("guideline"), sentence: z.string().min(1) })
      .nullable()
      .default(null),
    // skipped: outside the manuscript scope, not checked against the manuscript.
    // unsure: an LLM check could not judge some paragraphs and found no violation.
    status: z
      .enum(["unchecked", "ok", "violations", "error", "skipped", "unsure"])
      .default("unchecked"),
    violation_count: z.number().int().nonnegative().default(0),
    checked_at: z.string().nullable().default(null),
    // L10: violations of the last check keyed by paragraph hash, and the
    // hashes of every paragraph that check covered (null: never checked).
    violations: z.array(Violation).default([]),
    unsure: z.array(Unsure).default([]),
    checked_hashes: z.array(z.string()).nullable().default(null),
  });

// A rule as it comes in (ledger_add, a ledger_update that carries a check,
// an import): its regex must compile.
export const LedgerItemSchema = ledgerItem(CompiledRegexCheck);
// A rule as the file holds it. A file written before the schema compiled
// the pattern can hold one that does not; refusing it here would lock the
// author out of the whole review file (loadRevision and saveRevision both
// parse with RevisionSchema). The file loads and saves, the ledger check
// reports the rule as an error, and a tracked edit is refused while such a
// forbid rule is active (ledger-check.mjs brokenForbidRules).
const StoredLedgerItem = ledgerItem(RegexCheck);

const External = z.object({
  id: z.string().regex(/^X\d+$/),
  detected_at: z.string(),
  kind: z.enum(["modified", "inserted", "deleted"]),
  paragraph_index: z.number().int().nullable(),
  old_index: z.number().int().nullable(),
  before: z.string(),
  after: z.string(),
  lost_facts: z.array(z.string()).default([]),
});

export const RevisionSchema = z.object({
  version: z.literal(REVISION_VERSION),
  manuscript: z.string(),
  updated_at: z.string(),
  comments: z.array(Comment).default([]),
  ledger: z.array(StoredLedgerItem).default([]),
  // Highest L-number ever given out. null: a file from before the counter.
  ledger_seq: z.number().int().nonnegative().nullable().default(null),
  // L4: paragraph texts (reviewed view) at the last status refresh, and the
  // changes found since then that no review change explains.
  snapshot: z
    .object({
      taken_at: z.string(),
      paragraphs: z.array(z.string()),
      // Paragraph ids of the task pane at that refresh, same order (null per
      // paragraph in a document with index ids; null: a file from before).
      // A paragraph_hint of a tracked edit is one of them (review-tools.mjs
      // editTarget).
      ids: z.array(z.string().nullable()).nullable().default(null),
    })
    .nullable()
    .default(null),
  external: z.array(External).default([]),
  // Journal profile of this paper and its own values on top of it
  // (journal-profile.mjs effectiveProfile).
  journal_profile: z.string().nullable().default(null),
  profile_overrides: z.record(z.string(), z.unknown()).default({}),
  // The rendered PDF the line ranges were read from, and the manuscript text
  // it was rendered from (response-quotes.mjs snapshotHash).
  response_render: z
    .object({
      pdf_sha: z.string(),
      snapshot_hash: z.string(),
      at: z.string(),
      lines_indexed: z.number().int(),
    })
    .nullable()
    .default(null),
  response_export: z
    .object({
      path: z.string(),
      exported_at: z.string(),
      fail: z.number().int(),
      warn: z.number().int(),
    })
    .nullable()
    .default(null),
});

// ---- paths ----------------------------------------------------------------

// Local path of the document in this pane, or null when the daemon cannot
// open: unsaved document, SharePoint/OneDrive URL, UNC share outside WSL.
export function localDocPath(activeDoc, env = process.env) {
  if (typeof activeDoc !== "string" || !activeDoc.trim()) return null;
  const local = toLocalPath(activeDoc.trim(), env);
  if (local === null) return null;
  if (/^file:\/\//i.test(local)) {
    try {
      return fileURLToPath(local);
    } catch {
      return null;
    }
  }
  if (/^[a-z][a-z0-9+.-]*:\/\//i.test(local)) return null;
  return local;
}

const DOC_EXT = /\.(docx?|docm|dotx?|dotm)$/i;

export function sidecarPath(docPath, suffix) {
  const stem = basename(docPath).replace(DOC_EXT, "") || "document";
  return join(dirname(docPath), `${stem}${suffix}`);
}

export const revisionPathFor = (docPath) => sidecarPath(docPath, ".revision.json");
export const commentsBackupPathFor = (docPath) => sidecarPath(docPath, ".comments_backup.json");
// LLM decision receipts (decision-log.mjs), append only.
export const decisionsPathFor = (docPath) => sidecarPath(docPath, ".decisions.jsonl");

// ---- load / save ------------------------------------------------------------

export function emptyRevision(manuscript, now = new Date()) {
  return {
    version: REVISION_VERSION,
    manuscript,
    updated_at: now.toISOString(),
    comments: [],
    ledger: [],
    snapshot: null,
    external: [],
    journal_profile: null,
    profile_overrides: {},
    response_render: null,
    response_export: null,
  };
}

export async function loadRevision(path, manuscript) {
  let raw;
  try {
    raw = await readFile(path, "utf8");
  } catch (e) {
    if (e.code === "ENOENT") return emptyRevision(manuscript);
    throw e;
  }
  return withCounters(RevisionSchema.parse(JSON.parse(raw)));
}

export async function writeJsonAtomic(path, data) {
  const tmp = join(dirname(path), `.${basename(path)}.${process.pid}.${randomUUID()}.tmp`);
  await writeFile(tmp, JSON.stringify(data, null, 2) + "\n", "utf8");
  try {
    await rename(tmp, path);
  } catch (e) {
    await unlink(tmp).catch(() => {});
    throw e;
  }
}

export async function saveRevision(path, revision, now = new Date()) {
  const data = RevisionSchema.parse({ ...revision, updated_at: now.toISOString() });
  await writeJsonAtomic(path, data);
  return data;
}

const chains = new Map();

// Load, apply `fn`, save; one at a time per path. `fn` returns the next
// revision, or null to leave the file untouched.
export function withRevision(path, manuscript, fn) {
  const prev = chains.get(path) ?? Promise.resolve();
  const run = prev
    .catch(() => {})
    .then(async () => {
      const current = await loadRevision(path, manuscript);
      const next = await fn(current);
      return next ? saveRevision(path, next) : current;
    });
  chains.set(path, run);
  return run;
}

// ---- pure operations (return a new revision) --------------------------------

const clone = (x) => structuredClone(x);

export function findComment(rev, ref) {
  return rev.comments.find((c) => c.ref === ref) ?? null;
}

function requireComment(rev, ref) {
  const c = findComment(rev, ref);
  if (!c) throw new Error(`No reviewer comment ${ref}`);
  return c;
}

// Replace the comment list with the user-edited one. Changes, blocked edits
// and status of a ref that survives are kept. Dropping a ref that already
// has changes is refused: its edits are in the document. A dropped ref in
// `keep` (work on it is still open, review-service.mjs heldRefs) stays, after
// the listed ones.
export function replaceComments(rev, incoming, keep = new Set()) {
  const next = clone(rev);
  const seen = new Set();
  const list = [];
  for (const c of incoming) {
    if (seen.has(c.ref)) throw new Error(`Duplicate comment ref ${c.ref}`);
    seen.add(c.ref);
    const old = findComment(rev, c.ref);
    const section = c.section ?? old?.section;
    list.push({
      ref: c.ref,
      reviewer: c.reviewer,
      text: c.text,
      ...(section && { section }),
      requirements: Array.isArray(c.requirements)
        ? numberRequirements(old?.requirements ?? [], c.requirements, [`q${requirementSeq(old)}`])
        : clone(old?.requirements ?? []),
      status: c.status ?? old?.status ?? "pending",
      changes: old ? clone(old.changes) : [],
      blocked: old ? clone(old.blocked) : [],
      blocked_seq: old ? blockedSeq(old) : 0,
      response: old ? clone(old.response ?? null) : null,
      declined: old ? clone(old.declined ?? null) : null,
      questions: old ? clone(old.questions ?? []) : [],
    });
    const row = list[list.length - 1];
    row.requirement_seq = requirementSeq({ ...row, requirement_seq: requirementSeq(old) });
  }
  for (const old of rev.comments) {
    if (seen.has(old.ref)) continue;
    if (old.changes.length > 0) {
      throw new Error(`Comment ${old.ref} has ${old.changes.length} recorded change(s); keep it`);
    }
    if (keep.has(old.ref)) list.push(clone(old));
  }
  next.comments = list;
  return next;
}

// Put a comment back so a change or blocked edit can be recorded under it
// when the comment was removed anyway. from: { reviewer, text, section, requirements }
// the daemon still holds, or null for the ref alone.
export function restoreComment(rev, ref, from = null) {
  const m = /^(?:R(\d+)|E)\.\d+$/.exec(ref);
  if (!m) throw new Error(`Bad comment ref ${ref}`);
  const next = clone(rev);
  next.comments.push(
    Comment.parse({
      ref,
      reviewer: from?.reviewer ?? (m[1] ? Number(m[1]) : 0),
      text: from?.text ?? "",
      ...(from?.section && { section: from.section }),
      requirements: clone(from?.requirements ?? []),
      status: "in_progress",
    }),
  );
  return next;
}

// Requirement ids that recorded changes of this comment cite, and those that
// blocked edits cite (blocked, applying or outcome_unknown, with the change
// an unanswered attempt would record): an approval or a late answer records
// them later, so their ids stay taken (Codex review 9).
function citedIds(comment) {
  const ids = (r) => (Array.isArray(r?.requirement_ids) ? r.requirement_ids : []);
  return [
    ...(comment?.changes ?? []).flatMap((ch) => ids(ch.rationale)),
    ...(comment?.blocked ?? []).flatMap((b) => [
      ...ids(b.args?.rationale),
      ...ids(b.attempt?.change?.rationale),
    ]),
  ].filter((id) => typeof id === "string");
}

// ---- id counters -------------------------------------------------------------
// Ledger ids (L<n>), requirement ids (q<n>, per comment) and blocked-edit ids
// (<ref>-b<n>, per comment) are never given out twice: a stale reference to a
// removed item (blocked_by of an older blocked edit, an Allow exception
// button in a pane that was not redrawn, a running turn that was told q2)
// would otherwise mean the item that took its id (Codex review, P2: delete
// L1, add a forbid rule, it became L1 and the old exception approved it).
// The highest number given out is stored; a file without it gets the highest
// number among the ids it still names.

function idNumber(id, prefix) {
  if (typeof id !== "string" || !id.startsWith(prefix)) return 0;
  const rest = id.slice(prefix.length);
  return /^\d+$/.test(rest) ? Number(rest) : 0;
}

// Every string anywhere in the file: blocked_by, override_of, the change an
// unanswered attempt would record, and fields added later.
function maxIdAnywhere(value, prefix) {
  let max = 0;
  const walk = (v) => {
    if (typeof v === "string") max = Math.max(max, idNumber(v, prefix));
    else if (v && typeof v === "object") Object.values(v).forEach(walk);
  };
  walk(value);
  return max;
}

const ledgerSeq = (rev) => Math.max(rev.ledger_seq ?? 0, maxIdAnywhere(rev, "L"));

function requirementSeq(comment) {
  const ids = [...(comment?.requirements ?? []).map((r) => r.id), ...citedIds(comment)];
  return Math.max(comment?.requirement_seq ?? 0, ...ids.map((id) => idNumber(id, "q")));
}

function blockedSeq(comment) {
  const ids = (comment?.blocked ?? []).map((b) => b.id);
  return Math.max(comment?.blocked_seq ?? 0, ...ids.map((id) => idNumber(id, `${comment.ref}-b`)));
}

// On load: write the counters down, so they hold after the ids they were
// read from are gone.
function withCounters(rev) {
  rev.ledger_seq = ledgerSeq(rev);
  for (const c of rev.comments) {
    c.requirement_seq = requirementSeq(c);
    c.blocked_seq = blockedSeq(c);
  }
  return rev;
}

// L5. texts: strings or { text }. A text that was already listed keeps its
// id, so changes that point at it stay linked; new texts get the next q-id.
// The next q-id is above every listed id and every id in `used` (the
// comment's counter as q<n>), so a removed requirement's id is never given
// to a new one.
export function numberRequirements(old, texts, used = []) {
  const byText = new Map(old.map((r) => [r.text, r.id]));
  let n = nextNumber([...old.map((r) => r.id), ...used], "q");
  const out = [];
  const seen = new Set();
  for (const t of texts) {
    const text = String(typeof t === "string" ? t : (t?.text ?? "")).trim();
    if (!text || seen.has(text)) continue;
    seen.add(text);
    out.push({ id: byText.get(text) ?? `q${n++}`, text });
  }
  return out;
}

export function setRequirements(rev, ref, texts) {
  const next = clone(rev);
  const c = requireComment(next, ref);
  const seq = requirementSeq(c);
  c.requirements = numberRequirements(c.requirements, texts, [`q${seq}`]);
  c.requirement_seq = requirementSeq({ ...c, requirement_seq: seq });
  return next;
}

// A question for the author under a comment: ids a1, a2, ... in the order
// asked (questions are never removed, so the next id is the count plus one
// and no counter is kept). The same question asked again while it has no
// answer gives the row that is there.
export function addQuestion(
  rev,
  ref,
  { requirement_id = null, question, why = "" },
  now = new Date(),
) {
  const next = clone(rev);
  const c = requireComment(next, ref);
  c.questions ??= [];
  const text = String(question ?? "").trim();
  const same = c.questions.find((q) => !q.answer && q.question === text);
  if (same) return { revision: next, question: same };
  const row = Question.parse({
    id: `a${nextNumber(
      c.questions.map((q) => q.id),
      "a",
    )}`,
    requirement_id: requirement_id ?? null,
    question: text,
    why: String(why ?? "").trim(),
    asked_at: now.toISOString(),
  });
  c.questions.push(row);
  return { revision: next, question: row };
}

// The author's answer to a question; an empty one takes the answer back.
export function answerQuestion(rev, ref, id, answer, now = new Date()) {
  const next = clone(rev);
  const q = (requireComment(next, ref).questions ?? []).find((x) => x.id === id);
  if (!q) throw new Error(`No question ${id} under ${ref}`);
  q.answer = String(answer ?? "").trim();
  q.answered_at = q.answer ? now.toISOString() : null;
  return next;
}

// What the author answered, anywhere in the review.
export const authorAnswers = (rev) =>
  (rev?.comments ?? []).flatMap((c) => (c.questions ?? []).map((q) => q.answer).filter(Boolean));

export function setCommentStatus(rev, ref, status) {
  const next = clone(rev);
  requireComment(next, ref).status = status;
  return next;
}

function nextNumber(ids, prefix) {
  let max = 0;
  for (const id of ids) {
    if (id.startsWith(prefix)) {
      const n = Number(id.slice(prefix.length));
      if (Number.isInteger(n) && n > max) max = n;
    }
  }
  return max + 1;
}

export function addChange(rev, ref, change, now = new Date()) {
  const next = clone(rev);
  const c = requireComment(next, ref);
  const id = `${ref}-${nextNumber(
    c.changes.map((x) => x.id),
    `${ref}-`,
  )}`;
  const row = ChangeSchema.parse({ ...change, id, created_at: now.toISOString() });
  c.changes.push(row);
  if (c.status === "pending") c.status = "in_progress";
  return { revision: next, change: row };
}

// A change that was applied without its rationale comment (comment_missing)
// got the comment later (review-service.mjs review_comment_set): its id is
// recorded, once, and never an id another change holds.
export function setChangeComment(rev, ref, changeId, wordCommentId) {
  const next = clone(rev);
  const ch = requireComment(next, ref).changes.find((x) => x.id === changeId);
  if (!ch) throw new Error(`No change ${changeId} under ${ref}`);
  if (ch.word_comment_id != null) {
    throw new Error(`${changeId} already has its rationale comment`);
  }
  const id = wordCommentId == null ? "" : String(wordCommentId);
  if (!id) throw new Error("a Word comment id is required");
  if (next.comments.some((c) => c.changes.some((x) => x.word_comment_id === id))) {
    throw new Error(`Word comment ${id} belongs to another change`);
  }
  ch.word_comment_id = id;
  ch.comment_missing = null;
  return next;
}

// ids: Word comment ids the task pane's cleanup deleted. Marks the changes
// whose rationale comment is one of them; returns null when none is.
export function markCommentsCleaned(rev, ids, now = new Date()) {
  const gone = new Set((ids ?? []).map(String));
  const next = clone(rev);
  let hit = false;
  for (const c of next.comments) {
    for (const ch of c.changes) {
      if (ch.word_comment_id == null || ch.comment_cleaned_at) continue;
      if (!gone.has(String(ch.word_comment_id))) continue;
      ch.comment_cleaned_at = now.toISOString();
      hit = true;
    }
  }
  return hit ? next : null;
}

// rows: [{ ref, change_id, state, paragraph_index }] from classifyRevision.
export function applyStates(rev, rows, now = new Date()) {
  const next = clone(rev);
  const at = now.toISOString();
  for (const r of rows) {
    const ch = findComment(next, r.ref)?.changes.find((x) => x.id === r.change_id);
    if (!ch) continue;
    ch.state = r.state;
    if (r.paragraph_index != null) ch.paragraph_index = r.paragraph_index;
    ch.checked_at = at;
  }
  return next;
}

export function addBlocked(rev, ref, { args, blocked_by }, now = new Date()) {
  const next = clone(rev);
  const c = requireComment(next, ref);
  c.blocked_seq = blockedSeq(c) + 1;
  const id = `${ref}-b${c.blocked_seq}`;
  const row = { id, args: clone(args), blocked_by: [...blocked_by], created_at: now.toISOString() };
  c.blocked.push(row);
  return { revision: next, blocked: row };
}

// A tracked edit no forbid rule refused, written BEFORE Word is called: from
// then on it may be in the document. Kept like a claimed exception (state
// applying, with the attempt), with no rule in blocked_by. Its id comes from
// the blocked-edit counter of the comment, so no two attempts share one.
export function addAttempt(rev, ref, { args, attempt }, now = new Date()) {
  const next = clone(rev);
  const c = requireComment(next, ref);
  c.blocked_seq = blockedSeq(c) + 1;
  const row = {
    id: `${ref}-b${c.blocked_seq}`,
    args: clone(args),
    blocked_by: [],
    created_at: now.toISOString(),
    state: "applying",
    attempt: Attempt.parse(clone(attempt)),
  };
  c.blocked.push(row);
  return { revision: next, blocked: row };
}

// The same for a claimed exception (claimBlocked), before Word is called.
export function setBlockedAttempt(rev, ref, blockedId, attempt) {
  const next = clone(rev);
  const b = requireComment(next, ref).blocked.find((x) => x.id === blockedId);
  if (!b) throw new Error(`No blocked edit ${blockedId} under ${ref}`);
  b.attempt = Attempt.parse(clone(attempt));
  return next;
}

export function markBlockedUnknown(rev, ref, blockedId, attempt) {
  const next = clone(rev);
  const b = requireComment(next, ref).blocked.find((x) => x.id === blockedId);
  if (!b) throw new Error(`No blocked edit ${blockedId} under ${ref}`);
  b.state = "outcome_unknown";
  b.attempt = Attempt.parse(clone(attempt));
  return next;
}

// Start an override: the item turns applying so an overlapping override of
// the same item is refused instead of editing Word twice. An unanswered one
// (outcome_unknown) goes again only with force.
export function claimBlocked(rev, ref, blockedId, { force = false } = {}) {
  const next = clone(rev);
  const b = findComment(next, ref)?.blocked.find((x) => x.id === blockedId);
  if (!b) throw new Error(`No blocked edit ${blockedId} under ${ref}`);
  if (b.state === "applying") {
    throw new Error(`${b.id}: this exception is already being applied. Wait for it to finish.`);
  }
  if (b.state === "outcome_unknown" && !force) {
    throw new Error(
      `${b.id}: Word did not answer the last attempt, so it may already be applied. Refresh status, then check the paragraph. Re-apply only if the change is not there.`,
    );
  }
  const before = clone(b);
  b.state = "applying";
  return { revision: next, blocked: before };
}

// End an override that did not apply: back to the state it had before, and
// to the attempt it had then when one is given (`attempt`, null for none).
export function releaseBlocked(rev, ref, blockedId, state, attempt) {
  const next = clone(rev);
  const b = findComment(next, ref)?.blocked.find((x) => x.id === blockedId);
  if (!b || b.state !== "applying") return null;
  b.state = state;
  if (attempt !== undefined) b.attempt = clone(attempt);
  return next;
}

// The revision without that blocked edit, or null when it is not there.
export function dropBlocked(rev, ref, blockedId) {
  return findComment(rev, ref)?.blocked.some((b) => b.id === blockedId)
    ? takeBlocked(rev, ref, blockedId).revision
    : null;
}

// applying items no running override owns (the daemon stopped mid-call) may
// be in Word: they turn outcome_unknown. `active`: "<ref>\0<id>" keys.
export function markStaleApplying(rev, active) {
  const next = clone(rev);
  let n = 0;
  for (const c of next.comments) {
    for (const b of c.blocked) {
      if (b.state === "applying" && !active.has(`${c.ref}\0${b.id}`)) {
        b.state = "outcome_unknown";
        n++;
      }
    }
  }
  return n ? next : null;
}

export function takeBlocked(rev, ref, blockedId) {
  const next = clone(rev);
  const c = requireComment(next, ref);
  const i = c.blocked.findIndex((b) => b.id === blockedId);
  if (i === -1) throw new Error(`No blocked edit ${blockedId} under ${ref}`);
  const [blocked] = c.blocked.splice(i, 1);
  return { revision: next, blocked };
}

export function addLedgerItem(rev, item) {
  const next = clone(rev);
  next.ledger_seq = ledgerSeq(next) + 1;
  const id = `L${next.ledger_seq}`;
  const row = LedgerItemSchema.parse({ ...item, id });
  next.ledger.push(row);
  return { revision: next, item: row };
}

export function updateLedgerItem(rev, id, patch) {
  const next = clone(rev);
  const i = next.ledger.findIndex((x) => x.id === id);
  if (i === -1) throw new Error(`No ledger item ${id}`);
  const old = next.ledger[i];
  // A new rule text or check invalidates the last check (L10 needs a full run).
  const ruleChanged =
    ("text" in patch && patch.text !== old.text) ||
    ("check" in patch && JSON.stringify(patch.check) !== JSON.stringify(old.check));
  const reset = ruleChanged
    ? {
        status: "unchecked",
        violation_count: 0,
        violations: [],
        unsure: [],
        checked_hashes: null,
        checked_at: null,
      }
    : {};
  // A patch without a check (a check result, a new text) is stored on the
  // rule as the file holds it, also on one whose regex does not compile.
  const schema = "check" in patch ? LedgerItemSchema : StoredLedgerItem;
  next.ledger[i] = schema.parse({ ...old, ...patch, ...reset, id });
  return next;
}

export function removeLedgerItem(rev, id) {
  const next = clone(rev);
  const before = next.ledger.length;
  next.ledger = next.ledger.filter((x) => x.id !== id);
  if (next.ledger.length === before) throw new Error(`No ledger item ${id}`);
  return next;
}

// ---- P15: the ledger as a JSON file -------------------------------------------
// A file holds rules only ({ kind, text, scope, check }): ids and check
// results belong to one manuscript. tests/fixtures/review_ledger.json has the
// same item shape, so a hand-written list imports too.

const RULE_KEYS = ["kind", "text", "scope", "check", "source"];
const ruleOf = (item) => Object.fromEntries(RULE_KEYS.map((k) => [k, item?.[k]]));
const ruleKey = (item) => JSON.stringify(RULE_KEYS.map((k) => item[k]));

export function ledgerExportDoc(rev, now = new Date()) {
  return {
    format: "manuscript-pane-ledger",
    version: 1,
    exported_at: now.toISOString(),
    items: rev.ledger.map(ruleOf),
  };
}

// mode "append": rules the ledger already holds are skipped. "replace": the
// ledger is emptied first. Every row is checked before anything is added, so
// a bad row adds nothing. Ids come from addLedgerItem.
export function importLedgerItems(rev, items, { mode = "append" } = {}) {
  if (!Array.isArray(items)) throw new Error("ledger file needs an items array");
  if (mode !== "append" && mode !== "replace") throw new Error(`unknown import mode ${mode}`);
  const rows = items.map((it, n) => {
    const parsed = LedgerItemSchema.safeParse({ ...ruleOf(it), id: "L0" });
    if (!parsed.success) {
      const issue = parsed.error.issues[0];
      const where = issue?.path?.length ? `${issue.path.join(".")}: ` : "";
      throw new Error(`item ${n + 1}: ${where}${issue?.message ?? "invalid"}`);
    }
    return ruleOf(parsed.data);
  });
  let next = mode === "replace" ? { ...clone(rev), ledger: [] } : rev;
  const have = new Set(next.ledger.map(ruleKey));
  const added = [];
  const skipped = [];
  for (const row of rows) {
    if (have.has(ruleKey(row))) {
      skipped.push(row);
      continue;
    }
    const r = addLedgerItem(next, row);
    next = r.revision;
    have.add(ruleKey(row));
    added.push(r.item);
  }
  return { revision: next, added, skipped };
}

// ---- L4: snapshot and outside changes ----------------------------------------

export const EXTERNAL_KEEP = 100;

export function setSnapshot(rev, paragraphs, now = new Date(), ids = null) {
  const next = clone(rev);
  next.snapshot = {
    taken_at: now.toISOString(),
    paragraphs: [...paragraphs],
    ids: ids ? [...ids] : null,
  };
  return next;
}

// Append outside changes with X-ids; only the newest EXTERNAL_KEEP stay.
export function addExternal(rev, items, now = new Date()) {
  const next = clone(rev);
  let n = nextNumber(
    next.external.map((x) => x.id),
    "X",
  );
  const added = items.map((x) => ({ ...clone(x), id: `X${n++}`, detected_at: now.toISOString() }));
  next.external = [...next.external, ...added].slice(-EXTERNAL_KEEP);
  return { revision: next, added };
}

export function dismissExternal(rev, id) {
  const next = clone(rev);
  const before = next.external.length;
  next.external = next.external.filter((x) => x.id !== id);
  if (next.external.length === before) throw new Error(`No outside change ${id}`);
  return next;
}

// ---- response letter ---------------------------------------------------------

// The prose of an answer. patch: { text?, gloss?, source?, previous?,
// declined_requirements? }; nothing else is taken (citations and stale marks
// have their own function below). A new text keeps the one it replaces as
// `previous` and stamps drafted_at.
const RESPONSE_PROSE = ["text", "gloss", "source", "previous", "declined_requirements"];
export function setResponse(rev, ref, patch = {}, now = new Date()) {
  const next = clone(rev);
  const c = requireComment(next, ref);
  const old = c.response ?? Response.parse({});
  const take = Object.fromEntries(
    RESPONSE_PROSE.filter((k) => k in patch).map((k) => [k, patch[k]]),
  );
  const textChanged = "text" in take && take.text !== old.text;
  c.response = Response.parse({
    ...old,
    ...take,
    previous: textChanged && old.text ? old.text : (take.previous ?? old.previous),
    drafted_at: textChanged ? now.toISOString() : old.drafted_at,
  });
  return next;
}

// declined: { reason }; a blank reason takes the decline back.
export function setDeclined(rev, ref, declined) {
  const next = clone(rev);
  const reason = String(declined?.reason ?? "").trim();
  requireComment(next, ref).declined = reason ? { reason } : null;
  return next;
}

// The computed citations of a comment and, when given, its stale marks.
export function setCitations(rev, ref, citations, stale) {
  const next = clone(rev);
  const c = requireComment(next, ref);
  c.response = Response.parse({ ...(c.response ?? {}), citations, ...(stale ? { stale } : {}) });
  return next;
}

// The profile name and the paper's overrides. An override merges one level
// deep, like the profile merge; a null value removes its key, also inside a
// dict, and a dict left empty goes away.
export function setProfile(rev, { journal_profile, overrides } = {}) {
  const next = clone(rev);
  if (journal_profile !== undefined) next.journal_profile = journal_profile;
  const all = next.profile_overrides;
  for (const [k, v] of Object.entries(overrides ?? {})) {
    if (v === null) delete all[k];
    else if (typeof v === "object" && !Array.isArray(v)) {
      const base = all[k] && typeof all[k] === "object" && !Array.isArray(all[k]) ? all[k] : {};
      const merged = Object.fromEntries(
        Object.entries({ ...base, ...clone(v) }).filter(([, x]) => x !== null),
      );
      if (Object.keys(merged).length) all[k] = merged;
      else delete all[k];
    } else all[k] = v;
  }
  return next;
}

export const setRender = (rev, render) => ({ ...clone(rev), response_render: render });
export const setExport = (rev, exported) => ({ ...clone(rev), response_export: exported });
