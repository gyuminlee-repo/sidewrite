// Review-turn tool policy and the office_tracked_edit tool.
//
// In a review turn the only way to change the document is
// office_tracked_edit, which applies one tracked change and anchors one
// "[R1.3] what / why / addresses" comment on it. Every other document or
// comment writer is refused, so "changes == rationale comments" holds by
// construction. disallowedTools cannot do this per turn: it is fixed when
// query() starts (daemon/index.mjs:1062) and the session lives across turns.

import { z } from "zod";
import { forbidHits, forbidHitsAfterEdit, brokenForbidRules } from "./ledger-check.mjs";
import {
  lostFacts,
  unsourcedTerms,
  narrowEdit,
  startsSentence,
  extractFacts,
  factPresent,
} from "./edit-facts.mjs";
import { changeContext, normalizeParaText } from "./revision-classify.mjs";
import { WHOLE_AUTHOR_INSERTION } from "../taskpane/shared/review-view.js";
import { paragraphHash } from "./doc-diff.mjs";
import { newNumbers } from "./number-guard.mjs";
import { rewriteSegments } from "./word-diff.mjs";
import {
  FIELDS_UNREAD,
  fieldRefusal,
  fieldsUnread,
  fieldVerdict,
  fieldView,
  fieldViews,
  hasFieldMarks,
  spliceShown,
} from "../taskpane/shared/field-text.js";

export const REVIEW_WRITE_TOOL = "office_tracked_edit";
// The ways a review turn changes the document: one tracked edit, or a
// paragraph rewritten as a whole (applyTrackedRewrite), which reaches Word
// as tracked edits too.
export const REVIEW_WRITE_TOOLS = new Set([REVIEW_WRITE_TOOL, "office_tracked_rewrite"]);
// What a task pane must name in its hello to be sent the parts of a
// rewrite: it can add an edit without a comment and refuse a paragraph
// that is not plain text.
export const REWRITE_CAP = "tracked_rewrite";
// The capability a task pane must name in its hello to be sent a tracked
// edit (taskpane/shared/tools-word.js PANE_CAPS): it compares the paragraph
// with expected_before_hash before it edits.
export const TARGET_HASH_CAP = "target_hash";
// What a task pane must name to be sent a deletion that takes a field with
// it (a whole sentence with its citation): it checks the fields of the
// paragraph after the edit and takes the edit back when one is not whole.
export const FIELD_CAP = "field_guard";

// Read-only Word tools stay available in a review turn.
export const REVIEW_READ_TOOLS = new Set([
  "office_get_selection",
  "office_read_paragraphs",
  "office_get_document_text",
  "office_get_outline",
  "office_find",
  "office_list_comments",
]);

// Every other Word tool registered in office-tools.mjs writes the document,
// its formatting, or its comments. The test compares this list with the
// wordTools array so a newly added tool must be classified.
export const REVIEW_BLOCKED_TOOLS = new Set([
  "office_insert_paragraphs",
  "office_replace_paragraphs",
  "office_replace_text",
  "office_replace_section",
  "office_delete_paragraphs",
  "office_highlight",
  "office_clear_highlights",
  "office_add_comment",
  "office_clear_comments",
  "office_apply_style",
  "office_set_font",
  "office_set_paragraph_formatting",
  "office_insert_table",
  "office_set_table_cell",
  "office_set_list",
  "office_insert_image",
  "office_insert_hyperlink",
  "office_insert_bookmark",
  "office_reply_to_comment",
  "office_resolve_comment",
  "office_header_footer",
]);

// Tools of a review turn that write the review file and nothing in Word:
// a question for the author, and a request that belongs to another file.
export const REVIEW_NOTE_TOOLS = new Set(["office_ask_author", "office_requirement_outside"]);

export function reviewTurnToolAllowed(name) {
  return REVIEW_READ_TOOLS.has(name) || REVIEW_NOTE_TOOLS.has(name) || REVIEW_WRITE_TOOLS.has(name);
}

// Wrap the bridge call used by every office/excel tool handler.
export function guardedCall(rawCall, getReviewTurn) {
  return (name, args) => {
    const turn = getReviewTurn?.() ?? null;
    if (turn && !reviewTurnToolAllowed(name)) {
      return Promise.reject(
        new Error(
          `${name} is disabled during review turn ${turn.ref}. Use office_tracked_edit (or office_tracked_rewrite for a whole paragraph): every change needs its rationale comment.`,
        ),
      );
    }
    return rawCall(name, args);
  };
}

export const TRACKED_EDIT_DESCRIPTION =
  "Review turns only. Apply ONE tracked change (replace, delete, or pure insertion) and attach its rationale comment. " +
  'Replace/delete: give `find` and `replace` ("" deletes). Insertion: give `find` (anchor) and `insert_after`. ' +
  "`find` must occur exactly once in the document; if it occurs more than once, pass `paragraph_hint` (a paragraph ID from office_find). " +
  "A citation or another field is never edited: an edit inside its text, across its edge, or a replacement over it is refused (FieldNotEditable). To remove a citation, delete its whole sentence in ONE edit. " +
  "Returns ok:false with blocked_by when the new text, alone or together with the words next to it, matches the forbid list of the author; then rephrase, do not retry the same text.";

export const trackedEditShape = {
  comment_ref: z
    .string()
    .regex(/^(?:R\d+|E)\.\d+$/)
    .describe(
      'Reviewer comment this change answers, e.g. "R1.3". Must be the comment of this turn.',
    ),
  find: z
    .string()
    .min(1)
    .max(255)
    .describe(
      "Exact text to change, or the anchor to insert after. Word search limit: 255 characters.",
    ),
  replace: z
    .string()
    .optional()
    .describe('Replacement text. "" deletes `find`. Give exactly one of replace or insert_after.'),
  insert_after: z
    .string()
    .min(1)
    .optional()
    .describe(
      "Text inserted right after `find` (pure insertion). Include a leading space if needed.",
    ),
  paragraph_hint: z
    .string()
    .optional()
    .describe("Paragraph ID that narrows `find` when it occurs more than once."),
  rationale: z
    .object({
      what: z.string().min(1).describe("What changed."),
      why: z.string().min(1).describe("Why."),
      addresses: z.string().min(1).describe("Which part of the reviewer comment this answers."),
      requirement_ids: z
        .array(z.string())
        .optional()
        .describe(
          "Ids of the listed comment requirements this change handles (q1, q2). Use only ids from the prompt. If the change handles only part of a request, still put its id. Leave it empty only when the change answers none of them.",
        ),
    })
    .describe(
      "In the language the prompt asks for. Facts from the manuscript or the review only; say so when inferring.",
    ),
};

const commentRef = (what) =>
  z
    .string()
    .regex(/^(?:R\d+|E)\.\d+$/)
    .describe(`${what} Must be the comment of this turn.`);

export const ASK_AUTHOR_DESCRIPTION =
  "Review turns only. Ask the author ONE question instead of changing the manuscript, when what the comment needs cannot be told from the manuscript or the review: " +
  "which of two numbers that differ is right, how something was measured, a statistic that would have to be computed again, a count or a direction that is only in the author's records. " +
  "The question is kept on the card of the comment; the author's answer comes with the next turn. Do not edit the text the question is about in this turn, " +
  "except for a request the reviewer made conditional on evidence the manuscript does not give: there the change the reviewer asked for is made and this question goes with it (same requirement_id).";

export const askAuthorShape = {
  comment_ref: commentRef('Reviewer comment the question belongs to, e.g. "R1.3".'),
  requirement_id: z
    .string()
    .optional()
    .describe("Id of the listed request the question holds up (q2), when there is one."),
  question: z
    .string()
    .min(1)
    .describe(
      "The question, in the language of the rationale. Name the places and the values it is about, so the author can answer without searching.",
    ),
  why: z.string().min(1).describe("Why the manuscript and the review cannot answer it."),
  paragraph_id: z
    .string()
    .optional()
    .describe(
      "Id of the paragraph the answer would go into, when there is one such paragraph. Left out when the answer has no one place.",
    ),
};

// The turn and the request id of a note tool, checked as office_tracked_edit
// checks them. Returns { error } or { at, ctx }.
async function noteTurn(deps, tool, args, turn) {
  if (!turn)
    return { error: fail(`${tool} works only in a review turn started from the Review tab.`) };
  if (args.comment_ref !== turn.ref) {
    return {
      error: fail(
        `comment_ref ${args.comment_ref} does not match the comment of this turn ${turn.ref}.`,
      ),
    };
  }
  const at = deps.begin?.();
  const ctx = (await deps.getReviewContext?.(turn.ref, at)) ?? {};
  const known = (ctx.requirements ?? []).map((r) => r.id);
  if (args.requirement_id != null && !known.includes(args.requirement_id)) {
    return {
      error: fail(
        `Unknown requirement id ${args.requirement_id} for ${turn.ref}. Valid ids: ${known.join(", ") || "(none listed)"}.`,
      ),
    };
  }
  return { at, ctx };
}

// office_ask_author. deps: { begin(), getReviewContext(ref),
// onQuestion({ref, question: {requirement_id, question, why, paragraph_id?}, at})
// -> {id}, log }.
// The log names the question by id; its text stays in the review file.
// paragraph_id: kept when the manuscript of the last status refresh has that
// paragraph (its task pane id, else "p" and its index); any other id is
// dropped and the question is recorded without a place.
export async function askAuthor(deps, args, turn) {
  const t = await noteTurn(deps, "office_ask_author", args, turn);
  if (t.error) return t.error;
  const asked = args.paragraph_id ?? null;
  const place =
    asked !== null &&
    (t.ctx.paragraphs ?? []).some((_, i) => (t.ctx.paragraph_ids?.[i] ?? `p${i}`) === asked)
      ? asked
      : null;
  const saved = await deps.onQuestion({
    ref: turn.ref,
    question: {
      requirement_id: args.requirement_id ?? null,
      question: args.question,
      why: args.why,
      ...(place ? { paragraph_id: place } : {}),
    },
    at: t.at,
  });
  turn.follow?.recorded?.questions.push(saved.id);
  deps.log?.(
    `[review] question ${turn.ref} ${saved.id} requirement=${args.requirement_id ?? "-"}${place ? ` place=${place}` : ""}`,
  );
  const unknown =
    asked !== null && !place
      ? ` paragraph_id ${asked} is not a paragraph of the manuscript: the question is recorded without a place.`
      : "";
  return {
    ok: true,
    question_id: saved.id,
    message: `Recorded as ${saved.id} on the card of ${turn.ref}. Leave the text it is about as it is in this turn (unless the question goes with the change of a request made conditional on evidence the manuscript does not give); the author's answer comes with the next turn.${unknown}`,
  };
}

export const REQUIREMENT_OUTSIDE_DESCRIPTION =
  "Review turns only. Say that ONE listed request of the comment cannot be handled in this file because what it asks for is in another file (a supplementary file, a figure file, the response letter). " +
  "The author then sees the request as outside this file instead of open. Not for a request that needs something only the author knows (office_ask_author), and not for a request you chose not to handle.";

export const requirementOutsideShape = {
  comment_ref: commentRef('Reviewer comment the request belongs to, e.g. "R1.3".'),
  requirement_id: z.string().describe("Id of the listed request (q2)."),
  why: z
    .string()
    .min(1)
    .describe("Which file or place it belongs to, and what has to be done there."),
  part: z
    .string()
    .optional()
    .describe(
      "Only when the rest of the request is changed in this file: which part of it belongs to the other file. The author then sees both, the change here and this part as outside this file.",
    ),
};

// office_requirement_outside. deps: { begin(), getReviewContext(ref),
// onRequirementOutside({ref, note: {id, why, part}, at}), log }.
export async function requirementOutside(deps, args, turn) {
  const t = await noteTurn(deps, "office_requirement_outside", args, turn);
  if (t.error) return t.error;
  await deps.onRequirementOutside({
    ref: turn.ref,
    note: { id: args.requirement_id, why: args.why, ...(args.part ? { part: args.part } : {}) },
    at: t.at,
  });
  deps.log?.(
    `[review] requirement outside ${turn.ref} ${args.requirement_id}${args.part ? " part=1" : ""}`,
  );
  return {
    ok: true,
    message: args.part
      ? `The part of ${args.requirement_id} you named is shown to the author as outside this file, next to what is changed here.`
      : `${args.requirement_id} is shown to the author as outside this file.`,
  };
}

export const TOOL_COMMENT_PREFIX = /^\s*\[(?:R\d+|E)\.\d+\]/;

// L1: facts the change deletes go on a last line, so the author sees them in
// Word next to the change. L5: requirement ids follow the addresses part.
// fromReview: words and quantities of the new text that the review has and
// the manuscript does not: the change took them from the reviewer, and the
// author sees that they were not checked against a source.
export function formatRationaleComment(
  ref,
  { what, why, addresses, requirement_ids },
  { lost = [], fromReview = [] } = {},
) {
  const one = (s) => String(s).replace(/\s+/g, " ").trim();
  const ids = requirement_ids?.length ? ` [${requirement_ids.join(",")}]` : "";
  const facts = lost.length ? `\n사라지는 사실: ${lost.join("; ")}` : "";
  const taken = fromReview.length
    ? `\nFrom the review, not checked against the source: ${fromReview.join("; ")}`
    : "";
  return `[${ref}] ${one(what)} / ${one(why)} / ${one(addresses)}${ids}${facts}${taken}`;
}

const NO_ANSWER_CODES = new Set(["ToolTimeout", "TaskpaneDisconnected"]);

function fail(message, extra = {}) {
  return { ok: false, message, ...extra };
}

// The texts the recorded changes put into paragraph `target` (ctx.inserted:
// { paragraph_index, paragraph_id, text } of each change of this file not
// accepted or rejected). The task pane lets a deletion take a whole pending
// insertion only when its text is one of these (tools-word.js
// WholeAuthorInsertion). By paragraph id when both have one, else by index.
export function ownInsertions(inserted, target, ids) {
  if (target === null || !Array.isArray(inserted)) return [];
  const id = ids?.[target] ?? null;
  return [
    ...new Set(
      inserted
        .filter((c) =>
          id !== null && c.paragraph_id != null
            ? c.paragraph_id === id
            : c.paragraph_index === target,
        )
        .map((c) => c.text),
    ),
  ];
}

// deps: { begin(), call(name, args), getLedger(), getReviewContext(ref),
//         onAttempt({ref, args, attempt}) -> {id, end()}, onEdit({ref, change, attempt_id}),
//         onNotApplied({ref, attempt_id}), onBlocked({ref, args, blocked_by}),
//         onOutcomeUnknown({id, ref, record, attempt, attempt_id}),
//         onRecordFailed({ref, attempt_id, change}), onStaleTarget({ref}),
//         onUndoTookMore({ref, paragraph_index}),
//         paneChecksTarget() -> boolean, paneCan(cap) -> boolean,
//         holdRef(ref) -> release, log }
// onAttempt: called before Word, with the edit and the form it would have in
// the document; it writes them to the review file and returns the id of
// that attempt (end(): the call is over). When it throws, Word is not
// called. onEdit settles the attempt in the save that records the change;
// onNotApplied removes it; onOutcomeUnknown keeps it with the call id.
// begin (optional): called once, before anything else. What it returns
// (review-service.mjs: the document of the pane, as it is now) is handed to
// every other function of this call as its last argument or as `at`, also
// when Word answers late: none of them reads the pane's context again.
// holdRef (optional): marks the comment as in use while Word is called, so a
// comment-list save does not drop it; the returned function ends that.
// onOutcomeUnknown: the task pane did not answer. `args` is the edit as it
// would be stored for a blocked one; `id` is the bridge call id
// a late tool_result will carry; `record(result)` records that result the
// same way as an answer in time (review-service.mjs lateToolResult);
// `attempt` is the change as it would read in the document (unknownAttempt).
// getReviewContext returns { requirements, comment_text, paragraphs,
// paragraph_ids, recent_changes, author_texts } for the L2 sources (manuscript snapshot
// texts and the comment), the L5 ids, and what is judged in the paragraph of
// the edit before Word is called (paragraph_ids: the task pane ids of the
// snapshot paragraphs, for paragraph_hint; recent_changes: changes recorded
// since that snapshot, oldest first; author_texts: what the author answered
// to questions of this review, a source of numbers like the comment;
// inserted: what each recorded change not accepted or rejected put in, by
// paragraph, for own_insertions).
// turn: { ref } of the current review turn, or null. A follow-through turn
// ("Check other places", review-service.mjs review_propose) also has
// follow: { candidate_ids, sources }. candidate_ids: the paragraphs it was
// given; an edit of any other paragraph is refused (NotAFollowPlace), also
// one whose paragraph cannot be told from here. sources: what the changes
// of this comment put in and took out. A number that stands there went in
// through this guard or was in the manuscript, so the turn may carry it to
// the heading, legend or methods paragraph that states the same thing; no
// other guard reads it. An edit of such a turn that a forbid rule blocks
// keeps both (onBlocked), for the exception the author may allow later.
// recorded: optional in-memory receipt lists of this live turn. They are
// never copied to a blocked item or to an author's later approval request.
// override: set only by the daemon after the user approved an exception in
// the task pane; it is not part of the tool schema the model sees.
// part: set only by applyTrackedRewrite, for one edit of a paragraph that is
// rewritten as a whole: { rewrite_id, comment_text, comment (true: this part
// carries the rationale comment of the paragraph; false: it adds none and
// is recorded under comment_id), comment_id, plain_only, lost, unsourced,
// fromReview }. The number guard, the forbid list and the fact checks have
// run on the paragraph as a whole there; here they are not run again (a
// number moved from one sentence to another is no new number, though the
// part that puts it in does not hold it).
// The refusal of an edit a follow-through turn may not make. index: the
// paragraph of the edit, or null when it cannot be told; then `why` is the
// code of editTarget and `hint` the paragraph_hint that was given, and the
// message says what to give next.
function notAFollowPlace(deps, turn, index, why = null, hint = null) {
  deps.log?.(`[review] not a follow place ${turn.ref} paragraph=${index ?? "?"}`);
  const what =
    index !== null
      ? "This paragraph is not one of them."
      : why === "NotFound"
        ? "The words to change are in no paragraph as they were given: copy them exactly from a listed place."
        : why === "HintNotFound"
          ? `paragraph_hint ${hint} names no paragraph: give the id of a listed place.`
          : hint
            ? `The words to change occur more than once in ${hint}: give more of the words around them.`
            : "The words to change occur in more than one place: give paragraph_hint with the id of a listed place.";
  return fail(
    `Not applied: this turn checks the other places of ${turn.ref} and may change only the paragraphs listed for it (${turn.follow.candidate_ids.join(", ")}). ${what} Leave any other paragraph as it is.`,
    { code: "NotAFollowPlace", paragraph_index: index },
  );
}

export async function applyTrackedEdit(deps, args, turn, { override = null, part = null } = {}) {
  if (!turn)
    return fail("office_tracked_edit works only in a review turn started from the Review tab.");
  if (args.comment_ref !== turn.ref) {
    return fail(
      `comment_ref ${args.comment_ref} does not match the comment of this turn ${turn.ref}.`,
    );
  }
  const hasReplace = typeof args.replace === "string";
  const hasInsert = typeof args.insert_after === "string";
  if (hasReplace === hasInsert) return fail("Give exactly one of replace or insert_after.");
  const newText = hasReplace ? args.replace : args.insert_after;
  const kind = hasInsert ? "insert" : args.replace === "" ? "delete" : "replace";
  const edit = {
    comment_ref: args.comment_ref,
    find: args.find,
    ...(hasReplace ? { replace: args.replace } : { insert_after: args.insert_after }),
    ...(args.paragraph_hint ? { paragraph_hint: args.paragraph_hint } : {}),
    rationale: args.rationale,
  };

  const at = deps.begin?.();
  const ctx = (await deps.getReviewContext?.(turn.ref, at)) ?? {};
  const known = new Set((ctx.requirements ?? []).map((r) => r.id));
  const unknown = (args.rationale?.requirement_ids ?? []).filter((id) => !known.has(id));
  if (unknown.length) {
    return fail(
      `Unknown requirement id(s) ${unknown.join(", ")} for ${turn.ref}. Valid ids: ${[...known].join(", ") || "(none listed)"}.`,
    );
  }

  // The paragraph Word will edit, when it can be told from here (editTarget).
  const current = currentTexts(ctx.paragraphs, ctx.recent_changes);
  const aim = editTarget(current, ctx.paragraph_ids, args.find, args.paragraph_hint);
  const target = aim.index;
  // A follow-through turn changes the places it was given and no other. The
  // parts of a rewrite were judged with their paragraph (applyTrackedRewrite).
  if (turn.follow && !part) {
    const id = target === null ? null : (ctx.paragraph_ids?.[target] ?? `p${target}`);
    if (!turn.follow.candidate_ids.includes(id))
      return notAFollowPlace(deps, turn, target, aim.code, args.paragraph_hint ?? null);
  }
  // The paragraphs as they are shown: the status refresh gives a field with
  // its instruction (field-text.js), `find` is text as Word shows it.
  const views = fieldViews(current);
  const shown = views.map((v) => v.text);
  const narrow = hasReplace ? narrowEdit(args.find, args.replace) : null;
  // Fields. An edit leaves every field as it was or removes one as a whole:
  // a range inside the text of a field, across its edge, or a replacement
  // over one is refused here, before anything else is judged and before
  // Word is called. A deletion that holds a field whole (a sentence with
  // its citation) goes on, to a task pane that checks the paragraph after
  // the edit. The parts of a rewrite are not judged here: the task pane
  // refuses the paragraph as a whole (ParagraphNotPlain).
  // The fields of the paragraph could not be read at the last refresh (its
  // reviewed text is missing, or that of a paragraph before it, where a
  // field over several paragraphs would begin): nothing here can tell
  // whether the edit cuts through one, so it is refused, a part of a rewrite
  // too. Not when the document is known to hold no field.
  if (target !== null && fieldsUnread(ctx.fields_read, target)) {
    deps.log?.(`[review] fields unread ${turn.ref} paragraph=${target} by=daemon`);
    return fail(FIELDS_UNREAD, { code: "FieldsUnread", reason: "unread", paragraph_index: target });
  }
  if (!part && target !== null) {
    const v = fieldVerdict(
      views[target],
      args.find,
      hasReplace ? { replace: args.replace } : { insert_after: args.insert_after },
      narrow,
    );
    const reason = !v.ok
      ? v.reason
      : v.remove && deps.paneCan && !deps.paneCan(FIELD_CAP, at)
        ? "pane-outdated"
        : null;
    if (reason) {
      deps.log?.(
        `[review] field edit refused ${turn.ref} why=${reason} by=daemon paragraph=${target}`,
      );
      return fail(fieldRefusal(reason), {
        code: "FieldNotEditable",
        reason,
        paragraph_index: target,
      });
    }
  }
  // Forbid list: the new text on its own, and the paragraph as it will read
  // after the edit (a forbidden phrase formed with the words next to it).
  // Without a known target every occurrence of `find` is tried.
  const ledger = (await deps.getLedger?.(at)) ?? [];
  // A forbid rule that cannot be compiled cannot be checked: the edit is
  // refused, not passed. An approved exception does not lift it; the rule
  // is fixed or removed in the Review tab.
  const broken = brokenForbidRules(ledger);
  if (broken.length) {
    const ids = broken.map((b) => b.id);
    deps.log?.(`[review] refused ${turn.ref}: forbid rule(s) ${ids.join(",")} do not compile`);
    return fail(
      `Not applied: forbidden rule(s) ${ids.join(", ")} cannot be checked (invalid regex: ${broken.map((b) => `${b.id} ${b.error}`).join("; ")}). No edit is applied until the author fixes or removes the rule in the Review tab. Do not retry; tell the author.`,
      { code: "ForbidRuleBroken", broken_rules: ids },
    );
  }
  // A number the change adds (number-guard.mjs) must be given: by the text
  // it replaces, the paragraph it lands in, the reviewer comment, an author
  // instruction or an author answer (in a follow-through turn also by a
  // change of the comment: turn.follow.sources). Any other number is a fact only the
  // author can confirm, so the edit is refused and the model asks. Unit
  // notation changes no number; a deletion adds none. An approved forbid
  // exception does not lift this. The log holds the count, not the number.
  const numbers = part
    ? []
    : newNumbers({
        kind,
        find: args.find,
        text: newText,
        sources: [
          ctx.comment_text ?? "",
          ...(ctx.author_texts ?? []),
          ...(turn.follow?.sources ?? []),
          ...ledger.filter((i) => i.kind === "instruction").map((i) => i.text ?? ""),
          ...(target === null ? [] : [current[target]]),
        ],
      });
  if (numbers.length) {
    deps.log?.(`[review] number needs author ${turn.ref} count=${numbers.length}`);
    return fail(
      `Not applied: the change puts a number into the manuscript (${numbers.join(", ")}) that is not in the text it replaces, its paragraph, the reviewer comment, an author instruction or an author answer. Which value is right is a fact only the author can confirm. Leave the text as it is and ask with office_ask_author: say which values disagree and where. Do not try another number.`,
      { code: "NumberNeedsAuthor", numbers },
    );
  }
  const around = forbidHitsAfterEdit(ledger, target === null ? shown : [shown[target]], {
    find: args.find,
    text: newText,
    kind,
  });
  const hits = part ? [] : [...new Set([...forbidHits(ledger, newText), ...around])];
  const excepted = override?.blocked_by ?? [];
  const open = hits.filter((id) => !excepted.includes(id));
  if (open.length) {
    const { candidate_ids, sources } = turn.follow ?? {};
    const blocked = await deps.onBlocked?.({
      ref: turn.ref,
      args: edit,
      blocked_by: open,
      at,
      ...(turn.follow ? { follow: { candidate_ids, sources } } : {}),
    });
    deps.log?.(`[review] blocked ${turn.ref} by=${open.join(",")}`);
    return fail(
      `Not applied: the new text, alone or with the words next to it, matches forbidden rule(s) ${open.join(", ")}. Rephrase without it; the author can approve an exception in the Review tab.`,
      { blocked_by: open, blocked_id: blocked?.id ?? null },
    );
  }

  const rationale = excepted.length
    ? {
        ...args.rationale,
        why: `${args.rationale.why} (금지 규칙 ${excepted.join(", ")} 예외 승인)`,
      }
    : args.rationale;
  // L1-L3 (edit-facts.mjs). Sources for L2: the manuscript snapshot of the
  // last status refresh, the comment, and the text being replaced.
  const texts = ctx.paragraphs ?? [];
  const lost = part
    ? part.lost
    : kind === "insert"
      ? []
      : lostFacts(args.find, args.replace, {
          sentenceStart: startsSentence(args.find, target === null ? texts : [shown[target]]),
        });
  const unsourced = part
    ? part.unsourced
    : kind === "delete"
      ? []
      : unsourcedTerms(newText, [...texts, ctx.comment_text ?? "", args.find]);
  // Facts of the new text (quantities, names: edit-facts.mjs extractFacts)
  // that the review has and the manuscript does not: taken from the
  // reviewer. They pass as sourced above; the rationale comment tells the
  // author they were not checked against a source. A plain word the reviewer
  // suggested is wording, not a fact, and is not marked.
  const own = [...texts, args.find].join("\n");
  const fromReview = part
    ? part.fromReview
    : kind === "delete"
      ? []
      : [
          ...new Set(
            extractFacts(newText)
              .map((f) => f.text)
              .filter((t) => !factPresent(t, own) && factPresent(t, ctx.comment_text ?? "")),
          ),
        ];

  const commentText =
    part?.comment_text ?? formatRationaleComment(turn.ref, rationale, { lost, fromReview });
  // The edit as it would read in the document (unknownAttempt). It goes into
  // the review file BEFORE Word is called (onAttempt), so the stretch between
  // "Word applied it" and "the file says so" is never empty: a daemon that
  // stops there, or a record that cannot be saved, leaves the attempt for
  // the next refresh. id: the bridge call id once a call got no answer.
  const attemptOf = (id) => ({
    tool_call_id: id,
    unknown_at: new Date().toISOString(),
    comment_text: commentText,
    change: unknownAttempt(shown, target, ctx.paragraph_ids, {
      kind,
      find: args.find,
      text: newText,
      rationale,
      override_of: excepted,
      lost_facts: lost,
      unsourced_terms: unsourced,
      from_review: fromReview,
      ...(part ? { rewrite_id: part.rewrite_id } : {}),
      // A part without a comment of its own: the comment of its paragraph.
      ...(part && !part.comment ? { word_comment_id: part.comment_id } : {}),
      // The narrower form the task pane applies when it can (L3). Word's
      // answer says which one went in; when that answer is never saved, a
      // refresh looks for both (narrowedAttempt).
      ...(narrow ? { narrow } : {}),
    }),
  });
  // A task pane that does not compare the paragraph before the edit (it was
  // loaded before the daemon was updated; its hello names no target_hash)
  // would apply the edit unchecked. It is sent nothing. A pane that is not
  // connected reads the same here; the answer holds for both.
  if (deps.paneChecksTarget && !deps.paneChecksTarget(at)) {
    deps.log?.(`[review] refused ${turn.ref}: task pane without ${TARGET_HASH_CAP}`);
    return fail(
      "Not applied: the task pane is older than the daemon (or not connected) and does not check the paragraph before an edit, so nothing was sent to Word. Tell the author to reload the task pane (close and reopen it); then propose the edit again.",
      { code: "TaskPaneOutdated" },
    );
  }
  let pre;
  try {
    pre =
      (await deps.onAttempt?.({ ref: turn.ref, args: edit, attempt: attemptOf(null), at })) ?? null;
  } catch (e) {
    deps.log?.(`[review] attempt not recorded ${turn.ref}: ${e?.message ?? e}`);
    return fail(
      `Not applied: the edit could not be written to the review file first (${e?.message ?? e}), so nothing was sent to Word. Tell the author; retry only after that is fixed.`,
      { code: "AttemptNotRecorded" },
    );
  }
  const attemptId = pre?.id ?? null;
  // Word answered that the edit is not in the document: its attempt goes. A
  // failed removal only leaves it for the refresh, which drops it.
  const notApplied = () =>
    Promise.resolve(deps.onNotApplied?.({ ref: turn.ref, attempt_id: attemptId, at })).catch((e) =>
      deps.log?.(`[review] attempt ${attemptId} not removed: ${e?.message ?? e}`),
    );

  // The answer of the task pane, now or late: record it (or return its error).
  const record = async (r) => {
    // No result at all says nothing about the document: the attempt stays
    // and reads as unanswered once this call is over.
    if (!r) return fail("empty result from the task pane", { code: "EmptyResult" });
    // applied:true with ok:false: the task pane changed the document but could
    // not add the rationale comment (CommentAnchorMissing: no range to anchor
    // it; CommentInsertFailed: a step after the edit threw;
    // CommentWaitingForWord: Word refused the comment insert with
    // AccessDenied, minimized or not ready, and the task pane adds it when
    // Word takes it). The change is
    // still recorded, with no comment id, so the change list and the
    // "changes == rationale comments" check show it instead of losing it.
    const commentMissing = r.ok === false && r.applied === true ? (r.code ?? "applied") : null;
    if (r.ok === false && !commentMissing) {
      await notApplied();
      // The task pane refused the edit for a field (before the edit, or
      // after it, when Word did not keep a field whole and the edit was
      // taken back). The log says which; the model gets the pane's message.
      if (r.code === "FieldNotEditable") {
        deps.log?.(
          `[review] field edit refused ${turn.ref} why=${r.reason ?? "?"} by=pane paragraph=${r.paragraph_index ?? "?"}` +
            `${r.detail ? ` found=${String(r.detail).replace(/\s+/g, "_")}` : ""}${r.undone ? " undone=1" : ""}` +
            `${r.took_more ? " took_more=1" : ""}`,
        );
        // The undo took a pending change next to the edit with it (the task
        // pane read the paragraph before the edit, after it and after the
        // undo). The edit is not in the document and nothing is recorded;
        // the change that vanished is on record as pending, so the task
        // pane is asked for a status refresh, which reads it from the
        // document again, and the author is told on the card of the comment.
        if (r.took_more) {
          await Promise.resolve(
            deps.onUndoTookMore?.({
              ref: turn.ref,
              paragraph_index: r.paragraph_index ?? null,
              at,
            }),
          ).catch(() => {});
        }
      }
      // The deletion would have taken the whole of an insertion the author
      // made and has not accepted: Word keeps no mark of that. The model is
      // told to ask, with the paragraph as it knows it.
      if (r.code === "WholeAuthorInsertion") {
        const index = r.paragraph_index ?? target;
        deps.log?.(`[review] whole author insertion ${turn.ref} paragraph=${index ?? "?"}`);
        const id =
          index === null || index === undefined
            ? null
            : (ctx.paragraph_ids?.[index] ?? `p${index}`);
        return fail(
          `${WHOLE_AUTHOR_INSERTION}${id ? ` Paragraph: ${id} (give it as paragraph_id).` : ""}`,
          { code: "WholeAuthorInsertion", applied: false, paragraph_index: index ?? null },
        );
      }
      if (r.code !== "StaleTarget") return r;
      // The task pane found the paragraph not as it was judged here and
      // changed nothing. The judgement (forbid list, lost facts) was made
      // on a text that is gone, so nothing is recorded; the task pane is
      // asked for a status refresh (onStaleTarget), which brings the
      // paragraph as it is now, and the model proposes again on that.
      deps.log?.(
        `[review] stale target ${turn.ref} paragraph=${r.paragraph_index ?? "?"} expected=${target ?? "unknown"}`,
      );
      await Promise.resolve(deps.onStaleTarget?.({ ref: turn.ref, at })).catch(() => {});
      return fail(
        target === null
          ? `Not applied: the paragraph of "${args.find}" could not be told from the last status refresh, so the edit was not checked and Word was left unchanged. Read the paragraph again and propose the edit with a find that occurs once (or a paragraph_hint). The task pane is refreshing the status; if this repeats, tell the author to press Refresh status.`
          : `Not applied: the paragraph changed in Word after it was last read, so the edit was checked against a text that is no longer there. Word was left unchanged. Read the paragraph again and propose the edit against its current text. The task pane is refreshing the status; if this repeats, tell the author to press Refresh status.`,
        { code: "StaleTarget", paragraph_index: r.paragraph_index ?? null },
      );
    }
    const out = await recordApplied(deps, turn.ref, r, {
      at,
      attemptId,
      texts: shown,
      kind,
      newText,
      narrow,
      rationale,
      excepted,
      lost,
      unsourced,
      fromReview,
      part,
      find: args.find,
      commentMissing,
      // change: the change as Word applied it (narrowed or not, with its
      // comment id). The attempt on file was written before the call and
      // does not know either, so the change is handed over for a later save.
      // The task pane is told (its status line). The call is over first, so
      // the attempt reads as unowned when the pane loads the file again.
      onFailed: async (change) => {
        pre?.end?.();
        await Promise.resolve(
          deps.onRecordFailed?.({ ref: turn.ref, attempt_id: attemptId, change, at }),
        ).catch(() => {});
      },
    });
    if (out.change_id) turn.follow?.recorded?.changes.push(out.change_id);
    return out;
  };

  const ownTexts = kind === "insert" ? [] : ownInsertions(ctx.inserted, target, ctx.paragraph_ids);
  // The comment stays in the list until the answer is recorded (or handed
  // to onOutcomeUnknown, which holds it on its own).
  const release = deps.holdRef?.(turn.ref, at);
  try {
    let r;
    try {
      r = await deps.call(REVIEW_WRITE_TOOL, {
        find: args.find,
        ...(hasReplace ? { replace: args.replace } : { insert_after: args.insert_after }),
        ...(args.paragraph_hint ? { paragraph_hint: args.paragraph_hint } : {}),
        ...(narrow ? { narrow } : {}),
        comment_text: commentText,
        // A part of a rewrite: no comment of its own, or (the first one
        // sent) only for a paragraph of plain text.
        ...(part && !part.comment ? { no_comment: true } : {}),
        ...(part?.plain_only ? { plain_only: true } : {}),
        // The paragraph everything above was judged on: its text as it
        // reads now here (reviewed view: the snapshot with the recorded
        // changes in it), hashed. The task pane hashes the paragraph it is
        // about to edit the same way and changes nothing when they differ
        // (StaleTarget). Without a known target no hash goes out and the
        // task pane refuses any paragraph it finds: nothing was judged.
        // Comment ids already on record: the task pane's second try for the
        // comment reuses none of them (tools-word.js existingComment).
        ...(ctx.comment_ids?.length ? { known_comment_ids: ctx.comment_ids } : {}),
        // What the tool itself put into the paragraph: a deletion may take
        // such an insertion as a whole, and no other (WholeAuthorInsertion).
        ...(ownTexts.length ? { own_insertions: ownTexts } : {}),
        ...(target === null
          ? {}
          : {
              expected_before_hash: paragraphHash(current[target]),
              expected_paragraph: { index: target, id: ctx.paragraph_ids?.[target] ?? null },
            }),
      });
    } catch (e) {
      // An error the task pane sent back: its edit batch failed, nothing was
      // applied (tools-word.js toolTrackedEdit reports every failure after
      // the edit as a result with applied:true, never as an error).
      if (!NO_ANSWER_CODES.has(e?.code)) {
        await notApplied();
        throw e;
      }
      // No answer (bridge.mjs timeout or pane disconnect): the edit may be in
      // the document with nothing recorded. A retry could apply it twice.
      const id = e.toolCallId ?? null;
      deps.log?.(`[review] outcome unknown ${turn.ref} code=${e.code} id=${id}`);
      try {
        await deps.onOutcomeUnknown?.({
          id,
          ref: turn.ref,
          record,
          attempt: attemptOf(id),
          attempt_id: attemptId,
          args: edit,
          at,
        });
      } catch (err) {
        // The attempt written before the call is still on file.
        deps.log?.(`[review] outcome unknown ${turn.ref} not saved: ${err?.message ?? err}`);
      }
      return fail(
        `The task pane did not answer (${e.message}). The edit may already be in the document but was not recorded. Do not repeat it: read the paragraph first, and tell the author to check it.`,
        { code: "OutcomeUnknown", tool_call_id: id },
      );
    }
    return await record(r);
  } finally {
    pre?.end?.();
    release?.();
  }
}

export const TRACKED_REWRITE_DESCRIPTION =
  "Review turns only. Rewrite ONE paragraph as a whole: give its paragraph id and the full new text of the paragraph. " +
  "Use it when three or more places of a paragraph have to change or its argument changes; for one or two words use office_tracked_edit. " +
  "Only the words that differ are put into the document as tracked changes, with one rationale comment for the paragraph. " +
  "Start from the paragraph as it reads now (pending changes of other comments included) and keep every sentence you do not mean to change exactly as it is. " +
  "A paragraph with fields, pictures or footnotes is refused (ParagraphNotPlain): use office_tracked_edit there.";

export const trackedRewriteShape = {
  comment_ref: commentRef('Reviewer comment the rewrite answers, e.g. "R1.3".'),
  paragraph_id: z
    .string()
    .describe("Id of the paragraph, as the manuscript in the prompt gives it."),
  text: z.string().min(1).describe("The whole paragraph as it should read."),
  what: z.string().min(1).describe("What changed in the paragraph."),
  why: z.string().min(1).describe("Why."),
  addresses: z.string().min(1).describe("Which part of the reviewer comment this answers."),
  requirement_ids: z
    .array(z.string())
    .optional()
    .describe("Ids of the listed requests the rewrite handles (q1, q2)."),
};

// office_tracked_rewrite. deps: those of applyTrackedEdit, and paneCan(cap)
// -> boolean. The paragraph as it reads now (the last status refresh with
// the changes recorded since) and the new text are compared word by word
// (word-diff.mjs); each run of words that differs is one office_tracked_edit
// through applyTrackedEdit, from the last to the first, so every one of them
// is written to the review file before Word is called, is checked against
// the paragraph as it reads then (expected_before_hash) and is recorded on
// its own. What is judged on text is judged once, on the paragraph as a
// whole, before the first of them: the number guard (sources: the comment,
// author instructions and answers, and the paragraph itself), the forbid
// list (a rule the paragraph already broke and still breaks does not count;
// a hit is refused here and is not kept for an exception, which only a
// single edit can get), facts that are lost, words without a source. The
// first edit sent carries the one rationale comment of the paragraph and
// asks the task pane to refuse a paragraph that is not plain text; the
// others add no comment and are recorded under that comment. When Word
// refuses a part, the rewrite stops there: what went in stays on record.
export async function applyTrackedRewrite(deps, args, turn) {
  const t = await noteTurn(deps, "office_tracked_rewrite", args, turn);
  if (t.error) return t.error;
  const { ctx, at } = t;
  const known = new Set((ctx.requirements ?? []).map((r) => r.id));
  const unknown = (args.requirement_ids ?? []).filter((id) => !known.has(id));
  if (unknown.length) {
    return fail(
      `Unknown requirement id(s) ${unknown.join(", ")} for ${turn.ref}. Valid ids: ${[...known].join(", ") || "(none listed)"}.`,
    );
  }
  if (deps.paneCan && !deps.paneCan(REWRITE_CAP, at)) {
    return fail(
      "Not applied: the task pane is older than the daemon (or not connected) and cannot take a rewritten paragraph. Use office_tracked_edit, or tell the author to reload the task pane.",
      { code: "TaskPaneOutdated" },
    );
  }
  const current = currentTexts(ctx.paragraphs, ctx.recent_changes);
  const ids = ctx.paragraph_ids ?? [];
  const byId = ids.some((id) => id != null) ? ids.indexOf(args.paragraph_id) : -1;
  const m = /^p(\d+)$/.exec(args.paragraph_id);
  const index = byId !== -1 ? byId : m && Number(m[1]) < current.length ? Number(m[1]) : -1;
  if (index === -1) {
    return fail(`No paragraph ${args.paragraph_id} in the manuscript of the last status refresh.`, {
      code: "ParagraphNotFound",
    });
  }
  if (turn.follow && !turn.follow.candidate_ids.includes(ids[index] ?? `p${index}`))
    return notAFollowPlace(deps, turn, index);
  // Fields, before anything is cut: a paragraph whose fields could not be
  // read (as in applyTrackedEdit), and a paragraph of a field that runs over
  // several paragraphs (a reference list). The task pane refuses a paragraph
  // that holds a field of its own (ParagraphNotPlain); the middle paragraph
  // of such a list holds no marker at all, so only the carry from the
  // paragraphs before it tells, and that is known here.
  const refusedForFields = (code, message) => {
    deps.log?.(`[review] rewrite ${turn.ref} paragraph=${index} parts=0 applied=0 stopped=${code}`);
    return fail(message, { code, applied_parts: 0, parts: 0, paragraph_index: index });
  };
  if (fieldsUnread(ctx.fields_read, index)) return refusedForFields("FieldsUnread", FIELDS_UNREAD);
  const views = fieldViews(current);
  if ((index > 0 && views[index - 1].carry > 0) || views[index].carry > 0)
    return refusedForFields(
      "ParagraphNotPlain",
      "Not applied: the paragraph is part of a field that runs over several paragraphs (a reference list written by the reference manager), so it is not rewritten as a whole. Nothing was changed. Do not edit it with office_tracked_edit either: a field is written by the reference manager.",
    );
  const before = current[index];
  const after = normalizeParaText(args.text);
  const cut = rewriteSegments(before, after);
  if (cut.error) {
    return fail(
      `Not applied: ${cut.error}. Change this paragraph with office_tracked_edit, one place at a time.`,
      { code: "RewriteNotExpressible" },
    );
  }
  if (!cut.segments.length) {
    return fail(`Nothing to change: paragraph ${args.paragraph_id} already reads the same.`);
  }
  const ledger = (await deps.getLedger?.(at)) ?? [];
  const broken = brokenForbidRules(ledger);
  if (broken.length) {
    const bad = broken.map((b) => b.id);
    deps.log?.(`[review] refused ${turn.ref}: forbid rule(s) ${bad.join(",")} do not compile`);
    return fail(
      `Not applied: forbidden rule(s) ${bad.join(", ")} cannot be checked. No edit is applied until the author fixes or removes the rule in the Review tab. Do not retry; tell the author.`,
      { code: "ForbidRuleBroken", broken_rules: bad },
    );
  }
  const had = forbidHits(ledger, before);
  const hits = forbidHits(ledger, after).filter((id) => !had.includes(id));
  if (hits.length) {
    deps.log?.(`[review] rewrite blocked ${turn.ref} by=${hits.join(",")}`);
    return fail(
      `Not applied: the new paragraph matches forbidden rule(s) ${hits.join(", ")}. Write it without that; an exception can be approved only for a single office_tracked_edit.`,
      { code: "RewriteBlocked", blocked_by: hits },
    );
  }
  const numbers = newNumbers({
    kind: "replace",
    find: before,
    text: after,
    sources: [
      ctx.comment_text ?? "",
      ...(ctx.author_texts ?? []),
      ...(turn.follow?.sources ?? []),
      ...ledger.filter((i) => i.kind === "instruction").map((i) => i.text ?? ""),
    ],
  });
  if (numbers.length) {
    deps.log?.(`[review] number needs author ${turn.ref} count=${numbers.length}`);
    return fail(
      `Not applied: the new paragraph holds a number (${numbers.join(", ")}) that is not in the paragraph as it reads now, the reviewer comment, an author instruction or an author answer. Which value is right is a fact only the author can confirm. Leave it as it is and ask with office_ask_author. Do not try another number.`,
      { code: "NumberNeedsAuthor", numbers },
    );
  }
  const texts = ctx.paragraphs ?? [];
  const lost = lostFacts(before, after, { sentenceStart: true });
  const unsourced = unsourcedTerms(after, [...texts, ctx.comment_text ?? "", before]);
  const own = [...texts, before].join("\n");
  const fromReview = [
    ...new Set(
      extractFacts(after)
        .map((f) => f.text)
        .filter((x) => !factPresent(x, own) && factPresent(x, ctx.comment_text ?? "")),
    ),
  ];
  const rationale = {
    what: args.what,
    why: args.why,
    addresses: args.addresses,
    ...(args.requirement_ids?.length ? { requirement_ids: args.requirement_ids } : {}),
  };
  const commentText = formatRationaleComment(turn.ref, rationale, { lost, fromReview });
  const rewriteId = `${turn.ref}-w${Date.now().toString(36)}`;
  const parts = cut.segments;
  const changeIds = [];
  let commentId = null;
  let stopped = null;
  // From the last to the first: an edit never moves one that is still to come.
  for (let k = parts.length - 1; k >= 0; k--) {
    const first = k === parts.length - 1;
    const r = await applyTrackedEdit(
      deps,
      {
        comment_ref: args.comment_ref,
        find: parts[k].find,
        replace: parts[k].replace,
        paragraph_hint: args.paragraph_id,
        rationale,
      },
      turn,
      {
        part: {
          rewrite_id: rewriteId,
          comment_text: commentText,
          comment: first,
          comment_id: commentId,
          plain_only: first,
          lost: first ? lost : [],
          unsourced: first ? unsourced : [],
          fromReview: first ? fromReview : [],
        },
      },
    );
    if (r?.change_id) changeIds.push(r.change_id);
    if (first && r?.word_comment_id != null) commentId = String(r.word_comment_id);
    if (!r?.ok) {
      stopped = r ?? { code: "EmptyResult", message: "no answer" };
      break;
    }
  }
  const applied = changeIds.length;
  deps.log?.(
    `[review] rewrite ${turn.ref} paragraph=${index} parts=${parts.length} applied=${applied}` +
      `${stopped ? ` stopped=${stopped.code ?? "refused"}` : ""}`,
  );
  if (stopped) {
    const now = currentTexts(
      ctx.paragraphs,
      ((await deps.getReviewContext?.(turn.ref, at)) ?? ctx).recent_changes,
    )[index];
    if (!applied) {
      return fail(
        `Not applied: ${stopped.message ?? stopped.code ?? "the task pane refused the first part"} Nothing of the rewrite is in the document.${stopped.code === "ParagraphNotPlain" ? " Change this paragraph with office_tracked_edit, one place at a time." : ""}`,
        { code: stopped.code ?? "RewriteStopped", applied_parts: 0, parts: parts.length },
      );
    }
    // Marked, not undone: an edit that takes the parts back would be one
    // more tracked change on a paragraph Word already found different. The
    // card of the comment shows the record until those changes are rejected.
    try {
      await deps.onRewriteStopped?.({
        ref: turn.ref,
        record: {
          rewrite_id: rewriteId,
          paragraph_index: index,
          paragraph_id: args.paragraph_id,
          applied,
          parts: parts.length,
          code: stopped.code ?? "refused",
          change_ids: changeIds,
        },
        at,
      });
    } catch (e) {
      deps.log?.(`[review] rewrite stop not recorded ${turn.ref}: ${e?.message ?? e}`);
    }
    return fail(
      `The rewrite stopped: ${applied} of ${parts.length} parts are in the document (changes ${changeIds.join(", ")}), then a part was refused (${stopped.code ?? stopped.message ?? "refused"}). Do not send the rewrite again. The paragraph now reads as given in paragraph_now: finish it with office_tracked_edit, or tell the author.`,
      {
        code: "RewriteStopped",
        applied_parts: applied,
        parts: parts.length,
        change_ids: changeIds,
        paragraph_now: now,
      },
    );
  }
  return {
    ok: true,
    change_ids: changeIds,
    parts: parts.length,
    paragraph_index: index,
    word_comment_id: commentId,
    lost_facts: lost,
    unsourced_terms: unsourced,
    from_review: fromReview,
    ...(unsourced.length || fromReview.length
      ? {
          warning: [
            unsourced.length
              ? `Not in the manuscript or the comment: ${unsourced.join(", ")}. Name them in your answer so the author can check where they come from.`
              : "",
            fromReview.length
              ? `From the review, not checked against the source: ${fromReview.join(", ")}.`
              : "",
          ]
            .filter(Boolean)
            .join(" "),
        }
      : {}),
  };
}

// The paragraph a tracked edit lands in, by the rule of toolTrackedEdit
// (taskpane/shared/tools-word.js): with one hit of `find` in the document
// that hit, the hint unread; with several hits, the paragraph the hint names
// when it holds exactly one. tests/edit-target.test.mjs gives both sides the
// same documents and compares the answers.
// texts: paragraph texts in document order; ids: their task pane ids (null,
// or null per paragraph, in a document with p<N> index ids).
// Returns { index, code }: code is null with an index, else the error Word
// answers with (NotFound, Ambiguous, HintNotFound) and index is null.
// A paragraph with a field is searched as Word shows it, without the
// instruction of the field (shownText): Word's search finds a sentence
// across its citation, and so does this.
export function editTarget(texts, ids, find, hint) {
  const none = (code) => ({ index: null, code });
  const f = String(find ?? "")
    .replace(/\u0005/g, "")
    .replace(/\s+/g, " ");
  if (!f.trim()) return none("NotFound");
  const counts = (texts ?? []).map((t) => shownText(t).split(f).length - 1);
  const total = counts.reduce((a, b) => a + b, 0);
  let index;
  let hits = total;
  if (total > 1 && hint) {
    if ((ids ?? []).some((id) => id != null)) index = ids.indexOf(hint);
    else {
      const m = /^p(\d+)$/.exec(hint);
      index = m && Number(m[1]) < counts.length ? Number(m[1]) : -1;
    }
    if (index === -1) return none("HintNotFound");
    hits = counts[index];
  } else {
    index = counts.findIndex((n) => n > 0);
  }
  if (hits === 0) return none("NotFound");
  if (hits > 1) return none("Ambiguous");
  return { index, code: null };
}

// A paragraph as Word shows it: one with a field (the status refresh gives
// it with the instruction of the field, field-text.js) without the
// instruction and the marks; any other as it is.
const shownText = (t) => (hasFieldMarks(t) ? fieldView(t).text : String(t ?? ""));

// The paragraphs as they read now: the texts of the last status refresh with
// the changes recorded since then put in (a narrowed change by the find and
// replace the model asked for). A change whose find is not exactly once in
// the texts is left out. In a paragraph with a field the find is looked for
// in the text as it is shown, and a deletion that held a field whole takes
// the field out with its instruction, as the task pane reads the paragraph
// after it.
export function currentTexts(texts, changes) {
  const out = (texts ?? []).map(normalizeParaText);
  for (const ch of changes ?? []) {
    const f = String(ch.narrowed_from?.find ?? ch.find ?? "")
      .replace(/\u0005/g, "")
      .replace(/\s+/g, " ");
    if (!f.trim()) continue;
    // The paragraph Word reported for the change; without one, the only
    // paragraph that holds its find.
    const at = out.flatMap((t, k) => (shownText(t).includes(f) ? [k] : []));
    const k = at.includes(ch.paragraph_index) ? ch.paragraph_index : at.length === 1 ? at[0] : null;
    if (k === null) continue;
    const t = out[k];
    if (hasFieldMarks(t)) {
      const next = spliceShown(
        t,
        f,
        ch.narrowed_from?.text ?? ch.text ?? "",
        !ch.narrowed_from && ch.kind === "insert",
      );
      if (next !== null) out[k] = next;
      continue;
    }
    const i = t.indexOf(f);
    if (t.indexOf(f, i + 1) !== -1) continue;
    const cut = !ch.narrowed_from && ch.kind === "insert" ? i + f.length : i;
    out[k] = t.slice(0, cut) + (ch.narrowed_from?.text ?? ch.text ?? "") + t.slice(i + f.length);
  }
  return out;
}

// The change an unanswered tracked edit would have recorded, with the
// paragraph text it expects before and after: texts are the paragraphs as
// they read now (currentTexts), target the index of the one Word edits
// (editTarget) or null. Without a target that holds `find` once, before and
// after stay empty and context null: a refresh then cannot tell whether the
// edit is in the document (review-service.mjs resolveUnknownOverrides).
export function unknownAttempt(texts, target, ids, change) {
  const t = target === null ? null : texts[target];
  const ctx = t == null ? null : changeContext(t, change.find, { kind: change.kind });
  const one = ctx ? { index: target, t, ctx } : null;
  let after = "";
  if (one) {
    const f = String(change.find)
      .replace(/\u0005/g, "")
      .replace(/\s+/g, " ");
    const i = one.t.indexOf(f);
    const cut = change.kind === "insert" ? i + f.length : i;
    after = normalizeParaText(one.t.slice(0, cut) + change.text + one.t.slice(i + f.length));
  }
  return {
    ...change,
    word_comment_id: change.word_comment_id ?? null,
    paragraph_index: one ? one.index : null,
    paragraph_id: one ? (ids?.[one.index] ?? null) : null,
    before: one ? one.t : "",
    duplicate_before: one ? hasTwin(texts, one.t) : false,
    after,
    narrowed_from: null,
    context: one ? one.ctx : null,
  };
}

// The narrower form of an attempt (attempt.change.narrow), as recordApplied
// records an edit the task pane narrowed: the cut find and replacement, the
// words around the cut range, and the full call under narrowed_from. null
// when the attempt has none, or its place in `before` cannot be told.
export function narrowedAttempt(change) {
  const n = change?.narrow;
  if (change?.kind !== "replace" || typeof n?.find !== "string" || !n.find) return null;
  const context = changeContext(change.before, change.find, {
    kind: change.kind,
    narrowFind: n.find,
  });
  if (!context) return null;
  const { narrow, ...rest } = change;
  const text = String(narrow.replace ?? "");
  return {
    ...rest,
    kind: text === "" ? "delete" : "replace",
    find: narrow.find,
    text,
    narrowed_from: { find: change.find, text: change.text },
    context,
  };
}

// Is the paragraph text `before` there more than once?
const hasTwin = (texts, before) => {
  const b = normalizeParaText(before);
  return !!b && (texts ?? []).filter((t) => normalizeParaText(t) === b).length > 1;
};

// One failed comment step of the task pane (tools-word.js commentError) as
// the words of a log line: step, Word's code and place, message, and what
// Word said more (debugInfo message, inner error) and whether the task pane
// was hidden, when the task pane sent them. No text of the manuscript.
export function commentErrorWords(e) {
  const one = (v, n) => String(v).replace(/\s+/g, " ").slice(0, n);
  return (
    `step=${e?.step ?? "-"} word=${e?.code ?? "-"} at=${e?.location ?? "-"} msg=${one(e?.message ?? "-", 200)}` +
    (e?.debug_message ? ` debug=${one(e.debug_message, 200)}` : "") +
    (e?.inner ? ` inner=${one(e.inner, 100)}` : "") +
    (typeof e?.hidden === "boolean" ? ` hidden=${e.hidden ? 1 : 0}` : "")
  );
}

// Record a result the task pane reported as applied (ok, or applied without
// its rationale comment) and build the tool answer.
async function recordApplied(
  deps,
  ref,
  r,
  {
    at,
    attemptId,
    texts,
    kind,
    newText,
    narrow,
    rationale,
    excepted,
    lost,
    unsourced,
    fromReview = [],
    part = null,
    find,
    commentMissing,
    onFailed,
  },
) {
  // The task pane reports whether it could apply the narrowed range (the cut
  // find must occur once inside the located hit); otherwise the full range.
  const narrowed = !!(narrow && r.narrowed === true);
  const appliedKind = narrowed ? (narrow.replace === "" ? "delete" : "replace") : kind;
  const change = {
    // A part of a rewrite that added no comment is under the comment of its paragraph.
    word_comment_id:
      part && !part.comment
        ? (part.comment_id ?? null)
        : r.word_comment_id == null
          ? null
          : String(r.word_comment_id),
    ...(part ? { rewrite_id: part.rewrite_id } : {}),
    paragraph_index: r.paragraph_index ?? null,
    paragraph_id: r.paragraph_id ?? null,
    kind: appliedKind,
    find: narrowed ? narrow.find : find,
    text: narrowed ? narrow.replace : newText,
    before: r.before ?? "",
    after: r.after ?? "",
    rationale,
    override_of: excepted,
    lost_facts: lost,
    unsourced_terms: unsourced,
    from_review: fromReview,
    narrowed_from: narrowed ? { find: find, text: newText } : null,
    comment_missing: commentMissing,
    // Why it is missing, when Word said: the last failed try of the task
    // pane (commentErrorWords), so the review file tells it too.
    ...(commentMissing && Array.isArray(r.comment_errors) && r.comment_errors.length
      ? { comment_error: commentErrorWords(r.comment_errors.at(-1)) }
      : {}),
    duplicate_before: hasTwin(texts, r.before),
    // A deletion or a replace whose old words lay inside a tracked
    // insertion, when the task pane could tell (tools-word.js
    // insertionAround); left out when it could not. Only a deletion is read
    // with it (revision-classify.mjs inInsertion); a replace has its own
    // tracked new words, and the flag is there for the note on its card.
    ...(appliedKind !== "insert" && typeof r.in_insertion === "boolean"
      ? { in_insertion: r.in_insertion }
      : {}),
    // What of that insertion stays on each side of the old words.
    ...(appliedKind !== "insert" &&
    r.in_insertion === true &&
    typeof r.insertion_left === "string" &&
    typeof r.insertion_right === "string"
      ? { insertion_left: r.insertion_left, insertion_right: r.insertion_right }
      : {}),
    // Words around the range the task pane changed (revision-classify.mjs).
    context: changeContext(r.before, find, {
      kind,
      narrowFind: narrowed ? narrow.find : null,
    }),
  };
  let saved;
  try {
    // One save: the change, and its attempt out of the blocked list.
    saved = await deps.onEdit?.({ ref: ref, change, at, attempt_id: attemptId });
  } catch (e) {
    // The edit is in the document and the file does not say so. Not thrown:
    // an error would read as "not applied" and invite a second apply. The
    // attempt written before the call stays; a refresh records the change.
    deps.log?.(`[review] applied, record failed ${attemptId ?? ref}: ${e?.message ?? e}`);
    await onFailed?.(change);
    return {
      ok: false,
      applied: true,
      code: "AppliedNotRecorded",
      attempt_id: attemptId,
      kind: appliedKind,
      word_comment_id: change.word_comment_id,
      paragraph_index: change.paragraph_index,
      before: r.before,
      after: r.after,
      message: `The edit was applied in Word but could not be written to the review file (${e?.message ?? e}). Do not repeat it: it is in the document. Tell the author to press Refresh status in the Review tab, which records it${attemptId ? ` from attempt ${attemptId}` : ""}.`,
    };
  }
  const id = saved?.id ?? "?";
  deps.log?.(
    `[review] change recorded ${id} kind=${appliedKind} comment=${change.word_comment_id} paragraph=${change.paragraph_index}` +
      ` narrowed=${narrowed ? 1 : 0} lost=${lost.length} unsourced=${unsourced.length}`,
  );
  if (lost.length) deps.log?.(`[review] lost facts ${id}: ${lost.join("; ")}`);
  if (unsourced.length) deps.log?.(`[review] unsourced ${id}: ${unsourced.join("; ")}`);
  if (commentMissing) deps.log?.(`[review] comment missing ${id} code=${commentMissing}`);
  // The task pane found a field of the paragraph not whole after the edit
  // and could not take the edit back (tools-word.js). The change is on
  // record, so the author sees it; the model is told to leave it alone.
  const fieldBroken = commentMissing === "FieldEditNotUndone";
  if (fieldBroken) {
    deps.log?.(
      `[review] field edit not undone ${id} paragraph=${change.paragraph_index} found=${String(r.detail ?? "?").replace(/\s+/g, "_")}`,
    );
  }
  // What Word answered when the comment step failed (tools-word.js
  // afterEditApplied): one line per failed try, no text of the manuscript;
  // the last one also on the change (comment_error) when no try worked.
  // Without it nothing says why a change has no rationale comment.
  const commentErrors = Array.isArray(r.comment_errors) ? r.comment_errors : [];
  commentErrors.forEach((e, n) =>
    deps.log?.(`[review] comment error ${id} try=${n + 1} ${commentErrorWords(e)}`),
  );
  if (!commentMissing && commentErrors.length) {
    deps.log?.(`[review] comment added on try ${r.comment_attempts ?? "?"} ${id}`);
  }
  // A task pane that compares the paragraph before the edit says so
  // (tools-word.js target_checked). One that does not is older than this
  // daemon: its edit is in the document and is recorded, but it was not
  // held against the paragraph the checks above were made on.
  const unchecked = r.target_checked !== true;
  if (unchecked) deps.log?.(`[review] target not checked ${id}: task pane older than the daemon`);
  const warnings = [
    ...(unchecked
      ? [
          `The task pane applied this change but did not compare the paragraph with the one the edit was checked against (it is older than the daemon). Tell the author to reload the task pane and to look at change ${id}.`,
        ]
      : []),
    ...(fieldBroken
      ? [
          `The change is in the document, but it did not leave a citation or another field of its paragraph whole (${r.detail ?? "field changed"}) and Word could not take it back. Do not repeat it and do not try to repair it with another edit; tell the author to reject change ${id} in Word. It has no rationale comment.`,
        ]
      : commentMissing === "CommentWaitingForWord"
        ? [
            `The change was applied; Word could not take its rationale comment now (the Word window is minimized or not ready). The task pane adds it when the Word window is open. Do not repeat this edit; tell the author that the rationale comment of change ${id} is waiting for the Word window to be opened.`,
          ]
        : commentMissing
          ? [
              `The change was applied but no rationale comment could be anchored on it (${commentMissing}). Do not repeat this edit; tell the author that change ${id} has no rationale comment.`,
            ]
          : []),
    ...(unsourced.length
      ? [
          `Not in the manuscript or the comment: ${unsourced.join(", ")}. Name them in your answer so the author can check where they come from.`,
        ]
      : []),
    ...(fromReview.length
      ? [
          `From the review, not checked against the source: ${fromReview.join(", ")}. The rationale comment says so; name them in your answer.`,
        ]
      : []),
  ];
  return {
    ok: true,
    ...(commentMissing ? { applied: true, comment_missing: commentMissing } : {}),
    change_id: saved?.id ?? null,
    kind: appliedKind,
    word_comment_id: change.word_comment_id,
    paragraph_index: change.paragraph_index,
    before: r.before,
    after: r.after,
    applied_find: change.find,
    lost_facts: lost,
    unsourced_terms: unsourced,
    from_review: fromReview,
    ...(warnings.length ? { warning: warnings.join(" ") } : {}),
  };
}
