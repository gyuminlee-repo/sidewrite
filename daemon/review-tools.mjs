// Review-turn tool policy and the office_tracked_edit tool.
//
// In a review turn the only way to change the document is
// office_tracked_edit, which applies one tracked change and anchors one
// "[R1.3] what / why / addresses" comment on it. Every other document or
// comment writer is refused, so "changes == rationale comments" holds by
// construction. disallowedTools cannot do this per turn: it is fixed when
// query() starts (daemon/index.mjs:1062) and the session lives across turns.

import { z } from "zod";
import { forbidHits } from "./ledger-check.mjs";
import { lostFacts, unsourcedTerms, narrowEdit, startsSentence } from "./edit-facts.mjs";
import { changeContext } from "./revision-classify.mjs";

export const REVIEW_WRITE_TOOL = "office_tracked_edit";

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

export function reviewTurnToolAllowed(name) {
  return REVIEW_READ_TOOLS.has(name) || name === REVIEW_WRITE_TOOL;
}

// Wrap the bridge call used by every office/excel tool handler.
export function guardedCall(rawCall, getReviewTurn) {
  return (name, args) => {
    const turn = getReviewTurn?.() ?? null;
    if (turn && !reviewTurnToolAllowed(name)) {
      return Promise.reject(
        new Error(
          `${name} is disabled during review turn ${turn.ref}. Use office_tracked_edit: every change needs its own rationale comment.`,
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
  "Returns ok:false with blocked_by when the new text matches the forbid list of the author; then rephrase, do not retry the same text.";

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
          "Ids of the listed comment requirements this change answers (q1, q2). Use only ids from the prompt.",
        ),
    })
    .describe(
      "Korean by default. Facts from the manuscript or the review only; say so when inferring.",
    ),
};

export const TOOL_COMMENT_PREFIX = /^\s*\[(?:R\d+|E)\.\d+\]/;

// L1: facts the change deletes go on a last line, so the author sees them in
// Word next to the change. L5: requirement ids follow the addresses part.
export function formatRationaleComment(
  ref,
  { what, why, addresses, requirement_ids },
  { lost = [] } = {},
) {
  const one = (s) => String(s).replace(/\s+/g, " ").trim();
  const ids = requirement_ids?.length ? ` [${requirement_ids.join(",")}]` : "";
  const facts = lost.length ? `\n사라지는 사실: ${lost.join("; ")}` : "";
  return `[${ref}] ${one(what)} / ${one(why)} / ${one(addresses)}${ids}${facts}`;
}

function fail(message, extra = {}) {
  return { ok: false, message, ...extra };
}

// deps: { call(name, args), getLedger(), getReviewContext(ref), onEdit({ref, change}),
//         onBlocked({ref, args, blocked_by}), log }
// getReviewContext returns { requirements, comment_text, paragraphs } for the
// L2 sources (manuscript snapshot texts and the comment) and the L5 ids.
// turn: { ref } of the current review turn, or null.
// override: set only by the daemon after the user approved an exception in
// the task pane; it is not part of the tool schema the model sees.
export async function applyTrackedEdit(deps, args, turn, { override = null } = {}) {
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

  const ctx = (await deps.getReviewContext?.(turn.ref)) ?? {};
  const known = new Set((ctx.requirements ?? []).map((r) => r.id));
  const unknown = (args.rationale?.requirement_ids ?? []).filter((id) => !known.has(id));
  if (unknown.length) {
    return fail(
      `Unknown requirement id(s) ${unknown.join(", ")} for ${turn.ref}. Valid ids: ${[...known].join(", ") || "(none listed)"}.`,
    );
  }

  const ledger = (await deps.getLedger?.()) ?? [];
  const hits = forbidHits(ledger, newText);
  const excepted = override?.blocked_by ?? [];
  const open = hits.filter((id) => !excepted.includes(id));
  if (open.length) {
    const blocked = await deps.onBlocked?.({ ref: turn.ref, args: edit, blocked_by: open });
    deps.log?.(`[review] blocked ${turn.ref} by=${open.join(",")}`);
    return fail(
      `Not applied: the new text matches forbidden rule(s) ${open.join(", ")}. Rephrase without it; the author can approve an exception in the Review tab.`,
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
  const lost =
    kind === "insert"
      ? []
      : lostFacts(args.find, args.replace, { sentenceStart: startsSentence(args.find, texts) });
  const unsourced =
    kind === "delete" ? [] : unsourcedTerms(newText, [...texts, ctx.comment_text ?? "", args.find]);
  const narrow = hasReplace ? narrowEdit(args.find, args.replace) : null;

  const r = await deps.call(REVIEW_WRITE_TOOL, {
    find: args.find,
    ...(hasReplace ? { replace: args.replace } : { insert_after: args.insert_after }),
    ...(args.paragraph_hint ? { paragraph_hint: args.paragraph_hint } : {}),
    ...(narrow ? { narrow } : {}),
    comment_text: formatRationaleComment(turn.ref, rationale, { lost }),
  });
  if (!r) return fail("empty result from the task pane");
  // applied:true with ok:false: the task pane changed the document but could
  // not anchor the rationale comment (CommentAnchorMissing). The change is
  // still recorded, with no comment id, so the change list and the
  // "changes == rationale comments" check show it instead of losing it.
  const commentMissing = r.ok === false && r.applied === true ? (r.code ?? "applied") : null;
  if (r.ok === false && !commentMissing) return r;

  // The task pane reports whether it could apply the narrowed range (the cut
  // find must occur once inside the located hit); otherwise the full range.
  const narrowed = !!(narrow && r.narrowed === true);
  const appliedKind = narrowed ? (narrow.replace === "" ? "delete" : "replace") : kind;
  const change = {
    word_comment_id: r.word_comment_id == null ? null : String(r.word_comment_id),
    paragraph_index: r.paragraph_index ?? null,
    paragraph_id: r.paragraph_id ?? null,
    kind: appliedKind,
    find: narrowed ? narrow.find : args.find,
    text: narrowed ? narrow.replace : newText,
    before: r.before ?? "",
    after: r.after ?? "",
    rationale,
    override_of: excepted,
    lost_facts: lost,
    unsourced_terms: unsourced,
    narrowed_from: narrowed ? { find: args.find, text: newText } : null,
    comment_missing: commentMissing,
    // Words around the range the task pane changed (revision-classify.mjs).
    context: changeContext(r.before, args.find, {
      kind,
      narrowFind: narrowed ? narrow.find : null,
    }),
  };
  const saved = await deps.onEdit?.({ ref: turn.ref, change });
  const id = saved?.id ?? "?";
  deps.log?.(
    `[review] change recorded ${id} kind=${appliedKind} comment=${change.word_comment_id} paragraph=${change.paragraph_index}` +
      ` narrowed=${narrowed ? 1 : 0} lost=${lost.length} unsourced=${unsourced.length}`,
  );
  if (lost.length) deps.log?.(`[review] lost facts ${id}: ${lost.join("; ")}`);
  if (unsourced.length) deps.log?.(`[review] unsourced ${id}: ${unsourced.join("; ")}`);
  if (commentMissing) deps.log?.(`[review] comment missing ${id} code=${commentMissing}`);
  const warnings = [
    ...(commentMissing
      ? [
          `The change was applied but no rationale comment could be anchored on it (${commentMissing}). Do not repeat this edit; tell the author that change ${id} has no rationale comment.`,
        ]
      : []),
    ...(unsourced.length
      ? [
          `Not in the manuscript or the comment: ${unsourced.join(", ")}. Name them in your answer so the author can check where they come from.`,
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
    ...(warnings.length ? { warning: warnings.join(" ") } : {}),
  };
}
