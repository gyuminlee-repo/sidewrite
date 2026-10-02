// The submission checklist of the Review tab ("Before submission", Check now).
// Pure: no I/O. The service (aux-service.mjs submission_check) reads the disk
// and hands the hashes in.
//
// submissionRows({ revision, profile, live, files, file }) returns rows
//   { id, file, state: "ready" | "not_ready" | "not_applicable", text, go }
// in the order the author works through them. A row is computed from the
// record of the paper, the hashes of the files and the document as the pane
// read it just now; nothing of a row is stored. A row that cannot be computed
// from what was handed in is left out: it is never guessed.
//
//   revision  the record, brought up to date with `live.snapshot` in memory
//             (response-service.mjs refreshed)
//   profile   with defaults filled (journal-profile.mjs effectiveProfile)
//   live      { snapshot, headings, format } as the pane read them: the
//             snapshot of collectReviewSnapshot (paragraphs with w_ins and
//             w_del, comments), the usable headings, the format read of
//             readFormatSnapshot (null when it could not be read)
//   files     { <path as the record holds it>: sha256 | null } for the paths
//             of submissionPaths(revision): the file as it is on the disk
//             now, null when it is not there. A path that was not handed in
//             counts as not there. The saved manuscript itself is under the
//             key MANUSCRIPT.
//   file      the name of the open manuscript; every row is about that file
//
//   go        where the row is fixed, or null:
//             { tab, target }   a tab of the pane and the id of an element
//             { paragraph }     a paragraph of the document, by index
//
// The rows:
//   answers          every comment has an answer that passes the pure checks
//   requests         no request is open, no question to the author unanswered
//   changes_decided  no tracked change is waiting for a decision
//   tool_comments    the tool's own comments are removed
//   file_clean       no tracked change and no comment is left (the state the
//                    two rows above lead to, plus the comments of people)
//   format_counts    no format finding and no failing count rule
//   letter           the exported letter (.docx) still says what the record
//                    says now. Compared by hash, not by time: the export
//                    recorded the hash of the letter model and of the file
//                    (response_export.model_sha, docx_sha), and the model
//                    built now must hash the same. A later accept or reject,
//                    an edited answer, a profile change or new line numbers
//                    all change the model.
//   marked_copy      a marked copy is on record (revision.artifacts.marked),
//                    it is still the file that was made, and the saved
//                    manuscript is still the file it was made from. Removing
//                    the tool's comments and saving changes the manuscript,
//                    so the cleanup rows stand before this one and the row
//                    says to make the copy after the cleanup.
//   anonymized_copy  the same for the anonymized copy, when the profile asks
//                    for one (submission.anonymized_copy: true); false or
//                    null: not applicable

import { checkResponses } from "./response-check.mjs";
import { buildLetterModel, letterModelSha } from "./response-model.mjs";
import { inScope } from "./ledger-check.mjs";
import { findFormatIssues, formatRulesOn } from "./format-check.mjs";
import { checkCountItem } from "./count-check.mjs";
import { TOOL_COMMENT_PREFIX } from "./review-tools.mjs";

export const READY = "ready";
export const NOT_READY = "not_ready";
export const NOT_APPLICABLE = "not_applicable";

// "1 comment", "3 comments".
const n = (count, one, many = `${one}s`) => `${count} ${count === 1 ? one : many}`;
const uniq = (list) => [...new Set(list)];
// At most four names in a sentence of a 328 px pane.
function names(list) {
  const all = uniq(list.filter(Boolean));
  return all.length > 4
    ? `${all.slice(0, 4).join(", ")} and ${all.length - 4} more`
    : all.join(", ");
}

const TO_RESPONSE_CHECK = { tab: "response", target: "response-check" };
const TO_COMMENTS = { tab: "review", target: "review-comments" };
const TO_CLEAN_TOOL = { tab: "review", target: "review-clean-tool" };
const TO_DELETE_ALL = { tab: "review", target: "review-delete-all" };
const TO_TOOLS_CHECK = { tab: "tools", target: "tools-check" };
const TO_EXPORT = { tab: "response", target: "response-export" };
const TO_MARKED = { tab: "tools", target: "tools-marked-build" };
const TO_ANONYMIZE = { tab: "tools", target: "tools-anonymize" };

// The key of the saved manuscript in `files` (no path of a record reads so).
export const MANUSCRIPT = ":manuscript";

// The last part of a path, whichever separator it was written with.
const nameOf = (path) =>
  String(path ?? "")
    .split(/[\\/]/)
    .pop();

// The files whose hashes the rows compare, as the record names them. The
// service hashes each one and hands the result in as `files`.
export function submissionPaths(revision) {
  return uniq(
    [
      revision?.response_export?.path,
      revision?.artifacts?.marked?.path,
      revision?.artifacts?.anonymized?.path,
    ].filter(Boolean),
  );
}

// Rows 1 and 2: the pure checks of the response letter. The `requirements`
// rows are the open requests; every other failing row is an answer that does
// not pass.
function answerRows({ revision, profile, live }) {
  const fails = checkResponses({
    revision,
    snapshot: live.snapshot,
    headings: live.headings ?? [],
    profile,
  }).filter((r) => r.level === "fail");
  const answers = fails.filter((r) => r.id !== "requirements");
  const requests = fails.filter((r) => r.id === "requirements");
  const unanswered = (revision.comments ?? []).flatMap((c) =>
    (c.questions ?? []).filter((q) => !String(q.answer ?? "").trim()).map(() => c.ref),
  );
  // Rules a model checks are not run here: the row says so, and the letter
  // row needs an export that ran them.
  const byModel = (revision.ledger ?? []).filter(
    (i) => inScope(i, "response") && i.check?.type === "llm",
  ).length;
  const modelNote = byModel
    ? ` ${n(byModel, "rule")} checked by a model ${byModel === 1 ? "is" : "are"} not run here; run the checks in the Response tab for ${byModel === 1 ? "it" : "them"}.`
    : "";

  const refs = uniq(answers.map((r) => r.ref));
  const answersRow = refs.length
    ? {
        id: "answers",
        state: NOT_READY,
        text: `${n(refs.length, "answer")} ${refs.length === 1 ? "is" : "are"} missing or ${refs.length === 1 ? "fails" : "fail"} a check: ${names(refs)}.${modelNote}`,
        go: TO_RESPONSE_CHECK,
      }
    : {
        id: "answers",
        state: READY,
        text: `Every comment has an answer that passes the checks.${modelNote}`,
        go: byModel ? TO_RESPONSE_CHECK : null,
      };

  const open = [
    requests.length &&
      `${n(requests.length, "request")} ${requests.length === 1 ? "has" : "have"} no accepted change`,
    unanswered.length &&
      `${n(unanswered.length, "question")} to you ${unanswered.length === 1 ? "has" : "have"} no answer`,
  ].filter(Boolean);
  const requestsRow = open.length
    ? {
        id: "requests",
        state: NOT_READY,
        text: `${open.join(" and ")}: ${names([...requests.map((r) => r.ref), ...unanswered])}.`,
        go: TO_COMMENTS,
      }
    : {
        id: "requests",
        state: READY,
        text: "No request is open and no question to you is waiting.",
        go: null,
      };
  return [answersRow, requestsRow];
}

// Rows 3 to 5: what is still in the open document. Without the comment list
// of the snapshot only the tracked changes can be told.
function documentRows({ live }) {
  const paragraphs = live.snapshot.paragraphs ?? [];
  const tracked = paragraphs.filter((p) => (p.w_ins ?? 0) + (p.w_del ?? 0) > 0);
  const toFirst = tracked.length
    ? { paragraph: tracked[0].index ?? paragraphs.indexOf(tracked[0]) }
    : null;
  const rows = [
    tracked.length
      ? {
          id: "changes_decided",
          state: NOT_READY,
          text: `${n(tracked.length, "paragraph")} still ${tracked.length === 1 ? "holds" : "hold"} tracked changes. Accept or reject them in Word.`,
          go: toFirst,
        }
      : {
          id: "changes_decided",
          state: READY,
          text: "No tracked change is waiting for a decision in this file.",
          go: null,
        },
  ];
  const comments = live.snapshot.comments;
  if (!Array.isArray(comments)) return rows;
  const tool = comments.filter((c) => TOOL_COMMENT_PREFIX.test(String(c?.content ?? ""))).length;
  const other = comments.length - tool;
  rows.push(
    tool
      ? {
          id: "tool_comments",
          state: NOT_READY,
          text: `${n(tool, "comment")} of the tool ${tool === 1 ? "is" : "are"} still in this file. Remove them before you make the marked copy.`,
          go: TO_CLEAN_TOOL,
        }
      : { id: "tool_comments", state: READY, text: "The tool's comments are removed.", go: null },
  );
  const left = [
    tracked.length && `${n(tracked.length, "paragraph")} with tracked changes`,
    tool && `${n(tool, "comment")} of the tool`,
    other && `${n(other, "comment")} of yours or the reviewers`,
  ].filter(Boolean);
  rows.push(
    left.length
      ? {
          id: "file_clean",
          state: NOT_READY,
          text: `This file is not clean yet. Left in it: ${left.join(", ")}.`,
          // The comments of people have no row of their own: this one leads
          // to where they are deleted. Otherwise to what a row above names.
          go: other ? TO_DELETE_ALL : tool ? TO_CLEAN_TOOL : toFirst,
        }
      : {
          id: "file_clean",
          state: READY,
          text: "This file is clean: no tracked changes and no comments are left.",
          go: null,
        },
  );
  return rows;
}

// Row 6: the format rules of the profile and the count rules of the ledger,
// as the Tools tab's "Run checks" reads them.
function formatRow({ revision, profile, live }) {
  const id = "format_counts";
  const counts = (revision.ledger ?? []).filter((i) => i.check?.type === "count");
  const rules = formatRulesOn(profile);
  if (!rules.length && !counts.length)
    return {
      id,
      state: NOT_APPLICABLE,
      text: "No format rule and no count rule is set for this paper.",
      go: null,
    };
  const paragraphs = live.snapshot.paragraphs ?? [];
  const headings = live.headings ?? [];
  const read = live.format?.paragraphs;
  // A format read that does not fit the text was not checked: never "ready".
  if (!Array.isArray(read) || read.length !== paragraphs.length)
    return {
      id,
      state: NOT_READY,
      text: "The formatting of the document could not be read, so the format rules were not checked. Press Check now again, or run the checks in the Tools tab.",
      go: TO_TOOLS_CHECK,
    };
  const findings = findFormatIssues(read, profile, { headings }).length;
  const results = counts.map((item) => checkCountItem(item, paragraphs, { headings, profile }));
  const failing = results.filter((r) => r.status === "violations").length;
  const unchecked = results.filter((r) => r.status === "error").length;
  const bad = [
    findings && n(findings, "format finding"),
    failing && `${n(failing, "count rule")} ${failing === 1 ? "fails" : "fail"}`,
    unchecked && `${n(unchecked, "count rule")} could not be checked`,
  ].filter(Boolean);
  return bad.length
    ? {
        id,
        state: NOT_READY,
        text: `${bad.join(", ")}. Run the checks in the Tools tab to see where.`,
        go: TO_TOOLS_CHECK,
      }
    : { id, state: READY, text: "The format and count checks have no failing rule.", go: null };
}

// Row 7: the response letter that was exported, against the record as it is
// now. Not ready is the answer whenever one of the compared values is missing.
function letterRow({ revision, profile, files }) {
  const id = "letter";
  const notReady = (text) => ({ id, state: NOT_READY, text, go: TO_EXPORT });
  const x = revision.response_export;
  if (!x) return notReady("The response letter (.docx) has not been exported yet.");
  const name = nameOf(x.path);
  if (!x.model_sha || !x.docx_sha)
    return notReady(
      `Export the letter again: ${name} was written before the tool recorded what it said, so it cannot be compared with the answers as they are now.`,
    );
  if (x.fail > 0)
    return notReady(
      `${name} was exported while ${n(x.fail, "check")} failed. Fix ${x.fail === 1 ? "it" : "them"} in the Response tab, then export the letter again.`,
    );
  const onDisk = files[x.path] ?? null;
  if (onDisk === null) return notReady(`The exported letter is gone: ${name}. Export it again.`);
  if (onDisk !== x.docx_sha)
    return notReady(
      `${name} was changed after it was exported, so it is no longer the letter the tool wrote. Export it again.`,
    );
  const now = letterModelSha(
    buildLetterModel(revision, profile, { gloss: !!profile.response_letter.gloss_enabled }),
  );
  if (now !== x.model_sha)
    return notReady(
      `The answers, the accepted changes or the line numbers changed after ${name} was exported. Export the letter again.`,
    );
  return {
    id,
    state: READY,
    text: `The exported letter says what the answers and the accepted changes say now: ${name}.`,
    go: null,
  };
}

// Rows 8 and 9: a copy made from the saved manuscript, against its record.
// what: "marked copy" or "anonymized copy". No record (never made, made by
// an earlier version, or dropped by an earlier daemon) is not ready: make it.
function copyRow(id, what, artifact, files, go) {
  const notReady = (text) => ({ id, state: NOT_READY, text, go });
  if (!artifact?.path || !artifact.sha256 || !artifact.source_sha256)
    return notReady(
      `No ${what} is on record for this file. Finish the cleanup above and save the manuscript first (removing comments changes the saved file), then make the copy.`,
    );
  const name = nameOf(artifact.path);
  const onDisk = files[artifact.path] ?? null;
  if (onDisk === null) return notReady(`The ${what} is gone: ${name}. Make it again.`);
  if (onDisk !== artifact.sha256)
    return notReady(
      `${name} was changed after it was made, so it is no longer the ${what} the tool wrote. Make it again.`,
    );
  if ((files[MANUSCRIPT] ?? null) !== artifact.source_sha256)
    return notReady(
      `The manuscript was saved again after ${name} was made (removing the tool's comments counts too). Finish the cleanup above, save, then make the ${what} again.`,
    );
  return {
    id,
    state: READY,
    text: `The ${what} was made from the manuscript as it is saved now: ${name}.`,
    go: null,
  };
}

function copyRows({ revision, profile, files }) {
  const wanted = profile.submission?.anonymized_copy;
  return [
    copyRow("marked_copy", "marked copy", revision.artifacts?.marked, files, TO_MARKED),
    wanted === true
      ? copyRow(
          "anonymized_copy",
          "anonymized copy",
          revision.artifacts?.anonymized,
          files,
          TO_ANONYMIZE,
        )
      : {
          id: "anonymized_copy",
          state: NOT_APPLICABLE,
          text:
            wanted === false
              ? "The journal profile asks for no anonymized copy."
              : "The journal profile does not say that an anonymized copy is needed.",
          go: null,
        },
  ];
}

export function submissionRows({ revision, profile, live, files = {}, file = null }) {
  const input = { revision, profile, live, files };
  return [
    ...answerRows(input),
    ...documentRows(input),
    formatRow(input),
    letterRow(input),
    ...copyRows(input),
  ].map((row) => ({ ...row, file }));
}
