// Write an anonymized copy of the document (P10) and check it.
//
// Two steps through the injected Python runner: tools/anonymize_docx.py
// writes the copy, scripts/verify-anonymized.py counts what is left and
// compares the text with the source. The result is the verifier's report.
// When the verifier does not pass, the copy is renamed to
// "<name>_anonymized.INCOMPLETE.docx", so no file that failed the check
// carries the name of a finished copy.
//
// The source is never the output, and no existing file is overwritten: a
// taken name gets a time stamp and then a counter, the Python tool creates
// its output exclusively (it refuses a name that exists by the time it
// writes, and another name is tried), and the move to the INCOMPLETE name is
// exclusive too.
//
// Not wired yet. The caller passes runPy (daemon/py-tool.mjs), the local
// path of the saved document and existsSync, and shows ANONYMIZE_NOTICE with
// the report.

import { join, resolve } from "node:path";
import {
  pyJson,
  pyError,
  siblingPath,
  freeName,
  outputTaken,
  moveExclusive,
  REPO_ROOT,
  SAVE_FIRST,
} from "./doc-tools.mjs";

export const ANONYMIZE_NOTICE =
  "Author marks and file metadata are cleared in the copy. The text itself is not changed: a name written in the manuscript stays there. Hidden text, embedded objects and TIFF pictures are reported, not cleared.";

export const anonymizedPath = (docPath, deps) => siblingPath(docPath, "anonymized", deps);

const TRIES = 5;

const samePath = (a, b) =>
  resolve(a) === resolve(b) || resolve(a).toLowerCase() === resolve(b).toLowerCase();

// What the verifier found, as one sentence for the pane.
function shortfall(report) {
  const left = Object.entries(report.categories ?? {})
    .filter(([, c]) => c.status === "not cleared" && c.default !== "report")
    .map(([id, c]) => `${id} (${c.count})`);
  const parts = [];
  if (left.length) parts.push(`still present: ${left.join(", ")}`);
  if (report.body_equal === false) parts.push("the text changed");
  if (report.markup_equal === false) parts.push("the number of tracked changes changed");
  if (report.structure_ok === false) parts.push("a reference points at a missing part");
  if (report.unreadable_parts?.length) parts.push("a part is not well-formed");
  return parts.join("; ") || "the check did not pass";
}

export async function runAnonymize({
  src,
  out,
  author = "Anonymous",
  initials = "A",
  keep = [],
  also = [],
  dates,
  runPy,
  exists,
  move = moveExclusive,
  sameFile,
  now,
  root = REPO_ROOT,
} = {}) {
  if (!src) return { ok: false, error: SAVE_FIRST };
  const who = ["--author", author, "--initials", initials];
  const which = [
    ...(keep.length ? ["--keep", keep.join(",")] : []),
    ...(also.length ? ["--also", also.join(",")] : []),
  ];
  // The tool creates its output exclusively. When the name that was free a
  // moment ago is taken by the time the tool writes, another name is picked.
  const taken = new Set();
  const isTaken = (p) => taken.has(p) || (exists?.(p) ?? false);
  let target;
  let made;
  for (let attempt = 0; attempt < TRIES; attempt++) {
    target = out ?? anonymizedPath(src, { exists: isTaken, now });
    if (samePath(src, target) || (sameFile && (await sameFile(src, target)))) {
      return { ok: false, error: `The output is the same file as the document: ${target}` };
    }
    made = await runPy(join(root, "tools", "anonymize_docx.py"), [
      "--in",
      src,
      "--out",
      target,
      ...who,
      ...which,
      ...(dates ? ["--dates", dates] : []),
    ]);
    if (out || !outputTaken(made)) break;
    taken.add(target);
  }
  if (made.code !== 0) return { ok: false, error: pyError(made) };
  const checked = await runPy(join(root, "scripts", "verify-anonymized.py"), [
    "--before",
    src,
    "--after",
    target,
    "--json",
    ...who,
    ...which,
  ]);
  const report = pyJson(checked);
  if (checked.code === 0 && report) {
    return {
      ok: true,
      path: target,
      report,
      changed: pyJson(made)?.changed ?? {},
      notice: ANONYMIZE_NOTICE,
    };
  }
  // Not verified: the copy must not keep the name of a finished one. The
  // move is exclusive, so an earlier set-aside copy is not replaced either.
  const ext = /(\.[^./\\]+)?$/.exec(target)[0];
  const stem = `${target.slice(0, target.length - ext.length)}.INCOMPLETE`;
  let path = null;
  let failure = null;
  for (let attempt = 0; attempt < TRIES && !path; attempt++) {
    const aside = freeName(stem, ext, isTaken);
    try {
      await move(target, aside);
      path = aside;
    } catch (e) {
      failure = e;
      if (e?.code !== "EEXIST") break;
      taken.add(aside);
    }
  }
  if (!path) {
    return {
      ok: false,
      path: target,
      report,
      error: `The copy failed its check and could not be set aside (${String(failure?.message ?? failure)}). Do not use ${target}.`,
    };
  }
  if (checked.code === 1 && report) {
    return { ok: false, path, report, error: `The copy is incomplete: ${shortfall(report)}.` };
  }
  return { ok: false, path, error: `The copy could not be checked: ${pyError(checked)}` };
}
