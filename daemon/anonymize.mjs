// Write an anonymized copy of the document (P10) and check it.
//
// Two steps through the injected Python runner: tools/anonymize_docx.py
// writes the copy, scripts/verify-anonymized.py counts what is left and
// compares the text with the source. The result is the verifier's report.
// When the verifier does not pass, the copy is renamed to
// "<name>_anonymized.INCOMPLETE.docx", so no file that failed the check
// carries the name of a finished copy.
//
// The source is never the output, and an existing copy is never
// overwritten (a taken name gets a time stamp).
//
// Not wired yet. The caller passes runPy (daemon/py-tool.mjs), the local
// path of the saved document, existsSync and fs rename, and shows
// ANONYMIZE_NOTICE with the report.

import { join, resolve } from "node:path";
import { pyJson, pyError, siblingPath, REPO_ROOT, SAVE_FIRST } from "./doc-tools.mjs";

export const ANONYMIZE_NOTICE =
  "Author marks and file metadata are cleared in the copy. The text itself is not changed: a name written in the manuscript stays there. Hidden text, embedded objects and TIFF pictures are reported, not cleared.";

export const anonymizedPath = (docPath, deps) => siblingPath(docPath, "anonymized", deps);

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
  rename,
  sameFile,
  now,
  root = REPO_ROOT,
} = {}) {
  if (!src) return { ok: false, error: SAVE_FIRST };
  const target = out ?? anonymizedPath(src, { exists, now });
  if (samePath(src, target) || (sameFile && (await sameFile(src, target)))) {
    return { ok: false, error: `The output is the same file as the document: ${target}` };
  }
  const who = ["--author", author, "--initials", initials];
  const which = [
    ...(keep.length ? ["--keep", keep.join(",")] : []),
    ...(also.length ? ["--also", also.join(",")] : []),
  ];
  const made = await runPy(join(root, "tools", "anonymize_docx.py"), [
    "--in",
    src,
    "--out",
    target,
    ...who,
    ...which,
    ...(dates ? ["--dates", dates] : []),
  ]);
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
  // Not verified: the copy must not keep the name of a finished one.
  const aside = target.replace(/(\.[^./\\]+)?$/, (ext) => `.INCOMPLETE${ext}`);
  let path = aside;
  try {
    await rename(target, aside);
  } catch (e) {
    return {
      ok: false,
      path: target,
      report,
      error: `The copy failed its check and could not be set aside (${String(e?.message ?? e)}). Do not use ${target}.`,
    };
  }
  if (checked.code === 1 && report) {
    return { ok: false, path, report, error: `The copy is incomplete: ${shortfall(report)}.` };
  }
  return { ok: false, path, error: `The copy could not be checked: ${pyError(checked)}` };
}
