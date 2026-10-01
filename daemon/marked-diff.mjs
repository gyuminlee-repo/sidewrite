// Build a colour-marked copy of the document against a baseline (P4).
//
// One step through the injected Python runner: tools/build_marked_diff.py
// --json. Its exit code decides the result: 0 the copy was written, 1 the
// tool's own check failed and nothing was written, 2 the request was
// refused (the reason is the tool's). The copy goes next to the document;
// a taken name gets a time stamp and then a counter, and the tool creates
// its output exclusively, so an earlier copy is never overwritten.
// Colours, mode and author come from the journal profile (colors.
// marked_insert, colors.marked_delete, marked_mode, marked_author).
//
// The tool needs the Python package lxml. This module installs nothing: a
// missing package is reported with what to install.
//
// Not wired yet. The caller passes runPy (daemon/py-tool.mjs), the path of
// the uploaded baseline (daemon/uploads.mjs), the path of the current
// document's bytes written to a temporary file (daemon/doc-bytes.mjs, with
// its `source`), the local path of the document and existsSync.

import { join, resolve } from "node:path";
import { pyJson, pyError, siblingPath, outputTaken, REPO_ROOT, SAVE_FIRST } from "./doc-tools.mjs";
import { DISK_NOTICE } from "./doc-bytes.mjs";

export const MARKED_NOTICE =
  "Pending tracked changes were read as accepted in both files before they were compared.";
export const NEEDS_LXML =
  "The marked copy needs the Python package lxml, which is not installed. Install it yourself: pip install lxml";

export const markedOutPath = (docPath, deps) => siblingPath(docPath, "marked", deps);

const samePath = (a, b) =>
  resolve(a) === resolve(b) || resolve(a).toLowerCase() === resolve(b).toLowerCase();
const RGB = /^[0-9A-Fa-f]{6}$/;
const MODES = ["color", "track_changes"];

// The tool options a profile gives, or { error }.
function profileArgs(profile) {
  const args = [];
  for (const [key, flag] of [
    ["marked_insert", "--insert-rgb"],
    ["marked_delete", "--delete-rgb"],
  ]) {
    const value = profile.colors?.[key];
    if (value === undefined || value === null) continue;
    if (!RGB.test(String(value))) {
      return { error: `The profile colour ${key} must be six hex digits, for example FF0000.` };
    }
    args.push(flag, String(value).toUpperCase());
  }
  if (profile.marked_mode !== undefined && profile.marked_mode !== null) {
    if (!MODES.includes(profile.marked_mode)) {
      return { error: `The profile marked_mode must be one of: ${MODES.join(", ")}.` };
    }
    args.push("--mode", profile.marked_mode);
  }
  if (profile.marked_author) args.push("--author", String(profile.marked_author));
  return { args };
}

export async function runMarkedDiff({
  baseline,
  clean,
  docPath,
  out,
  profile = {},
  source,
  runPy,
  exists,
  now,
  root = REPO_ROOT,
} = {}) {
  if (!baseline) return { ok: false, error: "Choose the baseline file first." };
  if (!clean || !(out ?? docPath)) return { ok: false, error: SAVE_FIRST };
  const fromProfile = profileArgs(profile);
  if (fromProfile.error) return { ok: false, error: fromProfile.error };
  // The tool creates its output exclusively. When the name that was free a
  // moment ago is taken by the time the tool writes, another name is picked.
  const taken = new Set();
  const isTaken = (p) => taken.has(p) || (exists?.(p) ?? false);
  let target;
  let r;
  for (let attempt = 0; attempt < 5; attempt++) {
    target = out ?? markedOutPath(docPath, { exists: isTaken, now });
    if (samePath(target, baseline) || samePath(target, clean)) {
      return { ok: false, error: `The output is the same file as an input: ${target}` };
    }
    r = await runPy(join(root, "tools", "build_marked_diff.py"), [
      "--baseline",
      baseline,
      "--clean",
      clean,
      "--out",
      target,
      "--json",
      ...fromProfile.args,
    ]);
    if (out || !outputTaken(r)) break;
    taken.add(target);
  }
  const report = pyJson(r);
  if (r.code === 0 && report?.written) {
    return {
      ok: true,
      path: target,
      report,
      notice: source === "disk" ? `${MARKED_NOTICE} ${DISK_NOTICE}` : MARKED_NOTICE,
    };
  }
  if (/No module named '?lxml'?/.test(String(r.stderr ?? ""))) {
    return { ok: false, error: NEEDS_LXML };
  }
  if (r.code === 1 && report) {
    return {
      ok: false,
      report,
      error: `The marked copy did not pass its own check (${report.failures} of ${report.checks} checks failed); nothing was written.`,
    };
  }
  return { ok: false, error: pyError(r) };
}
