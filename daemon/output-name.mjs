// Names for files the daemon writes next to a manuscript for the author to
// keep (the response letter, its plain-text copies). The rule is one: a file
// that is already there is never replaced unless the request says so. The
// writer creates its file exclusively and the caller walks this sequence
// until a name is free.
//
// This is the one place the sequence is defined, so another naming helper
// can take its place without touching the callers.

import { access } from "node:fs/promises";

// The nth name of `path`: itself for n = 1, then "-2", "-3" ... in front of
// the last extension ("paper.response.docx" -> "paper.response-2.docx").
export function outputName(path, n) {
  if (n <= 1) return path;
  const slash = Math.max(path.lastIndexOf("/"), path.lastIndexOf("\\"));
  const dot = path.lastIndexOf(".");
  return dot > slash + 1 ? `${path.slice(0, dot)}-${n}${path.slice(dot)}` : `${path}-${n}`;
}

export const exists = (path) =>
  access(path).then(
    () => true,
    () => false,
  );

export const MAX_OUTPUT_NAMES = 200;

// The first n >= from for which none of group(outputName(path, n)) exists.
// group: the files written together under one name (default: the file alone).
export async function freeOutputIndex(path, { from = 1, group = (p) => [p] } = {}) {
  for (let n = from; n < from + MAX_OUTPUT_NAMES; n++) {
    const taken = await Promise.all(group(outputName(path, n)).map(exists));
    if (!taken.some(Boolean)) return n;
  }
  throw new Error(`No free name for ${path}: ${MAX_OUTPUT_NAMES} files of that name are there.`);
}
