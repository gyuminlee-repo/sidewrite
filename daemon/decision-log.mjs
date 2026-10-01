// Decision receipts: one JSON line per LLM decision in
// `<manuscript stem>.decisions.jsonl` next to the document
// (revision-store.mjs decisionsPathFor). Kinds: comment_split,
// requirement_extract, ledger_check, requirement_link, ledger_suggest,
// proposal_batch (one per proposal turn, proposal-turns.mjs: the outcome of
// each proposal by id). A later gold-set tool reads the file; a person fills
// `expected` with the right answer.
//
// A record holds a sha256 fingerprint of the input, never the input text:
// no manuscript or review words go into the file.
//
// Append only. Each call writes its lines with one appendFile, and calls are
// serialized per file inside this process (like withRevision), so lines
// from overlapping calls do not interleave.

import { appendFile } from "node:fs/promises";
import { createHash } from "node:crypto";

export function fingerprint(input) {
  const s = typeof input === "string" ? input : JSON.stringify(input);
  return createHash("sha256").update(s).digest("hex");
}

export function decisionRecord({ kind, input, model = null, output, now = new Date() }) {
  return {
    ts: now.toISOString(),
    kind,
    input_fingerprint: fingerprint(input),
    model,
    output,
    expected: null,
  };
}

// Split warnings quote review sentences in double quotes: keep the text
// before the first quote only.
export function redactQuotes(warning) {
  const s = String(warning);
  const i = s.indexOf('"');
  return i === -1 ? s : `${s.slice(0, i).trimEnd()} [quote omitted]`;
}

const chains = new Map();

export function appendDecisions(path, records) {
  const text = records.map((r) => JSON.stringify(r) + "\n").join("");
  const prev = chains.get(path) ?? Promise.resolve();
  const run = prev.catch(() => {}).then(() => (text ? appendFile(path, text, "utf8") : undefined));
  chains.set(path, run);
  return run;
}
