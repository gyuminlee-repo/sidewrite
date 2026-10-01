// A stand-in for the provider of a proposal turn, for the acceptance run
// (scripts/e2e) only. With CC_OFFICE_E2E_PROPOSAL_FIXTURE=<file.json> in the
// environment of the daemon, a proposal turn takes its reply from that file
// and no provider is asked: the stretch from the runner to Word (checks,
// tracked edits, stops, the receipt) can be run without a model. Nothing
// else of the turn changes. Without the variable this file is not used
// (proposal-turns.mjs proposalTurnsFor is its one caller).
//
// The file: { "<comment ref>": { "proposals": [...], "note": "" }, ... }.
// A proposal is as a provider writes it (proposal.mjs), except that its
// paragraph may be given by text instead of by id, because ids are Word's:
//   "paragraph_contains": "<text>"   the first paragraph shown that holds it
// The file is read on every turn, so it can be changed between two turns.
// Receipts of such a turn carry the model "fixture:<file name>".

import { readFile as defaultReadFile } from "node:fs/promises";
import { basename } from "node:path";

const fail = (error) => ({ ok: false, error });

// (choice) => { provider, model, run(req) }, the shape of
// oneshot-router.mjs handleFor. run never throws.
export function fixtureHandleFor(file, { readFile = defaultReadFile, log = console.log } = {}) {
  const model = `fixture:${basename(file)}`;
  async function reply(req) {
    if (req?.signal?.aborted) return { ok: false, aborted: true, error: "stopped" };
    const ref = /^\[Proposal turn ([^\]]+)\]/.exec(String(req?.prompt ?? ""))?.[1];
    if (!ref) return fail("the proposal fixture could not tell the comment of the prompt");
    const data = JSON.parse(await readFile(file, "utf8"));
    const entry = data?.[ref];
    if (!entry) return fail(`the proposal fixture has no reply for ${ref}`);
    // The paragraphs the turn shows: "[id] text" lines of the prompt.
    const shown = [...String(req.prompt).matchAll(/^\[([^\]]+)\] (.*)$/gm)].map((m) => ({
      id: m[1],
      text: m[2],
    }));
    const proposals = [];
    for (const [n, p] of (entry.proposals ?? []).entries()) {
      const { paragraph_contains: holds, ...rest } = p;
      const id = holds ? shown.find((s) => s.text.includes(holds))?.id : rest.paragraph_id;
      if (!id) {
        return fail(
          `the proposal fixture names a paragraph that is not shown (proposal ${n + 1} of ${ref})`,
        );
      }
      proposals.push({ ...rest, paragraph_id: id });
    }
    const json = { proposals, note: entry.note ?? "" };
    log(`[proposal] fixture reply ref=${ref} proposals=${proposals.length}`);
    return { ok: true, text: JSON.stringify(json), json };
  }
  return () => ({
    provider: "fixture",
    model,
    run: (req) =>
      reply(req).then(undefined, (e) =>
        fail(`the proposal fixture could not be read: ${String(e?.message ?? e)}`),
      ),
  });
}
