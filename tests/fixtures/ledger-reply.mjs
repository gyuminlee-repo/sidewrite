// A fake model for the ledger LLM check (daemon/ledger-check.mjs): reads the
// request the way a model would and answers in the form it asks for.
import { normalizeParaText } from "../../daemon/revision-classify.mjs";

// { batch_id, paragraphs: [{ id, text }] } of a ledgerPrompt (also of the
// retry prompt, which repeats it).
export function readLedgerPrompt(prompt) {
  const batch_id = /"batch_id":"([^"]+)"/.exec(prompt)?.[1] ?? null;
  const paragraphs = [];
  const lines = String(prompt).split("\n");
  for (let i = lines.indexOf("Paragraphs:") + 1; i > 0 && i < lines.length; i++) {
    const m = /^\[(\d+)\] ?(.*)$/.exec(lines[i]);
    if (!m) break;
    paragraphs.push({ id: Number(m[1]), text: m[2] });
  }
  return { batch_id, paragraphs };
}

// The reply of a model that keeps to the form: one verdict per paragraph of
// the batch, "ok" unless named. paras: the document paragraphs ({ index,
// text, reviewed_text }); violations: [{ paragraph_index, quote, reason }]
// and unsure: [{ paragraph_index, reason }] by DOCUMENT index, as the tests
// think of them. An entry for a paragraph that is not in this batch is left
// out (the model was not asked about it).
export function ledgerJson(req, paras, { violations = [], unsure = [] } = {}) {
  const { batch_id, paragraphs } = readLedgerPrompt(req.prompt);
  const idOf = (index) => {
    const p = paras.find((x) => x.index === index);
    const text = p ? normalizeParaText(p.reviewed_text ?? p.text) : null;
    return paragraphs.find((q) => q.text === text)?.id ?? null;
  };
  const here = (list) =>
    list
      .map(({ paragraph_index, ...rest }) => ({ paragraph_id: idOf(paragraph_index), ...rest }))
      .filter((x) => x.paragraph_id !== null);
  const v = here(violations);
  const u = here(unsure);
  const verdict = (id) =>
    v.some((x) => x.paragraph_id === id)
      ? "violation"
      : u.some((x) => x.paragraph_id === id)
        ? "unsure"
        : "ok";
  return {
    batch_id,
    verdicts: paragraphs.map((p) => [p.id, verdict(p.id)]),
    violations: v,
    unsure: u,
  };
}

export const ledgerReply = (req, paras, marks) => ({
  ok: true,
  json: ledgerJson(req, paras, marks),
});
