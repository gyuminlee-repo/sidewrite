import { emptyRevision, RevisionSchema } from "../../daemon/revision-store.mjs";

// Invented workshop prose only. About 1 MB on disk, 20 requests, 120 edits,
// and 150 paragraphs, including the reviewed and tracked views Word sends.
export function largeRevision(doc) {
  const paragraphs = Array.from({ length: 150 }, (_, index) => {
    const lead = `Workshop note ${index}: the blue cart crossed the square. `;
    const text = lead + "The clerk checked the lamps and stacked the empty boxes. ".repeat(34);
    return { index, id: `p${index}`, text, reviewed_text: text, w_ins: 0, w_del: 0, tracked: [] };
  });
  const comments = Array.from({ length: 20 }, (_, i) => ({
    ref: `R1.${i + 1}`,
    reviewer: 1,
    text: `Clarify workshop notes ${i * 6} through ${i * 6 + 5}.`,
    requirements: [],
    status: "in_progress",
    blocked: [],
    changes: Array.from({ length: 6 }, (_, j) => {
      const index = i * 6 + j;
      const p = paragraphs[index];
      const before = p.text;
      p.reviewed_text = before.replace("blue cart", "green cart");
      p.text = before.replace("blue cart", "green cartblue cart");
      p.w_ins = 1;
      p.w_del = 1;
      p.tracked = [
        { kind: "ins", text: "green cart", start: before.indexOf("blue cart"), marks: 1 },
        { kind: "del", text: "blue cart", start: before.indexOf("blue cart") + 10, marks: 1 },
      ];
      return {
        id: `R1.${i + 1}-${j + 1}`,
        word_comment_id: String(index),
        paragraph_index: index,
        paragraph_id: p.id,
        kind: "replace",
        find: "blue cart",
        text: "green cart",
        before,
        after: p.reviewed_text,
        context: { left: `${index}: the `, right: " crossed the square." },
        rationale: {
          what: "Clarified the cart.",
          why: "Use the intended color.",
          addresses: "The note.",
        },
        created_at: "2026-10-15T00:00:00Z",
      };
    }),
  }));
  for (const c of comments) {
    c.response = {
      text: "The workshop notes now identify the intended cart.",
      citations: c.changes.map((ch) => ({
        change_ids: [ch.id],
        paragraph_index: ch.paragraph_index,
        quote: paragraphs[ch.paragraph_index].reviewed_text.split(".")[0] + ".",
      })),
    };
  }
  const revision = RevisionSchema.parse({
    ...emptyRevision(doc),
    comments,
    snapshot: {
      taken_at: "2026-10-15T00:00:00Z",
      paragraphs: paragraphs.map((p) => p.reviewed_text),
      ids: paragraphs.map((p) => p.id),
    },
  });
  const snapshot = {
    paragraphs,
    comments: paragraphs.slice(0, 120).map((p) => ({
      id: String(p.index),
      paragraph_index: p.index,
      paragraph_id: p.id,
      content: `[R1.${Math.floor(p.index / 6) + 1}] Clarified the cart.`,
    })),
    has_fields: false,
  };
  return { revision, snapshot };
}
