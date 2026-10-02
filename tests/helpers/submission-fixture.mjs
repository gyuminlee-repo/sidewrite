// An invented paper for the submission checklist tests: one comment with one
// request, answered by one accepted change, with a letter that passes every
// pure check. Nothing here is from a real manuscript.
import { snapshotHash } from "../../daemon/response-quotes.mjs";
import { withDefaults } from "../../daemon/journal-profile.mjs";

export const QUOTE = "The incubation temperature was held at 30 °C.";
export const SECTION = "2. Materials and Methods";
export const FILE = "paper.docx";

// What the pane read: the text, the comments of the document, the headings.
export const liveOf = ({ paragraph = {}, comments = [], format = null } = {}) => ({
  snapshot: {
    paragraphs: [
      { index: 0, text: SECTION, w_ins: 0, w_del: 0 },
      {
        index: 1,
        text: `Flasks were shaken. ${QUOTE} Samples were drawn daily.`,
        w_ins: 0,
        w_del: 0,
        ...paragraph,
      },
    ],
    comments,
  },
  headings: [{ index: 0, level: 1, text: SECTION }],
  format,
  // Word said the document has no unsaved changes.
  saved: true,
});

// The record of the paper in the state where every row is ready. mutate
// plants one defect.
export function recordOf(mutate = () => {}, live = liveOf()) {
  const revision = {
    comments: [
      {
        ref: "R1.1",
        reviewer: 1,
        text: "Please state the unit.",
        requirements: [{ id: "q1", text: "state the unit" }],
        questions: [],
        declined: null,
        changes: [
          {
            id: "R1.1-1",
            kind: "replace",
            find: "degrees",
            text: "°C",
            state: "accepted",
            paragraph_index: 1,
            rationale: { requirement_ids: ["q1"] },
          },
        ],
        figure_changes: [],
        response: {
          text: "We now state the unit.",
          previous: null,
          declined_requirements: [],
          stale: [],
          citations: [
            {
              change_ids: ["R1.1-1"],
              kind: "revised",
              section: SECTION,
              paragraph_index: 1,
              quote: QUOTE,
              lines: { start: 9, end: 9, start_page: 1, end_page: 1, match_count: 1 },
              lines_pdf: "sha",
              problem: null,
            },
          ],
        },
      },
    ],
    ledger: [],
    response_render: {
      pdf_sha: "sha",
      snapshot_hash: snapshotHash(live.snapshot),
      at: "2000-01-01T00:00:00.000Z",
      lines_indexed: 40,
    },
    response_export: null,
    artifacts: null,
  };
  mutate(revision);
  return revision;
}

export const profileOf = (extra = {}) => withDefaults(extra);

// A comment of the document as the snapshot lists it.
export const docComment = (id, content) => ({ id, content, resolved: false });

export const rowOf = (rows, id) => rows.find((r) => r.id === id);
