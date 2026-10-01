// Inputs of buildReviewPrompt whose whole output is kept in
// tests/fixtures/review_prompt_outputs.json (written from the code before the
// comment and ledger blocks were shared with the proposal prompt). Invented
// text.
const comment = {
  ref: "R1.3",
  reviewer: 1,
  section: "Round 2 / Minor comments",
  text: "In section 3, please say how growth was followed,\nand describe the substrate as novel.",
  requirements: [
    { id: "q1", text: "say how growth was followed" },
    { id: "q2", text: "describe the substrate as novel" },
  ],
};
const ledger = [
  { id: "L1", kind: "instruction", text: "in vitro 는 하이픈 없이 쓴다", scope: "manuscript" },
  { id: "L2", kind: "forbid", text: "novel 을 쓰지 않는다", scope: "manuscript" },
  { id: "L3", kind: "instruction", text: "다중비교 문장을 지우지 않는다", scope: "manuscript" },
  { id: "L4", kind: "instruction", text: "응답서는 두괄식", scope: "response" },
];
const paragraphs = [
  { id: "U9", index: 9, text: "Growth was followed at 600 nm." },
  { index: 10, text: "The substrate was a synthetic carbon source." },
];

export const CASES = {
  full: { comment, ledger, paragraphs },
  no_selection_english: { comment, ledger, lang: "English" },
  forbid_omitted: { comment, ledger, paragraphs, includeForbid: false },
  editor_bare: { comment: { ref: "E.1", reviewer: 0, text: "Shorten the abstract." } },
};
