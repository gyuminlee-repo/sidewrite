// Read-only synthetic reproduction for Sidewrite PR36 head e01aecc9e34860288bac93023a4a23a09059efa2.
// Usage: node repro-pr36-answer-starvation.mjs /path/to/sidewrite-checkout
// Imports production selection helpers. No model, Office, network or document writes.
import assert from "node:assert/strict";
import { resolve, join } from "node:path";
import { pathToFileURL } from "node:url";
const root = resolve(process.argv[2] ?? ".");
const { followPlaces, followState, MAX_PLACES } = await import(
  pathToFileURL(join(root, "taskpane/shared/follow-places.js"))
);
function scenario(name, paragraphs, expectedFirst) {
  const ids = paragraphs.map((_, i) => `k${i}`);
  const comment = {
    ref: "R1.1",
    changes: [],
    questions: ids.map((id, i) => ({
      id: `a${i + 1}`,
      paragraph_id: id,
      answer: "Author answer.",
    })),
  };
  const first = followPlaces(comment, paragraphs, ids);
  assert.equal(
    first.length,
    expectedFirst,
    "use a checkout of the reviewed PR36 head; initial candidate selection differs",
  );
  comment.follow_through = {
    candidate_ids: first.map((p) => p.id),
    change_ids: [],
    questions: 0,
    seen_change_ids: [],
    seen_answer_ids: first.flatMap((p) => p.answer_ids),
  };
  const unseen = comment.questions
    .filter((q) => !comment.follow_through.seen_answer_ids.includes(q.id))
    .map((q) => q.id);
  const next = followState(comment, { paragraphs, ids });
  assert.ok(unseen.length, "fixture must retain an unchecked answer");
  assert.equal(
    next.offer,
    null,
    "reviewed head incorrectly hides the unchecked answer",
  );
  return { name, firstCount: first.length, unseen, next };
}
console.log(
  JSON.stringify(
    [
      scenario(
        "place limit",
        Array.from(
          { length: MAX_PLACES + 1 },
          (_, i) => `This is paragraph ${i} with a sentence.`,
        ),
        MAX_PLACES,
      ),
      scenario(
        "character limit",
        ["a".repeat(25000) + ".", "b".repeat(25000) + "."],
        1,
      ),
    ],
    null,
    2,
  ),
);
