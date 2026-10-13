import assert from "node:assert/strict";
import { trackedRuns, countRevisionMarks } from "../../../taskpane/shared/review-view.js";
import {
  changeContext,
  classifyChange,
  classifyRevision,
} from "../../../daemon/revision-classify.mjs";
import { outsideChanges } from "../../../daemon/doc-diff.mjs";
import {
  emptyRevision,
  replaceComments,
  addExternal,
  addChange,
  addAttempt,
  markBlockedUnknown,
} from "../../../daemon/revision-store.mjs";
import { resolveUnknownOverrides } from "../../../daemon/review-service.mjs";

const rationale = {
  what: "Shorten the bridge description.",
  why: "Remove the extra word.",
  addresses: "Use a short description.",
};
const stamp = new Date("2026-10-13T00:00:00.000Z");
const prefix = "Crews installed ";
const inserted = "new steel supports ";
const old = "old piers";
const suffix = " beside the road.";
const before = prefix + inserted + old + suffix;
const after = before.replace(" supports", "");
const make = (kind, find, text, input, id) => ({
  id,
  kind,
  find,
  text,
  before: input,
  after: input.replace(find, text),
  paragraph_index: 4,
  paragraph_id: "P4",
  word_comment_id: null,
  rationale,
  context: changeContext(input, find),
});
const deletion = make("delete", " supports", "", before, "R1.1-2");
const sibling = make("replace", old, inserted, prefix + old + suffix, "R1.1-1");
const xml = (ia, da) =>
  `<pkg:package><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body><w:p><w:r><w:t>${prefix}</w:t></w:r><w:ins ${ia}><w:r><w:t>new steel </w:t></w:r></w:ins><w:del ${da}><w:r><w:delText>${old}</w:delText></w:r></w:del><w:r><w:t>${suffix}</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
const sameAuthor = xml(
  'w:id="10" w:author="Bridge editor" w:date="2026-10-12T09:00:00Z"',
  'w:id="11" w:author="Bridge editor" w:date="2026-10-12T09:00:00Z"',
);
const unrelated = xml(
  'w:id="23" w:author="Catalog editor" w:date="2026-10-12T10:00:00Z"',
  'w:id="45" w:author="Road editor" w:date="2026-10-12T11:00:00Z"',
);
const paragraph = (ooxml) => ({
  index: 4,
  id: "P4",
  text: after,
  reviewed_text: prefix + "new steel" + suffix,
  ...countRevisionMarks(ooxml),
  tracked: trackedRuns(ooxml),
});
const authorSnapshot = { paragraphs: [paragraph(sameAuthor)], comments: [] };
const unrelatedSnapshot = { paragraphs: [paragraph(unrelated)], comments: [] };
assert.deepEqual(authorSnapshot, unrelatedSnapshot);
assert.equal(classifyChange(deletion, authorSnapshot).state, "accepted");
assert.equal(classifyChange(deletion, unrelatedSnapshot).state, "accepted");
assert.equal(classifyChange(deletion, authorSnapshot, [sibling, deletion]).state, "pending");
console.log(
  "Equal classifier inputs after author/date/revision-ID differences are discarded: true",
);
console.log(
  "Author-replace tail deletion without recorded sibling: accepted (known limitation; desired pending)",
);
console.log("Independent ins/del with accepted deletion: accepted (control)");
console.log("Recorded same-paragraph replace sibling: pending (control)");
console.log("Tracked runs: " + JSON.stringify(authorSnapshot.paragraphs[0].tracked));

let revision = replaceComments(emptyRevision("synthetic-bridge.docx", stamp), [
  { ref: "R1.1", reviewer: 1, text: "Shorten the bridge description." },
]);
revision = addChange(revision, "R1.1", deletion, stamp).revision;
const headings = ["Bridge report", "Inspection notes", "Site description", "Deck summary"];
const externalItems = outsideChanges(
  [...headings, prefix + old + suffix],
  [...headings, prefix + inserted.trimEnd() + suffix],
);
assert.deepEqual(
  Object.keys(externalItems[0]).sort(),
  ["after", "before", "kind", "lost_facts", "old_index", "paragraph_index"].sort(),
);
assert.equal(externalItems[0].paragraph_index, 4);
assert.equal(externalItems[0].old_index, 4);
const withExternal = addExternal(revision, externalItems, stamp).revision;
assert.deepEqual(
  classifyRevision(revision, authorSnapshot),
  classifyRevision(withExternal, authorSnapshot),
);
console.log("External before/after entry changes classifier result: false");
console.log("External fields: " + Object.keys(withExternal.external[0]).join(", "));

// 결과 불명 시도는 의도다. 해당 추적 run을 만들었다는 확인이 아니다.
const attempt = {
  tool_call_id: "synthetic-call",
  unknown_at: stamp.toISOString(),
  comment_text: "[R1.1] Replace old piers.",
  change: sibling,
};
const added = addAttempt(
  revision,
  "R1.1",
  { args: { find: old, replace: inserted }, attempt },
  stamp,
);
const unknown = markBlockedUnknown(added.revision, "R1.1", added.blocked.id, attempt);
const recovered = resolveUnknownOverrides(unknown, authorSnapshot);
assert.equal(recovered.resolved.length, 0);
assert.equal(recovered.revision.comments[0].blocked.length, 1);
assert.equal(classifyRevision(recovered.revision, authorSnapshot)[0].state, "accepted");
assert.equal(classifyChange(deletion, unrelatedSnapshot, [sibling, deletion]).state, "pending");
console.log("Unknown replace without full new text or rationale comment: unresolved");
console.log(
  "Blindly treating that attempted replace as applied sibling changes the independent control to: pending (unsafe)",
);

// 근거 주석이 있으면 현재 복구 경로가 replace를 기록할 수 있다.
const commented = {
  ...authorSnapshot,
  comments: [{ id: "91", content: attempt.comment_text, paragraph_id: "P4", paragraph_index: 4 }],
};
const proven = resolveUnknownOverrides(unknown, commented);
assert.equal(proven.resolved.length, 1);
assert.equal(proven.revision.comments[0].blocked.length, 0);
assert.equal(classifyRevision(proven.revision, commented)[0].state, "pending");
console.log(
  "Unknown replace with matching rationale comment: resolved; subsequent classification pending",
);
console.log("All synthetic evidence assertions passed.");
