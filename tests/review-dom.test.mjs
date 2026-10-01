// Static check: every element id the Review tab script looks up exists in
// the Review section of the Word page. A typo here would fail only in Word.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const read = (p) => readFile(new URL(p, import.meta.url), "utf8");

test("review-pane.js ids all exist in the Review section of word/index.html", async () => {
  const js = await read("../taskpane/shared/review-pane.js");
  const html = await read("../taskpane/word/index.html");
  const start = html.indexOf('data-tab="review" hidden');
  const end = html.indexOf("<!-- Setup tab content -->");
  assert.ok(start !== -1 && end > start, "Review section not found");
  const section = html.slice(start, end);
  const wanted = new Set([
    ...[...js.matchAll(/\$\("([^"]+)"\)/g)].map((m) => m[1]),
    ...[...js.matchAll(/getElementById\("([^"]+)"\)/g)].map((m) => m[1]),
  ]);
  const present = new Set([...section.matchAll(/id="([^"]+)"/g)].map((m) => m[1]));
  const missing = [...wanted].filter((id) => !present.has(id));
  assert.ok(wanted.size > 0);
  assert.deepEqual(missing, []);
  console.log(`checked ${wanted.size} ids`);
});

// An LLM check judges each paragraph on its own (daemon/ledger-check.mjs):
// the ledger form says so, next to where the check type is chosen.
test("the ledger form says that an LLM check reads paragraph by paragraph", async () => {
  const html = await read("../taskpane/word/index.html");
  const start = html.indexOf('id="ledger-check-type"');
  const end = html.indexOf('id="review-ledger"');
  assert.ok(start !== -1 && end > start);
  const text = html.slice(start, end).replace(/\s+/g, " ");
  assert.match(text, /An LLM check reads the manuscript one paragraph at a time/);
});
