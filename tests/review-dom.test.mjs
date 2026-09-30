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
