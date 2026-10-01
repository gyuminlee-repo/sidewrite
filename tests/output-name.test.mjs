import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { outputName, freeOutputIndex, exists } from "../daemon/output-name.mjs";

test("outputName: the counter goes in front of the last extension", () => {
  assert.equal(outputName("/m/paper.response.docx", 1), "/m/paper.response.docx");
  assert.equal(outputName("/m/paper.response.docx", 2), "/m/paper.response-2.docx");
  assert.equal(outputName("/m/paper.response.R1.txt", 12), "/m/paper.response.R1-12.txt");
  assert.equal(outputName("C:\\m\\paper.response.docx", 3), "C:\\m\\paper.response-3.docx");
  // No extension, a dot only in a folder name, a dot file.
  assert.equal(outputName("/m/letter", 2), "/m/letter-2");
  assert.equal(outputName("/m.d/letter", 2), "/m.d/letter-2");
  assert.equal(outputName("/m/.letter", 2), "/m/.letter-2");
});

test("freeOutputIndex: the first name none of whose group files is there", async () => {
  const dir = await mkdtemp(join(tmpdir(), "outname-"));
  const path = join(dir, "paper.response.docx");
  const group = (p) => [p, p.replace(/\.docx$/, ".model.json")];
  assert.equal(await freeOutputIndex(path, { group }), 1);
  await writeFile(path, "x");
  assert.equal(await freeOutputIndex(path, { group }), 2);
  // A sidecar alone takes the name as well.
  await writeFile(join(dir, "paper.response-2.model.json"), "x");
  assert.equal(await freeOutputIndex(path, { group }), 3);
  assert.equal(await freeOutputIndex(path), 2);
  assert.equal(await freeOutputIndex(path, { from: 5 }), 5);
  assert.deepEqual([await exists(path), await exists(join(dir, "nope"))], [true, false]);
});
