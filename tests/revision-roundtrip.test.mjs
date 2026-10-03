import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { loadRevision, saveRevision } from "../daemon/revision-store.mjs";

// The fixture was saved by the daemon of 89cdedb, before the optional fields
// of the third replay (requirements_outside[].part, rewrites_stopped). It
// holds a rewrite part, a question, an outside note and a paragraph check.
// The repository formatter lays its arrays out anew, so the bytes compared
// are those of the daemon's own layout of the same data.
test("a review file from before the replay-3 fields loads and saves byte for byte", async () => {
  const fixture = new URL("./fixtures/revision_before_replay3.json", import.meta.url);
  const raw = JSON.parse(await readFile(fixture, "utf8"));
  const text = JSON.stringify(raw, null, 2) + "\n";
  for (const c of raw.comments) {
    assert.equal("rewrites_stopped" in c, false, "the fixture is the old format");
    for (const n of c.requirements_outside) assert.equal("part" in n, false);
  }
  const dir = await mkdtemp(join(tmpdir(), "replay3-old-"));
  try {
    const path = join(dir, "paper.revision.json");
    await writeFile(path, text);
    const loaded = await loadRevision(path, "x");
    assert.equal(loaded.version, 1);
    await saveRevision(path, loaded, new Date(raw.updated_at));
    assert.equal(await readFile(path, "utf8"), text);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
