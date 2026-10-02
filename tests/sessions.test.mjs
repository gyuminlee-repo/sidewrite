// Unit tests for daemon/sessions.mjs — per-(host, workspace) session-id
// bookkeeping. Tests override $HOME via the environment so they touch a
// fresh tmpdir instead of ~/.claude/. sessions.mjs reads homedir() at
// import time, so the override must happen before the import — we use
// dynamic import() inside the test to do that cleanly.

import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, mkdir, writeFile, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

async function withFakeHome(fn) {
  const fakeHome = await mkdtemp(join(tmpdir(), "cc-office-sess-test-"));
  const prevHome = process.env.HOME;
  const prevUserProfile = process.env.USERPROFILE;
  process.env.HOME = fakeHome;
  process.env.USERPROFILE = fakeHome;
  try {
    // Cache-bust the dynamic import so each test gets a freshly-evaluated
    // module reading the new $HOME.
    const cacheBuster = "?t=" + Date.now() + Math.random();
    const mod = await import("../daemon/sessions.mjs" + cacheBuster);
    await fn(mod, fakeHome);
  } finally {
    if (prevHome === undefined) delete process.env.HOME;
    else process.env.HOME = prevHome;
    if (prevUserProfile === undefined) delete process.env.USERPROFILE;
    else process.env.USERPROFILE = prevUserProfile;
    await rm(fakeHome, { recursive: true, force: true });
  }
}

async function readStateFile(fakeHome) {
  try {
    const raw = await readFile(join(fakeHome, ".claude", "office-addins", "sessions.json"), "utf8");
    return JSON.parse(raw);
  } catch (e) {
    if (e.code === "ENOENT") return null;
    throw e;
  }
}

test("getSessionId returns null when no state file exists", async () => {
  await withFakeHome(async ({ getSessionId }) => {
    assert.equal(await getSessionId("word", "/tmp/whatever"), null);
  });
});

test("save → get round-trip is keyed by host AND cwd", async () => {
  await withFakeHome(async ({ saveSessionId, getSessionId }) => {
    await saveSessionId("word", "/tmp/folderA", "word-sid");
    await saveSessionId("excel", "/tmp/folderA", "excel-sid");

    // Same folder, different host → independent session ids.
    assert.equal(await getSessionId("word", "/tmp/folderA"), "word-sid");
    assert.equal(await getSessionId("excel", "/tmp/folderA"), "excel-sid");
    // A host with no saved session here → null.
    assert.equal(await getSessionId("word", "/tmp/folderB"), null);
  });
});

test("normalizeHost: a bad host is a no-op / null", async () => {
  await withFakeHome(async ({ saveSessionId, getSessionId }) => {
    await saveSessionId("powerpoint", "/tmp/x", "nope");
    assert.equal(await getSessionId("powerpoint", "/tmp/x"), null);
    assert.equal(await getSessionId("word", "/tmp/x"), null);
  });
});

test("touchFolder does not clobber a saved session id", async () => {
  await withFakeHome(async ({ touchFolder, saveSessionId, getSessionId }) => {
    await touchFolder("/tmp/folderB");
    assert.equal(await getSessionId("word", "/tmp/folderB"), null);

    await saveSessionId("word", "/tmp/folderB", "w1");
    await touchFolder("/tmp/folderB");
    assert.equal(
      await getSessionId("word", "/tmp/folderB"),
      "w1",
      "touch must not blow away a saved session id",
    );
  });
});

test("touchFolder silently drops disallowed system home children (macOS only)", async () => {
  if (process.platform !== "darwin") return; // The deny-list is macOS-specific.
  await withFakeHome(async ({ touchFolder }, fakeHome) => {
    await touchFolder(join(fakeHome, "Library"));
    const state = await readStateFile(fakeHome);
    // touchFolder returns early for $HOME children — nothing persisted.
    assert.ok(
      state === null || !state.folders || Object.keys(state.folders).length === 0,
      "an OS-managed $HOME child must never be persisted",
    );
  });
});

test("migrates a v1 sessions.json: the un-hostable old id is dropped", async () => {
  await withFakeHome(async ({ getSessionId }, fakeHome) => {
    const dir = join(fakeHome, ".claude", "office-addins");
    await mkdir(dir, { recursive: true });
    const v1 = {
      version: 1,
      folders: {
        "/tmp/legacy": {
          session_id: "old-uuid",
          last_used: "2026-05-01T00:00:00.000Z",
          display_name: "legacy",
        },
      },
    };
    await writeFile(join(dir, "sessions.json"), JSON.stringify(v1));
    // The v1 id can't be attributed to a host, so it's intentionally not
    // resumable post-migration.
    assert.equal(await getSessionId("word", "/tmp/legacy"), null);
    assert.equal(await getSessionId("excel", "/tmp/legacy"), null);
  });
});

// ---- a hint per document, for the files of a revision with parts ----------------

test("a document with a hint of its own never reads the hint of its folder, and does not change it", async () => {
  await withFakeHome(async ({ saveSessionId, getSessionId }, fakeHome) => {
    const cwd = "/tmp/paperdir";
    const main = `${cwd}/paper.docx`;
    const si = `${cwd}/SI.docx`;
    // The manuscript ran a turn before it had a supplementary file.
    await saveSessionId("word", cwd, "folder-sid");
    // The supplement asks by its own path: nothing yet, not the folder's.
    assert.equal(await getSessionId("word", cwd, si), null);
    assert.equal(await getSessionId("word", cwd, main), null);
    await saveSessionId("word", cwd, "si-sid", si);
    await saveSessionId("word", cwd, "main-sid", main);
    assert.equal(await getSessionId("word", cwd, si), "si-sid");
    assert.equal(await getSessionId("word", cwd, main), "main-sid");
    // Any other document of the folder still gets the hint stored before.
    assert.equal(await getSessionId("word", cwd), "folder-sid");
    assert.equal(await getSessionId("word", cwd, null), "folder-sid");
    // Per host, as the folder hint is.
    assert.equal(await getSessionId("excel", cwd, si), null);
    const state = await readStateFile(fakeHome);
    assert.deepEqual(state.folders[cwd].sessions, { word: "folder-sid" });
    assert.deepEqual(state.folders[cwd].docs, {
      [si]: { word: "si-sid" },
      [main]: { word: "main-sid" },
    });
    assert.equal(state.version, 2);
  });
});

test("a sessions.json from before the hints per document is read as it is", async () => {
  await withFakeHome(async ({ getSessionId, saveSessionId }, fakeHome) => {
    const dir = join(fakeHome, ".claude", "office-addins");
    await mkdir(dir, { recursive: true });
    const before = {
      version: 2,
      folders: {
        "/tmp/old": {
          last_used: "2026-05-01T00:00:00.000Z",
          display_name: "old",
          sessions: { word: "kept-sid" },
        },
      },
    };
    await writeFile(join(dir, "sessions.json"), JSON.stringify(before, null, 2) + "\n");
    assert.equal(await getSessionId("word", "/tmp/old"), "kept-sid");
    assert.equal(await getSessionId("word", "/tmp/old", "/tmp/old/SI.docx"), null);
    // Reading writes nothing: the file is byte for byte what it was.
    assert.deepEqual(await readStateFile(fakeHome), before);
    // A folder hint saved now adds no member.
    await saveSessionId("word", "/tmp/old", "next-sid");
    assert.deepEqual(Object.keys((await readStateFile(fakeHome)).folders["/tmp/old"]).sort(), [
      "display_name",
      "last_used",
      "sessions",
    ]);
  });
});

test("the line of a turn names the document, the session, the id asked for, and how the session came about", async () => {
  await withFakeHome(async ({ turnSessionLine }) => {
    // A new session: no id was asked for.
    assert.equal(
      turnSessionLine({ doc: "SI.docx", sessionId: "aaa" }),
      "[agent] turn doc=SI.docx session=aaa asked=- (started)",
    );
    // The first turn of a loop that resumed the stored session.
    assert.equal(
      turnSessionLine({ doc: "paper.docx", sessionId: "bbb", asked: "bbb" }),
      "[agent] turn doc=paper.docx session=bbb asked=bbb (resumed)",
    );
    // The SDK answered with another id than the one asked for: the line
    // still says whose transcript was asked for.
    assert.equal(
      turnSessionLine({ doc: "SI.docx", sessionId: "ccc", asked: "bbb" }),
      "[agent] turn doc=SI.docx session=ccc asked=bbb (started)",
    );
    // A later turn of the same loop.
    assert.equal(
      turnSessionLine({ doc: "SI.docx", sessionId: "aaa", known: "aaa" }),
      "[agent] turn doc=SI.docx session=aaa asked=- (continued)",
    );
    assert.equal(
      turnSessionLine({ sessionId: "aaa" }),
      "[agent] turn doc=- session=aaa asked=- (started)",
    );
  });
});
