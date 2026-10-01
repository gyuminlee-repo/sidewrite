import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, readdir } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createUploads } from "../daemon/uploads.mjs";
import { toChunks } from "../daemon/doc-bytes.mjs";

const data = Buffer.from("PK\u0003\u0004 an invented baseline file ".repeat(40));
const newDir = () => mkdtemp(join(tmpdir(), "uploads-"));
const send = async (uploads, id, chunks) => {
  let r;
  for (const c of chunks) {
    r = await uploads.accept({ upload_id: id, ...c });
    if (!r.ok) break;
  }
  return r;
};

test("chunks in order become one file in the upload folder; the path comes with the last chunk", async () => {
  const dir = await newDir();
  const uploads = createUploads({ dir });
  const chunks = toChunks(data, 300);
  assert.ok(chunks.length > 2);
  const first = await uploads.accept({ upload_id: "base-1", ...chunks[0] });
  assert.deepEqual(first, { ok: true, received: 300 });
  assert.deepEqual(await readdir(dir), []);
  const done = await send(uploads, "base-1", chunks.slice(1));
  assert.deepEqual(done, { ok: true, received: data.length, path: join(dir, "base-1.docx") });
  assert.ok((await readFile(done.path)).equals(data));
  assert.equal(uploads.pathOf("base-1"), done.path);
  assert.equal(uploads.pathOf("other"), null);
});

test("two uploads at once do not mix", async () => {
  const dir = await newDir();
  const uploads = createUploads({ dir });
  const a = toChunks(Buffer.from("aaaaaaaa"), 4);
  const b = toChunks(Buffer.from("bbbbbbbbbbbb"), 4);
  await uploads.accept({ upload_id: "a", ...a[0] });
  await uploads.accept({ upload_id: "b", ...b[0] });
  await uploads.accept({ upload_id: "a", ...a[1] });
  await send(uploads, "b", b.slice(1));
  assert.equal(String(await readFile(join(dir, "a.docx"))), "aaaaaaaa");
  assert.equal(String(await readFile(join(dir, "b.docx"))), "bbbbbbbbbbbb");
});

test("a chunk out of order or over the size cap is refused and nothing is written", async () => {
  const dir = await newDir();
  const uploads = createUploads({ dir, maxBytes: 500 });
  const chunks = toChunks(data, 300);
  await uploads.accept({ upload_id: "gap", ...chunks[0] });
  const gap = await uploads.accept({ upload_id: "gap", ...chunks[2] });
  assert.deepEqual([gap.ok, /expected/.test(gap.error)], [false, true]);
  // The upload is over: its next chunk starts nothing.
  assert.equal((await uploads.accept({ upload_id: "gap", ...chunks[1] })).ok, false);
  const big = await send(uploads, "big", chunks);
  assert.deepEqual([big.ok, /larger than 500 bytes/.test(big.error)], [false, true]);
  assert.deepEqual(await readdir(dir), []);
  assert.equal(uploads.pathOf("big"), null);
});

test("an upload id is a plain name: it cannot leave the folder", async () => {
  const dir = await newDir();
  const uploads = createUploads({ dir });
  const [only] = toChunks(Buffer.from("x"), 4);
  for (const upload_id of ["../x", "a/b", "a\\b", "", "x".repeat(65), null, 5, "a.b", ".."]) {
    const r = await uploads.accept({ upload_id, ...only });
    assert.equal(r.ok, false, String(upload_id));
  }
  assert.deepEqual(await readdir(dir), []);
  assert.equal((await uploads.accept(null)).ok, false);
});

test("only a few uploads are held at a time; a finished or dropped one makes room", async () => {
  const dir = await newDir();
  const uploads = createUploads({ dir, maxOpen: 2 });
  const chunks = toChunks(Buffer.from("12345678"), 4);
  assert.equal((await uploads.accept({ upload_id: "u1", ...chunks[0] })).ok, true);
  assert.equal((await uploads.accept({ upload_id: "u2", ...chunks[0] })).ok, true);
  const third = await uploads.accept({ upload_id: "u3", ...chunks[0] });
  assert.deepEqual([third.ok, /too many uploads/.test(third.error)], [false, true]);
  await uploads.accept({ upload_id: "u1", ...chunks[1] });
  assert.equal((await uploads.accept({ upload_id: "u3", ...chunks[0] })).ok, true);
  uploads.drop("u2");
  assert.equal((await uploads.accept({ upload_id: "u4", ...chunks[0] })).ok, true);
});

test("a file that cannot be written is an error, not a path", async () => {
  const uploads = createUploads({
    dir: "/nowhere",
    writeFile: async () => {
      throw new Error("EACCES: permission denied");
    },
    mkdir: async () => {},
  });
  const r = await send(uploads, "u", toChunks(Buffer.from("x"), 4));
  assert.deepEqual([r.ok, /EACCES/.test(r.error), "path" in r], [false, true, false]);
  assert.equal(uploads.pathOf("u"), null);
});

test("a finished upload that is sent again keeps its record: the same bytes succeed, other bytes are refused", async () => {
  const dir = await newDir();
  const uploads = createUploads({ dir });
  const chunks = toChunks(data, 300);
  const first = await send(uploads, "base-1", chunks);
  assert.equal(first.ok, true);
  // The answer to the last chunk was lost, so the pane sends the file again.
  const begun = await uploads.accept({ upload_id: "base-1", ...chunks[0] });
  assert.equal(begun.ok, true);
  assert.equal(
    uploads.pathOf("base-1"),
    first.path,
    "the record is gone while the file is sent again",
  );
  const again = await send(uploads, "base-1", chunks.slice(1));
  assert.deepEqual(again, first);
  assert.equal(uploads.pathOf("base-1"), first.path);
  assert.ok((await readFile(first.path)).equals(data));
  // Other bytes under the same id: an error that says so, and the first file and its record stay.
  const other = await send(uploads, "base-1", toChunks(Buffer.from("something else"), 300));
  assert.deepEqual([other.ok, /different content/.test(other.error)], [false, true]);
  assert.equal(uploads.pathOf("base-1"), first.path);
  assert.ok((await readFile(first.path)).equals(data));
  // A second attempt that fails half way does not touch the record either.
  await uploads.accept({ upload_id: "base-1", ...chunks[0] });
  assert.equal((await uploads.accept({ upload_id: "base-1", ...chunks[2] })).ok, false);
  assert.equal(uploads.pathOf("base-1"), first.path);
  assert.deepEqual(await readdir(dir), ["base-1.docx"]);
});
