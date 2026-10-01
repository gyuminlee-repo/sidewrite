import { test } from "node:test";
import assert from "node:assert/strict";
import {
  toChunks,
  createAssembler,
  getDocBytes,
  DOC_MAX_BYTES,
  SAVE_FIRST,
  DISK_NOTICE,
} from "../daemon/doc-bytes.mjs";

const bytes = (n, seed = 7) =>
  Buffer.from(Array.from({ length: n }, (_, i) => (i * seed + 3) % 256));

test("bytes cut into chunks come back the same through the assembler", () => {
  for (const [n, size] of [
    [0, 4],
    [1, 4],
    [4, 4],
    [5, 4],
    [10_000, 999],
  ]) {
    const data = bytes(n);
    const chunks = toChunks(data, size);
    assert.equal(chunks.length, Math.max(1, Math.ceil(n / size)));
    assert.deepEqual(
      chunks.map((c) => c.seq),
      chunks.map((_, i) => i),
    );
    assert.deepEqual(
      chunks.map((c) => c.last),
      chunks.map((_, i) => i === chunks.length - 1),
    );
    const asm = createAssembler();
    let done;
    for (const c of chunks) done = asm.add(c);
    assert.equal(done.ok, true);
    assert.equal(done.done, true);
    assert.ok(done.bytes.equals(data));
  }
});

test("chunks are counted as they come in", () => {
  const asm = createAssembler();
  const [a, b, c] = toChunks(bytes(10), 4);
  assert.deepEqual(asm.add(a), { ok: true, done: false, received: 4 });
  assert.deepEqual(asm.add(b), { ok: true, done: false, received: 8 });
  const end = asm.add(c);
  assert.deepEqual([end.done, end.received, end.bytes.length], [true, 10, 10]);
});

test("a chunk out of order, repeated, after the end, or not base64 is refused and ends the transfer", () => {
  const chunks = toChunks(bytes(12), 4);
  const cases = {
    "a gap": [chunks[0], chunks[2]],
    "a repeat": [chunks[0], chunks[0]],
    "a start past zero": [chunks[1]],
    "a chunk after the last": [
      chunks[0],
      chunks[1],
      chunks[2],
      { seq: 3, last: true, data_b64: "AAAA" },
    ],
    "text that is not base64": [{ seq: 0, last: true, data_b64: "not base64!" }],
    "base64 with a bad length": [{ seq: 0, last: true, data_b64: "AAA" }],
    "a missing sequence number": [{ last: true, data_b64: "AAAA" }],
    "data that is not text": [{ seq: 0, last: true, data_b64: 42 }],
    nothing: [null],
  };
  for (const [name, list] of Object.entries(cases)) {
    const asm = createAssembler();
    const results = list.map((c) => asm.add(c));
    const bad = results.at(-1);
    assert.equal(bad.ok, false, name);
    assert.equal(typeof bad.error, "string", name);
    // Once refused, nothing more is taken.
    assert.equal(asm.add(chunks[0]).ok, false, name);
  }
});

test("the size cap counts decoded bytes and refuses the chunk that crosses it", () => {
  const asm = createAssembler({ maxBytes: 10 });
  const chunks = toChunks(bytes(12), 4);
  assert.equal(asm.add(chunks[0]).ok, true);
  assert.equal(asm.add(chunks[1]).ok, true);
  const over = asm.add(chunks[2]);
  assert.deepEqual([over.ok, /larger than 10 bytes/.test(over.error)], [false, true]);
  assert.equal(DOC_MAX_BYTES, 100 * 1024 * 1024);
  // A single oversized chunk is refused before it is decoded.
  const big = createAssembler({ maxBytes: 3 });
  assert.equal(big.add({ seq: 0, last: true, data_b64: "A".repeat(4000) }).ok, false);
});

test("the document comes from the pane when the pane can give it", async () => {
  const data = bytes(50);
  let read = 0;
  const r = await getDocBytes({
    fromPane: async () => ({ ok: true, chunks: toChunks(data, 16) }),
    readFile: async () => (read++, Buffer.alloc(0)),
    docPath: "/work/paper.docx",
  });
  assert.equal(r.ok, true);
  assert.equal(r.source, "pane");
  assert.ok(r.bytes.equals(data));
  assert.equal(r.notice, undefined);
  assert.equal(read, 0);
});

test("when the pane cannot give it, the saved file is read and the result says so", async () => {
  const data = bytes(30);
  for (const fromPane of [
    async () => ({ ok: false, error: "getFileAsync is not supported" }),
    async () => {
      throw new Error("pane closed");
    },
    async () => ({ ok: true, chunks: [{ seq: 1, last: true, data_b64: "AAAA" }] }),
    undefined,
  ]) {
    const seen = [];
    const r = await getDocBytes({
      fromPane,
      readFile: async (p) => (seen.push(p), data),
      docPath: "/work/paper.docx",
    });
    assert.deepEqual([r.ok, r.source, r.notice], [true, "disk", DISK_NOTICE]);
    assert.ok(r.bytes.equals(data));
    assert.deepEqual(seen, ["/work/paper.docx"]);
  }
  assert.match(DISK_NOTICE, /saved/);
});

test("no pane bytes and no saved file: a plain refusal", async () => {
  const none = await getDocBytes({ fromPane: async () => ({ ok: false }), docPath: null });
  assert.deepEqual(none, { ok: false, error: SAVE_FIRST });
  assert.equal(SAVE_FIRST, "Save the document to a local folder first.");
  const gone = await getDocBytes({
    docPath: "/work/paper.docx",
    readFile: async () => {
      throw new Error("ENOENT: no such file");
    },
  });
  assert.deepEqual([gone.ok, /ENOENT/.test(gone.error)], [false, true]);
  const big = await getDocBytes({
    docPath: "/work/paper.docx",
    readFile: async () => bytes(11),
    maxBytes: 10,
  });
  assert.deepEqual([big.ok, /larger than 10 bytes/.test(big.error)], [false, true]);
});
