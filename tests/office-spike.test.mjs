import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { spikeRenderPdfPath, saveSpikeRenderPdf } from "../daemon/office-tools.mjs";

test("spikeRenderPdfPath derives <stem>.render.pdf from each doc path form", () => {
  const cwd = "/work";
  assert.equal(spikeRenderPdfPath(cwd, "D:\\papers\\draft v2.docx"), "/work/draft v2.render.pdf");
  assert.equal(spikeRenderPdfPath(cwd, "/home/u/ms.DOCX"), "/work/ms.render.pdf");
  assert.equal(
    spikeRenderPdfPath(cwd, "file:///C:/a/my%20paper.docm"),
    "/work/my paper.render.pdf",
  );
  assert.equal(spikeRenderPdfPath(cwd, "https://x/y/z.docx?web=1"), "/work/z.render.pdf");
  assert.equal(spikeRenderPdfPath(cwd, null), "/work/document.render.pdf");
  assert.equal(spikeRenderPdfPath(cwd, "bad%E0.docx"), "/work/bad%E0.render.pdf");
});

test("saveSpikeRenderPdf writes decoded bytes and returns metadata only", async () => {
  const dir = await mkdtemp(join(tmpdir(), "spike-"));
  const bytes = Buffer.from("%PDF-1.7\nfake\n%%EOF\n");
  const r = await saveSpikeRenderPdf({
    cwd: dir,
    activeDoc: "C:\\x\\ms.docx",
    result: {
      ok: true,
      pdf_base64: bytes.toString("base64"),
      byte_length: bytes.length,
      slice_count: 1,
      body_text_length: 42,
    },
  });
  assert.equal(r.ok, true);
  assert.equal(r.path, join(dir, "ms.render.pdf"));
  assert.equal(r.bytes, bytes.length);
  assert.equal(r.slice_count, 1);
  assert.equal(r.body_text_length, 42);
  assert.equal("pdf_base64" in r, false);
  assert.deepEqual(await readFile(r.path), bytes);
});

test("saveSpikeRenderPdf passes a taskpane failure through unchanged", async () => {
  const fail = { ok: false, code: 5001, message: "Internal Error", stage: "getFileAsync" };
  assert.deepEqual(await saveSpikeRenderPdf({ cwd: "/nope", activeDoc: null, result: fail }), fail);
});
