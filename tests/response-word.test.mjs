// exportRenderPdf and collectHeadings (taskpane/shared/tools-word.js) on fake
// Office and Word objects: the PDF comes back whole, the file handle is always
// closed, and the headings carry the paragraph index of the review snapshot.
import { test } from "node:test";
import assert from "node:assert/strict";

const FAILED = "failed";
// slices: the byte arrays getSliceAsync gives; failAt: "open" | slice index | null.
function fakeOffice(slices, { failAt = null, closeFails = false } = {}) {
  const state = { closed: 0, asked: null };
  const ok = (value) => ({ status: "succeeded", value });
  const bad = (message) => ({ status: FAILED, error: new Error(message) });
  const file = {
    sliceCount: slices.length,
    getSliceAsync: (i, cb) => cb(failAt === i ? bad(`slice ${i} failed`) : ok({ data: slices[i] })),
    closeAsync: (cb) => {
      state.closed += 1;
      cb(closeFails ? bad("close failed") : ok(undefined));
    },
  };
  globalThis.Office = {
    AsyncResultStatus: { Failed: FAILED },
    FileType: { Pdf: "pdf" },
    context: {
      document: {
        getFileAsync: (type, options, cb) => {
          state.asked = { type, options };
          cb(failAt === "open" ? bad("export refused") : ok(file));
        },
      },
    },
  };
  return state;
}

function fakeWord(paragraphs) {
  const items = paragraphs.map(([style, text], i) => ({ style, text, uniqueLocalId: `P${i}` }));
  const context = {
    document: { load() {}, changeTrackingMode: "Off", body: { paragraphs: { items, load() {} } } },
    sync: async () => {},
  };
  globalThis.Word = { run: (fn) => fn(context) };
}

const { exportRenderPdf, collectHeadings } = await import("../taskpane/shared/tools-word.js");

test("exportRenderPdf joins the slices in order and closes the file", async () => {
  const pdf = Buffer.from("%PDF-1.7\n" + "x".repeat(5000));
  // Office gives a slice as a plain array of numbers or as a Uint8Array.
  const state = fakeOffice([Array.from(pdf.subarray(0, 3000)), new Uint8Array(pdf.subarray(3000))]);
  const out = await exportRenderPdf();
  assert.deepEqual([out.byte_length, out.slice_count], [pdf.length, 2]);
  assert.ok(Buffer.from(out.pdf_base64, "base64").equals(pdf));
  assert.equal(state.asked.type, "pdf");
  assert.equal(state.asked.options.sliceSize, 4 * 1024 * 1024);
  assert.equal(state.closed, 1);
});

test("exportRenderPdf throws what Word reported and still closes the file", async () => {
  const refused = fakeOffice([], { failAt: "open" });
  await assert.rejects(exportRenderPdf(), /export refused/);
  assert.equal(refused.closed, 0); // nothing was opened
  const state = fakeOffice([[1], [2]], { failAt: 1 });
  await assert.rejects(exportRenderPdf(), /slice 1 failed/);
  assert.equal(state.closed, 1);
  // A close that fails does not hide a good export.
  const closing = fakeOffice([[37, 80, 68, 70]], { closeFails: true });
  assert.equal((await exportRenderPdf()).byte_length, 4);
  assert.equal(closing.closed, 1);
});

test("collectHeadings lists heading paragraphs with their index, level and text", async () => {
  fakeWord([
    ["Title", "A fictional study"],
    ["Heading 1", " 1. Introduction "],
    ["Normal", "Flasks are common."],
    ["Acme-Heading 2", "1.1 Scope"],
    ["Normal", "Heading 3 is only a word here."],
    ["heading 1", "2. Materials and Methods"],
    ["Heading 1", "   "],
  ]);
  assert.deepEqual(await collectHeadings(), [
    { index: 1, level: 1, text: "1. Introduction" },
    { index: 3, level: 2, text: "1.1 Scope" },
    { index: 5, level: 1, text: "2. Materials and Methods" },
  ]);
  fakeWord([["Normal", "No headings."]]);
  assert.deepEqual(await collectHeadings(), []);
});
