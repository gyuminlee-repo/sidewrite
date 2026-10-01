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

// paragraphs: [style name, text, built-in style id?, outline level?]. The
// last two are what Word reports whatever its language (styleBuiltIn,
// outlineLevel). refuse: a property this Word cannot load; a load that asks
// for it fails at the sync, as Office.js does.
function fakeWord(paragraphs, { refuse = null } = {}) {
  const items = paragraphs.map(([style, text, styleBuiltIn, outlineLevel], i) => ({
    style,
    text,
    uniqueLocalId: `P${i}`,
    ...(styleBuiltIn === undefined ? {} : { styleBuiltIn }),
    ...(outlineLevel === undefined ? {} : { outlineLevel }),
  }));
  const loads = [];
  let failing = false;
  const context = {
    document: {
      load() {},
      changeTrackingMode: "Off",
      body: {
        paragraphs: {
          items,
          load(props) {
            loads.push(props);
            failing = !!refuse && props.includes(refuse);
          },
        },
      },
    },
    sync: async () => {
      if (failing) throw new Error(`property ${refuse} is not available`);
    },
  };
  globalThis.Word = { run: (fn) => fn(context) };
  return loads;
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

// The first acceptance run was on a Korean Word: the built-in heading styles
// are named in Korean there, and no heading was found by name.
test("collectHeadings: headings of a Word in another language, by built-in style id", async () => {
  const loads = fakeWord([
    ["제목", "A fictional study", "Title", 10],
    ["제목 1", "1. Introduction", "Heading1", 1],
    ["표준", "Flasks are common.", "Normal", 10],
    ["제목 2", "1.1 Scope", "Heading2", 2],
    ["표준", "Text.", "Normal", 10],
    ["제목 1", "2. Materials and Methods", "Heading1", 1],
  ]);
  assert.deepEqual(await collectHeadings(), [
    { index: 1, level: 1, text: "1. Introduction" },
    { index: 3, level: 2, text: "1.1 Scope" },
    { index: 5, level: 1, text: "2. Materials and Methods" },
  ]);
  assert.match(loads[0], /items\/styleBuiltIn/);
  assert.match(loads[0], /items\/outlineLevel/);
});

test("collectHeadings: a custom heading style is found by its outline level, in any language", async () => {
  fakeWord([
    ["논문 제목", "A fictional study", "Other", 10],
    ["장 제목", "1. Introduction", "Other", 1],
    ["본문", "Flasks are common.", "Other", 10],
    ["절 제목", "1.1 Scope", "Other", 2],
  ]);
  assert.deepEqual(await collectHeadings(), [
    { index: 1, level: 1, text: "1. Introduction" },
    { index: 3, level: 2, text: "1.1 Scope" },
  ]);
  // The built-in id wins over the outline level of the paragraph.
  fakeWord([
    ["제목 2", "1.1 Scope", "Heading2", 4],
    ["표준", "Text.", "Normal", 10],
  ]);
  assert.deepEqual(await collectHeadings(), [{ index: 0, level: 2, text: "1.1 Scope" }]);
});

test("collectHeadings: an outline level that every paragraph has says nothing", async () => {
  // A host that reports a level from 1 to 9 for every paragraph would turn
  // the whole manuscript into headings: the level is then not used.
  fakeWord([
    ["Heading 1", "1. Introduction", "Heading1", 1],
    ["Normal", "Flasks are common.", "Normal", 1],
    ["Normal", "Text.", "Normal", 1],
  ]);
  assert.deepEqual(await collectHeadings(), [{ index: 0, level: 1, text: "1. Introduction" }]);
});

test("collectHeadings: a Word that cannot load the two properties still finds headings by name", async () => {
  for (const refuse of ["styleBuiltIn", "outlineLevel"]) {
    const loads = fakeWord(
      [
        ["Heading 1", "1. Introduction"],
        ["Normal", "Text."],
      ],
      { refuse },
    );
    assert.deepEqual(await collectHeadings(), [{ index: 0, level: 1, text: "1. Introduction" }]);
    // Asked once with them, then as before; the paragraph ids are kept.
    assert.equal(loads.length, 2);
    assert.match(loads[1], /uniqueLocalId/);
    assert.doesNotMatch(loads[1], /styleBuiltIn|outlineLevel/);
  }
});

test("the outline of the chat tools uses the same heading test", async () => {
  const { toolGetOutline } = await import("../taskpane/shared/tools-word.js");
  fakeWord([
    ["제목 1", "1. Introduction", "Heading1", 1],
    ["표준", "Text.", "Normal", 10],
    ["제목 2", "1.1 Scope", "Heading2", 2],
  ]);
  const out = await toolGetOutline();
  assert.deepEqual(
    out.outline.map((h) => [h.level, h.text]),
    [
      [1, "1. Introduction"],
      [2, "1.1 Scope"],
    ],
  );
});

test("no style name is compared to find a heading outside the one helper (source)", async () => {
  const { readFile } = await import("node:fs/promises");
  const src = await readFile(new URL("../taskpane/shared/tools-word.js", import.meta.url), "utf8");
  // headingLevel(name) is the last resort inside paragraphLevel only.
  assert.equal((src.match(/headingLevel\(/g) ?? []).length, 2, "its definition and one call");
});
