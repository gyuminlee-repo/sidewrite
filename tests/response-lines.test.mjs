import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  fmtLines,
  lineRequests,
  applyLineResults,
  renderFrame,
  savePdf,
  locate,
  renderPdfPathFor,
} from "../daemon/response-lines.mjs";
import { runPy } from "../daemon/py-tool.mjs";

const L = { line_one: "line {a}", line_range: "lines {a}-{b}" };

test("line wording comes from the profile", () => {
  assert.equal(fmtLines({ start: 9, end: 9 }, L), "line 9");
  assert.equal(fmtLines({ start: 9, end: 11 }, L), "lines 9-11");
  assert.equal(
    fmtLines({ start: 9, end: 11 }, { line_one: "Zeile {a}", line_range: "Zeilen {a}-{b}" }),
    "Zeilen 9-11",
  );
  assert.equal(fmtLines(null, L), null);
});

test("a frame without a line range drops the lines part", () => {
  assert.equal(
    renderFrame("[Revised in {section}, {lines}]", { section: "3. Results", lines: "line 4" }),
    "[Revised in 3. Results, line 4]",
  );
  assert.equal(
    renderFrame("[Revised in {section}, {lines}]", { section: "3. Results", lines: null }),
    "[Revised in 3. Results]",
  );
  assert.equal(renderFrame("Reviewer #{n}", { n: 2 }), "Reviewer #2");
});

test("a section name is written as it is, also with $ signs in it", () => {
  assert.equal(
    renderFrame("[Revised in {section}, {lines}]", {
      section: "Cost in $& and $1",
      lines: "line 4",
    }),
    "[Revised in Cost in $& and $1, line 4]",
  );
});

test("requests carry the paragraph text; results land on the right citation with the PDF hash", () => {
  const comments = [
    {
      ref: "R1.1",
      response: {
        citations: [
          { kind: "revised", quote: "q one", paragraph_index: 1 },
          { kind: "removed", quote: "", paragraph_index: 1 },
        ],
      },
    },
    { ref: "R1.2", response: null },
  ];
  const snap = { paragraphs: [{ index: 1, text: "before q one after" }] };
  const reqs = lineRequests(comments, snap);
  assert.deepEqual(reqs, [{ id: "R1.1#0", quote: "q one", within: "before q one after" }]);
  const out = applyLineResults(
    comments,
    [
      {
        id: "R1.1#0",
        match: { start: { page: 1, line_no: 4 }, end: { page: 1, line_no: 5 }, match_count: 1 },
      },
    ],
    "sha",
  );
  assert.deepEqual(out[0].citations[0].lines, {
    start: 4,
    end: 5,
    start_page: 1,
    end_page: 1,
    match_count: 1,
  });
  assert.equal(out[0].citations[0].lines_pdf, "sha");
  assert.equal(out[0].citations[1].lines, null);
  assert.deepEqual(out[1], { ref: "R1.2", citations: [] });
});

test("a quote the PDF does not hold loses its old range and is marked as looked up", () => {
  const comments = [
    {
      ref: "R1.1",
      response: {
        citations: [
          {
            kind: "revised",
            quote: "q",
            paragraph_index: 1,
            lines: { start: 1, end: 1 },
            lines_pdf: "old",
          },
        ],
      },
    },
  ];
  const [c] = applyLineResults(comments, [{ id: "R1.1#0", match: null }], "new");
  assert.deepEqual([c.citations[0].lines, c.citations[0].lines_pdf], [null, "new"]);
});

test("savePdf keeps the PDF next to the manuscript and refuses anything else", async () => {
  const dir = await mkdtemp(join(tmpdir(), "lines-"));
  const doc = join(dir, "paper.docx");
  const bytes = Buffer.from("%PDF-1.7\nfake body\n");
  const saved = await savePdf(doc, bytes.toString("base64"));
  assert.equal(saved.path, join(dir, "paper.render.pdf"));
  assert.equal(saved.path, renderPdfPathFor(doc));
  assert.equal(saved.bytes, bytes.length);
  assert.match(saved.sha, /^[0-9a-f]{64}$/);
  assert.deepEqual(await readFile(saved.path), bytes);
  await assert.rejects(savePdf(doc, Buffer.from("PK\u0003\u0004 docx").toString("base64")), /PDF/);
  await assert.rejects(savePdf(doc, undefined), /PDF/);
});

test("locate passes the requests to the line map and names what is wrong with the PDF", async () => {
  const dir = await mkdtemp(join(tmpdir(), "lines-"));
  const tmpJson = join(dir, "q.json");
  const calls = [];
  const fake = (result) => async (script, args) => {
    calls.push([script, args]);
    return result;
  };
  const good = { lines_indexed: 3, monotonic: true, results: [{ id: "R1.1#0", match: null }] };
  const reqs = [{ id: "R1.1#0", quote: "q", within: "p" }];
  const out = await locate("/x/m.render.pdf", reqs, {
    run: fake({ ok: true, code: 0, json: good }),
    tmpJson,
  });
  assert.deepEqual(out, good);
  assert.match(calls[0][0], /scripts[\\/]pdf-line-map\.py$/);
  assert.deepEqual(calls[0][1], ["/x/m.render.pdf", "--quotes-json", tmpJson]);
  assert.deepEqual(JSON.parse(await readFile(tmpJson, "utf8")), reqs);
  await assert.rejects(
    locate("/x/m.pdf", reqs, {
      run: fake({ ok: false, code: 2, json: { error: "no_line_numbers" } }),
      tmpJson,
    }),
    /no margin line numbers/,
  );
  await assert.rejects(
    locate("/x/m.pdf", reqs, {
      run: fake({ ok: true, code: 0, json: { ...good, monotonic: false } }),
      tmpJson,
    }),
    /Continuous/,
  );
  await assert.rejects(
    locate("/x/m.pdf", reqs, {
      run: fake({ ok: false, code: null, error: "python3 not found" }),
      tmpJson,
    }),
    /python3 not found/,
  );
});

test("runPy: JSON from the last stdout line, exit code kept, missing interpreter is a plain error", async () => {
  const ok = await runPy("-c", ["print('x'); print('{\"a\":1}')"]);
  assert.deepEqual([ok.ok, ok.code, ok.json, ok.error], [true, 0, { a: 1 }, null]);
  const bad = await runPy("-c", ["import sys; print('{\"e\":1}'); sys.exit(2)"]);
  assert.deepEqual([bad.ok, bad.code, bad.json], [false, 2, { e: 1 }]);
  const noJson = await runPy("-c", ["import sys; sys.stderr.write('boom\\n'); sys.exit(1)"]);
  assert.deepEqual(
    [noJson.ok, noJson.code, noJson.json, noJson.error],
    [false, 1, undefined, "boom"],
  );
  const missing = await runPy("x.py", [], { python: "/no/such/python" });
  assert.deepEqual([missing.ok, missing.code], [false, null]);
  assert.match(missing.error, /not found/);
});

test("runPy: a tool that runs too long is stopped", async () => {
  const slow = await runPy("-c", ["import time; time.sleep(30)"], { timeoutMs: 300 });
  assert.equal(slow.ok, false);
  assert.match(slow.error, /timed out after 300 ms/);
});
