// Line ranges of the quoted sentences, read from the margin line numbers of
// the PDF that Word rendered from the manuscript (scripts/pdf-line-map.py).
// The numbers are never typed by a person or written by the model.

import { createHash, randomUUID } from "node:crypto";
import { writeFile, rename, rm } from "node:fs/promises";
import { basename, dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { sidecarPath } from "./revision-store.mjs";
import { viewText } from "./response-quotes.mjs";
import { runPy } from "./py-tool.mjs";

const LINE_MAP = join(dirname(fileURLToPath(import.meta.url)), "..", "scripts", "pdf-line-map.py");

export const renderPdfPathFor = (doc) => sidecarPath(doc, ".render.pdf");

// lines: { start, end } of a citation; L: response_letter of the profile.
export const fmtLines = (lines, L) =>
  !lines
    ? null
    : lines.start === lines.end
      ? L.line_one.replace("{a}", () => lines.start)
      : L.line_range.replace("{a}", () => lines.start).replace("{b}", () => lines.end);

// Fill a profile frame. Without a line range the ", {lines}" part is left out.
// Values go in as they are (a "$" in a section name is not a pattern).
export function renderFrame(frame, { section = "", lines = null, n = "" } = {}) {
  const f =
    lines == null
      ? String(frame).replace(/,?\s*\{lines\}/, "")
      : String(frame).replace("{lines}", () => lines);
  return f.replace("{section}", () => section ?? "").replace("{n}", () => n);
}

// The location line of a citation. Without a section name (no heading above
// the sentence) the section-less frame of the profile is used, so the line
// never reads "in ,".
export function locationLine(cit, L) {
  const removed = cit.kind === "removed";
  const frame = cit.section
    ? removed
      ? L.removed_line
      : L.location_line
    : removed
      ? L.removed_line_no_section
      : L.location_line_no_section;
  return renderFrame(frame, { section: cit.section, lines: fmtLines(cit.lines, L) });
}

// One request per quoted sentence, with the text of its paragraph so a phrase
// that also occurs elsewhere is looked up inside that paragraph.
// The id is "<ref>#<position of the citation>".
export function lineRequests(comments, snapshot) {
  const byIndex = new Map((snapshot.paragraphs ?? []).map((p) => [p.index, viewText(p)]));
  return comments.flatMap((c) =>
    (c.response?.citations ?? []).flatMap((cit, i) =>
      cit.kind === "revised" && cit.quote
        ? [
            {
              id: `${c.ref}#${i}`,
              quote: cit.quote,
              within: byIndex.get(cit.paragraph_index) ?? null,
            },
          ]
        : [],
    ),
  );
}

// [{ ref, citations }] with the line ranges of `results` filled in. A quote
// that was looked up and not found loses its range; pdfSha records which PDF
// each answer came from.
export function applyLineResults(comments, results, pdfSha) {
  const by = new Map(results.map((r) => [r.id, r.match]));
  return comments.map((c) => ({
    ref: c.ref,
    citations: (c.response?.citations ?? []).map((cit, i) => {
      const id = `${c.ref}#${i}`;
      const m = by.get(id);
      if (!m)
        return { ...cit, lines: null, lines_pdf: by.has(id) ? pdfSha : (cit.lines_pdf ?? null) };
      return {
        ...cit,
        lines_pdf: pdfSha,
        lines: {
          start: m.start.line_no,
          end: m.end.line_no,
          start_page: m.start.page,
          end_page: m.end.page,
          match_count: m.match_count,
        },
      };
    }),
  }));
}

// Keep the PDF Word rendered as <stem>.render.pdf next to the manuscript.
export async function savePdf(doc, base64) {
  const buf = Buffer.from(String(base64 ?? ""), "base64");
  if (buf.length < 5 || buf.subarray(0, 5).toString("latin1") !== "%PDF-")
    throw new Error("Word did not return a PDF");
  const path = renderPdfPathFor(doc);
  // A temp file of this call's own, created exclusively (the naming of
  // revision-store.mjs writeJsonAtomic): two saves never share one, and a
  // file that was already there is never touched.
  const tmp = join(dirname(path), `.${basename(path)}.${process.pid}.${randomUUID()}.tmp`);
  await writeFile(tmp, buf, { flag: "wx" });
  try {
    await rename(tmp, path);
  } catch (e) {
    await rm(tmp, { force: true });
    throw e;
  }
  return { path, sha: createHash("sha256").update(buf).digest("hex"), bytes: buf.length };
}

// Look the requests up in the PDF. tmpJson: where to write the request list.
// Returns the line map output { lines_indexed, first_line, last_line,
// monotonic, results }.
export async function locate(pdfPath, requests, { run = runPy, tmpJson } = {}) {
  await writeFile(tmpJson, JSON.stringify(requests));
  const r = await run(LINE_MAP, [pdfPath, "--quotes-json", tmpJson]);
  if (r.json?.error === "no_line_numbers")
    throw new Error(
      "The PDF has no margin line numbers. Turn on Layout > Line Numbers > Continuous, then try again.",
    );
  if (!r.ok || !r.json || !Array.isArray(r.json.results))
    throw new Error(`pdf-line-map failed: ${r.error ?? "no output"}`);
  if (r.json.monotonic === false)
    throw new Error(
      "The line numbers restart inside the document. Set Layout > Line Numbers to Continuous, then try again.",
    );
  return r.json;
}
