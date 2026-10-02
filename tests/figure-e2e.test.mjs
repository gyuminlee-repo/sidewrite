// The figure lane end to end without Word: the real bridge with a WebSocket
// client as the pane, the figures routes over HTTP, the real Python renderer
// and the invented fixture figure (scripts/e2e/make_figure_fixture.py).
// propose -> render -> adopt, then the copy on disk is the recorded one and
// the response letter cites it.
import { test } from "node:test";
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { missing } from "./fixtures/py-packages.mjs";
import { figureWorld, sha256, listing } from "./helpers/figure-world.mjs";
import { buildLetterModel } from "../daemon/response-model.mjs";
import { checkResponses } from "../daemon/response-check.mjs";
import { figureFileRows } from "../daemon/response-service.mjs";
import { withDefaults } from "../daemon/journal-profile.mjs";

const FIXTURE = fileURLToPath(new URL("../scripts/e2e/make_figure_fixture.py", import.meta.url));
const PNG = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
const python = process.env.CC_OFFICE_PYTHON || "python3";

test(
  "propose, render and adopt the fixture figure over HTTP with the real renderer",
  { skip: missing("fitz") },
  async () => {
    const w = await figureWorld({});
    try {
      const made = JSON.parse(
        execFileSync(python, [FIXTURE, "--out", join(w.dir, "figures")], { encoding: "utf8" }),
      );
      assert.equal(made.pdf, true);
      const sources = ["fig2_before.pdf", "fig2_after.pdf"].map((f) => join(w.dir, "figures", f));
      const start = await Promise.all(sources.map(sha256));

      // The page first asks which manuscripts are open, then for their state.
      const open = (await w.call("manuscripts")).json.manuscripts;
      assert.deepEqual(
        open.map((m) => m.doc),
        [w.doc],
      );
      const p = await w.propose({
        before: "figures/fig2_before.pdf",
        after: "figures/fig2_after.pdf",
      });
      assert.equal(p.status, 200, JSON.stringify(p.json));
      assert.deepEqual(p.json.problems, []);
      const id = p.json.id;

      // Both pictures are PNG files, and wider after than before.
      const size = (png) => [png.readUInt32BE(16), png.readUInt32BE(20)];
      const pictures = {};
      for (const side of ["before", "after"]) {
        const r = await w.call("render", { query: { doc: w.doc, ref: "R1.1", id, side } });
        assert.equal(r.status, 200);
        assert.deepEqual(r.body.subarray(0, 8), PNG);
        pictures[side] = size(r.body);
      }
      assert.deepEqual(pictures, { before: [667, 500], after: [834, 500] });

      const state = (await w.call("state", { query: { doc: w.doc } })).json;
      assert.equal(state.renderer.available, true);
      const f = state.comments[0].figures[0];
      assert.equal(f.facts.after.type, "PDF");
      assert.ok(f.facts.before.words >= 2, "the old file has live text (OD600, Time)");
      assert.ok(f.facts.after.words > f.facts.before.words, "the legend added words");
      assert.deepEqual(
        f.facts.after.fonts.map((x) => x.embedded),
        [true],
      );
      assert.deepEqual([f.renders.before.width, f.renders.after.width], [667, 834]);
      assert.match(f.findings.join(" "), /different pixel sizes: 667 x 500 and 834 x 500/);

      const a = await w.call("adopt", { body: { doc: w.doc, ref: "R1.1", id } });
      assert.equal(a.status, 200, JSON.stringify(a.json));
      const record = (await w.rev()).comments[0].figure_changes[0];
      assert.equal(record.state, "adopted");
      assert.equal(record.adopted.path, "figures/submission/fig2_after.pdf");
      // The copy on disk is the file that was proposed, byte for byte.
      assert.equal(await sha256(join(w.dir, record.adopted.path)), record.adopted.sha256);
      assert.equal(record.adopted.sha256, start[1]);
      assert.deepEqual(await Promise.all(sources.map(sha256)), start, "the sources are untouched");
      const sourcesFile = await readFile(
        join(w.dir, "figures", "submission", "FILE_SOURCES.md"),
        "utf8",
      );
      assert.match(
        sourcesFile,
        new RegExp(
          `\\| Figure 2 \\| R1\\.1 \\| adopted \\| figures/fig2_after\\.pdf \\| ${start[1]} \\| fig2_after\\.pdf \\| ${start[1]} \\|`,
        ),
      );
      // The letter names the adopted copy, and its request is answered.
      const rev = await w.rev();
      const profile = withDefaults({ colors: { response_revised: "0070C0" } });
      assert.deepEqual(
        buildLetterModel(rev, profile).paragraphs.filter((x) => x.kind === "figure"),
        [
          {
            kind: "figure",
            text: "[Revised Figure 2, file fig2_after.pdf]",
            color: "0070C0",
            ref: "R1.1",
            reviewer: 1,
          },
        ],
      );
      assert.deepEqual(await figureFileRows(rev, w.doc), []);
      const unanswered = checkResponses({ revision: rev, snapshot: { paragraphs: [] }, profile })
        .filter((x) => x.id === "requirements")
        .map((x) => x.detail);
      assert.deepEqual(unanswered, [
        "q2 is neither cited nor declined: give the unit of the time axis",
      ]);
      const files = await listing(w.dir);
      assert.equal(files.filter((x) => x.startsWith("figures/.renders/")).length, 4);
    } finally {
      await w.cleanup();
    }
  },
);
