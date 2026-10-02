// The figure service on a temp manuscript folder, with a stand-in for the
// Python renderer: what it records, what it refuses, and where it writes.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdir, mkdtemp, readFile, rm, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createFigureService, figureFindings, SOURCES_FILE } from "../daemon/figure-service.mjs";
import { RevisionSchema, saveRevision } from "../daemon/revision-store.mjs";
import { ENV, fakeRunPy, listing, manuscriptFolder, sha256, svg } from "./helpers/figure-world.mjs";

const NOW = new Date("2026-03-04T05:06:07.000Z");
async function world({ missing = false } = {}) {
  const m = await manuscriptFolder();
  const runPy = fakeRunPy({ missing });
  const service = createFigureService({ runPy, env: ENV, log: () => {}, now: () => NOW });
  const input = (extra = {}) => ({
    ref: "R1.1",
    figure: "Figure 2",
    kind: "revised",
    requirement_ids: ["q1"],
    summary: "The control curve was added.",
    before: "figures/fig2_before.svg",
    after: "figures/fig2_after.svg",
    ...extra,
  });
  const figure = async (id) =>
    (await m.rev()).comments.flatMap((c) => c.figure_changes).find((f) => f.id === id);
  return { ...m, service, runPy, input, figure, fig: (name) => join(m.dir, "figures", name) };
}

test("a proposed change records both files with their hashes, relative to the manuscript folder", async () => {
  const w = await world();
  try {
    const { id, problems } = await w.service.propose(w.doc, w.input());
    assert.equal(id, "R1.1-f1");
    assert.deepEqual(problems, []);
    const f = await w.figure(id);
    assert.equal(f.state, "proposed");
    assert.equal(f.before.source.path, "figures/fig2_before.svg");
    assert.equal(f.after.source.path, "figures/fig2_after.svg");
    assert.equal(f.before.source.sha256, await sha256(w.fig("fig2_before.svg")));
    assert.equal(f.after.source.sha256, await sha256(w.fig("fig2_after.svg")));
    assert.equal(f.after.source.bytes, Buffer.byteLength(svg(true)));
    // The pictures are named by the hash of the file they show.
    assert.equal(f.renders.after.path, `figures/.renders/${f.after.source.sha256}-150.png`);
    assert.equal(f.renders.after.of, f.after.source.sha256);
    assert.deepEqual([f.renders.before.width, f.renders.after.width], [667, 834]);
    assert.equal(f.facts.after.words, 1);
    const state = await w.service.state(w.doc);
    const shown = state.comments[0].figures[0];
    assert.match(shown.findings.join(" "), /different pixel sizes: 667 x 500 and 834 x 500/);
    assert.deepEqual(shown.problems, []);
    assert.deepEqual(
      state.comments[0].requirements.map((q) => [q.id, q.outside]),
      [
        ["q1", true],
        ["q2", false],
      ],
    );
    assert.equal(state.renderer.available, true);
  } finally {
    await w.cleanup();
  }
});

test("the files the author named are the same bytes after every operation", async () => {
  const w = await world();
  try {
    const sources = ["fig2_before.svg", "fig2_after.svg"].map(w.fig);
    const hashes = () => Promise.all(sources.map(sha256));
    const start = await hashes();
    const same = async (step) => assert.deepEqual(await hashes(), start, `after ${step}`);
    const { id } = await w.service.propose(w.doc, w.input());
    await same("propose");
    await w.service.render(w.doc, { ref: "R1.1", id, side: "before" });
    await w.service.render(w.doc, { ref: "R1.1", id, side: "after" });
    await same("render");
    await w.service.update(w.doc, { ref: "R1.1", id, patch: { summary: "Other words." } });
    await same("update");
    await w.service.adopt(w.doc, { ref: "R1.1", id });
    await same("adopt");
    await w.service.reopen(w.doc, { ref: "R1.1", id });
    await same("reopen");
    await w.service.reject(w.doc, { ref: "R1.1", id });
    await same("reject");
    await w.service.state(w.doc);
    await w.service.files(w.doc);
    await same("state and files");
  } finally {
    await w.cleanup();
  }
});

test("Adopt copies the new file into the submission folder and records the copy and its hash", async () => {
  const w = await world();
  try {
    const { id } = await w.service.propose(w.doc, w.input());
    const out = await w.service.adopt(w.doc, { ref: "R1.1", id });
    assert.equal(out.adopted.path, "figures/submission/fig2_after.svg");
    const f = await w.figure(id);
    assert.equal(f.state, "adopted");
    assert.deepEqual(f.adopted, out.adopted);
    assert.equal(f.adopted.at, NOW.toISOString());
    assert.equal(f.decided_at, NOW.toISOString());
    assert.equal(await sha256(join(w.dir, f.adopted.path)), f.adopted.sha256);
    assert.equal(f.adopted.sha256, f.after.source.sha256);
    // A second Adopt of the same change is refused.
    await assert.rejects(w.service.adopt(w.doc, { ref: "R1.1", id }), /is adopted already/);
  } finally {
    await w.cleanup();
  }
});

test("Adopt refuses a new file that changed since it was proposed, and copies nothing", async () => {
  const w = await world();
  try {
    const { id } = await w.service.propose(w.doc, w.input());
    await writeFile(w.fig("fig2_after.svg"), svg(true).replace("OD600", "OD600 (a.u.)"));
    await assert.rejects(
      w.service.adopt(w.doc, { ref: "R1.1", id }),
      /The new file changed since it was proposed: figures\/fig2_after\.svg/,
    );
    assert.equal((await w.figure(id)).state, "proposed");
    assert.deepEqual(
      (await listing(w.dir)).filter((p) => p.startsWith("figures/submission/")),
      [],
    );
    // The page says so before Adopt is pressed.
    const shown = (await w.service.state(w.doc)).comments[0].figures[0];
    assert.match(shown.problems.join(" "), /The new file was saved again since it was proposed/);
    // And a render of the changed file is refused as well.
    await assert.rejects(
      w.service.render(w.doc, { ref: "R1.1", id, side: "after" }),
      /changed since it was proposed/,
    );
  } finally {
    await w.cleanup();
  }
});

test("a file of the same name in the submission folder is kept and the copy is named -2", async () => {
  const w = await world();
  try {
    const sub = join(w.dir, "figures", "submission");
    await mkdir(sub);
    await writeFile(join(sub, "fig2_after.svg"), "the figure of an earlier round");
    const { id } = await w.service.propose(w.doc, w.input());
    const out = await w.service.adopt(w.doc, { ref: "R1.1", id });
    assert.equal(out.adopted.path, "figures/submission/fig2_after-2.svg");
    assert.equal(
      await readFile(join(sub, "fig2_after.svg"), "utf8"),
      "the figure of an earlier round",
    );
    assert.equal(await sha256(join(sub, "fig2_after-2.svg")), out.adopted.sha256);
    assert.equal(await readFile(join(sub, "fig2_after-2.svg"), "utf8"), svg(true));
  } finally {
    await w.cleanup();
  }
});

test("FILE_SOURCES.md gets one row per decision: adopted, withdrawn, adopted again without a second copy", async () => {
  const w = await world();
  try {
    const { id } = await w.service.propose(w.doc, w.input({ figure: "Figure 2 | growth" }));
    const first = await w.service.adopt(w.doc, { ref: "R1.1", id });
    const path = join(w.dir, "figures", "submission", SOURCES_FILE);
    const rows = async () =>
      (await readFile(path, "utf8"))
        .split("\n")
        .filter((l) => l.startsWith("| ") && !l.startsWith("| Date") && !l.startsWith("| ---"));
    const after = (await w.figure(id)).after.source;
    const row = (decision) =>
      `| 2026-03-04 | Figure 2 \\| growth | R1.1 | ${decision} | figures/fig2_after.svg | ${after.sha256} | fig2_after.svg | ${first.adopted.sha256} |`;
    assert.deepEqual(await rows(), [row("adopted")]);
    const text = await readFile(path, "utf8");
    assert.match(text, /^# Figure files for submission\n/);
    assert.match(
      text,
      /\| Date \| Figure \| Comment \| Decision \| Source \| Source SHA-256 \| File in this folder \| SHA-256 \|/,
    );

    await w.service.reopen(w.doc, { ref: "R1.1", id });
    assert.deepEqual(await rows(), [row("adopted"), row("withdrawn")]);
    assert.equal((await w.figure(id)).adopted, null);
    // The copy stays; the same file adopted again is that copy, not a "-2".
    const again = await w.service.adopt(w.doc, { ref: "R1.1", id });
    assert.equal(again.adopted.path, "figures/submission/fig2_after.svg");
    assert.deepEqual(await rows(), [row("adopted"), row("withdrawn"), row("adopted")]);
    assert.deepEqual(
      (await listing(w.dir)).filter((p) => p.startsWith("figures/submission/")),
      [`figures/submission/${SOURCES_FILE}`, "figures/submission/fig2_after.svg"],
    );
  } finally {
    await w.cleanup();
  }
});

test("nothing is written outside .renders, submission and revision.json, also for a file from another folder", async () => {
  const w = await world();
  const elsewhere = await mkdtemp(join(tmpdir(), "figure-elsewhere-"));
  try {
    await writeFile(join(elsewhere, "fig3_new.svg"), svg(true));
    const before = await listing(w.dir);
    const outsideBefore = await listing(elsewhere);
    const a = await w.service.propose(w.doc, w.input());
    const b = await w.service.propose(
      w.doc,
      w.input({
        kind: "new",
        figure: "Figure 3",
        before: "",
        after: join(elsewhere, "fig3_new.svg"),
      }),
    );
    await w.service.render(w.doc, { ref: "R1.1", id: a.id, side: "after" });
    await w.service.adopt(w.doc, { ref: "R1.1", id: a.id });
    await w.service.adopt(w.doc, { ref: "R1.1", id: b.id });
    await w.service.reopen(w.doc, { ref: "R1.1", id: a.id });
    await w.service.reject(w.doc, { ref: "R1.1", id: a.id });
    await w.service.state(w.doc);
    await w.service.files(w.doc);

    const added = (await listing(w.dir)).filter((p) => !before.includes(p));
    assert.ok(added.length >= 6, `files were written: ${added.join(", ")}`);
    const allowed = (p) => p.startsWith("figures/.renders/") || p.startsWith("figures/submission/");
    assert.deepEqual(
      added.filter((p) => !allowed(p)),
      [],
    );
    assert.ok(before.includes("paper.revision.json"));
    assert.deepEqual(
      added.filter((p) => p.includes(".tmp")),
      [],
      "no temp file is left",
    );
    assert.deepEqual(await listing(elsewhere), outsideBefore);
    // The file from the other folder is recorded by its full path and copied in.
    const f = await w.figure(b.id);
    assert.equal(f.after.source.path, join(elsewhere, "fig3_new.svg"));
    assert.equal(f.before, null);
    assert.equal(f.adopted.path, "figures/submission/fig3_new.svg");
  } finally {
    await w.cleanup();
    await rm(elsewhere, { recursive: true, force: true });
  }
});

test("a figures or submission folder that leads out of the manuscript folder is refused", async () => {
  const w = await world();
  const elsewhere = await mkdtemp(join(tmpdir(), "figure-elsewhere-"));
  try {
    const { id } = await w.service.propose(w.doc, w.input());
    await symlink(elsewhere, join(w.dir, "figures", "submission"));
    await assert.rejects(
      w.service.adopt(w.doc, { ref: "R1.1", id }),
      /leads outside the manuscript folder/,
    );
    assert.deepEqual(await listing(elsewhere), []);
    assert.equal((await w.figure(id)).state, "proposed");

    const rev = await w.rev();
    await saveRevision(
      join(w.dir, "paper.revision.json"),
      RevisionSchema.parse({ ...rev, figure_lane: { dir: "../other" } }),
    );
    await assert.rejects(
      w.service.state(w.doc),
      /The figures folder must be inside the manuscript folder/,
    );
  } finally {
    await w.cleanup();
    await rm(elsewhere, { recursive: true, force: true });
  }
});

test("what is not a figure change is refused with words for the author", async () => {
  const w = await world();
  try {
    const no = (extra, re) => assert.rejects(w.service.propose(w.doc, w.input(extra)), re);
    await no({ figure: " " }, /Enter the label of the figure or file/);
    await no({ after: "" }, /Choose the new file/);
    await no({ before: "" }, /Choose the old file/);
    await no({ after: "figures/none.svg" }, /No file at figures\/none\.svg/);
    await no({ after: "paper.docx" }, /paper\.docx is the manuscript itself/);
    await no({ after: "figures/fig2_before.svg" }, /are the same file/);
    await no({ kind: "new" }, /A new figure or file has no old file/);
    await no({ ref: "R7.7" }, /No reviewer comment R7\.7/);
    await no({ requirement_ids: ["q9"] }, /No requirement q9/);
    await no({ after: "https://example.org/fig.svg" }, /not a file this computer can open/);
    assert.equal((await w.rev()).comments[0].figure_changes.length, 0);
    await assert.rejects(
      w.service.render(w.doc, { ref: "R1.1", id: "R1.1-f1", side: "after" }),
      /No figure change R1\.1-f1/,
    );
    await assert.rejects(
      w.service.render(w.doc, { ref: "R1.1", id: "x", side: "../after" }),
      /side must be before or after/,
    );
  } finally {
    await w.cleanup();
  }
});

test("without PyMuPDF the lane still records, adopts and says plainly that there are no pictures", async () => {
  const w = await world({ missing: true });
  try {
    const { id, problems } = await w.service.propose(w.doc, w.input());
    assert.equal(problems.length, 1);
    assert.match(problems[0], /PyMuPDF is not installed/);
    const f = await w.figure(id);
    assert.deepEqual(f.renders, { before: null, after: null });
    const state = await w.service.state(w.doc);
    assert.equal(state.renderer.available, false);
    assert.match(
      state.renderer.reason,
      /PyMuPDF is not installed .* still recorded, adopted and cited/,
    );
    await assert.rejects(
      w.service.render(w.doc, { ref: "R1.1", id, side: "after" }),
      (e) => e.status === 503 && /PyMuPDF is not installed/.test(e.message),
    );
    const out = await w.service.adopt(w.doc, { ref: "R1.1", id });
    assert.equal(await sha256(join(w.dir, out.adopted.path)), out.adopted.sha256);
    assert.deepEqual(
      (await listing(w.dir)).filter((p) => p.startsWith("figures/.renders/")),
      [],
    );
  } finally {
    await w.cleanup();
  }
});

test("a picture is made once per file content and found again by its hash", async () => {
  const w = await world();
  try {
    const { id } = await w.service.propose(w.doc, w.input());
    const renders = () => w.runPy.calls.filter((a) => a.includes("--in")).length;
    assert.equal(renders(), 2);
    const { png } = await w.service.render(w.doc, { ref: "R1.1", id, side: "after" });
    assert.equal(renders(), 2, "served from the cache");
    assert.match(await readFile(png, "utf8"), /^PNG of .*fig2_after\.svg$/);
    // A second change with the same files makes no new picture.
    await w.service.propose(w.doc, w.input({ figure: "Figure 2 again" }));
    assert.equal(renders(), 2);
  } finally {
    await w.cleanup();
  }
});

test("an .ai source with a PDF saved from it: the PDF is shown and submitted, the .ai file is the source", async () => {
  const w = await world();
  try {
    await writeFile(w.fig("fig2_after.ai"), "an .ai file");
    await writeFile(w.fig("fig2_after.pdf"), "the PDF saved from it");
    const { id } = await w.service.propose(
      w.doc,
      w.input({ after: "figures/fig2_after.ai", after_view: "figures/fig2_after.pdf" }),
    );
    const f = await w.figure(id);
    assert.equal(f.after.source.path, "figures/fig2_after.ai");
    assert.equal(f.after.view.path, "figures/fig2_after.pdf");
    assert.equal(f.renders.after.of, f.after.view.sha256);
    const out = await w.service.adopt(w.doc, { ref: "R1.1", id });
    assert.equal(out.adopted.path, "figures/submission/fig2_after.pdf");
    assert.equal(out.adopted.sha256, f.after.view.sha256);
    const sources = await readFile(join(w.dir, "figures", "submission", SOURCES_FILE), "utf8");
    assert.match(
      sources,
      new RegExp(
        `figures/fig2_after\\.ai \\| ${f.after.source.sha256} \\| fig2_after\\.pdf \\| ${f.after.view.sha256} \\|`,
      ),
    );
  } finally {
    await w.cleanup();
  }
});

test("the file list holds the files of the figures folder, newest first, without the pictures and the list of sources", async () => {
  const w = await world();
  try {
    const { id } = await w.service.propose(w.doc, w.input());
    await w.service.adopt(w.doc, { ref: "R1.1", id });
    await writeFile(w.fig("Table_S3_v2.xlsx"), "an invented table");
    await writeFile(w.fig(".hidden.xlsx"), "not shown");
    const out = await w.service.files(w.doc);
    assert.equal(out.folder, "figures");
    // FILE_SOURCES.md is written by Adopt and is not a file to record.
    assert.ok((await listing(w.dir)).includes("figures/submission/FILE_SOURCES.md"));
    assert.deepEqual(out.files.map((f) => f.path).sort(), [
      "figures/Table_S3_v2.xlsx",
      "figures/fig2_after.svg",
      "figures/fig2_before.svg",
      "figures/submission/fig2_after.svg",
    ]);
    assert.deepEqual(
      out.files.filter((f) => f.submitted).map((f) => f.path),
      ["figures/submission/fig2_after.svg"],
    );
  } finally {
    await w.cleanup();
  }
});

test("findings: pure, from the stored facts", () => {
  const ref = (c) => ({ path: "x", sha256: c.repeat(64), bytes: 1, mtime: "t" });
  const facts = (extra) => ({
    type: "PDF",
    pages: 1,
    words: 4,
    images: 0,
    paths: 1,
    fonts: [],
    ...extra,
  });
  const render = (width) => ({ path: "p", of: "h", dpi: 150, width, height: 10 });
  const f = (extra = {}) => ({
    before: { source: ref("a"), view: null },
    after: { source: ref("b"), view: null },
    renders: { before: render(10), after: render(10) },
    facts: { before: facts(), after: facts() },
    ...extra,
  });
  assert.deepEqual(figureFindings(f()), []);
  assert.match(figureFindings(f({ after: { source: ref("a"), view: null } }))[0], /identical/);
  assert.match(
    figureFindings(f({ renders: { before: render(10), after: render(12) } }))[0],
    /different pixel sizes: 10 x 10 and 12 x 10/,
  );
  const changed = figureFindings(
    f({
      facts: {
        before: facts(),
        after: facts({
          type: "PNG",
          pages: 3,
          words: 0,
          fonts: [
            { name: "Arial", embedded: false },
            { name: "Symbol", embedded: true },
          ],
        }),
      },
    }),
  );
  assert.deepEqual(changed, [
    "The file type changed from PDF to PNG.",
    "The new file has 3 pages. Only the first is shown.",
    "The old file has live text and the new file has none.",
    "Fonts not embedded in the new file: Arial.",
  ]);
  // A new figure, and a change that was never rendered.
  assert.deepEqual(figureFindings({ before: null, after: { source: ref("b") } }), []);
});
