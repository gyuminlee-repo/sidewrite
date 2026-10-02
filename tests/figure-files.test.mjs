// Step A0: a supplementary file that was changed outside Word (a table, a
// data file, a document) is recorded on the figures page like a figure:
// hashes, an adopted copy, and the response check counts its requests. Only
// the picture is missing. Invented files; the renderer is a stand-in.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdir, mkdtemp, readFile, rm, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createFigureService, NO_PREVIEW } from "../daemon/figure-service.mjs";
import { checkResponses } from "../daemon/response-check.mjs";
import { buildLetterModel } from "../daemon/response-model.mjs";
import { withDefaults } from "../daemon/journal-profile.mjs";
import { RevisionSchema, setResponse } from "../daemon/revision-store.mjs";
import {
  ENV,
  fakeRunPy,
  figureWorld,
  listing,
  manuscriptFolder,
  sha256,
} from "./helpers/figure-world.mjs";
import { figuresPage, findAll, textOf } from "./helpers/figures-page.mjs";

const NOW = new Date("2026-03-04T05:06:07.000Z");
const COMMENTS = [
  {
    ref: "R1.2",
    reviewer: 1,
    text: "Add the control strain to Table S3.",
    requirements: [{ id: "q1", text: "add the control strain to Table S3" }],
    requirements_outside: [
      { id: "q1", why: "The table is a file of its own.", at: "2026-01-01T00:00:00.000Z" },
    ],
  },
];
const V1 = "strain,rate\nA,0.31\n";
const V2 = "strain,rate\nA,0.31\ncontrol,0.02\n";

async function world() {
  const m = await manuscriptFolder({ comments: COMMENTS });
  const fig = (name) => join(m.dir, "figures", name);
  await writeFile(fig("Table_S3_v1.xlsx"), V1);
  await writeFile(fig("Table_S3_v2.xlsx"), V2);
  const runPy = fakeRunPy();
  const service = createFigureService({ runPy, env: ENV, log: () => {}, now: () => NOW });
  const input = (extra = {}) => ({
    ref: "R1.2",
    figure: "Table S3",
    kind: "revised",
    requirement_ids: ["q1"],
    summary: "The control strain was added as a row.",
    before: "figures/Table_S3_v1.xlsx",
    after: "figures/Table_S3_v2.xlsx",
    ...extra,
  });
  const record = async (id) => (await m.rev()).comments[0].figure_changes.find((f) => f.id === id);
  return { ...m, service, runPy, input, record, fig };
}

test("a file of any type is recorded with its hashes; a folder, the manuscript and a relative path out of the folder are refused", async () => {
  const w = await world();
  const elsewhere = await mkdtemp(join(tmpdir(), "figure-elsewhere-"));
  try {
    await writeFile(w.fig("Supplement_v2.docx"), "an invented supplement");
    await writeFile(w.fig("Data_S1.pdf"), "an invented data sheet");
    await writeFile(w.fig("raw counts"), "a file without a type in its name");
    await writeFile(join(elsewhere, "Table_S4.xlsx"), "outside");
    const kept = [];
    for (const [figure, after] of [
      ["Supplementary text", "figures/Supplement_v2.docx"],
      ["Data S1", "figures/Data_S1.pdf"],
      ["Raw counts", "figures/raw counts"],
    ]) {
      const { id } = await w.service.propose(
        w.doc,
        w.input({ kind: "new", before: "", figure, after }),
      );
      kept.push((await w.record(id)).after.source.path);
    }
    assert.deepEqual(kept, [
      "figures/Supplement_v2.docx",
      "figures/Data_S1.pdf",
      "figures/raw counts",
    ]);

    const { id } = await w.service.propose(w.doc, w.input());
    const f = await w.record(id);
    assert.equal(f.kind, "revised");
    assert.equal(f.state, "proposed");
    assert.deepEqual(
      [f.before.source.path, f.after.source.path],
      ["figures/Table_S3_v1.xlsx", "figures/Table_S3_v2.xlsx"],
    );
    assert.equal(f.after.source.sha256, await sha256(w.fig("Table_S3_v2.xlsx")));
    assert.equal(f.after.source.bytes, V2.length);
    assert.equal(f.before.source.sha256, await sha256(w.fig("Table_S3_v1.xlsx")));

    const count = async () => (await w.rev()).comments[0].figure_changes.length;
    const before = await count();
    const no = (extra, re) => assert.rejects(w.service.propose(w.doc, w.input(extra)), re);
    await no({ after: "figures" }, /figures is a folder\. Choose a file\./);
    await no({ after: "figures/submission/.." }, /is a folder/);
    await no({ after: "paper.docx" }, /paper\.docx is the manuscript itself/);
    await no({ after: w.doc }, /is the manuscript itself/);
    // A path of the list never leaves the manuscript folder, also through "..".
    const rel = join("..", elsewhere.split("/").pop(), "Table_S4.xlsx");
    await no(
      { after: rel },
      /is outside the manuscript folder\. Paste the full path of the file\./,
    );
    await no({ after: "figures/../../nowhere.xlsx" }, /is outside the manuscript folder/);
    await no({ after: "figures/none.xlsx" }, /No file at figures\/none\.xlsx/);
    assert.equal(await count(), before, "a refused file records nothing");

    // A file from another folder is named by its full path, as a figure is.
    const far = await w.service.propose(
      w.doc,
      w.input({
        kind: "new",
        before: "",
        figure: "Table S4",
        after: join(elsewhere, "Table_S4.xlsx"),
      }),
    );
    assert.equal((await w.record(far.id)).after.source.path, join(elsewhere, "Table_S4.xlsx"));
    assert.deepEqual(await listing(elsewhere), ["Table_S4.xlsx"]);
  } finally {
    await w.cleanup();
    await rm(elsewhere, { recursive: true, force: true });
  }
});

test("no picture is asked for of a table: the record has no renders and the page is told there is no preview", async () => {
  const w = await world();
  try {
    const { id, problems } = await w.service.propose(w.doc, w.input());
    assert.deepEqual(problems, []);
    // The renderer was not run at all, and nothing was written for a picture.
    assert.deepEqual(w.runPy.calls, []);
    const f = await w.record(id);
    assert.deepEqual(f.renders, { before: null, after: null });
    assert.deepEqual(f.facts, { before: null, after: null });
    assert.deepEqual(
      (await listing(w.dir)).filter((p) => p.includes(".renders")),
      [],
    );
    const shown = (await w.service.state(w.doc)).comments[0].figures[0];
    assert.equal(shown.preview, false);
    assert.deepEqual(shown.problems, []);
    // A page that asks for the picture all the same gets the plain sentence.
    for (const side of ["before", "after"])
      await assert.rejects(
        w.service.render(w.doc, { ref: "R1.2", id, side }),
        (e) => e.status === 415 && e.message === NO_PREVIEW,
      );
    assert.equal(NO_PREVIEW, "No preview for this file type.");
    // (The state asks once whether the renderer is there; no file was given to it.)
    assert.deepEqual(
      w.runPy.calls.filter((a) => a.includes("--in")),
      [],
    );

    // A figure next to it is rendered as before, and says so.
    const fig = await w.service.propose(
      w.doc,
      w.input({
        figure: "Figure 2",
        before: "figures/fig2_before.svg",
        after: "figures/fig2_after.svg",
      }),
    );
    assert.equal(w.runPy.calls.filter((a) => a.includes("--in")).length, 2);
    const both = (await w.service.state(w.doc)).comments[0].figures;
    assert.deepEqual(
      both.map((x) => x.preview),
      [false, true],
    );
    assert.equal((await w.record(fig.id)).renders.after.width, 834);
    // A figure replaced by a table: one side has no picture, so there is none to compare.
    await w.service.propose(
      w.doc,
      w.input({
        figure: "Figure 3",
        before: "figures/fig2_before.svg",
        after: "figures/Table_S3_v1.xlsx",
      }),
    );
    assert.equal((await w.service.state(w.doc)).comments[0].figures[2].preview, false);
  } finally {
    await w.cleanup();
  }
});

test("Adopt copies the table and records the copy and its hash; a table saved again since is refused", async () => {
  const w = await world();
  try {
    const { id } = await w.service.propose(w.doc, w.input());
    const out = await w.service.adopt(w.doc, { ref: "R1.2", id });
    assert.deepEqual(out.adopted, {
      path: "figures/submission/Table_S3_v2.xlsx",
      sha256: await sha256(w.fig("Table_S3_v2.xlsx")),
      at: NOW.toISOString(),
    });
    assert.equal(await readFile(join(w.dir, out.adopted.path), "utf8"), V2);
    assert.equal(await sha256(join(w.dir, out.adopted.path)), out.adopted.sha256);
    const f = await w.record(id);
    assert.equal(f.state, "adopted");
    assert.deepEqual(f.adopted, out.adopted);
    // The source is as it was, and the row of the copy names the label.
    assert.equal(await readFile(w.fig("Table_S3_v2.xlsx"), "utf8"), V2);
    const sources = await readFile(join(w.dir, "figures/submission/FILE_SOURCES.md"), "utf8");
    assert.match(
      sources,
      new RegExp(
        `\\| 2026-03-04 \\| Table S3 \\| R1\\.2 \\| adopted \\| figures/Table_S3_v2\\.xlsx \\| ${out.adopted.sha256} \\| Table_S3_v2\\.xlsx \\|`,
      ),
    );

    // Edited after it was recorded: the copy would not be the recorded file.
    await w.service.reopen(w.doc, { ref: "R1.2", id });
    await writeFile(w.fig("Table_S3_v2.xlsx"), `${V2}late,0.5\n`);
    await assert.rejects(
      w.service.adopt(w.doc, { ref: "R1.2", id }),
      (e) => e.status === 409 && /The new file changed since it was proposed/.test(e.message),
    );
    assert.equal((await w.record(id)).state, "proposed");
  } finally {
    await w.cleanup();
  }
});

test("the copy never leaves the manuscript folder, also for a table (a submission folder that is a link out)", async () => {
  const w = await world();
  const elsewhere = await mkdtemp(join(tmpdir(), "figure-elsewhere-"));
  try {
    const { id } = await w.service.propose(w.doc, w.input());
    await symlink(elsewhere, join(w.dir, "figures", "submission"));
    await assert.rejects(
      w.service.adopt(w.doc, { ref: "R1.2", id }),
      /leads outside the manuscript folder/,
    );
    assert.deepEqual(await listing(elsewhere), []);
    assert.equal((await w.record(id)).state, "proposed");
  } finally {
    await w.cleanup();
    await rm(elsewhere, { recursive: true, force: true });
  }
});

// ---- the letter ----------------------------------------------------------------

test("the response check counts a request named by an adopted table record, and still fails it for a proposed or rejected one", async () => {
  const w = await world();
  try {
    const profile = withDefaults({});
    const snapshot = { paragraphs: [] };
    const fails = async () => {
      const revision = setResponse(RevisionSchema.parse(await w.rev()), "R1.2", {
        text: "We added the control strain to Table S3.",
        source: "author",
      });
      return checkResponses({ revision, snapshot, profile })
        .filter((r) => r.id === "requirements")
        .map((r) => r.detail);
    };
    // Before any record: the request blocks the letter.
    assert.deepEqual(await fails(), [
      "q1 is neither cited nor declined: add the control strain to Table S3",
    ]);
    const { id } = await w.service.propose(w.doc, w.input());
    assert.deepEqual(await fails(), [
      "q1 has only a figure change that is not adopted yet: add the control strain to Table S3",
    ]);
    await w.service.adopt(w.doc, { ref: "R1.2", id });
    assert.deepEqual(await fails(), []);

    // The letter line of a table reads as one, with the name of the copy.
    const lines = buildLetterModel(await w.rev(), profile)
      .paragraphs.filter((p) => p.kind === "figure")
      .map((p) => p.text);
    assert.deepEqual(lines, ["[Revised Table S3, file Table_S3_v2.xlsx]"]);

    await w.service.reopen(w.doc, { ref: "R1.2", id });
    await w.service.reject(w.doc, { ref: "R1.2", id });
    assert.deepEqual(await fails(), [
      "q1 is neither cited nor declined: add the control strain to Table S3",
    ]);
    // A new file has its own frame.
    const added = await w.service.propose(
      w.doc,
      w.input({ kind: "new", before: "", figure: "Table S5", after: "figures/Table_S3_v1.xlsx" }),
    );
    await w.service.adopt(w.doc, { ref: "R1.2", id: added.id });
    assert.deepEqual(
      buildLetterModel(await w.rev(), profile)
        .paragraphs.filter((p) => p.kind === "figure")
        .map((p) => p.text),
      ["[New Table S5, file Table_S3_v1.xlsx]"],
    );
  } finally {
    await w.cleanup();
  }
});

// ---- over HTTP and on the page -------------------------------------------------

test("over HTTP: the list of files holds the table, the picture route answers 415, Adopt copies it", async () => {
  const w = await figureWorld({ runPy: fakeRunPy(), comments: COMMENTS });
  try {
    await mkdir(join(w.dir, "figures"), { recursive: true });
    await writeFile(join(w.dir, "figures", "Table_S3_v2.xlsx"), V2);
    const files = (await w.call("files", { query: { doc: w.doc } })).json.files.map((f) => f.path);
    assert.ok(files.includes("figures/Table_S3_v2.xlsx"), files.join(", "));
    const p = await w.propose({
      ref: "R1.2",
      kind: "new",
      figure: "Table S3",
      before: undefined,
      after: "figures/Table_S3_v2.xlsx",
    });
    assert.equal(p.status, 200, JSON.stringify(p.json));
    const id = p.json.id;
    const picture = await w.call("render", {
      query: { doc: w.doc, ref: "R1.2", id, side: "after" },
    });
    assert.equal(picture.status, 415);
    assert.deepEqual(picture.json, { ok: false, error: "No preview for this file type." });
    const adopt = await w.call("adopt", { body: { doc: w.doc, ref: "R1.2", id } });
    assert.equal(adopt.status, 200);
    assert.equal(adopt.json.adopted.path, "figures/submission/Table_S3_v2.xlsx");
    assert.equal(await readFile(join(w.dir, adopt.json.adopted.path), "utf8"), V2);
  } finally {
    await w.cleanup();
  }
});

test("the page shows a table record with its name, size and type, says there is no preview and asks for no picture", async () => {
  const sha = "a".repeat(64);
  const file = (path, bytes) => ({ path, sha256: sha, bytes, mtime: "2026-01-01T00:00:00.000Z" });
  const record = {
    id: "R1.2-f1",
    figure: "Table S3",
    kind: "revised",
    requirement_ids: ["q1"],
    summary: "The control strain was added as a row.",
    before: { source: file("figures/Table_S3_v1.xlsx", 2048), view: null },
    after: { source: file("figures/Table_S3_v2.xlsx", 4096), view: null },
    renders: { before: null, after: null },
    facts: { before: null, after: null },
    state: "proposed",
    adopted: null,
    created_at: "2026-01-01T00:00:00.000Z",
    preview: false,
    findings: [],
    problems: [],
  };
  const comments = [
    {
      ref: "R1.2",
      reviewer: 1,
      section: null,
      text: "Add the control strain to Table S3.",
      requirements: [{ id: "q1", text: "add the control strain", outside: true, why: "A file." }],
      figures: [record],
    },
  ];
  const t = await figuresPage("", {
    manuscripts: [{ doc: "/m/paper.docx", path: "/m/paper.docx" }],
    comments,
  });
  const card = t.card("R1.2");
  const text = textOf(card);
  assert.match(text, /Table S3 \(revised\)/);
  assert.match(text, /Waiting for your decision/);
  assert.match(text, /No preview for this file type\./);
  assert.match(text, /Table_S3_v1\.xlsx \(2 KB\)/);
  assert.match(text, /Table_S3_v2\.xlsx \(4 KB\)/);
  assert.equal(text.split("XLSX file").length - 1, 2);
  // Nothing that only a figure has: no page size, no fonts, no comparison modes.
  for (const never of [
    "Page size",
    "Fonts",
    "Live text",
    "Overlay",
    "Not read",
    "Loading the pictures",
  ])
    assert.equal(text.includes(never), false, never);
  assert.equal(t.asked.filter((u) => u.includes("/api/render")).length, 0);
  // The decision is offered as for a figure.
  const buttons = findAll(card, (n) => n.tagName === "BUTTON").map(textOf);
  assert.ok(buttons.includes("Adopt") && buttons.includes("Reject"), buttons.join(", "));
  // The words of the page name files, not only figures.
  assert.ok(buttons.includes("Add a figure or file change"));
  assert.match(text, /A figure or file waits for a decision/);
  // No id and no hash outside the folded details.
  const details = findAll(card, (n) => n.tagName === "DETAILS")
    .map(textOf)
    .join("");
  const main = text.replace(details, "");
  assert.equal(main.includes(sha), false);
  assert.equal(main.includes("R1.2-f1"), false);
  assert.ok(details.includes(sha));
});
