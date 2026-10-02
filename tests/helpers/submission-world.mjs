// An invented paper in a folder of its own with the two services the
// submission checklist reads from: the Response tab's (the letter) and the
// Tools tab's (the copies and the checklist itself). Fake Python tools, no
// Word, no port. Nothing here is from a real manuscript.
import { existsSync } from "node:fs";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, join } from "node:path";
import { createResponseService } from "../../daemon/response-service.mjs";
import { createAuxService } from "../../daemon/aux-service.mjs";
import { emptyRevision, loadRevision, saveRevision } from "../../daemon/revision-store.mjs";

export const ENV = { CC_OFFICE_WSL_PATHS: "0", CC_OFFICE_CONFIG_DIR: "/nonexistent-config" };
export const SECTION = "2. Materials and Methods";
export const BEFORE =
  "Flasks were shaken. The incubation temperature was held at 30 degrees. Samples were drawn daily.";
export const AFTER =
  "Flasks were shaken. The incubation temperature was held at 30 °C. Samples were drawn daily.";

// The manuscript as the pane sends it. paragraph: the second paragraph.
export const snapshotOf = (paragraph = { text: AFTER }, comments = []) => ({
  paragraphs: [
    { index: 0, id: "P0", text: SECTION, w_ins: 0, w_del: 0 },
    { index: 1, id: "P1", w_ins: 0, w_del: 0, ...paragraph },
  ],
  comments,
});
// The one change of the paper: accepted, still tracked, or rejected in Word.
export const ACCEPTED = snapshotOf();
export const PENDING = snapshotOf({
  text: BEFORE.replace("30 degrees", "30 °C30 degrees"),
  reviewed_text: AFTER,
  w_ins: 1,
  w_del: 1,
});
export const REJECTED = snapshotOf({ text: BEFORE });
export const HEADINGS = [{ index: 0, level: 1, text: SECTION }];
export const PDF = Buffer.from("%PDF-1.7\nfake\n").toString("base64");

const CHANGE = {
  id: "R1.1-1",
  word_comment_id: null,
  paragraph_index: 1,
  paragraph_id: "P1",
  kind: "replace",
  find: "30 degrees",
  text: "30 °C",
  before: BEFORE,
  after: AFTER,
  rationale: {
    what: "Unit written as °C.",
    why: "The reviewer asked for it.",
    addresses: "unit",
    requirement_ids: ["q1"],
  },
  context: { left: "was held at ", right: ". Samples were" },
  created_at: "2026-01-01T00:00:00.000Z",
};
export const commentsOf = (over = {}) => [
  {
    ref: "R1.1",
    reviewer: 1,
    text: "Please state the unit.",
    requirements: [{ id: "q1", text: "state the unit" }],
    changes: [CHANGE],
    ...over,
  },
];

// The Python tools by script name. Each one writes what the real one writes;
// `tools` replaces one of them.
function fakePy(tools = {}) {
  return async (script, args) => {
    const name = basename(script);
    const arg = (flag) => args[args.indexOf(flag) + 1];
    if (tools[name]) return tools[name](args, arg);
    if (name === "pdf-line-map.py") {
      const requests = JSON.parse(await readFile(args[2], "utf8"));
      const results = requests.map((r) => ({
        id: r.id,
        match: { start: { page: 1, line_no: 9 }, end: { page: 1, line_no: 9 }, match_count: 1 },
      }));
      return { ok: true, code: 0, json: { lines_indexed: 40, monotonic: true, results } };
    }
    if (name === "response_docx.py") {
      const out = arg("--out");
      if (existsSync(out) && !args.includes("--overwrite"))
        return { ok: false, code: 4, json: { ok: false, code: "exists" } };
      await writeFile(out, `letter ${await readFile(arg("--model"), "utf8")}`);
      const paragraphs = JSON.parse(await readFile(arg("--model"), "utf8")).paragraphs.length;
      return { ok: true, code: 0, json: { ok: true, path: out, paragraphs } };
    }
    if (name === "response_verify.py")
      return { ok: true, code: 0, json: { ok: true, checked: {}, failures: [] } };
    throw new Error(`unexpected tool ${name}`);
  };
}

export async function world({ seed = commentsOf(), revision = {}, tools = {} } = {}) {
  const dir = await mkdtemp(join(tmpdir(), "submission-"));
  const doc = join(dir, "paper.docx");
  const revPath = join(dir, "paper.revision.json");
  await writeFile(doc, "manuscript as saved");
  await saveRevision(revPath, { ...emptyRevision(doc), comments: seed, ...revision });
  const waiters = new Map();
  const done = new Map();
  const sendEvent = (_key, e) => {
    if (e.event !== "job_done") return;
    if (waiters.has(e.job_id)) waiters.get(e.job_id)(e.result);
    else done.set(e.job_id, e.result);
  };
  const deps = {
    getContext: () => ({ activeDoc: doc }),
    sendEvent,
    runPy: fakePy(tools),
    log: () => {},
    env: ENV,
  };
  const handlers = {
    ...createResponseService({ ...deps, oneshotJob: async () => ({ run: null, model: null }) })
      .handlers,
    ...createAuxService(deps).handlers,
  };
  // What the pane gets: the reply of a request, or the result of a job.
  const ask = (type, payload = {}) =>
    new Promise((resolve) => {
      const reply = (r) => {
        if (!r.accepted) return resolve(r);
        if (done.has(r.job_id)) resolve(done.get(r.job_id));
        else waiters.set(r.job_id, resolve);
      };
      handlers[type]({ type, request_id: "q1", ...payload }, reply, "k", "word");
    });
  const withDoc = (snapshot = ACCEPTED) => ({ snapshot, headings: HEADINGS });
  return {
    dir,
    doc,
    revPath,
    ask,
    withDoc,
    rev: () => loadRevision(revPath, doc),
    raw: () => readFile(revPath, "utf8"),
    // The rows of "Check now" for the document as `snapshot` reads it.
    async check(snapshot = ACCEPTED, more = {}) {
      const r = await ask("submission_check", { ...withDoc(snapshot), ...more });
      if (!r.ok) throw new Error(r.error);
      return Object.fromEntries(r.rows.map((row) => [row.id, row]));
    },
    // The comment answered and located: the letter passes every check.
    async answered(snapshot = ACCEPTED) {
      const saved = await ask("response_save", {
        ref: "R1.1",
        text: "We now state the unit.",
        ...withDoc(snapshot),
      });
      if (!saved.ok) throw new Error(saved.error);
      const located = await ask("response_locate", { ...withDoc(snapshot), pdf_base64: PDF });
      if (!located.ok) throw new Error(located.error);
    },
    cleanup: () => rm(dir, { recursive: true, force: true }),
  };
}
