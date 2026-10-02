// A manuscript folder in a temp place with the figure lane around it, wired
// as in daemon/index.mjs but on ports of its own: the real bridge (port 0)
// with a real WebSocket client in the role of the Word pane, the figure
// service, and an HTTP server (port 0) with /bridge-token and the figures
// routes. Used by the figure tests and by a script that takes screenshots of
// the page; no Word and no model are needed.
import { createServer, request } from "node:http";
import { createHash } from "node:crypto";
import { mkdir, mkdtemp, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { WebSocket } from "ws";
import { createBridge } from "../../daemon/bridge.mjs";
import { createFigureService } from "../../daemon/figure-service.mjs";
import { createFigureHttp } from "../../daemon/figure-http.mjs";
import { emptyRevision, loadRevision, saveRevision } from "../../daemon/revision-store.mjs";

export const ENV = { CC_OFFICE_WSL_PATHS: "0", CC_OFFICE_CONFIG_DIR: "/nonexistent-config" };
export const TOKEN = "figure-test-token";
const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..", "..");

// An invented review: one request that only a figure can answer (marked as
// outside the Word file, as a review turn marks it) and one text request.
export const COMMENTS = [
  {
    ref: "R1.1",
    reviewer: 1,
    text: "The growth curve in Figure 2 has no control, and the time axis gives no unit.",
    requirements: [
      { id: "q1", text: "add the control curve to Figure 2" },
      { id: "q2", text: "give the unit of the time axis" },
    ],
    requirements_outside: [
      { id: "q1", why: "The figure is a picture file.", at: "2026-01-01T00:00:00.000Z" },
    ],
  },
  { ref: "R1.2", reviewer: 1, text: "Please state the incubation temperature." },
];

// Two small SVG files: an invented growth curve, the second with a control
// curve and wider.
export const svg = (after) =>
  `<svg xmlns="http://www.w3.org/2000/svg" width="${after ? 400 : 320}" height="240">` +
  `<rect width="100%" height="100%" fill="white"/>` +
  `<path d="M50 190 L${after ? 380 : 300} 40" stroke="blue" fill="none"/>` +
  (after ? `<path d="M50 190 L380 90" stroke="red" fill="none"/>` : "") +
  `<text x="14" y="125" font-size="12">OD600</text></svg>\n`;

export const sha256 = async (path) =>
  createHash("sha256")
    .update(await readFile(path))
    .digest("hex");

// Every file under dir, as sorted relative paths with "/" separators.
export async function listing(dir, prefix = "") {
  const out = [];
  for (const e of await readdir(dir, { withFileTypes: true })) {
    const rel = prefix ? `${prefix}/${e.name}` : e.name;
    if (e.isDirectory()) out.push(...(await listing(join(dir, e.name), rel)));
    else out.push(rel);
  }
  return out.sort();
}

// A stand-in for tools/figure_render.py: writes a file at --out and gives
// facts that depend on the name of the input. missing: PyMuPDF is not there.
export function fakeRunPy({ missing = false } = {}) {
  const calls = [];
  const run = async (script, args) => {
    calls.push(args);
    if (missing)
      return { ok: false, code: 3, json: { ok: false, code: "no_pymupdf", error: "no PyMuPDF" } };
    if (args[0] === "--probe") return { ok: true, code: 0, json: { ok: true } };
    const src = args[args.indexOf("--in") + 1];
    const out = args[args.indexOf("--out") + 1];
    await writeFile(out, `PNG of ${src}`, { flag: "wx" });
    const wide = /after/.test(src);
    return {
      ok: true,
      code: 0,
      json: {
        ok: true,
        type: "SVG",
        pages: 1,
        width_px: wide ? 834 : 667,
        height_px: 500,
        width_pt: wide ? 400 : 320,
        height_pt: 240,
        words: 1,
        images: 0,
        paths: wide ? 2 : 1,
        fonts: [],
      },
    };
  };
  run.calls = calls;
  return run;
}

// The manuscript folder alone: paper.docx (a stand-in), its revision.json
// and figures/fig2_before.svg, figures/fig2_after.svg.
export async function manuscriptFolder({ comments = COMMENTS } = {}) {
  const dir = await mkdtemp(join(tmpdir(), "figure-lane-"));
  const doc = join(dir, "paper.docx");
  await writeFile(doc, "not a real docx");
  await saveRevision(join(dir, "paper.revision.json"), { ...emptyRevision(doc), comments });
  await mkdir(join(dir, "figures"));
  await writeFile(join(dir, "figures", "fig2_before.svg"), svg(false));
  await writeFile(join(dir, "figures", "fig2_after.svg"), svg(true));
  return {
    dir,
    doc,
    rev: () => loadRevision(join(dir, "paper.revision.json"), doc),
    cleanup: () => rm(dir, { recursive: true, force: true }),
  };
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
async function until(fn, what, ms = 3000) {
  const end = Date.now() + ms;
  for (;;) {
    const v = await fn();
    if (v) return v;
    if (Date.now() > end) throw new Error(`timed out waiting for ${what}`);
    await sleep(10);
  }
}

// runPy: undefined runs the real Python tool. pane: false starts no pane.
export async function figureWorld({ comments = COMMENTS, runPy, pane = true } = {}) {
  const m = await manuscriptFolder({ comments });
  const lines = [];
  const bridge = createBridge({ port: 0, token: TOKEN, allowedOrigins: [], extraHandlers: {} });
  await until(() => bridge.address(), "bridge listening");
  const service = createFigureService({ runPy, env: ENV, log: (s) => lines.push(s) });

  let handle = null;
  const server = createServer(async (req, res) => {
    if (req.url === "/bridge-token") return res.writeHead(200).end(TOKEN);
    if (!(await handle(req, res))) res.writeHead(404).end("not a figures route");
  });
  await new Promise((r) => server.listen(0, "127.0.0.1", r));
  const origin = `http://127.0.0.1:${server.address().port}`;
  handle = createFigureHttp({
    origin,
    token: TOKEN,
    listPanes: () => bridge.listPanes(),
    service,
    pageDir: join(ROOT, "taskpane", "figures"),
    // As in daemon/index.mjs: what the page changed is told to the pane.
    sendEvent: (key, obj) => bridge.sendToTaskpane(obj, key),
    env: ENV,
    log: (s) => lines.push(s),
  });

  const sockets = [];
  async function openPane(activeDoc = m.doc, host = "word") {
    const ws = new WebSocket(`ws://127.0.0.1:${bridge.address().port}`);
    sockets.push(ws);
    let welcome = false;
    const frames = []; // every message the pane got after its welcome
    ws.on("message", (raw) => {
      const msg = JSON.parse(raw.toString());
      if (msg.type === "welcome") welcome = true;
      else frames.push(msg);
    });
    await new Promise((resolve, reject) => {
      ws.on("open", resolve);
      ws.on("error", reject);
    });
    ws.send(JSON.stringify({ type: "hello", token: TOKEN, host, active_doc: activeDoc }));
    await until(() => welcome, "welcome");
    return {
      frames,
      async close() {
        const closed = new Promise((r) => ws.on("close", r));
        ws.close();
        await closed;
        await until(
          () => !bridge.listPanes().some((p) => p.activeDoc === activeDoc && p.host === host),
          "the pane to be gone",
        );
      },
    };
  }
  const wordPane = pane ? await openPane() : null;

  // One request. headers: null sends no token; host: another Host header.
  // Returns { status, json, body (Buffer), headers }.
  function call(
    name,
    { query = null, body = null, token = TOKEN, headers = {}, host = null, path = null } = {},
  ) {
    const url = new URL(path ?? `/figures/api/${name}`, origin);
    for (const [k, v] of Object.entries(query ?? {})) url.searchParams.set(k, v);
    const data = body ? Buffer.from(JSON.stringify(body)) : null;
    return new Promise((resolve, reject) => {
      const req = request(
        url,
        {
          method: data ? "POST" : "GET",
          headers: {
            ...(token ? { "X-Sidewrite-Token": token } : {}),
            ...(data ? { "Content-Type": "application/json", "Content-Length": data.length } : {}),
            ...(host ? { Host: host } : {}),
            ...headers,
          },
        },
        (res) => {
          const chunks = [];
          res.on("data", (c) => chunks.push(c));
          res.on("end", () => {
            const buf = Buffer.concat(chunks);
            let json = null;
            try {
              json = JSON.parse(buf.toString("utf8"));
            } catch {
              /* not JSON */
            }
            resolve({ status: res.statusCode, json, body: buf, headers: res.headers });
          });
        },
      );
      req.on("error", reject);
      req.end(data ?? undefined);
    });
  }

  return {
    ...m,
    lines,
    bridge,
    service,
    origin,
    call,
    openPane,
    wordPane,
    // The usual figure change of the tests: Figure 2, before and after.
    propose: (extra = {}) =>
      call("propose", {
        body: {
          doc: m.doc,
          ref: "R1.1",
          figure: "Figure 2",
          kind: "revised",
          requirement_ids: ["q1"],
          summary: "The control curve was added.",
          before: "figures/fig2_before.svg",
          after: "figures/fig2_after.svg",
          ...extra,
        },
      }),
    async cleanup() {
      for (const ws of sockets) if (ws.readyState === WebSocket.OPEN) ws.close();
      await new Promise((r) => {
        server.close(r);
        server.closeAllConnections?.();
      });
      await bridge.close();
      await m.cleanup();
    },
  };
}
