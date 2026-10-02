// The figures page and its API on the daemon's HTTP server.
//
//   GET  /figures, /figures/           the page (taskpane/figures/index.html)
//   GET  /figures/figures.js, .css     its two files
//   GET  /figures/api/manuscripts      documents open in a connected Word pane
//   GET  /figures/api/state?doc=       comments, requests, figure changes
//   GET  /figures/api/files?doc=       figure files in the figures folder
//   POST /figures/api/propose          { doc, ref, figure, kind, ... }
//   POST /figures/api/update           { doc, ref, id, patch }
//   POST /figures/api/adopt | reject | reopen     { doc, ref, id }
//   GET  /figures/api/render?doc=&ref=&id=&side=  the PNG of one side
//
// Plain HTTP, not the WebSocket: a hello with the document of a Word pane
// would take the place of that pane (bridge.mjs).
//
// Who may ask. Every route answers only a request sent to this server by its
// own address (Host) and, when the browser names where the request comes from
// (Origin), only from the page of this server. The API also wants the bridge
// token in the X-Sidewrite-Token header; the page reads it from
// /bridge-token, which another site cannot read. `doc` must be the document
// of a Word pane that is connected now. A render is found by the ids of a
// recorded figure change: no request names a path to read back.
//
// A supplementary document in the revision of a manuscript takes no new
// record: the requests are those of the manuscript, whose letter counts the
// record, so propose answers "Record this under <manuscript>." there. What
// such a file recorded before it joined can still be decided.
//
// After a request that changed something, the Word panes of that manuscript
// are told ({ type: "review_event", event: "figures_changed", ref }), so the
// Review and Response tabs show the new state without a press on Refresh.

import { createReadStream } from "node:fs";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { timingSafeEqual } from "node:crypto";
import { loadRevision, localDocPath, revisionPathFor } from "./revision-store.mjs";
import { linkedSummary } from "./revision-parts.mjs";

export const FIGURES_PREFIX = "/figures";
const PAGE = {
  "/figures/": ["index.html", "text/html; charset=utf-8"],
  "/figures/figures.js": ["figures.js", "application/javascript; charset=utf-8"],
  "/figures/figures.css": ["figures.css", "text/css; charset=utf-8"],
};
const GETS = new Set(["state", "files", "render"]);
const POSTS = new Set(["propose", "update", "adopt", "reject", "reopen"]);
const MAX_BODY = 1 << 20;

const sameToken = (a, b) => {
  const x = Buffer.from(String(a ?? ""));
  const y = Buffer.from(String(b ?? ""));
  return x.length === y.length && timingSafeEqual(x, y);
};

function readBody(req) {
  return new Promise((done, fail) => {
    const chunks = [];
    let size = 0;
    req.on("data", (c) => {
      size += c.length;
      if (size > MAX_BODY) {
        fail(Object.assign(new Error("The request is too large."), { status: 413 }));
        req.destroy();
      } else chunks.push(c);
    });
    req.on("end", () => {
      try {
        const data = JSON.parse(Buffer.concat(chunks).toString("utf8") || "{}");
        if (!data || typeof data !== "object" || Array.isArray(data))
          throw new Error("not an object");
        done(data);
      } catch {
        fail(Object.assign(new Error("The request body is not JSON."), { status: 400 }));
      }
    });
    req.on("error", fail);
  });
}

// origin: the address of this HTTP server ("http://127.0.0.1:47834").
// token: the bridge token. listPanes: () => [{ key, host, activeDoc }] of the
// connected panes (bridge.listPanes). service: createFigureService().
// pageDir: the folder of the page files. sendEvent: (key, message) => sends
// a message to the pane `key` (bridge.sendToTaskpane).
export function createFigureHttp({
  origin,
  token,
  listPanes,
  service,
  pageDir,
  sendEvent = () => {},
  env = process.env,
  log = console.log,
}) {
  if (!token) throw new Error("createFigureHttp requires a token");
  const host = new URL(origin).host;

  // Documents of connected Word panes that the daemon can open, once each:
  // [{ doc, path, keys }]. doc: the name the pane gave (what a request sends
  // back). keys: the panes that have it open.
  function manuscripts() {
    const seen = new Map();
    for (const p of listPanes()) {
      if (p.host !== "word" || typeof p.activeDoc !== "string") continue;
      const path = localDocPath(p.activeDoc, env);
      if (!path) continue;
      if (!seen.has(p.activeDoc)) seen.set(p.activeDoc, { doc: p.activeDoc, path, keys: [] });
      seen.get(p.activeDoc).keys.push(p.key);
    }
    return [...seen.values()];
  }
  // The panes of a manuscript learn that its figure changes are not as they
  // were. A pane that is gone by now is nothing to report to the page.
  function tellPanes(doc, ref) {
    for (const key of manuscripts().find((m) => m.doc === doc)?.keys ?? []) {
      try {
        sendEvent(key, { type: "review_event", event: "figures_changed", ref: ref ?? null });
      } catch (e) {
        log(`[figures] no figures_changed for a pane: ${e?.message ?? e}`);
      }
    }
  }
  function openDoc(doc) {
    const hit = manuscripts().find((m) => m.doc === doc);
    if (!hit)
      throw Object.assign(
        new Error(
          "This manuscript is not open in a Word pane. Open it in Word with the Sidewrite pane, then reload.",
        ),
        { status: 403 },
      );
    return hit.path;
  }

  // A new record in the file of a supplement: refused, with where it goes.
  // A review file that cannot be read is left to the service to report.
  async function refuseInPart(path) {
    let main = null;
    try {
      main = linkedSummary(path, await loadRevision(revisionPathFor(path), path)).part_of;
    } catch {
      main = null;
    }
    if (main) throw Object.assign(new Error(`Record this under ${main.name}.`), { status: 409 });
  }

  const json = (res, status, body) => {
    res.writeHead(status, { "Content-Type": "application/json; charset=utf-8" });
    res.end(JSON.stringify(body));
  };

  async function api(req, res, name, query) {
    if (!sameToken(req.headers["x-sidewrite-token"], token))
      return json(res, 401, { ok: false, error: "Reload the page: it has no valid token." });
    if (req.method === "GET") {
      if (name === "manuscripts")
        return json(res, 200, {
          ok: true,
          manuscripts: manuscripts().map((m) => ({ doc: m.doc, path: m.path })),
        });
      if (!GETS.has(name)) return json(res, 404, { ok: false, error: "No such request." });
      const doc = openDoc(query.get("doc"));
      if (name === "state") return json(res, 200, { ok: true, ...(await service.state(doc)) });
      if (name === "files") return json(res, 200, { ok: true, ...(await service.files(doc)) });
      if (name === "render") {
        const { png } = await service.render(doc, {
          ref: query.get("ref"),
          id: query.get("id"),
          side: query.get("side"),
        });
        res.writeHead(200, { "Content-Type": "image/png" });
        createReadStream(png)
          .on("error", () => res.destroy())
          .pipe(res);
        return;
      }
    } else if (req.method === "POST" && POSTS.has(name)) {
      const body = await readBody(req);
      const doc = openDoc(body.doc);
      if (name === "propose") await refuseInPart(doc);
      const out = await service[name](doc, body);
      tellPanes(body.doc, typeof body.ref === "string" ? body.ref : null);
      return json(res, 200, { ok: true, ...out });
    }
    return json(res, 404, { ok: false, error: "No such request." });
  }

  // true: the request was a figures request and is answered.
  return async function handle(req, res) {
    const url = new URL(req.url || "/", origin);
    const path = url.pathname;
    if (path !== FIGURES_PREFIX && !path.startsWith(`${FIGURES_PREFIX}/`)) return false;
    try {
      if (req.headers.host !== host || (req.headers.origin && req.headers.origin !== origin)) {
        json(res, 403, { ok: false, error: "Open this page at " + origin + FIGURES_PREFIX + "/" });
        return true;
      }
      if (path.startsWith(`${FIGURES_PREFIX}/api/`)) {
        await api(req, res, path.slice(`${FIGURES_PREFIX}/api/`.length), url.searchParams);
        return true;
      }
      if (req.method !== "GET") {
        json(res, 405, { ok: false, error: "GET only." });
      } else if (path === FIGURES_PREFIX) {
        res.writeHead(302, { Location: `${FIGURES_PREFIX}/` }).end();
      } else if (PAGE[path]) {
        const [file, type] = PAGE[path];
        const data = await readFile(join(pageDir, file));
        res.writeHead(200, { "Content-Type": type, "X-Content-Type-Options": "nosniff" });
        res.end(data);
      } else json(res, 404, { ok: false, error: "Not found." });
    } catch (e) {
      const status = Number.isInteger(e?.status) ? e.status : 500;
      if (status === 500) log(`[figures] ${req.method} ${path} failed: ${e?.stack ?? e}`);
      if (!res.headersSent)
        json(res, status, { ok: false, error: e?.message ?? "The request failed." });
      else res.destroy();
    }
    return true;
  };
}
