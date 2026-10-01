// The office tools over local streamable HTTP, for review turns run by a CLI
// provider (Codex, Antigravity). index.mjs mounts handle() on the daemon's
// own HTTP server (127.0.0.1 only) under /mcp/<route>.
//
// What a route is: bind(route, { token, key, ref }) opens it for ONE review
// turn (pane key and reviewer comment ref) and returns the function that
// closes it. cli-turns.mjs binds a random route with a random token when a
// CLI turn starts and closes it when the turn ends.
//
// What is refused, in this order, before MCP sees the request:
//   403  the peer is not loopback; the Host header is not a loopback name
//        (DNS rebinding); any Origin header or a cross-site fetch (a page in
//        a browser; a CLI sends neither)
//   401  the route is not bound, or the Bearer token is not the route's
//        (compared in constant time). The answer has no body.
//   405  anything but POST
//   409  the pane has no review turn for the comment the route was bound to
//
// What is served: only the review-turn tool set (the read tools and
// office_tracked_edit), built here with the same createOfficeBridgeMcp, so a
// call goes through the same guardedCall and the same applyTrackedEdit as a
// Claude turn. The set is fixed in this file: nothing in a binding or a
// request can add a tool. Each request gets a fresh server and a stateless
// transport.

import { createHash, timingSafeEqual } from "node:crypto";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { createOfficeBridgeMcp } from "./office-tools.mjs";
import { REVIEW_READ_TOOLS, REVIEW_WRITE_TOOL } from "./review-tools.mjs";

export const REVIEW_TURN_TOOLS = Object.freeze([...REVIEW_READ_TOOLS, REVIEW_WRITE_TOOL]);

const MAX_BODY = 4 * 1024 * 1024;
const ROUTE_RE = /^[A-Za-z0-9_-]{1,64}$/;
const MIN_TOKEN_LENGTH = 32;
const LOOPBACK_PEERS = new Set(["127.0.0.1", "::1", "::ffff:127.0.0.1"]);
const LOOPBACK_HOSTS = new Set(["127.0.0.1", "localhost", "[::1]"]);

// Constant time for any two strings: the digests have one length.
export function sameToken(a, b) {
  const h = (s) =>
    createHash("sha256")
      .update(String(s ?? ""))
      .digest();
  return timingSafeEqual(h(a), h(b)) && String(a ?? "").length === String(b ?? "").length;
}

function hostName(header) {
  const h = String(header ?? "").toLowerCase();
  const m = /^(\[[0-9a-f:]+\]|[^:]+)(?::\d+)?$/.exec(h);
  return m ? m[1] : null;
}

async function readBody(req) {
  const chunks = [];
  let n = 0;
  for await (const c of req) {
    n += c.length;
    if (n > MAX_BODY) throw Object.assign(new Error("body too large"), { status: 413 });
    chunks.push(c);
  }
  const s = Buffer.concat(chunks).toString("utf8");
  try {
    return s ? JSON.parse(s) : undefined;
  } catch {
    throw Object.assign(new Error("body is not JSON"), { status: 400 });
  }
}

// bridge: bridge.mjs (callTaskpaneTool). getReviewTurn(key): the running
// review turn of a pane ({ ref }) or null. trackedEditDeps(key):
// review-service.mjs trackedEditDeps.
export function createMcpHttp({ bridge, getReviewTurn, trackedEditDeps, log = console.log }) {
  const routes = new Map();
  // The turn the route was opened for, or null once it ended or another
  // comment's turn runs on the pane.
  const liveTurn = (b) => {
    const t = getReviewTurn(b.key) ?? null;
    return t && t.ref === b.ref ? t : null;
  };
  const refuse = (res, status) => (res.writeHead(status).end(), true);

  return {
    // Only token, key and ref are read from the binding.
    bind(route, { token, key, ref } = {}) {
      if (!ROUTE_RE.test(String(route))) throw new Error("bad MCP route name");
      if (typeof token !== "string" || token.length < MIN_TOKEN_LENGTH) {
        throw new Error("the MCP token is too short");
      }
      if (!key || !ref) throw new Error("an MCP route needs a pane key and a comment ref");
      if (routes.has(route)) throw new Error(`MCP route ${route} is already bound`);
      const binding = Object.freeze({ token, key, ref });
      routes.set(route, binding);
      return () => routes.get(route) === binding && routes.delete(route);
    },
    bound: (route) => routes.has(route),
    // Returns false when the path is not an MCP route (the caller goes on).
    async handle(req, res, urlPath) {
      const m = /^\/mcp\/([^/]+)$/.exec(urlPath);
      if (!m) return false;
      if (!LOOPBACK_PEERS.has(req.socket?.remoteAddress)) return refuse(res, 403);
      if (!LOOPBACK_HOSTS.has(hostName(req.headers.host))) return refuse(res, 403);
      const site = req.headers["sec-fetch-site"];
      if (req.headers.origin !== undefined || (site && site !== "none")) return refuse(res, 403);
      const b = routes.get(m[1]);
      // An unbound route is compared too, so it takes the same time.
      const expected = `Bearer ${b?.token ?? "-".repeat(MIN_TOKEN_LENGTH)}`;
      if (!sameToken(req.headers.authorization, expected) || !b) return refuse(res, 401);
      if (req.method !== "POST") return refuse(res, 405);
      if (!liveTurn(b)) return refuse(res, 409);
      let body;
      try {
        body = await readBody(req);
      } catch (e) {
        return refuse(res, e.status ?? 400);
      }
      const server = createOfficeBridgeMcp(bridge, "word", b.key, {
        getReviewTurn: () => liveTurn(b),
        trackedEdit: trackedEditDeps(b.key),
        only: new Set(REVIEW_TURN_TOOLS),
      }).instance;
      const transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: undefined,
        enableJsonResponse: true,
      });
      res.on("close", () => {
        const quiet = (p) => Promise.resolve(p).then(undefined, () => undefined);
        quiet(transport.close());
        quiet(server.close());
      });
      log(`[mcp-http] ${req.method} route=${m[1]} key=${b.key} ref=${b.ref}`);
      await server.connect(transport);
      await transport.handleRequest(req, res, body);
      return true;
    },
  };
}
