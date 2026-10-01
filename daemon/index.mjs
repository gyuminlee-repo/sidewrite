import { query } from "@anthropic-ai/claude-agent-sdk";
import { readFile, writeFile, mkdir, chmod } from "node:fs/promises";
import { randomBytes } from "node:crypto";
import { createServer } from "node:http";
import { fileURLToPath } from "node:url";
import { dirname, resolve, extname, join, sep } from "node:path";
import { homedir } from "node:os";
import { createBridge } from "./bridge.mjs";
import { createOfficeBridgeMcp } from "./office-tools.mjs";
import {
  resolveWorkspaceRoot,
  suggestWorkspaceRoot,
  ensureWorkspaceMarker,
  toLocalPath,
} from "./workspace.mjs";
import { randomUUID } from "node:crypto";
import { getSessionId, saveSessionId, touchFolder } from "./sessions.mjs";
import { readTranscript } from "./transcript.mjs";
import { diag } from "./diag.mjs";
import { getContextEntries, setContextEntries } from "./context.mjs";
import {
  detectUsageLimit,
  turnCompleteEvents,
  turnOutcome,
  emptyTurnStats,
  noteMessage,
} from "./limit-detect.mjs";
import { createReviewService } from "./review-service.mjs";
import { createResponseService } from "./response-service.mjs";
import { createAuxService } from "./aux-service.mjs";
import { markStale } from "./response-quotes.mjs";
import { makeOneshot, ONESHOT_MODEL } from "./oneshot.mjs";
import {
  bestClaudeModel,
  createModelCache,
  refusalFallback,
  declineStep,
  staysOnFallback,
  modelChange,
  resolveModelChoice,
} from "./model-list.mjs";
import { createKeyStore } from "./providers/keys.mjs";
import { probeAll, createProbeCache } from "./providers/probe.mjs";
import { createOneshotRouter } from "./providers/oneshot-router.mjs";
import { createProviderSelection, providersListReply } from "./providers/selection.mjs";
import { chatOneshot } from "./providers/openai-compat.mjs";
import { codexOneshot } from "./providers/codex.mjs";
import { agyOneshot } from "./providers/agy.mjs";
import {
  buildQueryOptions,
  formatUsageLine,
  resolveSettingSources,
  userMcpEnabled,
} from "./query-options.mjs";
import { stat } from "node:fs/promises";
import { healthBody } from "./health.mjs";

const STARTED_AT = Date.now();
const __dirname = dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = resolve(__dirname, "..");

// Draftspect deliberately uses 47833/47834. Another local Office add-in
// on this machine may bind 47823/47824, so a distinct pair avoids a port
// clash when both run at once. Keep WS_PORT = HTTP_PORT - 1 (the taskpane
// derives nothing; the WS port is referenced explicitly in taskpane.js
// and the index.html CSP — change all three together).
const WS_PORT = 47833;
const HTTP_PORT = 47834;
const HTTP_ORIGIN = `http://127.0.0.1:${HTTP_PORT}`;
const TOKEN_FILE = join(homedir(), ".claude", "office-addins", "bridge-token");

// Bridge token — random per-daemon-start. The taskpane fetches it from the
// HTTP server's /bridge-token endpoint (same-origin, CORS-restricted) and
// includes it in the first WS hello. Any WS that doesn't present this token
// (or comes from an unknown origin) is closed.
const BRIDGE_TOKEN = randomBytes(24).toString("hex");
{
  await mkdir(dirname(TOKEN_FILE), { recursive: true });
  await writeFile(TOKEN_FILE, BRIDGE_TOKEN, { mode: 0o600 });
  try {
    await chmod(TOKEN_FILE, 0o600);
  } catch {}
  console.log(`[daemon] Bridge token written to ${TOKEN_FILE}`);
}

// ---------------------------------------------------------------------------
// CLI args
// ---------------------------------------------------------------------------
const matterFolder = process.argv[2]
  ? resolve(process.argv[2])
  : process.env.MATTER_FOLDER
    ? resolve(process.env.MATTER_FOLDER)
    : process.cwd();

console.log(`[daemon] Workspace folder (agent cwd): ${matterFolder}`);

// ---------------------------------------------------------------------------
// HTTP server: serve the taskpane assets so Word can load them.
// ---------------------------------------------------------------------------
const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".json": "application/json; charset=utf-8",
  ".xml": "application/xml; charset=utf-8",
};

const taskpaneDir = join(PROJECT_ROOT, "taskpane");

// A branded, actionable error page. Office renders whatever the manifest's
// SourceLocation returns inside the task pane, so a bare "Not found" (the
// old body) left users staring at two unhelpful words. The common real
// causes are all recoverable; spell them out. Served only to document
// navigations (Accept: text/html) — asset fetches still get terse text so
// nothing tries to parse HTML as JS/CSS.
const htmlEscape = (s) =>
  String(s).replace(
    /[&<>"']/g,
    (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c],
  );

function errorPageHtml(status, headline, requestedPath) {
  requestedPath = htmlEscape(requestedPath);
  return `<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Draftspect — ${status}</title>
<style>
  body{font:14px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
       color:#1a1a1a;margin:0;padding:28px 24px;background:#fff}
  h1{font-size:18px;margin:0 0 4px}
  .sub{color:#666;margin:0 0 18px}
  ol{padding-left:20px;margin:0 0 18px} li{margin:6px 0}
  code{background:#f2f2f2;padding:1px 5px;border-radius:4px;font-size:12px}
  .foot{color:#999;font-size:12px;border-top:1px solid #eee;padding-top:12px}
</style></head><body>
<h1>Draftspect couldn't load this panel</h1>
<p class="sub">${headline}</p>
<p>This usually means one of:</p>
<ol>
  <li>The <strong>Draftspect tray app isn't running</strong> — start it, then reopen this panel.</li>
  <li>Word/Excel cached an old add-in — <strong>fully quit the app (⌘Q / Alt+F4) and reopen it</strong> so it re-reads the add-in.</li>
  <li>This add-in's manifest points at a <strong>different port</strong> than the running Draftspect daemon (e.g. another add-in's daemon answered). Relaunch Draftspect, then quit &amp; reopen Word/Excel.</li>
  <li>If it persists, reinstall the add-in from the Draftspect tray menu.</li>
</ol>
<p class="foot">Draftspect daemon on <code>127.0.0.1:${HTTP_PORT}</code> · requested <code>${requestedPath}</code> · ${status}</p>
</body></html>`;
}

function sendError(req, res, status, headline, requestedPath) {
  const wantsHtml = (req.headers.accept || "").includes("text/html");
  if (wantsHtml) {
    res.writeHead(status, { "Content-Type": "text/html; charset=utf-8" });
    res.end(errorPageHtml(status, headline, requestedPath));
  } else {
    res.writeHead(status).end(headline);
  }
}

const http = createServer(async (req, res) => {
  // Restrictive CORS: only the taskpane's same origin (HTTP_ORIGIN) gets the
  // Access-Control-Allow-Origin header. Other origins (malicious web pages
  // running in a regular browser tab) hit a no-CORS-header response and the
  // browser blocks them from reading it. Same-origin requests from the
  // taskpane itself don't go through CORS at all.
  const reqOrigin = req.headers.origin || "";
  if (reqOrigin === HTTP_ORIGIN) {
    res.setHeader("Access-Control-Allow-Origin", reqOrigin);
  }
  // No-cache: Word's webview likes to cache aggressively. During dev we want
  // every reload to pick up the latest taskpane JS/CSS/HTML.
  res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0");
  res.setHeader("Pragma", "no-cache");
  res.setHeader("Expires", "0");
  if (req.method === "OPTIONS") {
    res.writeHead(204).end();
    return;
  }

  try {
    const urlPath = (req.url || "/").split("?")[0];

    // Up or not, for scripts (install, doctor). No token asked, none given.
    if (urlPath === "/health") {
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(
        JSON.stringify(
          healthBody({ startedAt: STARTED_AT, now: Date.now(), ws: WS_PORT, http: HTTP_PORT }),
        ),
      );
      return;
    }

    // Token endpoint: serves the per-daemon bridge token to the taskpane.
    // Only same-origin (i.e. cross-origin requests get blocked by CORS).
    if (urlPath === "/bridge-token") {
      res.writeHead(200, { "Content-Type": "text/plain; charset=utf-8" });
      res.end(BRIDGE_TOKEN);
      return;
    }

    const relPath = urlPath === "/" ? "/index.html" : urlPath;
    const fsPath = join(taskpaneDir, relPath);
    // Containment check. `join` already normalizes `../`, so the obvious
    // traversal is blocked — but a bare startsWith(taskpaneDir) would also
    // accept a sibling like `<…>/taskpane-evil/x`. Require an exact match
    // OR a path under `taskpaneDir` + separator.
    if (fsPath !== taskpaneDir && !fsPath.startsWith(taskpaneDir + sep)) {
      res.writeHead(403).end("Forbidden");
      return;
    }
    const ext = extname(fsPath).toLowerCase();
    const mime = MIME[ext] || "application/octet-stream";
    if (ext === ".html") {
      // Office's webview caches the JS/CSS bundle aggressively and ignores
      // our no-store headers — so a taskpane code change wouldn't take
      // effect even after reopening the pane. Cache-bust the local asset
      // refs with the per-start bridge token: every daemon restart yields a
      // fresh URL, forcing a re-fetch. (The handler strips the query string
      // before resolving the file, so `?v=` doesn't affect routing.)
      let html = await readFile(fsPath, "utf8");
      html = html.replace(/(\/shared\/(?:taskpane\.js|styles\.css))"/g, `$1?v=${BRIDGE_TOKEN}"`);
      res.writeHead(200, { "Content-Type": mime });
      res.end(html);
    } else {
      const data = await readFile(fsPath);
      res.writeHead(200, { "Content-Type": mime });
      res.end(data);
    }
  } catch (err) {
    const reqPath = (req.url || "/").split("?")[0];
    if (err.code === "ENOENT") {
      sendError(
        req,
        res,
        404,
        "That page or file isn’t served by this Draftspect daemon.",
        reqPath,
      );
    } else {
      console.error("[http]", err);
      sendError(
        req,
        res,
        500,
        "The Draftspect daemon hit an internal error serving this page.",
        reqPath,
      );
    }
  }
});

http.listen(HTTP_PORT, "127.0.0.1", () => {
  console.log(`[daemon] HTTP server listening on http://127.0.0.1:${HTTP_PORT}/`);
});

// ---------------------------------------------------------------------------
// IPC channel back to the Electron main process. Used to ask main.mjs to
// show a native macOS open panel for folder/file picking — synthetic
// in-page modals can't navigate to Google Drive, iCloud, recent items, or
// any of the other sources macOS users expect in NSOpenPanel. The channel
// is fd 3 (added in main.mjs's spawn options), wired through Node IPC.
// ---------------------------------------------------------------------------
const pendingPicks = new Map(); // id -> { resolve, reject, timer }
const PICK_TIMEOUT_MS = 5 * 60_000; // 5 min; the user might leave the dialog open

if (process.send) {
  process.on("message", (msg) => {
    if (msg?.type !== "pick_path_result") return;
    const entry = pendingPicks.get(msg.id);
    if (!entry) return;
    pendingPicks.delete(msg.id);
    clearTimeout(entry.timer);
    entry.resolve(msg);
  });
}

function pickPathFromMain({ include_files, default_path, title, button_label }) {
  if (!process.send) {
    return Promise.reject(new Error("No IPC channel to Electron main process"));
  }
  const id = randomUUID();
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      pendingPicks.delete(id);
      reject(new Error("Picker timed out"));
    }, PICK_TIMEOUT_MS);
    pendingPicks.set(id, { resolve, reject, timer });
    process.send({ type: "pick_path", id, include_files, default_path, title, button_label });
  });
}

// Word and Excel are fully independent surfaces: each host gets its own
// agent loop, session, message queue, transcript AND workspace. Nothing
// one host does ever touches the other (no shared "current session", no
// cross-host interrupt). Both maps are hoisted above createBridge so
// bridge handlers that fire during the top-level awaits below
// (preflightHttpMcpServers, etc.) don't hit a TDZ on these bindings.
//
//   sessions:        paneKey -> live session { key, host, cwd, sessionId,
//                    abortController, settled, interrupted }
//   workspaceByKey:  paneKey -> last-known cwd, so a pane that connects
//                    before its first message still resolves a workspace
//                    (and survives that pane's loop ending).
//
// paneKey identifies one open document (host + doc path), so two Word
// docs (or two workbooks) open at once each get their own independent
// loop/session/queue/workspace. The host is still carried for the tool
// family + per-host behavior. SDK-transcript resume is still keyed by
// (host, cwd) — two documents in the same folder share a resume hint;
// that's a deliberate, documented simplification, far better than the
// ping-pong of sharing a pane.
const sessions = new Map();
const workspaceByKey = new Map();
// Panes whose workspace the user set explicitly (Change workspace). Their
// workspace must NOT be re-derived from the document folder on reconnect.
const explicitWorkspaceKeys = new Set();
// Panes with a turn in flight (a user_message or a review turn was queued
// and no result/stop/stream end has come back yet). A review turn is only
// started on an idle pane, so the SDK cannot fold it into a running turn.
const turnBusy = new Set();
// paneKey -> { ref } while a Review-tab turn runs. office-tools.mjs reads it
// to refuse every document writer except office_tracked_edit.
const reviewTurnByKey = new Map();
// Per-pane sticky model choice (set by the taskpane composer). The SDK
// model is fixed per agent loop, so a change while a loop is live triggers
// a resuming restart (see the set_model handler).
const modelByKey = new Map();
// What the session of a pane was given of the manuscript in a review turn
// (review-prompt.mjs manuscriptBlock): pane key -> { hashes, fresh }. fresh:
// given in a turn that starts the session itself (a first turn, or the
// restart after a model fallback); that start keeps it. Any other start of
// a session, and a compaction of its history, forgets it: the next review
// turn then carries the whole manuscript again.
const manuscriptByKey = new Map();
// The models the signed-in account can use (model-list.mjs), fetched once at
// start and kept for ten minutes. Until it is known, and when the fetch
// fails, the three aliases and well-formed claude-* ids are accepted.
const modelCache = createModelCache();

// The model to pass to query(). Draftspect always pins an explicit model,
// never the CLI default: the one the pane chose, or, when it chose none (or
// has not said yet), the best the account lists (model-list.mjs
// bestClaudeModel).
function modelArgFor(key) {
  return resolveModelChoice(modelByKey.get(key), modelCache.current());
}
// Refusals of the chosen model in this pane (model-list.mjs declineStep):
// pane key -> record. After two in a row the session is left on the model
// the CLI moved it to, for as long as that choice stands.
const declinedByKey = new Map();
const staysOn = (key) => staysOnFallback(declinedByKey.get(key), modelArgFor(key));

function sessionFor(key) {
  return sessions.get(key) ?? null;
}

// This pane's workspace: its live session's cwd, else the last cwd we
// recorded for it, else the launch default.
function cwdForKey(key) {
  return sessionFor(key)?.cwd ?? workspaceByKey.get(key) ?? matterFolder;
}

// Resolve the session id to replay for this pane. Prefer the pane's live
// session id (set once the SDK reports init); otherwise the id persisted
// for (host, cwd) — covers the window before the SDK has re-inited.
async function resolveReplaySessionId(key, host, cwd) {
  const live = sessionFor(key);
  if (live?.sessionId) return live.sessionId;
  if (!host || !cwd) return null;
  try {
    return (await getSessionId(host, cwd)) ?? null;
  } catch {
    return null;
  }
}

// Reconstruct this pane's prior conversation from its .jsonl and push it
// to THAT pane only. Sent on every taskpane hello and after a workspace
// switch (cwd_changed). Empty events => fresh chat, no divider.
async function sendTranscriptReplayTo(key, host, cwd) {
  try {
    const sessionId = await resolveReplaySessionId(key, host, cwd);
    const { events, truncated } = sessionId
      ? await readTranscript(sessionId, { maxEvents: 200 })
      : { events: [], truncated: false };
    bridge.sendToTaskpane(
      {
        type: "transcript_replay",
        session_id: sessionId ?? null,
        truncated,
        events,
      },
      key,
    );
  } catch (err) {
    console.warn("[daemon] transcript replay failed:", err?.message ?? err);
  }
}

// Start (or resume) this pane's session on the next tick — past the
// current message handler / the agent loop's finally block (which clears
// sessions.get(key)), so startSessionForFolder builds cleanly. Every
// path that (re)starts a pane's session — first message, post-stream-end,
// post-Stop — funnels through here; `reason` only flavors the failure log.
// Per-key serialization + coalescing of (re)starts. Every restart trigger —
// first message, post-Stop, post-stream-end, model/context change, workspace
// switch — funnels through here. startSessionForFolder aborts whatever
// session is current for the key, so if several triggers fire close together
// (e.g. "slow first turn → Stop → switch model → resend") their setImmediate
// callbacks used to interleave: each abort killed the next start before it
// could `init`, leaving an orphaned loop with no consumer (the Excel hang).
//
// Now each key has a single-flight runner. While a start is in progress,
// later requests only overwrite `latest` — intermediate restarts collapse,
// and exactly one final start runs to completion uncontested, so it always
// reaches `init`.
const startQueue = new Map(); // key -> { running: boolean, latest: req | null }

function scheduleSessionStart(
  cwd,
  sessionId,
  key,
  host,
  reason,
  { replay = true, keepQueued = false } = {},
) {
  let q = startQueue.get(key);
  if (!q) {
    q = { running: false, latest: null };
    startQueue.set(key, q);
  }
  q.latest = { cwd, resumeId: sessionId, host: host ?? null, reason, replay, keepQueued };
  if (q.running) return; // the active runner will pick up `latest`
  q.running = true;
  setImmediate(() => runStartQueue(key));
}

async function runStartQueue(key) {
  const q = startQueue.get(key);
  if (!q) return;
  // startSessionForFolder returns once setup is done (it kicks the agent
  // loop off detached), so awaiting it serializes only the abort+create
  // step — exactly the part that must not interleave.
  while (q.latest) {
    const req = q.latest;
    q.latest = null;
    try {
      await startSessionForFolder(req.cwd, req.resumeId, {
        key,
        host: req.host,
        replay: req.replay,
        keepQueued: req.keepQueued,
      });
    } catch (err) {
      console.error(`[daemon] ${req.reason} session start failed:`, err?.message ?? err);
    }
  }
  q.running = false;
}

// Backoff for the post-stream-end auto-restart. When the SDK stream dies
// immediately without ever producing a `result` (persistent usage-limit,
// auth failure, or an SDK fault), the loop would otherwise respawn forever,
// spamming the pane with an error + turn_complete on every cycle. We track,
// per pane key, how many such restarts happened with no successful turn in
// between, inside a rolling window. Past the cap we stop auto-resurrecting
// that pane; an explicit new user message still gets a fresh attempt (it
// goes through the un-capped "user message" path) and a completed turn
// clears the record (see noteSuccessfulTurn).
const LOOP_RESTART_WINDOW_MS = 60_000;
const MAX_LOOP_RESTARTS = 5;
const loopRestartByKey = new Map(); // key -> { count, first }

// Records one failure restart for `key` and returns true if the loop may be
// auto-restarted, false once the cap is hit within the window.
function allowFailureRestart(key) {
  const now = Date.now();
  let rec = loopRestartByKey.get(key);
  if (!rec || now - rec.first > LOOP_RESTART_WINDOW_MS) {
    rec = { count: 0, first: now };
    loopRestartByKey.set(key, rec);
  }
  rec.count += 1;
  return rec.count <= MAX_LOOP_RESTARTS;
}

function noteSuccessfulTurn(key) {
  loopRestartByKey.delete(key);
}

// Called on every taskpane hello — for ANY open document, many panes
// possibly connected at once. Connecting a pane must NOT start the agent
// loop: connect-driven starts amplified the old connect/disconnect
// ping-pong and burn an unasked turn. We only re-render this pane's own
// transcript. The loop starts lazily on the first user message
// (onUserMessage → ensureLoopForMessage).
async function onPaneConnect(key, host, doc) {
  if (!key) return;
  // No provider is confirmed for this connection until the pane says which
  // one it stored (restore_providers) or the user picks one.
  selection.begin(key);
  // Resolve this pane's workspace from the open document's own folder,
  // server-side and immediately — deterministic, no loop start. Without
  // this, cwdForKey() falls back to matterFolder (the daemon's launch
  // cwd, often the repo root) until a message or the taskpane's set_cwd
  // round-trip lands, so get_context / replay read the WRONG folder's
  // CLAUDE.md (the symptom: context files bleeding across documents).
  // Skip if the user explicitly pinned this pane's workspace, or a live
  // session already owns the cwd.
  if (doc && !explicitWorkspaceKeys.has(key) && !sessionFor(key)) {
    try {
      const folder = await resolveWorkspaceRoot(doc);
      if (folder) workspaceByKey.set(key, folder);
    } catch {
      /* unresolvable (cloud/unsaved) — keep the fallback */
    }
  }
  const cwd = cwdForKey(key);
  diag(`hello → replay key=${key} cwd=${cwd} (no loop start on connect)`);
  await sendTranscriptReplayTo(key, host, cwd);
}

// Called when a pane's WebSocket closes for good (the bridge already
// freed its own pane/queue state). Prune the daemon-side per-key Maps so
// they don't grow for the life of the daemon — but ONLY when no live
// session owns this key. A session is deliberately kept alive across a
// transient disconnect (it resumes on reconnect with the same stable
// key); its own `finally` clears `sessions` when its loop actually ends.
function onPaneClose(key) {
  if (!key || sessionFor(key)) return;
  workspaceByKey.delete(key);
  explicitWorkspaceKeys.delete(key);
  loopRestartByKey.delete(key);
  modelByKey.delete(key);
  manuscriptByKey.delete(key);
  declinedByKey.delete(key);
  selection.drop(key);
  startQueue.delete(key);
}

// Called when a user message arrives from a pane, BEFORE it's queued.
// Each pane has its OWN loop — independent of every other pane. If this
// pane's loop is already live, do nothing (its userMessageStream will
// consume the message). Otherwise start it, resuming this (host, cwd)
// conversation. Deferred via setImmediate so it lands after the message
// is queued and after any in-flight finally; the new loop then drains
// this pane's queue. No other pane's loop is ever touched.
// idle: no turn of this pane was running when the message came.
async function ensureLoopForMessage(key, host, { idle = true } = {}) {
  if (!key) return;
  const cwd = cwdForKey(key);
  const live = sessionFor(key);
  if (live && !live.settled) {
    // (staysOn: the chosen model declined twice in a row; the session is
    // left where the CLI put it.)
    if (!live.fellBackTo || !idle || staysOn(key)) return; // the live loop will consume the message
    // The CLI moved this session to another model after a refusal
    // (model_refusal_fallback) and keeps it there. The retry of that turn
    // stays; this turn starts on the model the user chose again. The old
    // session's wait for a message ends here, before the caller queues the
    // new one, so the restarted session is the one that reads it.
    console.log(
      `[agent] model restored to ${modelArgFor(key)} after the fallback to ${live.fellBackTo}`,
    );
    bridge.clearUserMessages(key, { keepQueued: true });
    restartSession(key, host, { reason: "model_restored", keepQueued: true }).catch((err) =>
      console.warn("[daemon] restart after a model fallback failed:", err.message),
    );
    return;
  }
  let resumeId = null;
  try {
    resumeId = await getSessionId(host, cwd);
  } catch {
    /* fresh session if lookup fails */
  }
  diag(`message → ensure loop key=${key} cwd=${cwd} resume=${resumeId ?? "(new)"}`);
  // replay:false — the pane already shows the chat (incl. the message
  // that just triggered this). An empty transcript_replay here (a
  // brand-new session has no .jsonl yet) would wipe the user's prompt.
  scheduleSessionStart(cwd, resumeId, key, host, "user message", { replay: false });
}

// ---------------------------------------------------------------------------
// Review tab (C1/C3). Lambdas reach `bridge` lazily: it is created below.
// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------
// Providers (P14): which one each pane uses, the keys of the key tier, the
// availability probe and the routing of judgments.
// ---------------------------------------------------------------------------
const selection = createProviderSelection();
const providerKeys = createKeyStore({ log: console.log });
const probeCache = createProbeCache({
  probe: () => probeAll({ claudeModels: () => modelCache.get(), keys: providerKeys }),
});
const oneshotRouter = createOneshotRouter({
  claude: makeOneshot(),
  codex: (req) => codexOneshot(req),
  antigravity: (req) => agyOneshot(req),
  http: (req) => chatOneshot(req),
  keys: providerKeys,
  // Waits for the pane's confirmed choice and throws when there is none:
  // the router then answers with that error and sends nothing anywhere.
  getChoice: (key, role) => selection.ready(key, role),
  // A choice of Claude that names no model: writing gets the best model of
  // the account, a judgment the cheap one.
  claudeDefault: (role) =>
    role === "writing" ? bestClaudeModel(modelCache.current()) : ONESHOT_MODEL,
  // What each request really ran on, for the line of the pane that says so.
  onUsed: (key, used) => bridge.sendToTaskpane({ type: "model_used", ...used }, key),
  log: console.log,
});

// Will the next turn of this pane start (or restart) its session?
const startsSession = (key) => {
  const s = sessionFor(key);
  return !s || s.settled || (!!s.fellBackTo && !staysOn(key));
};

const review = createReviewService({
  getContext: (key) => bridge.getContext(key),
  callTool: (name, args, key) => bridge.callTaskpaneTool(name, args, key),
  pushTurn: (key, host, text, turn) => {
    if (turnBusy.has(key)) {
      throw new Error(
        "The agent is still working on the previous message. Wait until it is Ready.",
      );
    }
    // A review turn is a turn of the pane's Agent SDK session: Claude is the
    // only provider offered for review turns (probe.mjs; turnGate below has
    // waited for the pane's confirmed choice and refuses any other).
    turnBusy.add(key);
    reviewTurnByKey.set(key, turn);
    // Same order as the user_message path in the bridge: schedule the loop,
    // then queue the message it will consume.
    ensureLoopForMessage(key, host).catch((e) =>
      console.warn("[review] could not start the agent loop:", e.message),
    );
    bridge.pushUserMessage(text, key);
  },
  oneshot: makeOneshot(),
  oneshotModel: ONESHOT_MODEL,
  // Judgments go to the provider the pane chose (default: Claude, haiku),
  // fixed for each job when it starts.
  // (options: the paragraph check asks for the role and model of the edits.)
  oneshotJob: (key, options) => oneshotRouter.pin(key, options),
  // A review turn starts only with a confirmed provider for review turns.
  turnGate: (key) => selection.ready(key, "turn"),
  getModel: (key) => modelArgFor(key),
  // The model the session of the pane runs on, for receipts.
  getRunModel: (key) =>
    sessionFor(key)?.fellBackTo ?? sessionFor(key)?.ranModel ?? modelArgFor(key),
  manuscriptSeen: (key) => (startsSession(key) ? null : (manuscriptByKey.get(key)?.hashes ?? null)),
  manuscriptGiven: (key, hashes) => manuscriptByKey.set(key, { hashes, fresh: startsSession(key) }),
  sendEvent: (key, obj) => bridge.sendToTaskpane(obj, key),
  // The answers of the response letter that a change of the manuscript made
  // out of date (same rule as the Response tab).
  afterRefresh: (rev, snap) => markStale(rev, snap),
});

// Response tab (C2): the response letter. Drafts and LLM checks go through
// the same router as the judgments of the review: the provider and model the
// pane chose, fixed for each job, and nothing when the pane confirmed none.
const response = createResponseService({
  getContext: (key) => bridge.getContext(key),
  oneshotJob: (key, options) => oneshotRouter.pin(key, options),
  sendEvent: (key, obj) => bridge.sendToTaskpane(obj, key),
});

// Tools tab: selection, journal rules, checks and copies of the manuscript.
const aux = createAuxService({
  getContext: (key) => bridge.getContext(key),
  oneshotFor: (key, options) => oneshotRouter.pin(key, options),
  sendEvent: (key, obj) => bridge.sendToTaskpane(obj, key),
});

// Clear per-turn state when a turn ends for any reason. outcome: how it
// ended ({ ok: true } or { ok: false, error }); a review turn's turn_done
// carries it to the task pane.
function endTurnState(key, outcome = { ok: true }) {
  turnBusy.delete(key);
  const turn = reviewTurnByKey.get(key);
  if (turn) {
    reviewTurnByKey.delete(key);
    review.turnEnded(key, turn, outcome);
  }
}

// ---------------------------------------------------------------------------
// WebSocket bridge.
// ---------------------------------------------------------------------------
const bridge = createBridge({
  port: WS_PORT,
  token: BRIDGE_TOKEN,
  allowedOrigins: [HTTP_ORIGIN],
  onHello: (key, host, doc) => onPaneConnect(key, host, doc),
  onUserMessage: (key, host) => {
    const idle = !turnBusy.has(key);
    turnBusy.add(key);
    return ensureLoopForMessage(key, host, { idle });
  },
  onClose: (key) => onPaneClose(key),
  // A tracked edit answered after its timeout or disconnect: record it.
  onLateToolResult: (e) => review.lateToolResult(e.key, e),
  extraHandlers: {
    ...review.handlers,
    ...response.handlers,
    ...aux.handlers,
    pick_path: async (msg, reply) => {
      try {
        const result = await pickPathFromMain({
          include_files: !!msg.include_files,
          default_path: msg.default_path || null,
          title: msg.title || null,
          button_label: msg.button_label || null,
        });
        reply({ type: "pick_path_result", request_id: msg.request_id, ...result });
      } catch (e) {
        reply({
          type: "pick_path_result",
          ok: false,
          error: e.message,
          request_id: msg.request_id,
        });
      }
    },
    set_cwd: async (msg, reply, key, host) => {
      try {
        let cwd;
        let explicitPick = false;
        if (msg.autodetect_from_doc) {
          const detected = await resolveWorkspaceRoot(msg.autodetect_from_doc);
          if (!detected)
            throw new Error("Could not auto-detect a workspace folder from that doc path");
          cwd = detected;
        } else if (msg.cwd) {
          cwd = toLocalPath(msg.cwd);
          if (!cwd) throw new Error(`Folder is not reachable from WSL: ${msg.cwd}`);
          explicitPick = true;
        } else {
          throw new Error("set_cwd requires `cwd` or `autodetect_from_doc`");
        }
        // Remember an explicit pick so a later reconnect doesn't re-derive
        // this pane's workspace from the doc folder; an autodetect switch
        // clears that pin (the doc folder is authoritative again).
        if (explicitPick) explicitWorkspaceKeys.add(key);
        else explicitWorkspaceKeys.delete(key);
        const resolved = await switchFolder(cwd, key, host);
        // Drop a CLAUDE.md marker on explicit user pick so the next open of
        // any doc in this folder auto-detects silently.
        let markerCreated = false;
        if (explicitPick) {
          try {
            markerCreated = await ensureWorkspaceMarker(resolved);
          } catch (e) {
            console.warn(`[daemon] could not create CLAUDE.md in ${resolved}: ${e.message}`);
          }
        }
        reply({
          type: "set_cwd_result",
          ok: true,
          cwd: resolved,
          marker_created: markerCreated,
          request_id: msg.request_id,
        });
      } catch (e) {
        reply({ type: "set_cwd_result", ok: false, error: e.message, request_id: msg.request_id });
      }
    },
    suggest_workspace: async (msg, reply) => {
      try {
        const suggestion = await suggestWorkspaceRoot(msg.doc_path || null);
        reply({
          type: "suggest_workspace_result",
          ok: true,
          suggestion,
          request_id: msg.request_id,
        });
      } catch (e) {
        reply({
          type: "suggest_workspace_result",
          ok: false,
          error: e.message,
          request_id: msg.request_id,
        });
      }
    },
    get_cwd_state: async (msg, reply, key) => {
      reply({
        type: "get_cwd_state_result",
        ok: true,
        // This pane's own workspace, resolvable even before its first
        // message (lazy start ⇒ no session yet).
        current_cwd: cwdForKey(key),
        request_id: msg.request_id,
      });
    },
    stop_agent: async (msg, reply, key) => {
      // User clicked Stop in this pane. Abort only THIS pane's loop; the
      // query() iterator's catch path sees AbortError and (because we
      // flag the session interrupted) emits turn_complete interrupted so
      // the taskpane flips to Ready and auto-restarts a resuming loop.
      // Every other pane's loop is untouched.
      const s = sessionFor(key);
      if (s) {
        s.interrupted = true;
        s.abortController.abort();
        reply({ type: "stop_agent_result", ok: true, request_id: msg.request_id });
      } else {
        reply({
          type: "stop_agent_result",
          ok: false,
          error: "No active agent turn",
          request_id: msg.request_id,
        });
      }
    },
    get_context: async (msg, reply, key) => {
      try {
        const cwd = cwdForKey(key);
        const entries = cwd ? await getContextEntries(cwd) : [];
        reply({ type: "get_context_result", ok: true, cwd, entries, request_id: msg.request_id });
      } catch (e) {
        reply({
          type: "get_context_result",
          ok: false,
          error: e.message,
          request_id: msg.request_id,
        });
      }
    },
    // ---- providers (P14). No reply, log line or error text holds a key. ----
    providers_list: async (msg, reply, key) => {
      const rows = await probeCache.get({ refresh: !!msg.refresh });
      console.log(
        `[providers] list available=${rows
          .filter((r) => r.available)
          .map((r) => r.id)
          .join(",")} turn_ready=${rows
          .filter((r) => r.turn_ready)
          .map((r) => r.id)
          .join(",")}`,
      );
      reply({
        type: "providers_list_result",
        ok: true,
        request_id: msg.request_id,
        ...providersListReply(rows, selection.get(key)),
      });
    },
    // The pane's stored choices, sent right after hello and before any
    // provider list: { turn, oneshot }, each { id, model } or null. Until
    // this is answered no judgment and no review turn of the pane goes
    // anywhere (selection.mjs). The probe is asked only for a stored choice.
    restore_providers: async (msg, reply, key) => {
      const pick = (c) =>
        c && typeof c === "object" && c.id ? { id: c.id, model: c.model } : null;
      const state = await selection.restore(
        key,
        { turn: pick(msg.turn), writing: pick(msg.writing), oneshot: pick(msg.oneshot) },
        () => probeCache.get(),
      );
      console.log(
        `[providers] restore turn=${state.state.turn} writing=${state.state.writing} oneshot=${state.state.oneshot}` +
          ` provider=${state.oneshot?.id ?? (state.state.oneshot === "ready" ? "claude" : "-")}`,
      );
      reply({
        type: "restore_providers_result",
        ok: true,
        request_id: msg.request_id,
        selection: state,
      });
    },
    provider_set_key: async (msg, reply, key) => {
      const done = (body) =>
        reply({ type: "provider_set_key_result", request_id: msg.request_id, ...body });
      try {
        await providerKeys.set(String(msg.id ?? ""), { key: msg.key, base_url: msg.base_url });
        const rows = await probeCache.get({ refresh: true });
        done({ ok: true, ...providersListReply(rows, selection.get(key)) });
      } catch (e) {
        done({ ok: false, error: e?.message ?? String(e) });
      }
    },
    set_provider: async (msg, reply, key) => {
      const done = (body) =>
        reply({ type: "set_provider_result", request_id: msg.request_id, ...body });
      try {
        const rows = await probeCache.get();
        const chosen = selection.set(key, msg.role, { id: msg.id, model: msg.model }, rows);
        console.log(
          `[providers] set role=${msg.role} id=${msg.id} model=${msg.model ?? "default"} (key bound)`,
        );
        done({ ok: true, selection: chosen });
      } catch (e) {
        done({ ok: false, error: e?.message ?? String(e), selection: selection.get(key) });
      }
    },
    models_list: async (msg, reply) => {
      const r = await modelCache.get({ refresh: !!msg.refresh });
      console.log(
        `[models] source=${r.source} count=${r.models.length}${r.error ? ` error=${r.error}` : ""}`,
      );
      reply({ type: "models_list_result", ok: true, request_id: msg.request_id, ...r });
    },
    set_model: async (msg, reply, key, host) => {
      const requested = String(msg.model || "").trim();
      const prev = modelByKey.get(key);
      // Recorded before the list is awaited, so a first message sent
      // meanwhile does not start on "sonnet"; modelArgFor checks it again.
      if (prev === undefined) modelByKey.set(key, requested);
      const after = await modelCache.get();
      // Read after the wait: a session a first message started meanwhile
      // has recorded the model it runs on (session.model).
      const chosen = resolveModelChoice(requested, after);
      const { model, restart } = modelChange({
        requested,
        after,
        // A session the CLI moved to another model runs on that one.
        // One left on that model because this choice declines the
        // manuscript counts as running on the choice: no restart.
        running: staysOnFallback(declinedByKey.get(key), chosen)
          ? chosen
          : (sessionFor(key)?.fellBackTo ?? sessionFor(key)?.model ?? null),
      });
      // Another model chosen: what the old one declined is no longer of use.
      if (prev !== undefined && prev !== model) declinedByKey.delete(key);
      modelByKey.set(key, model);
      // A pane that connects again has lost what it was told: the state goes
      // again, without a chat notice.
      const stays = staysOn(key);
      if (stays) {
        bridge.sendAssistantEvent(
          {
            ...refusalFallback({
              original_model: stays.original,
              fallback_model: stays.fallback,
              api_refusal_category: stays.category,
            }).event,
            stay: true,
            quiet: true,
          },
          key,
        );
      }
      reply({ type: "set_model_result", ok: true, model, request_id: msg.request_id });
      // Relaunch when the model a session of this pane runs on is not the
      // one just decided: a later choice that differs, or, on the first
      // connect, a session that a message started while the list was awaited
      // and whose model the list then replaced. On the first connect the
      // restart is not the user's: a first message no session has read yet
      // stays for the new one (keepQueued).
      if (restart) {
        console.log(`[daemon] model changed → ${model} (key bound); restarting loop`);
        restartSession(key, host, {
          reason: "model_changed",
          keepQueued: prev === undefined,
        }).catch((err) => console.warn("[daemon] restart after model change failed:", err.message));
      }
    },
    set_context: async (msg, reply, key, host) => {
      try {
        const cwd = cwdForKey(key);
        if (!cwd) throw new Error("No workspace selected");
        const { saved, errors } = await setContextEntries(cwd, msg.entries || []);
        reply({
          type: "set_context_result",
          ok: errors.length === 0,
          cwd,
          saved,
          errors,
          request_id: msg.request_id,
        });
        // Restart THIS pane's loop so its agent re-reads CLAUDE.md and
        // picks up the updated context block on the next turn. Every
        // other pane is unaffected.
        restartSession(key, host, { reason: "context_changed" }).catch((err) =>
          console.warn("[daemon] restart failed:", err.message),
        );
      } catch (e) {
        reply({
          type: "set_context_result",
          ok: false,
          error: e.message,
          request_id: msg.request_id,
        });
      }
    },
  },
});

// ---------------------------------------------------------------------------
// Office-bridge MCP server (in-process; forwards tool calls to the taskpane).
// Built per session so the registered tool family matches the connected
// host (see startSessionForFolder / the host re-narrow on hello).
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Load the user's MCP servers from ~/.claude.json — the same file the
// interactive `claude` CLI uses. The Agent SDK doesn't read this file by
// default (it reads `~/.claude/settings.json` instead), so without this step
// the user's configured servers (visio, etc.) would be invisible to the
// daemon.
// ---------------------------------------------------------------------------
async function loadUserMcpServers() {
  const configPath = join(homedir(), ".claude.json");
  try {
    const raw = await readFile(configPath, "utf8");
    const parsed = JSON.parse(raw);
    const servers = parsed?.mcpServers ?? {};
    return servers;
  } catch (err) {
    if (err.code === "ENOENT") return {};
    console.warn(`[daemon] Could not load MCP servers from ${configPath}:`, err.message);
    return {};
  }
}

// Hoisted ABOVE the top-level awaits below on purpose. Module evaluation
// suspends at the first top-level `await` (loadUserMcpServers /
// preflightHttpMcpServers). The bridge WS server is already listening by
// then, so a taskpane `hello` can arrive mid-suspension and drive
// startSessionForFolder before the rest of the module body runs. Anything
// that path touches must be initialized first, or it hits a TDZ
// ReferenceError. Keep this (and any other start-path consts) up here.
const WORD_MCP_DISALLOWED = [
  "mcp__word-mcp__word_accept_revisions",
  "mcp__word-mcp__word_add_comment",
  "mcp__word-mcp__word_apply_style",
  "mcp__word-mcp__word_begin_transaction",
  "mcp__word-mcp__word_commit_transaction",
  "mcp__word-mcp__word_delete_comment",
  "mcp__word-mcp__word_delete_paragraphs",
  "mcp__word-mcp__word_delete_snapshot",
  "mcp__word-mcp__word_diff_snapshots",
  "mcp__word-mcp__word_emergency_recover",
  "mcp__word-mcp__word_export_pdf",
  "mcp__word-mcp__word_find_text",
  "mcp__word-mcp__word_get_document_info",
  "mcp__word-mcp__word_get_outline",
  "mcp__word-mcp__word_get_paragraph",
  "mcp__word-mcp__word_get_paragraphs",
  "mcp__word-mcp__word_get_section",
  "mcp__word-mcp__word_get_selection",
  "mcp__word-mcp__word_get_styles",
  "mcp__word-mcp__word_insert_paragraphs",
  "mcp__word-mcp__word_list_comments",
  "mcp__word-mcp__word_list_open_documents",
  "mcp__word-mcp__word_list_revisions",
  "mcp__word-mcp__word_list_transactions",
  "mcp__word-mcp__word_prune_snapshots",
  "mcp__word-mcp__word_reject_revisions",
  "mcp__word-mcp__word_replace_paragraphs",
  "mcp__word-mcp__word_replace_range",
  "mcp__word-mcp__word_replace_section",
  "mcp__word-mcp__word_replace_text",
  "mcp__word-mcp__word_restore_snapshot",
  "mcp__word-mcp__word_rollback_transaction",
  "mcp__word-mcp__word_toggle_track_changes",
  "mcp__word-mcp__word_undo_last_edit",
];

const userMcpServers = userMcpEnabled() ? await loadUserMcpServers() : {};
const userMcpNames = Object.keys(userMcpServers);
if (userMcpNames.length > 0) {
  console.log(
    `[daemon] Loaded ${userMcpNames.length} MCP server(s) from ~/.claude.json: ${userMcpNames.join(", ")}`,
  );
}
console.log(
  `[daemon] settingSources: ${JSON.stringify(resolveSettingSources() ?? "all")}; ` +
    `user MCP from ~/.claude.json: ${userMcpEnabled() ? "on" : "off (CC_OFFICE_LOAD_USER_MCP=1 to load)"}`,
);

// Preflight HTTP MCP servers. The SDK will silently drop any server whose
// initial handshake fails, with no retry for the lifetime of the session
// (see memory: sdk-silently-drops-failed-mcp). We can't fix that here, but
// we can surface the failure in the daemon log so it's obvious why a tool
// is missing — "restart the daemon when the server is back up" instead of
// "no idea why Visio doesn't work".
async function preflightHttpMcpServers(servers) {
  const entries = Object.entries(servers).filter(([, cfg]) => cfg?.type === "http" && cfg.url);
  await Promise.all(
    entries.map(async ([name, cfg]) => {
      const body = JSON.stringify({
        jsonrpc: "2.0",
        id: 1,
        method: "initialize",
        params: {
          protocolVersion: "2024-11-05",
          capabilities: {},
          clientInfo: { name: "claude-code-office-preflight", version: "0.1" },
        },
      });
      try {
        const res = await fetch(cfg.url, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Accept: "application/json, text/event-stream",
          },
          body,
          signal: AbortSignal.timeout(5000),
        });
        if (res.ok) {
          console.log(`[daemon] MCP preflight: ${name} reachable at ${cfg.url}`);
        } else {
          console.warn(
            `[daemon] MCP preflight: ${name} returned HTTP ${res.status} — tools may be missing until daemon restart`,
          );
        }
      } catch (err) {
        const reason =
          err.name === "TimeoutError" ? "timeout after 5s" : err.cause?.code || err.message;
        console.warn(
          `[daemon] MCP preflight: ${name} unreachable (${reason}) — tools will be missing until daemon restart`,
        );
      }
    }),
  );
}
await preflightHttpMcpServers(userMcpServers);

// ---------------------------------------------------------------------------
// System prompt: Claude Code default + Office-specific append.
// ---------------------------------------------------------------------------
// Shared base + the active host's section only. An Excel session never
// carries Word's surgical-edit / paragraph-reference rules (pure noise
// for it) and vice-versa. Re-read fresh on every session start so edits
// take effect when a session restarts (no daemon restart required).
async function buildSystemPromptAppend(host) {
  const base = await readFile(join(__dirname, "system-prompt.md"), "utf8");
  const files = host === "excel" ? ["system-prompt-excel.md"] : ["system-prompt-word.md"];
  // Degraded pre-bind case (host null) shouldn't happen now that sessions
  // start per pane, but fall back to including both to stay safe.
  if (host !== "word" && host !== "excel") {
    files.length = 0;
    files.push("system-prompt-word.md", "system-prompt-excel.md");
  }
  const parts = [base];
  for (const f of files) parts.push(await readFile(join(__dirname, f), "utf8"));
  return parts.join("\n");
}

// ---------------------------------------------------------------------------
// Disallow the out-of-process word-mcp tools. The user has word-mcp configured
// in ~/.claude (it surfaces as `mcp__word-mcp__*`), but those tools drive Word
// via AppleScript/COM and cause screen flicker on every edit — disqualifying
// for the live-edit path. We replace them with the in-process office_* tools
// that go through Office.js. The SDK's disallowedTools doesn't support globs,
// so we have to enumerate.
// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------
// Permission handler. Hard rule: filesystem tools must never WRITE to .docx
// files. The .docx the user has open in Word is held with unsaved changes; a
// filesystem write clobbers their work and can corrupt the file (Word holds a
// lock, OOXML cross-file references can break). Reading is allowed — the agent
// occasionally falls back to unzipping .docx XML to extract text, which is
// safe and read-only.
// ---------------------------------------------------------------------------
// Path/extension test for the Office-managed file types.
const OFFICE_FILE_EXT = /\.(docx?|xlsx?|docm|xlsm)\b/i;

// Bash is allowed to *read* Office files (the agent legitimately does
// `unzip -p draft.docx word/document.xml`, `cat`, `git log -- report.docx`,
// `ls *.docx`). We only deny commands that mutate one in place: a shell
// redirection whose target is an Office path, or a destructive verb used
// against one. This is an accident guard, not a security boundary — it
// stops the obvious foot-gun (`cp draft.docx active.docx`,
// `rm -rf workspace/*.docx`, `echo x > active.xlsx`, `sed -i … report.docx`),
// not a determined evasion (base64 payloads, here-docs, etc.).
const REDIR_TO_OFFICE = />>?\s*['"]?[^'"|;&\s]*\.(docx?|xlsx?|docm|xlsm)\b/i;
const DESTRUCTIVE_VERB = /\b(rm|mv|cp|dd|shred|truncate|install|tee|ln)\b/i;
const SED_IN_PLACE = /\bsed\b[^|;&]*\s-i\b/i;

function bashMutatesOfficeFile(cmd) {
  if (REDIR_TO_OFFICE.test(cmd)) return true;
  if (!OFFICE_FILE_EXT.test(cmd)) return false;
  return DESTRUCTIVE_VERB.test(cmd) || SED_IN_PLACE.test(cmd);
}

function denyWithOfficeMessage() {
  return {
    behavior: "deny",
    message:
      "Refusing to write/move/delete a Word/Excel file via filesystem tools. These files are " +
      "managed by Office and may have unsaved changes; a filesystem mutation can corrupt the " +
      "active document. Use the host's editing tools (office_* for Word, excel_* for Excel) " +
      "to change document contents.",
  };
}

function customPermissionHandler(toolName, input) {
  if (toolName === "Write" || toolName === "Edit" || toolName === "MultiEdit") {
    const path = input?.file_path ?? input?.path;
    if (typeof path === "string" && OFFICE_FILE_EXT.test(path)) {
      return Promise.resolve(denyWithOfficeMessage());
    }
  }
  if (toolName === "Bash") {
    const cmd = input?.command;
    if (typeof cmd === "string" && bashMutatesOfficeFile(cmd)) {
      return Promise.resolve(denyWithOfficeMessage());
    }
  }
  // Everything else: auto-approve.
  return Promise.resolve({ behavior: "allow", updatedInput: input ?? {} });
}

// ---------------------------------------------------------------------------
// Async iterable that pulls user messages from the bridge and yields them
// to the Agent SDK in the SDKUserMessage shape.
//
// We also prepend a context header to each turn so the agent always knows the
// active doc and selection without having to call office_get_doc_info first.
// ---------------------------------------------------------------------------
async function* userMessageStream(key, session) {
  while (true) {
    let msg;
    try {
      msg = await bridge.nextUserMessage(key);
    } catch {
      // Bridge rejected the waiter — session was aborted. Exit cleanly so
      // the underlying query() iterator can shut down without a stray error.
      return;
    }
    const { text, context } = msg;
    // The Agent SDK only treats a turn as a slash command (built-in or a
    // custom .claude/commands/*.md) when the message *starts with* "/".
    // Our per-turn context header (Host:/Doc:/Selection:) would otherwise
    // push the "/" off the front and the command would be sent to the
    // model as prose. So for a slash command, send it bare (leading
    // whitespace trimmed so detection works) and skip the header — the
    // command template is self-contained; it can call office_* tools if
    // it needs doc context.
    const trimmed = typeof text === "string" ? text.trimStart() : text;
    const isSlashCommand = typeof trimmed === "string" && trimmed.startsWith("/");
    const header = renderContextHeader(context);
    const content = isSlashCommand ? trimmed : header ? `${header}\n\n${text}` : text;
    // Per-turn tracking so a slash command that produces no assistant text
    // or tool call (terminal-only built-ins like /help, /context, /clear)
    // doesn't look like a dead chat. Reset at the start of every turn.
    if (session) {
      session.slashCommandPending = isSlashCommand;
      session.turnProducedOutput = false;
    }
    yield {
      type: "user",
      message: { role: "user", content },
      parent_tool_use_id: null,
    };
  }
}

function renderContextHeader(ctx) {
  const parts = [];
  // Host first so the agent immediately knows which tool family to use.
  // Both office_* (Word) and excel_* tools are registered simultaneously;
  // without this hint, the agent can pick the wrong family.
  if (ctx.host === "word" || ctx.host === "excel") {
    parts.push(`Host: ${ctx.host === "word" ? "Word" : "Excel"}`);
  }
  if (ctx.activeDoc) parts.push(`Doc: ${ctx.activeDoc}`);
  if (ctx.selection) {
    const s = ctx.selection;
    if (s.text) {
      const preview = s.text.length > 80 ? s.text.slice(0, 77) + "..." : s.text;
      parts.push(`Selection: "${preview}"`);
    } else if (s.para_id) {
      parts.push(`Cursor in paragraph ${s.para_id}`);
    }
  }
  // Only surface track-changes mode when it's NOT the default ("always"). The
  // system prompt says "always" by default; only deviation needs signaling.
  if (ctx.trackChangesMode && ctx.trackChangesMode !== "always") {
    parts.push(`Track changes: ${ctx.trackChangesMode}`);
  }
  return parts.length ? `[${parts.join(" · ")}]` : "";
}

// ---------------------------------------------------------------------------
// Session management. Each open document (paneKey) runs its own
// independent query() loop with its own cwd; switching workspace or
// reloading config restarts only that pane's loop. Histories live in
// ~/.claude/projects/<hash>/*.jsonl per the SDK's normal persistence.
// `sessions` / `workspaceByKey` are declared at the top of the module
// (above createBridge) so bridge handlers that fire during top-level
// awaits (e.g. preflightHttpMcpServers, which can block for 5s) don't
// hit a TDZ on those bindings.
// ---------------------------------------------------------------------------

async function startSessionForFolder(
  cwd,
  resumeSessionId = null,
  { key = null, host = null, replay = true, keepQueued = false } = {},
) {
  // Only when actually superseding a live session for THIS pane: abort it
  // and drain its queue (same-pane restart: workspace switch, config
  // reload, post-Stop/stream-end resume). No other pane's loop is ever
  // involved. On a FRESH start there is no prior session and the first
  // user message has already been enqueued for this pane (it's what
  // triggered the lazy start) — clearing here would drop it, leaving
  // userMessageStream awaiting forever and the SDK with no first input
  // (no init, taskpane stuck on "Working…").
  // The manuscript a review turn gave: this start keeps it only when that
  // turn is what starts the session.
  const given = manuscriptByKey.get(key);
  if (given?.fresh) given.fresh = false;
  else manuscriptByKey.delete(key);
  const prior = sessionFor(key);
  if (prior) {
    prior.abortController.abort();
    // keepQueued (the model correction of the first connect): a message the
    // prior session has not read stays in the queue, and its turn with it.
    const kept = bridge.clearUserMessages(key, { keepQueued });
    if (!kept) endTurnState(key, { ok: false, error: "The session was restarted." });
  }

  const abortController = new AbortController();
  const session = {
    key,
    cwd,
    sessionId: resumeSessionId,
    abortController,
    settled: false,
    host,
  };
  sessions.set(key, session);
  workspaceByKey.set(key, cwd);

  // Register only this host's tool family, routed to THIS pane. A session
  // is created lazily on the first user message (onUserMessage →
  // ensureLoopForMessage), so `host` is normally "word" or "excel".
  const officeMcp = createOfficeBridgeMcp(bridge, host, key, {
    getReviewTurn: () => reviewTurnByKey.get(key) ?? null,
    trackedEdit: review.trackedEditDeps(key),
  });

  await touchFolder(cwd);
  console.log(
    `[daemon] Starting session for ${cwd}` +
      (resumeSessionId ? ` (resuming ${resumeSessionId.slice(0, 8)}…)` : " (new session)"),
  );
  bridge.sendAssistantEvent({ event: "cwd_changed", cwd, resumed: !!resumeSessionId }, key);
  // Structured readiness signal to the Electron shell over the IPC
  // channel — the session loop is up. Lets main.mjs flip the tray to
  // "Ready" without sniffing our stdout for a log substring. No-op when
  // run via `npm run dev` (no IPC channel).
  if (process.send) {
    try {
      process.send({ type: "daemon_ready", cwd });
    } catch {
      /* channel gone */
    }
  }
  // Replay this host's transcript for the new workspace so its panel
  // reflects the workspace you just switched to (not the previous chat).
  // Skipped when this start was triggered by the user's own message
  // (replay:false): the pane already shows that message, and a fresh
  // session's empty replay would erase it.
  if (replay) sendTranscriptReplayTo(key, host, cwd).catch(() => {});

  // Re-read the drafting setup append fresh each session start.
  const append = await buildSystemPromptAppend(host);

  // Did this turn see a proper `result` message before the stream ended?
  // The SDK ends the stream with a `result` on normal completion. On a
  // usage-limit / quota hit (and some transport failures) the stream just
  // ends with no result and no thrown error — leaving the taskpane pinned
  // to "Working…". We track this to recover.
  let sawResult = false;
  // Best-effort usage-limit detection from the SDK CLI's stderr. The exact
  // phrasing varies by SDK version and limit kind (per-minute / daily /
  // weekly); match broadly.
  let rateLimitHint = null;
  const RATE_LIMIT_RE =
    /(usage limit|rate limit|daily limit|weekly limit|quota|too many requests|429|limit reached|limit will reset|resets? at|upgrade to|out of (?:credits|quota))/i;

  // The model this session runs on, recorded where query() is given it:
  // set_model compares it with the model it decides.
  session.model = modelArgFor(key);
  // Fire-and-forget; index.mjs keeps running while the agent loop iterates.
  (async () => {
    try {
      for await (const msg of query({
        prompt: userMessageStream(key, session),
        options: buildQueryOptions({
          cwd,
          append,
          officeMcp,
          userMcpServers,
          disallowedTools: WORD_MCP_DISALLOWED,
          canUseTool: customPermissionHandler,
          // User-chosen model (composer dropdown); always explicit.
          model: session.model,
          abortController,
          // Surface the SDK CLI's stderr in the daemon log and sniff it for
          // usage-limit phrasing (see RATE_LIMIT_RE above).
          stderr: (data) => {
            for (const line of String(data).split(/\r?\n/)) {
              if (!line.trim()) continue;
              console.error(`[sdk] ${line}`);
              if (!rateLimitHint && RATE_LIMIT_RE.test(line)) rateLimitHint = line.trim();
            }
          },
          resume: resumeSessionId,
        }),
      })) {
        if (sessionFor(key) !== session) break; // this pane's loop was restarted
        if (msg.type === "result") sawResult = true;
        const limit = detectUsageLimit(msg);
        if (limit.limited && !session.usageLimit) {
          session.usageLimit = limit;
          console.log(`[agent] ${limit.kind} detected (${limit.reason})`);
        }
        noteMessage((session.turnStats ??= emptyTurnStats()), msg);
        handleAgentMessage(msg, session);
      }
      // A turn that produced a `result` means the loop is healthy again —
      // clear any accumulated failure-restart count for this pane.
      if (sawResult) noteSuccessfulTurn(key);
      // Loop ended normally. If we never saw a `result`, the session is
      // still the live one, and the user didn't Stop, the stream died
      // unexpectedly — almost always a usage-limit hit. Tell the taskpane
      // (so it leaves "Working…") and auto-restart the loop so the next
      // message has a live consumer (same rationale as the Stop path).
      if (sessionFor(key) === session && !sawResult && !session.interrupted) {
        const friendly = rateLimitHint
          ? `Claude usage limit reached. ${rateLimitHint}`
          : "The agent stopped unexpectedly — this is usually a Claude usage limit. Wait for your limit to reset, or set ANTHROPIC_API_KEY to use an API key.";
        bridge.sendAssistantEvent({ event: "turn_complete", subtype: "stream_ended" }, key);
        endTurnState(key, { ok: false, error: friendly });
        const { key: rkey, cwd: rcwd, sessionId: rsid, host: rhost } = session;
        if (allowFailureRestart(key)) {
          bridge.sendAssistantEvent({ event: "error", error: friendly }, key);
          scheduleSessionStart(rcwd, rsid, rkey, rhost, "post-stream-end");
        } else {
          // Cap hit: the loop has died immediately too many times in a row.
          // Stop auto-resurrecting it (the respawn was achieving nothing but
          // error spam) and tell the user how to actually recover.
          console.error(
            `[daemon] post-stream-end restart cap hit for key=${key} — pausing auto-restart`,
          );
          bridge.sendAssistantEvent(
            {
              event: "error",
              error:
                "The agent keeps stopping immediately. This is almost always a Claude usage limit or an auth problem. Wait for your limit to reset (or set ANTHROPIC_API_KEY), then send a new message to retry.",
            },
            key,
          );
        }
      }
    } catch (err) {
      if (err.name === "AbortError" || /aborted/i.test(err.message ?? "")) {
        // Expected on a same-host restart OR when the user clicked Stop.
        // In the stop case (session.interrupted = true) we still need to
        // flip the taskpane's agent status back to Ready (otherwise it
        // stays pinned to "Working…") AND auto-restart a fresh resuming
        // loop — without a live query() iterator awaiting
        // bridge.nextUserMessage(), the user's next message would enqueue
        // with nobody to consume it. Let the finally block clear this
        // host's session first (via setImmediate); the fresh
        // startSessionForFolder then builds cleanly.
        if (sessionFor(key) === session && session.interrupted) {
          bridge.sendAssistantEvent({ event: "turn_complete", interrupted: true }, key);
          endTurnState(key, { ok: false, error: "The turn was stopped." });
          const { key: rkey, cwd: rcwd, sessionId: rsid, host: rhost } = session;
          scheduleSessionStart(rcwd, rsid, rkey, rhost, "post-stop");
        }
      } else if (sessionFor(key) === session) {
        console.error("[daemon] Agent loop crashed:", err);
        // Detect auth failures and surface them as a distinct event so the
        // taskpane can show a recoverable banner ("sign in to Claude Code")
        // instead of just dumping the SDK's raw error. Matched generously:
        // SDK error messages have varied across versions.
        const msgText = String(err?.message ?? err);
        const isAuth =
          /\b(authentication|unauthorized|credential|api[- ]?key|sign[- ]?in|401)\b/i.test(
            msgText,
          ) || /OAUTH/i.test(msgText);
        if (isAuth) {
          bridge.sendAssistantEvent({ event: "auth_error", error: msgText }, key);
        } else {
          bridge.sendAssistantEvent({ event: "error", error: msgText }, key);
        }
        endTurnState(key, { ok: false, error: msgText });
      }
    } finally {
      session.settled = true;
      if (sessionFor(key) === session) sessions.delete(key);
    }
  })();

  return session;
}

async function switchFolder(rawCwd, key, host = null) {
  const cwd = resolve(rawCwd);
  // Validate the path is a directory.
  const s = await stat(cwd);
  if (!s.isDirectory()) throw new Error(`Not a directory: ${cwd}`);
  // Switch ONLY the requesting pane to the target folder; resume that
  // (host, cwd)'s prior conversation if one is on record. Every other
  // pane stays in its own workspace, untouched.
  const resumeId = host ? await getSessionId(host, cwd) : null;
  await startSessionForFolder(cwd, resumeId, { key, host });
  return cwd;
}

// Re-launch one pane's loop (same cwd, resuming via session_id) so that
// changes to CLAUDE.md, the drafting setup, or other config loaded at
// session-init take effect without losing conversation history. No-op if
// that pane has no live loop.
async function restartSession(key, host, { reason = "config_changed", keepQueued = false } = {}) {
  const s = sessionFor(key);
  if (!s) return;
  const { cwd, sessionId } = s;
  console.log(`[daemon] Restarting session for ${cwd} (reason: ${reason})`);
  bridge.sendAssistantEvent({ event: "config_reloaded", reason }, key);
  // Funnel through the serialized per-key queue (not a direct
  // startSessionForFolder) so a config/model restart coalesces with any
  // concurrent post-Stop / first-message start instead of racing it.
  scheduleSessionStart(cwd, sessionId, key, host, reason, { replay: false, keepQueued });
}

function handleAgentMessage(msg, session) {
  switch (msg.type) {
    case "system": {
      if (msg.subtype === "init") {
        // asked: the model query() was given; model: the one the CLI runs.
        // effort: what the CLI will send (the SDK's default; none is set here).
        console.log(
          `[agent] init session ${msg.session_id} (model: ${msg.model}) asked=${session?.model ?? "-"} effort=${msg.effort ?? "-"}`,
        );
        try {
          const tn = Array.isArray(msg.tools) ? msg.tools : Object.keys(msg.tools ?? {});
          diag(
            `init tools (${tn.length}):`,
            tn.filter((t) => /office|excel|mcp__/.test(String(t))).join(", ") ||
              "(no office/excel/mcp tools in init list!)",
          );
        } catch (e) {
          diag("init tools introspection failed:", e?.message);
        }
        // For the receipts of a review turn (getRunModel).
        if (session) session.ranModel = msg.model;
        bridge.sendAssistantEvent(
          {
            event: "session_init",
            session_id: msg.session_id,
            model: msg.model,
            asked: session?.model ?? null,
          },
          session?.key,
        );
        // Record this session_id for THIS (host, cwd) so the next time
        // this pane connects (or you switch back) it resumes here.
        if (session && msg.session_id && msg.session_id !== session.sessionId) {
          session.sessionId = msg.session_id;
          saveSessionId(session.host, session.cwd, msg.session_id).catch((err) =>
            console.warn("[daemon] Could not save session id:", err.message),
          );
        }
      } else if (msg.subtype === "compact_boundary") {
        // The history was summarized: the manuscript an earlier turn gave
        // may be gone from it, so the next review turn gives it again.
        console.log(`[agent] system/${msg.subtype}`);
        if (session?.key) manuscriptByKey.delete(session.key);
      } else if (msg.subtype === "model_refusal_fallback") {
        // The CLI ran a refused request again on another model. The pane is
        // told at once. For the session the swap lasts: it is recorded
        // (session.fellBackTo) next to the model that was asked for
        // (session.model), and set_model compares the user's choice with
        // it, so choosing the model again restarts the session on it.
        // A second refusal in a row of the same kind: the session is left
        // on the other model from then on (ensureLoopForMessage).
        const f = refusalFallback(msg);
        console.log(f.log);
        let stay = false;
        if (f.session && session) {
          session.fellBackTo = msg.fallback_model;
          const d = declineStep(declinedByKey.get(session.key), {
            asked: session.model,
            original: msg.original_model,
            fallback: msg.fallback_model,
            category: msg.api_refusal_category ?? null,
          });
          declinedByKey.set(session.key, d);
          stay = d.stay;
          if (d.stay) {
            console.log(
              `[agent] model stays on ${d.fallback}: ${d.original} declined ${d.count} turns in a row (category=${d.category ?? "-"})`,
            );
          }
        }
        bridge.sendAssistantEvent({ ...f.event, stay }, session?.key);
      } else {
        console.log(`[agent] system/${msg.subtype}`);
      }
      break;
    }
    case "stream_event": {
      // includePartialMessages stream events. Forward text_delta to the
      // taskpane as assistant_text. Other event types (content_block_start,
      // content_block_stop, message_start/stop) we currently ignore.
      const delta = msg.event?.delta;
      if (delta?.type === "text_delta" && typeof delta.text === "string" && delta.text.length > 0) {
        if (session) session.turnProducedOutput = true;
        bridge.sendAssistantText(delta.text, session?.key);
      }
      break;
    }
    case "assistant": {
      // The complete assistant message arrives after streaming. We skip
      // text blocks (already streamed as deltas) and only forward tool_use
      // announces — those are atomic and not streamed.
      const blocks = msg.message?.content ?? [];
      for (const block of blocks) {
        if (block.type === "tool_use") {
          diag("model called tool:", block.name);
          if (session) session.turnProducedOutput = true;
          bridge.sendAssistantEvent(
            {
              event: "tool_use_announce",
              tool: block.name,
              input: block.input,
            },
            session?.key,
          );
        }
      }
      break;
    }
    case "result": {
      // A usage limit, a refusal, an API error or an empty turn can arrive as
      // subtype "success"; turnCompleteEvents turns those into
      // usage_limit / refusal / api_error / empty + error.
      const events = turnCompleteEvents(msg, session?.usageLimit ?? null, session?.turnStats);
      if (session) {
        session.usageLimit = null;
        session.turnStats = null;
      }
      console.log(
        `[agent] turn complete (${events[0].subtype}) slash=${!!session?.slashCommandPending} output=${!!session?.turnProducedOutput}`,
      );
      console.log(formatUsageLine(msg.usage));
      // A slash command that emitted neither assistant text nor a tool call
      // (terminal-only built-ins: /help, /context, /clear, …) would
      // otherwise complete silently and look broken. Surface a note.
      if (session?.slashCommandPending && !session.turnProducedOutput) {
        console.log("[agent] slash command produced no output — sending info note");
        bridge.sendAssistantEvent(
          {
            event: "info",
            message:
              "Command ran, but produced no chat output. Some built-in commands (e.g. /help, /context, /clear) are terminal-only and don't return anything here — custom .claude/commands and content-producing commands do.",
          },
          session?.key,
        );
      }
      if (session) {
        session.slashCommandPending = false;
        session.turnProducedOutput = false;
      }
      for (const ev of events) bridge.sendAssistantEvent(ev, session?.key);
      // The chosen model answered this turn (the CLI did not move the
      // session): a refusal counted before is not one of two in a row.
      if (session?.key && !session.fellBackTo) declinedByKey.delete(session.key);
      // text: the answer of the turn, for what a review turn reports in its
      // last line (review-service.mjs turnEnded).
      if (session?.key) {
        endTurnState(session.key, {
          ...turnOutcome(events, msg),
          text: typeof msg.result === "string" ? msg.result : null,
        });
      }
      break;
    }
    case "user": {
      // tool_result messages — we don't forward them; the bridge handles them.
      break;
    }
    default:
      // Other event types (api_retry, hook events, etc.) — not surfaced.
      break;
  }
}

// ---------------------------------------------------------------------------
// Kick off.
//
// No eager session: an agent session needs a host, and the host is only
// known once a taskpane connects and says hello (→ ensureSessionFor
// ActivePane). Starting both-tool sessions pre-hello is exactly what
// caused the per-host tool/transcript churn. The bridge buffers any
// user_message until a session's loop consumes it, so nothing is lost.
//
// Tell the Electron shell we're up now (servers listening) so the tray
// flips to "Ready" without waiting for a pane — independent of, and
// idempotent with, the daemon_ready that startSessionForFolder also
// emits on the first real session.
// ---------------------------------------------------------------------------
// Fetched ahead of the first pane; fetchClaudeModels never rejects.
modelCache.get();
console.log(`[daemon] Ready; waiting for a taskpane. Default workspace: ${matterFolder}`);
if (process.send) {
  try {
    process.send({ type: "daemon_ready" });
  } catch {
    /* no IPC channel (npm run dev) */
  }
}

// Keep process alive even when nothing is happening.
function shutdown() {
  console.log("\n[daemon] Shutting down");
  process.exit(0);
}
process.on("SIGINT", shutdown);
process.on("SIGTERM", shutdown);
