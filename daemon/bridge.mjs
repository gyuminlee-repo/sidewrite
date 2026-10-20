import { WebSocketServer } from "ws";
import { randomUUID } from "node:crypto";

// 60s, not 30s. Some Office.js ops legitimately run long on big docs —
// office_clear_highlights / a wide office_replace_section over a large
// document, batched office_highlight, a whole-sheet excel_read_range.
// 30s was tripping on real-sized documents. (Per-tool timeouts would be
// finer-grained but aren't worth the complexity yet.)
const TOOL_TIMEOUT_MS = 60_000;

// A write that got no answer (timeout or disconnect) may still land in Word
// and answer later. For these tools the bridge keeps the call for a while
// and hands a late tool_result to onLateToolResult instead of dropping it,
// so the daemon can record the edit. Every other tool keeps the old drop.
const LATE_RESULT_TOOLS = new Set(["office_tracked_edit"]);
const LATE_RESULT_TTL_MS = 10 * 60_000;

function normalizeHost(h) {
  return h === "word" || h === "excel" ? h : null;
}

// Hooks run immediately (session startup must precede queueing a turn), but
// neither a synchronous throw nor an asynchronous rejection may escape a
// WebSocket event listener and terminate the daemon.
function runHook(name, hook, ...args) {
  if (!hook) return;
  const failed = (err) => console.warn(`[bridge] ${name} failed:`, err?.message ?? err);
  try {
    Promise.resolve(hook(...args)).catch(failed);
  } catch (err) {
    failed(err);
  }
}

// A pane is keyed by (host, document), not just host: a user can have two
// Word docs (or two workbooks) open at once, each with its own task pane.
// Keying by host alone made them fight over a single slot — the same
// connect/disconnect ping-pong the host split originally fixed, just one
// level down. The host stays embedded in the key so the tool family and
// per-host behavior are still derivable from it.
const PANE_SEP = "\u0000";
function makePaneKey(host, doc, paneId) {
  let d;
  if (doc && String(doc).trim()) {
    d = String(doc);
  } else if (paneId && String(paneId).trim()) {
    // Unsaved / cloud doc with no filesystem path: key by the taskpane's
    // stable per-instance id (sent in hello) so a reconnect REUSES this
    // pane's key instead of minting a fresh random one every ~1.5s while
    // the daemon is down — which leaked per-pane state unbounded.
    d = "anon:" + String(paneId);
  } else {
    d = "anon:" + randomUUID();
  }
  return host + PANE_SEP + d;
}
function hostOfKey(key) {
  return typeof key === "string" ? key.split(PANE_SEP, 1)[0] : null;
}

export function createBridge({
  port,
  extraHandlers = {},
  token,
  allowedOrigins = [],
  toolTimeoutMs = TOOL_TIMEOUT_MS,
  lateResultTtlMs = LATE_RESULT_TTL_MS,
  onLateToolResult,
  onHello,
  onUserMessage,
  onClose,
}) {
  if (!token) throw new Error("createBridge requires a token");
  const wss = new WebSocketServer({
    port,
    // Bind to loopback only. The `ws` library defaults to 0.0.0.0 when
    // given just `port`, which exposes the bridge to the local network.
    // This is a local-only IPC channel (taskpane <-> daemon on the same
    // machine), so it should never be reachable off-host regardless of
    // the token + origin gates below. Matches the HTTP server in
    // index.mjs, which already binds 127.0.0.1.
    host: "127.0.0.1",
    // First gate: only allow upgrades from known origins (the taskpane's
    // origin = our own HTTP server's). Browsers honor the Origin header on
    // WebSocket upgrades; rejecting unknown origins blocks malicious local
    // web pages from driving the agent.
    verifyClient: (info, cb) => {
      const origin = info.req.headers.origin || "";
      if (allowedOrigins.length === 0 || allowedOrigins.includes(origin)) {
        cb(true);
      } else {
        console.warn(`[bridge] Rejecting WS upgrade from origin: ${origin || "(none)"}`);
        cb(false, 403, "Forbidden origin");
      }
    },
  });

  // One pane per (host, document). A connection is "unbound" until its
  // hello names a host + active doc; a reconnect with the SAME key
  // replaces the prior pane (Office reloaded that document's panel), any
  // DIFFERENT key — different host OR a different document of the same
  // host — coexists.
  const panes = new Map(); // paneKey -> { ws, context }

  function emptyContext(host = null, activeDoc = null) {
    // caps: what the task pane said it can do, in its hello (taskpane.js
    // PANE_CAPS). None for a task pane from before caps, and for no pane.
    return { activeDoc, selection: null, trackChangesMode: "always", host, caps: [] };
  }

  function paneWs(key) {
    const p = panes.get(key);
    return p && p.ws.readyState === p.ws.OPEN ? p.ws : null;
  }

  // Tool calls awaiting their tool_result. Keyed by id; each remembers the
  // ws it was dispatched on so a stale pane's disconnect only rejects its
  // own in-flight calls.
  const pendingTools = new Map();

  // Calls of LATE_RESULT_TOOLS that were rejected without an answer, by id:
  // { name, args, key, code, timer }. The timer drops the entry after
  // lateResultTtlMs; a result for it is handed over once.
  const lateTools = new Map();
  function keepForLateResult(id, p, code) {
    if (!onLateToolResult || !LATE_RESULT_TOOLS.has(p.name)) return;
    const timer = setTimeout(() => lateTools.delete(id), lateResultTtlMs);
    timer.unref?.();
    lateTools.set(id, { name: p.name, args: p.args, key: p.key, code, timer });
  }
  function noAnswer(message, code, id) {
    return Object.assign(new Error(message), { code, toolCallId: id });
  }

  // Per-pane queue of pending user messages. Waiters are {resolve, reject}
  // so we can reject pending awaits when that pane's session is aborted
  // (otherwise the suspended generator from the aborted session would
  // consume the next user message, starving the new session).
  const queues = new Map(); // paneKey -> { queue: [], waiters: [] }
  function queueFor(key) {
    let q = queues.get(key);
    if (!q) {
      q = { queue: [], waiters: [] };
      queues.set(key, q);
    }
    return q;
  }

  function pushUserMessage(text, key) {
    const ctx = panes.get(key)?.context ?? emptyContext(hostOfKey(key));
    const payload = { text, context: { ...ctx, host: ctx.host ?? hostOfKey(key) } };
    const q = queueFor(key);
    if (q.waiters.length > 0) {
      q.waiters.shift().resolve(payload);
    } else {
      q.queue.push(payload);
    }
  }

  function nextUserMessage(key) {
    const q = queueFor(key);
    if (q.queue.length > 0) return Promise.resolve(q.queue.shift());
    return new Promise((resolve, reject) => q.waiters.push({ resolve, reject }));
  }

  // Reject all pending waiters and drop queued messages for ONE pane.
  // Called before that pane's session swaps so the new session starts
  // clean with no zombie consumer — and, crucially, without touching any
  // OTHER pane's queue (a message just enqueued for the pane we're
  // switching TO must survive).
  // keepQueued: the messages no session has read yet stay for the next
  // one (a restart the user did not ask for). Returns how many stayed.
  function clearUserMessages(key, { keepQueued = false } = {}) {
    const q = queues.get(key);
    if (!q) return 0;
    while (q.waiters.length > 0) {
      q.waiters.shift().reject(new Error("Session aborted"));
    }
    if (!keepQueued) q.queue.length = 0;
    return q.queue.length;
  }

  // Merge inbound doc/selection/track-changes fields into a pane's
  // context. `hello` resets doc+selection (a fresh taskpane connection —
  // absent fields mean null), while `context_update` patches (only fields
  // present on the message change). track-changes is "set only when
  // valid" in both modes. host is fixed at bind time, never re-derived.
  function mergeContext(ctx, msg, { reset = false } = {}) {
    if (reset) {
      ctx.activeDoc = msg.active_doc ?? null;
      ctx.selection = msg.selection ?? null;
      ctx.caps = Array.isArray(msg.caps) ? msg.caps.filter((c) => typeof c === "string") : [];
    } else {
      if (msg.active_doc !== undefined) ctx.activeDoc = msg.active_doc;
      if (msg.selection !== undefined) ctx.selection = msg.selection;
    }
    if (typeof msg.track_changes_mode === "string") {
      ctx.trackChangesMode = msg.track_changes_mode;
    }
  }

  function send(obj, key) {
    const ws = paneWs(key);
    if (!ws) {
      console.warn(`[bridge] No pane for ${key ?? "?"}; dropping`, obj.type);
      return;
    }
    ws.send(JSON.stringify(obj));
  }

  async function callTaskpaneTool(name, args, key = null) {
    const ws = paneWs(key);
    if (!ws) {
      throw new Error(`Cannot call tool ${name}: no taskpane for ${key ?? "?"}`);
    }
    const id = randomUUID();
    // Capture the WS this call is dispatched on. On close we only reject
    // pending calls belonging to that specific WS — so a stale pane
    // disconnecting after a fresh one is active doesn't kill live work.
    const ownerWs = ws;
    const promise = new Promise((resolve, reject) => {
      // code: the call got no answer, so the caller cannot tell whether a
      // write was applied (office_tracked_edit reports OutcomeUnknown).
      // toolCallId: the id a late tool_result will carry (onLateToolResult).
      const timer = setTimeout(() => {
        const p = pendingTools.get(id);
        pendingTools.delete(id);
        if (p) keepForLateResult(id, p, "ToolTimeout");
        reject(noAnswer(`Tool ${name} timed out after ${toolTimeoutMs}ms`, "ToolTimeout", id));
      }, toolTimeoutMs);
      pendingTools.set(id, { ws: ownerWs, resolve, reject, timer, name, args, key });
    });
    ownerWs.send(JSON.stringify({ type: "tool_call", id, name, args }));
    return await promise;
  }

  function sendAssistantText(delta, key) {
    send({ type: "assistant_text", delta }, key);
  }

  function sendAssistantEvent(event, key) {
    send({ type: "assistant_event", ...event }, key);
  }

  wss.on("connection", (ws, req) => {
    console.log(
      `[bridge] WS connection from ${req.socket.remoteAddress} (origin: ${req.headers.origin || "(none)"})`,
    );

    // Second gate: first message must be a hello with the correct token.
    // Until that arrives we don't trust the connection — any other message
    // type is rejected and the socket is closed. The taskpane fetches the
    // token from /bridge-token over the same-origin HTTP server before
    // opening the WS. A connection is bound to exactly one (host,
    // document) pane key for its lifetime.
    let authed = false;
    let boundKey = null;
    let boundHost = null;

    // True iff this ws is still the live pane for the key it bound to.
    // A superseded (same-key reconnect) ws stays open just long enough
    // to be force-closed; drop any stale frames it sends meanwhile.
    function isLivePane() {
      return authed && panes.get(boundKey)?.ws === ws;
    }

    ws.on("message", (raw) => {
      let msg;
      try {
        msg = JSON.parse(raw.toString());
      } catch (e) {
        console.error("[bridge] Bad JSON from taskpane:", e.message);
        return;
      }

      // Valid JSON is not necessarily a protocol message (notably null).
      // Reject it before accessing fields or dispatching to a handler.
      if (!msg || typeof msg !== "object" || Array.isArray(msg) || typeof msg.type !== "string") {
        console.warn("[bridge] Invalid message envelope; ignoring");
        if (!authed) ws.close(4001, "Authenticate first");
        return;
      }

      if (!authed) {
        if (msg.type !== "hello") {
          console.warn(`[bridge] Pre-auth message ${msg.type}; closing`);
          ws.close(4001, "Authenticate first");
          return;
        }
        if (msg.token !== token) {
          console.warn("[bridge] hello with invalid token; closing");
          ws.close(4002, "Invalid token");
          return;
        }
        const host = normalizeHost(msg.host);
        if (!host) {
          console.warn(`[bridge] hello without a valid host (${msg.host}); closing`);
          ws.close(4005, "hello must name host word|excel");
          return;
        }
        boundHost = host;
        boundKey = makePaneKey(host, msg.active_doc, msg.pane_id);
        // Replace only a prior pane with the SAME key (same document's
        // panel reloaded / reopened). A different document — even of the
        // same host — gets its own pane and coexists.
        const prior = panes.get(boundKey);
        if (prior && prior.ws !== ws && prior.ws.readyState === prior.ws.OPEN) {
          console.log(`[bridge] New pane authed for ${boundKey}; closing prior`);
          try {
            prior.ws.close(4003, "Replaced by new pane");
          } catch {}
        }
        authed = true;
        panes.set(boundKey, { ws, context: emptyContext(host, msg.active_doc ?? null) });
      } else if (!isLivePane()) {
        // Superseded by a newer same-key pane; its close is in flight.
        // Drop stale frames rather than letting them mutate context or
        // queue a user message for a pane that's going away.
        console.warn(`[bridge] Post-auth ${msg.type} from superseded pane ${boundKey}; ignoring`);
        return;
      }

      const pane = panes.get(boundKey);

      switch (msg.type) {
        case "hello": {
          mergeContext(pane.context, msg, { reset: true });
          ws.send(
            JSON.stringify({
              type: "welcome",
              session_id: randomUUID(),
              server_version: "1.0.0",
            }),
          );
          console.log(
            `[bridge] hello received; host: ${boundHost}; active doc: ${pane.context.activeDoc}`,
          );
          // Let the daemon replay this pane's transcript (and anything
          // else it wants on connect). Fire-and-forget; never block.
          runHook("onHello", onHello, boundKey, boundHost, pane.context.activeDoc);
          break;
        }
        case "user_message": {
          console.log(`[bridge] user_message (${boundHost}):`, msg.text);
          // Give the daemon a chance to (lazily) start/swap the agent
          // loop for this pane BEFORE the message is queued, so the right
          // loop consumes it. Fire-and-forget; ordering holds because
          // session start is deferred to setImmediate, after this push.
          runHook("onUserMessage", onUserMessage, boundKey, boundHost, pane.context.activeDoc);
          pushUserMessage(msg.text, boundKey);
          break;
        }
        case "context_update": {
          mergeContext(pane.context, msg);
          break;
        }
        case "tool_result": {
          const pending = pendingTools.get(msg.id);
          // An answer counts only from the pane the call was sent to (same
          // pane key, i.e. the same document): the pane of another document
          // can neither answer a call nor use up its late result.
          const sent = pending ?? lateTools.get(msg.id);
          if (sent && sent.key !== boundKey) {
            console.warn(`[bridge] tool_result ${msg.id} from another pane; ignoring`);
            return;
          }
          if (!pending) {
            // Looked up by id (and pane key): after a WS blip the task pane
            // sends the answer on its new socket.
            const late = lateTools.get(msg.id);
            if (!late) {
              console.warn("[bridge] tool_result for unknown id:", msg.id);
              return;
            }
            clearTimeout(late.timer);
            lateTools.delete(msg.id);
            console.warn(`[bridge] late tool_result for ${late.name} ${msg.id} (${late.code})`);
            Promise.resolve()
              .then(() =>
                onLateToolResult({
                  id: msg.id,
                  name: late.name,
                  args: late.args,
                  key: late.key,
                  code: late.code,
                  ok: !!msg.ok,
                  result: msg.result,
                  error: msg.error ?? null,
                }),
              )
              .catch((err) =>
                console.warn("[bridge] onLateToolResult failed:", err?.message ?? err),
              );
            return;
          }
          clearTimeout(pending.timer);
          pendingTools.delete(msg.id);
          if (msg.ok) {
            pending.resolve(msg.result);
          } else {
            pending.reject(new Error(msg.error ?? "Unknown tool error"));
          }
          break;
        }
        case "ping": {
          ws.send(JSON.stringify({ type: "pong" }));
          break;
        }
        default: {
          // Custom handlers registered by the daemon (e.g. settings /
          // refs). They receive the message, a `reply(obj)` shortcut that
          // sends back over this same WS, and the bound pane key + host so
          // per-pane actions (set_cwd, stop_agent, …) act on the right
          // document.
          const handler = extraHandlers[msg.type];
          if (handler) {
            const reply = (obj) => ws.send(JSON.stringify(obj));
            Promise.resolve()
              .then(() => handler(msg, reply, boundKey, boundHost))
              .catch((err) => {
                console.error(`[bridge] handler for ${msg.type} threw:`, err);
                reply({
                  type: msg.type + "_result",
                  ok: false,
                  error: err.message ?? String(err),
                  request_id: msg.request_id,
                });
              });
          } else {
            console.warn("[bridge] Unknown message type:", msg.type);
          }
        }
      }
    });

    ws.on("close", () => {
      console.log(`[bridge] Taskpane disconnected${boundKey ? ` (${boundHost})` : ""}`);
      // Only clear the pane mapping if THIS ws still owns it (a superseded
      // ws closing must not evict the fresh same-key pane).
      if (boundKey && panes.get(boundKey)?.ws === ws) {
        panes.delete(boundKey);
        // Let the daemon prune per-key state for a pane that's gone.
        // Deliberately NOT touching this pane's queue or its live
        // session here: a transient WS blip auto-reconnects with the
        // SAME key (stable now), and a suspended userMessageStream is
        // awaiting the existing queue — dropping it would resurface the
        // "stuck on Working…" bug. The daemon's onClose only prunes the
        // lightweight Maps for keys with no live session.
        runHook("onClose", onClose, boundKey, boundHost);
      }
      // Reject only the pending tool calls dispatched on THIS ws. A
      // different (newer) WS may have its own in-flight calls; leave them.
      for (const [id, p] of pendingTools) {
        if (p.ws === ws) {
          clearTimeout(p.timer);
          keepForLateResult(id, p, "TaskpaneDisconnected");
          p.reject(noAnswer("Taskpane disconnected", "TaskpaneDisconnected", id));
          pendingTools.delete(id);
        }
      }
    });

    ws.on("error", (err) => {
      console.error("[bridge] WS error:", err.message);
    });
  });

  console.log(
    `[bridge] WebSocket server listening on ws://127.0.0.1:${port} (origin allowlist: ${allowedOrigins.join(", ") || "<empty>"})`,
  );

  return {
    nextUserMessage,
    // Daemon-originated turns (Review tab "Propose changes"): queued exactly
    // like a task pane user_message, with the current context of this pane.
    pushUserMessage,
    clearUserMessages,
    callTaskpaneTool,
    sendAssistantText,
    sendAssistantEvent,
    sendToTaskpane: send,
    getContext: (key) => ({ ...(panes.get(key)?.context ?? emptyContext(hostOfKey(key))) }),
    // The panes that are connected now: [{ key, host, activeDoc }]. For the
    // figures page (figure-http.mjs), which may name only their documents.
    listPanes: () =>
      [...panes.keys()]
        .filter((key) => !!paneWs(key))
        .map((key) => {
          const { host, activeDoc } = panes.get(key).context;
          return { key, host: host ?? hostOfKey(key), activeDoc };
        }),
    isTaskpaneConnected: (key) =>
      key ? !!paneWs(key) : [...panes.keys()].some((k) => !!paneWs(k)),
    // Test/observability surface: the bound address (null until listening)
    // and a clean shutdown. Used by the loopback-bind unit test.
    address: () => wss.address(),
    close: () => {
      for (const late of lateTools.values()) clearTimeout(late.timer);
      lateTools.clear();
      return new Promise((resolve) => wss.close(resolve));
    },
  };
}
