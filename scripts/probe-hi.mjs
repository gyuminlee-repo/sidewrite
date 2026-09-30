#!/usr/bin/env node
// Word-free token probe for a running daemon. Connects to the WebSocket
// bridge the way the task pane does, sends one message ("hi") from a fresh
// throwaway workspace folder (a fresh folder never resumes an old session),
// waits for turn_complete, then reads that session's .jsonl and prints the
// first assistant call's input-side usage and the number of hook lines.
//
//   node scripts/probe-hi.mjs                  live probe, daemon must run
//   node scripts/probe-hi.mjs --jsonl <file>   summarize an existing session
//   --text <msg> --model haiku|sonnet|opus --timeout <sec>
//   --max-total <n> --max-hooks <n>            exit 1 when exceeded
//
// Document tool calls are answered with an error (no Word here).
// Side effect: the daemon lists the temp folder in
// ~/.claude/office-addins/sessions.json (recent workspaces).

import { readFile, readdir, mkdtemp } from "node:fs/promises";
import { homedir, tmpdir } from "node:os";
import { join } from "node:path";
import { randomUUID } from "node:crypto";
import { pathToFileURL } from "node:url";
import { WebSocket } from "ws";

export const WS_URL = "ws://127.0.0.1:47833";
export const HTTP_ORIGIN = "http://127.0.0.1:47834";
const TOKEN_FILE = join(homedir(), ".claude", "office-addins", "bridge-token");
const PROJECTS_DIR = join(homedir(), ".claude", "projects");

export function summarizeSessionJsonl(text) {
  let usage = null;
  let model = null;
  let version = null;
  let hookEvents = 0;
  for (const line of text.split("\n")) {
    if (!line.trim()) continue;
    let o;
    try {
      o = JSON.parse(line);
    } catch {
      continue;
    }
    if (line.includes('"hookEvent"')) hookEvents += 1;
    if (!version && typeof o.version === "string") version = o.version;
    if (!usage && o.type === "assistant" && o.message?.usage) {
      usage = o.message.usage;
      model = o.message.model ?? null;
    }
  }
  const u = usage ?? {};
  const input = u.input_tokens ?? 0;
  const cacheCreation = u.cache_creation_input_tokens ?? 0;
  const cacheRead = u.cache_read_input_tokens ?? 0;
  return {
    total: input + cacheCreation + cacheRead,
    input,
    cacheCreation,
    cacheRead,
    output: u.output_tokens ?? 0,
    hookEvents,
    version,
    model,
    hasAssistant: !!usage,
  };
}

export function runProbe({
  url = WS_URL,
  origin = HTTP_ORIGIN,
  token,
  doc,
  text = "hi",
  model = "sonnet",
  timeoutMs = 180_000,
}) {
  return new Promise((resolve, reject) => {
    const out = { sessionId: null, model: null, cwd: null, toolCalls: 0, text: "" };
    let sent = false;
    const ws = new WebSocket(url, { origin });
    const timer = setTimeout(() => finish(new Error(`timed out after ${timeoutMs} ms`)), timeoutMs);
    function finish(err) {
      clearTimeout(timer);
      try {
        ws.close();
      } catch {}
      if (err) reject(err);
      else resolve(out);
    }
    const send = (obj) => ws.send(JSON.stringify(obj));
    ws.on("open", () =>
      send({
        type: "hello",
        token,
        host: "word",
        active_doc: doc,
        pane_id: `probe-${randomUUID()}`,
        selection: null,
        track_changes_mode: "always",
      }),
    );
    ws.on("error", (e) => finish(e));
    ws.on("close", (code, reason) => finish(new Error(`socket closed (${code} ${reason})`)));
    ws.on("message", (raw) => {
      let msg;
      try {
        msg = JSON.parse(raw.toString());
      } catch {
        return;
      }
      if (msg.type === "welcome") {
        send({ type: "set_model", model, request_id: randomUUID() });
      } else if (msg.type === "transcript_replay" && !sent) {
        sent = true;
        send({ type: "user_message", text });
      } else if (msg.type === "tool_call") {
        out.toolCalls += 1;
        send({
          type: "tool_result",
          id: msg.id,
          ok: false,
          error: "probe-hi: document tools are not available in the probe",
        });
      } else if (msg.type === "assistant_text") {
        out.text += msg.delta ?? "";
      } else if (msg.type === "assistant_event") {
        if (msg.event === "session_init") {
          out.sessionId = msg.session_id;
          out.model = msg.model;
        } else if (msg.event === "cwd_changed") {
          out.cwd = msg.cwd;
        } else if (msg.event === "turn_complete" && sent) {
          if ((msg.subtype && msg.subtype !== "success") || msg.interrupted) {
            finish(new Error(`turn ended abnormally (${msg.subtype ?? "interrupted"})`));
          } else {
            finish(null);
          }
        } else if (msg.event === "error" || msg.event === "auth_error") {
          finish(new Error(`daemon ${msg.event}: ${msg.error}`));
        }
      }
    });
  });
}

async function readToken() {
  try {
    return (await readFile(TOKEN_FILE, "utf8")).trim();
  } catch {
    const res = await fetch(`${HTTP_ORIGIN}/bridge-token`);
    if (!res.ok) throw new Error(`cannot read bridge token (HTTP ${res.status})`);
    return (await res.text()).trim();
  }
}

async function readSessionFile(sessionId) {
  const dirs = await readdir(PROJECTS_DIR, { withFileTypes: true });
  for (const d of dirs) {
    if (!d.isDirectory()) continue;
    try {
      return await readFile(join(PROJECTS_DIR, d.name, `${sessionId}.jsonl`), "utf8");
    } catch {}
  }
  return null;
}

async function waitForSummary(sessionId, tries = 20) {
  for (let i = 0; i < tries; i++) {
    const text = await readSessionFile(sessionId);
    if (text) {
      const s = summarizeSessionJsonl(text);
      if (s.hasAssistant) return s;
    }
    await new Promise((r) => setTimeout(r, 500));
  }
  throw new Error(`no assistant usage found for session ${sessionId}`);
}

function parseArgs(argv) {
  const a = {
    text: "hi",
    model: "sonnet",
    timeout: 180,
    jsonl: null,
    maxTotal: null,
    maxHooks: null,
  };
  for (let i = 0; i < argv.length; i++) {
    const k = argv[i];
    const v = argv[i + 1];
    if (k === "--jsonl") ((a.jsonl = v), i++);
    else if (k === "--text") ((a.text = v), i++);
    else if (k === "--model") ((a.model = v), i++);
    else if (k === "--timeout") ((a.timeout = Number(v)), i++);
    else if (k === "--max-total") ((a.maxTotal = Number(v)), i++);
    else if (k === "--max-hooks") ((a.maxHooks = Number(v)), i++);
    else throw new Error(`unknown argument ${k}`);
  }
  return a;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  let meta = {};
  let summary;
  if (args.jsonl) {
    summary = summarizeSessionJsonl(await readFile(args.jsonl, "utf8"));
  } else {
    const dir = await mkdtemp(join(tmpdir(), "probe-hi-"));
    meta = await runProbe({
      token: await readToken(),
      doc: join(dir, "probe.docx"),
      text: args.text,
      model: args.model,
      timeoutMs: args.timeout * 1000,
    });
    if (!meta.sessionId) throw new Error("turn completed without session_init");
    summary = await waitForSummary(meta.sessionId);
  }
  const r = {
    ...summary,
    sessionId: meta.sessionId ?? null,
    cwd: meta.cwd ?? null,
    toolCalls: meta.toolCalls ?? 0,
  };
  console.log(
    `total=${r.total} input=${r.input} cache_creation=${r.cacheCreation} cache_read=${r.cacheRead}` +
      ` output=${r.output} hookEvents=${r.hookEvents} version=${r.version} model=${r.model}` +
      ` toolCalls=${r.toolCalls} session=${r.sessionId} cwd=${r.cwd}`,
  );
  console.log(JSON.stringify(r));
  let fail = false;
  if (args.maxTotal != null && r.total > args.maxTotal) {
    console.error(`[probe-hi] FAIL total ${r.total} > ${args.maxTotal}`);
    fail = true;
  }
  if (args.maxHooks != null && r.hookEvents > args.maxHooks) {
    console.error(`[probe-hi] FAIL hookEvents ${r.hookEvents} > ${args.maxHooks}`);
    fail = true;
  }
  process.exit(fail ? 1 : 0);
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((e) => {
    console.error(`[probe-hi] ${e.message}`);
    process.exit(2);
  });
}
