// Antigravity CLI (agy) adapter (P14): argument lists, the output of print
// mode, and one-shot judgments.
//
// Shapes are those of agy 1.2.14, captured on 2026-10-01
// (tests/fixtures/agy_print_json.txt, agy_print_stream.jsonl):
//   --output-format json, one object:
//     {"conversation_id":"...","status":"SUCCESS","response":"...","usage":{...}}
//   --output-format stream-json, one object per line:
//     {"event":"init","conversation_id":"...","init":{"cwd":"...","tools":[...]}}
//     {"event":"step_update","step_update":{"step_index":1,"state":"DONE",
//        "step_type":"agent_response","text_delta":"..."}}
//     {"event":"result","result":{"status":"SUCCESS","response":"..."}}
// A tool step (tests/fixtures/agy_turn_denied.jsonl):
//     {"event":"step_update","step_update":{"state":"ACTIVE","step_type":"tool",
//        "tool_name":"call_mcp_tool","tool_info":{"name":"call_mcp_tool",
//        "parameters":{"Arguments":{},"ServerName":"office","ToolName":"..."}}}}
//   and the same with "state":"DONE". In that run agy denied the call: the
//   result carried "denied_actions":[{"action":"mcp","display_name":
//   "CallMcpTool"}] (init said "permission_mode":"request-review") and
//   nothing reached the daemon. A failed run was not captured: any status
//   other than SUCCESS is taken as a failure.
//
// No run carries --dangerously-skip-permissions or --mode accept-edits. agy
// denies the office tools without one of them, so it is a provider for
// judgments only and the task pane says why (probe.mjs TURN_BLOCKED).
//
// agy has no setting for a run that leaves no trace: every judgment is kept
// as a conversation in the user's Antigravity history.
//
// What a print-mode run can reach, measured on 2026-10-01 (agy 1.2.14,
// made-up text only; fixture agy_print_tool_probe.jsonl), with
// --disable-slash-commands --sandbox:
//   - run_command ran: `echo hi` came back with its output, and `ls -la`
//     too. Nothing asked for a permission and nothing was denied.
//   - read_url_content was denied (denied_actions: read_url) and the run
//     ended with an empty response.
//   - `--mode plan` changed nothing: "--mode plan has no effect while slash
//     command expansion is disabled", and the same commands ran.
//   - init lists 58 tools (browser_*, search_web, write_to_file, ...) and
//     agy has no flag that takes one away. A hook of the user ran too.
//   - view_file and search_web were not tried: the model used run_command.
//   - --json-schema <schema> worked for a judgment (PATH and HOME alone as
//     environment): the result carries "structured_output", the response
//     text has two more keys (toolAction, toolSummary) and the stream has a
//     step of type "finish". Not used here.
// So agy cannot be kept from running a command next to the text it is
// given. It is offered for nothing (probe.mjs says why). The code below
// still reads every run as a stream and fails one that shows a tool step,
// a denied action or another permission mode, for a caller that runs it
// anyway; that watch is an alarm after the fact, not a wall.

import { parseJsonReply } from "../oneshot.mjs";
import { runCli as defaultRunCli, emptyWorkDir, resolveCliBin } from "./cli-exec.mjs";

const LIMIT_RE = /(usage limit|rate limit|quota|too many requests|out of credits|\b429\b)/i;
// One argument of a process is at most 128 KiB on Linux; the prompt is one.
const MAX_PROMPT_BYTES = 120_000;

function printArgs({ prompt, model, timeoutS, format }) {
  return [
    "-p",
    String(prompt),
    "--output-format",
    format,
    ...(model ? ["--model", model] : []),
    "--print-timeout",
    `${timeoutS}s`,
    "--disable-slash-commands",
    "--sandbox",
  ];
}

export const agyOneshotArgs = ({ prompt, model, timeoutS }) =>
  printArgs({ prompt, model, timeoutS, format: "json" });

export const agyTurnArgs = ({ prompt, model, timeoutS }) =>
  printArgs({ prompt, model, timeoutS, format: "stream-json" });

// `agy mcp add` writes the user's Antigravity settings, so it is run only
// when the user asks for it in the task pane, never by a probe. Flags come
// before the name.
export function agyMcpAddArgs({ url, token }) {
  return [
    "mcp",
    "add",
    "--type",
    "http",
    "--header",
    `Authorization: Bearer ${token}`,
    "office",
    url,
  ];
}

function resultOf(o) {
  const text = [o?.response, o?.result, o?.text].find((v) => typeof v === "string");
  const status = typeof o?.status === "string" ? o.status : null;
  if (o?.is_error === true || (status && status !== "SUCCESS")) {
    const detail = String(o?.error ?? o?.message ?? text ?? "").trim();
    const error = status
      ? `agy ended with status ${status}${detail ? `: ${detail}` : ""}`
      : detail || "agy failed";
    return { ok: false, error };
  }
  if (text === undefined || !text.trim()) return { ok: false, error: "agy gave no answer" };
  return { ok: true, text: text.trim() };
}

// The stdout of --output-format json as { ok: true, text } or { ok: false, error }.
export function parseAgyResult(stdout) {
  const o = parseJsonReply(stdout);
  if (!o || typeof o !== "object") return { ok: false, error: "agy output was not JSON" };
  return resultOf(o);
}

// One line of --output-format stream-json as a list of events:
// { kind: "session", id } | { kind: "text", text } | { kind: "tool", name,
// input } | { kind: "tool_done", name } | { kind: "step", step_type, state }
// | { kind: "denied", actions } | { kind: "done", text } | { kind: "error",
// message, limited }. An unknown line is []. An MCP tool is named
// mcp__<server>__<tool>, as in a Claude or Codex turn; a tool of agy's own
// is agy_<name>.
export function parseAgyLine(line) {
  let o;
  try {
    o = JSON.parse(line);
  } catch {
    return [];
  }
  if (!o || typeof o !== "object") return [];
  if (o.event === "init") return [{ kind: "session", id: o.conversation_id ?? null }];
  if (o.event === "step_update") {
    const s = o.step_update ?? {};
    if (s.step_type === "user_input") return [];
    if (s.step_type === "agent_response") {
      return typeof s.text_delta === "string" && s.text_delta
        ? [{ kind: "text", text: s.text_delta }]
        : [];
    }
    if (s.step_type === "tool") {
      const p = s.tool_info?.parameters ?? {};
      const name =
        s.tool_name === "call_mcp_tool"
          ? `mcp__${p.ServerName}__${p.ToolName}`
          : `agy_${s.tool_name ?? s.tool_info?.name ?? "tool"}`;
      if (s.state === "DONE") return [{ kind: "tool_done", name }];
      return [
        {
          kind: "tool",
          name,
          input: s.tool_name === "call_mcp_tool" ? (p.Arguments ?? {}) : p,
        },
      ];
    }
    return [{ kind: "step", step_type: String(s.step_type ?? ""), state: s.state ?? null }];
  }
  if (o.event === "result") {
    const res = o.result ?? {};
    const r = resultOf(res);
    const denied = (Array.isArray(res.denied_actions) ? res.denied_actions : []).map((d) =>
      String(d?.display_name ?? d?.action ?? "action"),
    );
    return [
      ...(denied.length ? [{ kind: "denied", actions: denied }] : []),
      r.ok
        ? { kind: "done", text: r.text }
        : { kind: "error", message: r.error, limited: LIMIT_RE.test(r.error) },
    ];
  }
  return [];
}

// The contract of oneshot.mjs runOneshot: { ok, text, json?, error?,
// limited? }; never throws. System and prompt go as one argument (agy has no
// separate system prompt and print mode takes the prompt as an argument).
export async function agyOneshot(
  { system, prompt, model, expectJson = true, timeoutMs = 120_000 },
  {
    runCli = defaultRunCli,
    bin = resolveCliBin("agy", { envVar: "CC_OFFICE_AGY_BIN" }),
    emptyDir = emptyWorkDir,
  } = {},
) {
  const full = `${system}\n\n${prompt}`;
  const bytes = Buffer.byteLength(full, "utf8");
  if (bytes > MAX_PROMPT_BYTES) {
    return {
      ok: false,
      error: `The request is too long for agy (${bytes} bytes in one argument).`,
    };
  }
  const timeoutS = Math.ceil(timeoutMs / 1000);
  let r;
  try {
    r = await runCli(bin, agyOneshotArgs({ prompt: full, model, timeoutS }), {
      cwd: await emptyDir(),
      // agy stops itself at --print-timeout; the kill is a backstop.
      timeoutMs: timeoutMs + 10_000,
    });
  } catch (e) {
    return { ok: false, error: `agy could not be run: ${String(e?.message ?? e)}` };
  }
  if (r.timedOut) return { ok: false, error: `timed out after ${timeoutMs} ms` };
  if (r.code !== 0) {
    const detail = String(r.stderr ?? "").slice(0, 300);
    return detail === "not installed"
      ? { ok: false, error: "Antigravity CLI (agy) is not installed" }
      : { ok: false, error: `agy exited ${r.code}: ${detail}` };
  }
  const res = parseAgyResult(r.stdout);
  if (!res.ok) {
    return {
      ok: false,
      error: res.error,
      limited: LIMIT_RE.test(res.error)
        ? { limited: true, kind: "rate_limit", text: res.error }
        : null,
    };
  }
  if (!expectJson) return { ok: true, text: res.text };
  const json = parseJsonReply(res.text);
  return json === undefined
    ? { ok: false, error: "reply was not JSON", text: res.text }
    : { ok: true, text: res.text, json };
}
