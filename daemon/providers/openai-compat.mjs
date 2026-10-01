// One-shot judgments over an OpenAI-compatible /chat/completions endpoint
// (Ollama and the key-tier providers, P14). The same contract as
// oneshot.mjs runOneshot: { ok, text, json?, error?, limited? }, and it
// never throws. No tools, no stream, temperature 0.
//
// The API key goes into the Authorization header and nowhere else; a server
// that echoes it in an error message has it cut out of the returned error.

import { parseJsonReply } from "../oneshot.mjs";

// A reasoning model may print its thinking in the answer (qwen on Ollama).
function stripThinking(content) {
  return String(content)
    .replace(/<think>[\s\S]*?<\/think>/g, "")
    .replace(/<think>[\s\S]*$/, "")
    .trim();
}

async function errorDetail(res) {
  try {
    const body = await res.json();
    const e = body?.error;
    return String((typeof e === "object" ? e?.message : e) ?? "").slice(0, 300);
  } catch {
    return "";
  }
}

export async function chatOneshot(
  { baseUrl, apiKey = null, model, system, prompt, expectJson = true, timeoutMs = 120_000 },
  { fetchFn = fetch } = {},
) {
  if (!model) return { ok: false, error: "Choose a model for this provider first." };
  const scrub = (s) => (apiKey ? String(s).split(apiKey).join("[key]") : String(s));
  // A timer of its own, not AbortSignal.timeout: that one does not keep the
  // process alive, so a request nobody else waits on would never time out.
  const abort = new AbortController();
  const timer = setTimeout(
    () => abort.abort(Object.assign(new Error("timeout"), { name: "TimeoutError" })),
    timeoutMs,
  );
  let res;
  try {
    res = await fetchFn(`${String(baseUrl).replace(/\/+$/, "")}/chat/completions`, {
      method: "POST",
      headers: {
        "content-type": "application/json",
        ...(apiKey ? { authorization: `Bearer ${apiKey}` } : {}),
      },
      body: JSON.stringify({
        model,
        messages: [
          { role: "system", content: String(system ?? "") },
          { role: "user", content: String(prompt ?? "") },
        ],
        temperature: 0,
        stream: false,
      }),
      signal: abort.signal,
    });
  } catch (e) {
    if (e?.name === "TimeoutError" || e?.name === "AbortError") {
      return { ok: false, error: `timed out after ${timeoutMs} ms` };
    }
    return { ok: false, error: `Connection failed (${e?.cause?.code ?? e?.code ?? e?.name})` };
  } finally {
    clearTimeout(timer);
  }
  if (!res.ok) {
    const detail = scrub(await errorDetail(res));
    const error = `HTTP ${res.status}${detail ? `: ${detail}` : ""}`;
    return res.status === 429
      ? { ok: false, error, limited: { limited: true, kind: "rate_limit", text: detail } }
      : { ok: false, error };
  }
  let content;
  try {
    content = (await res.json())?.choices?.[0]?.message?.content;
  } catch {
    return { ok: false, error: "the reply was not JSON from the server" };
  }
  if (typeof content !== "string") return { ok: false, error: "the reply had no message" };
  const text = stripThinking(content);
  if (!expectJson) return { ok: true, text };
  const json = parseJsonReply(text);
  if (json === undefined) return { ok: false, error: "reply was not JSON", text };
  return { ok: true, text, json };
}
