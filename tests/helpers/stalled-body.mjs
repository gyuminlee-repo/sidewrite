// A real HTTP server that sends its headers at once and its body late, and
// the requests of this add-in against it. Run as a process of its own by
// the tests (providers-openai-compat, providers-probe):
//   node tests/helpers/stalled-body.mjs chat | probe | temperature
// (temperature: servers that answer at once, HTTP 400 for a request with a
// temperature; real requests all the same, so they run here too.)
// It prints one JSON line and ends with process.exit: on Node 20.11 a
// process that aborted a fetch in the middle of a body and then ends by
// itself can spin at 100% CPU while it shuts down (seen on 2026-10-01, about
// one run in three under load), which would hang the test runner.
import { createServer } from "node:http";
import { chatOneshot } from "../../daemon/providers/openai-compat.mjs";
import { probeAll } from "../../daemon/providers/probe.mjs";

const scenario = process.argv[2];
const BODY_AFTER_MS = scenario === "probe" ? 8_000 : 400;
const body = JSON.stringify({ choices: [{ message: { content: '{"ok":true}' } }] });
const server = createServer((req, res) => {
  req.resume();
  res.writeHead(req.url.startsWith("/bad") ? 500 : 200, { "content-type": "application/json" });
  res.flushHeaders();
  const late = setTimeout(() => res.end(body), BODY_AFTER_MS);
  res.on("close", () => clearTimeout(late));
});
await new Promise((r) => server.listen(0, "127.0.0.1", r));
const base = `http://127.0.0.1:${server.address().port}`;

async function timed(fn) {
  const t0 = Date.now();
  const result = await fn();
  return { result, ms: Date.now() - t0 };
}

let out;
if (scenario === "chat") {
  const req = { apiKey: "k", model: "m", system: "", prompt: "" };
  out = {
    // 70 ms allowed, the body comes after 400 ms: good status and error status.
    good: await timed(() => chatOneshot({ ...req, baseUrl: `${base}/v1`, timeoutMs: 70 })),
    bad: await timed(() => chatOneshot({ ...req, baseUrl: `${base}/bad/v1`, timeoutMs: 70 })),
    // The same server within a limit that allows the body.
    patient: await timed(() => chatOneshot({ ...req, baseUrl: `${base}/v1`, timeoutMs: 5_000 })),
  };
} else if (scenario === "probe") {
  // Ollama's model list from the stalling server; every other provider is
  // answered at once without a network.
  const rows = await timed(() =>
    probeAll({
      claudeModels: async () => ({ source: "sdk", models: [] }),
      execFileFn: async () => ({ code: 1, stdout: "" }),
      readFileFn: async () => "{}",
      keys: { get: async () => ({ key: null }) },
      env: {},
      fetchFn: (url, init) =>
        url.endsWith("/api/tags")
          ? fetch(`${base}/api/tags`, init)
          : Promise.resolve({ ok: false, status: 401, json: async () => ({}) }),
    }),
  );
  out = { ms: rows.ms, ollama: rows.result.find((r) => r.id === "ollama") };
} else if (scenario === "temperature") {
  const NO_TEMPERATURE = {
    error: {
      message:
        "Unsupported value: 'temperature' does not support 0 with this model. Only the default (1) value is supported.",
      param: "temperature",
    },
  };
  const ANSWER = { choices: [{ message: { content: '{"a":1}' } }] };
  // answer(body) -> [status, json]; what each request carried is kept.
  const provider = async (answer) => {
    const bodies = [];
    const s = createServer((req, res) => {
      const chunks = [];
      req.on("data", (c) => chunks.push(c));
      req.on("end", () => {
        const sent = JSON.parse(Buffer.concat(chunks).toString("utf8"));
        bodies.push(sent);
        const [status, json] = answer(sent);
        res.writeHead(status, { "content-type": "application/json" });
        res.end(JSON.stringify(json));
      });
    });
    await new Promise((r) => s.listen(0, "127.0.0.1", r));
    return { bodies, baseUrl: `http://127.0.0.1:${s.address().port}/v1` };
  };
  const ask = async (answer) => {
    const p = await provider(answer);
    const result = await chatOneshot({
      baseUrl: p.baseUrl,
      model: "o-reasoner",
      system: "S",
      prompt: "P",
    });
    return {
      result,
      temperatures: p.bodies.map((b) => ("temperature" in b ? b.temperature : "none")),
      same_messages: p.bodies.every(
        (b) => JSON.stringify(b.messages) === JSON.stringify(p.bodies[0].messages),
      ),
    };
  };
  out = {
    // 400 for a request with a temperature, an answer without.
    picky: await ask((b) => ("temperature" in b ? [400, NO_TEMPERATURE] : [200, ANSWER])),
    // 400 about temperature whatever is sent.
    never: await ask(() => [400, NO_TEMPERATURE]),
    // 400 for another reason.
    other: await ask(() => [400, { error: { message: "model not found" } }]),
  };
} else {
  out = { error: `unknown scenario ${scenario}` };
}
process.stdout.write(JSON.stringify(out) + "\n", () => process.exit(out.error ? 2 : 0));
