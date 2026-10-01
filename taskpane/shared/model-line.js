// The line that says what each kind of work runs on (taskpane.js shows it
// above every tab), and the label of a provider and model for a result line.
// Pure: no DOM.
//
// A role of the line has { used, choice }:
//   used    { provider, asked, ran, ok? } of the last request that was sent
//           (the daemon's model_used event, or session_init for edits), or
//           null while none was. ran: the model the provider said it ran on,
//           null when it does not say.
//   choice  { provider, model } the next request will use (the menus of Setup).
// models: the account's Claude models (models_list), or null.

import { modelDisplayName } from "./model-menu.js";

const PROVIDERS = {
  claude: "Claude",
  codex: "Codex",
  antigravity: "Antigravity",
  ollama: "Ollama",
  openrouter: "OpenRouter",
  deepseek: "DeepSeek",
  groq: "Groq",
  mistral: "Mistral",
  gemini: "Gemini",
  "openai-compatible": "Custom server",
};
const ROLES = ["turn", "writing", "oneshot"];
const WORDS = { turn: "Edits", writing: "Writing", oneshot: "Judgments" };

// "Claude · Opus 5.5": a provider and a model as a person reads them.
export function modelLabel({ provider, model }, models) {
  const name = PROVIDERS[provider] ?? String(provider ?? "");
  const m = provider === "claude" ? modelDisplayName(model, models) : model || "its default model";
  return m ? `${name} · ${m}` : name;
}

// A model name without what does not tell models apart: a context suffix
// ("[1m]") and a date at the end of an id ("-20251001").
const bare = (s) =>
  String(s ?? "")
    .replace(/\[[^\]]*\]$/, "")
    .replace(/-\d{8}$/, "")
    .toLowerCase();

// Did another model run than the one asked for? true only when that can be
// told: the account's list says what the asked one resolves to; without it,
// a full claude-* id must be the id that ran, and a short name (an alias, the
// tag of a local model) must be in it.
export function ranOther(asked, ran, models) {
  if (!asked || !ran) return false;
  const a = bare(asked);
  const r = bare(ran);
  if (a === r) return false;
  const row = (models ?? []).find((m) => bare(m.value) === a);
  if (row?.resolvedModel) return bare(row.resolvedModel) !== r;
  return a.startsWith("claude-") || !r.includes(a);
}

// One kind of work: what its last request ran on; while there was none, the
// choice. A model other than the one asked for, a request that failed and a
// choice changed since are said.
export function roleText(role, { used, choice }, models) {
  const head = `${WORDS[role]}: `;
  const next = modelLabel(choice, models);
  if (!used) return `${head}${next} (not used yet)`;
  const asked = modelLabel({ provider: used.provider, model: used.asked }, models);
  if (ranOther(used.asked, used.ran, models)) {
    const ran = modelLabel({ provider: used.provider, model: used.ran }, models);
    return `${head}asked for ${asked}, ran on ${ran}`;
  }
  const label = modelLabel({ provider: used.provider, model: used.ran ?? used.asked }, models);
  if (used.ok === false) return `${head}${label} (last request failed)`;
  return asked === next ? `${head}${label}` : `${head}${label} (next: ${next})`;
}

export function modelLineText(state, models) {
  return ROLES.map((role) => roleText(role, state[role], models)).join("   ");
}
