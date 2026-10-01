// One-shot judgments go to the provider the pane chose (P14).
//
// getChoice(key) is the pane's choice for the "oneshot" role: { id, model },
// or null for the default (Claude, haiku). forKey(key) gives a function with
// the contract of oneshot.mjs runOneshot ({ ok, text, json?, error?,
// limited? }, never throws); the choice is read each time it runs, so a
// change in the pane applies to the next request.
//
// Nothing here falls back to another provider: a provider that cannot be
// used (no key, no server URL, unknown id) is an error result, so the text
// of a manuscript never goes somewhere the user did not choose.
//
// modelLabel(key) is the model name written into decision receipts:
// "<provider>:<model>", or the bare alias for Claude as before this stage.

import { presetById } from "./registry.mjs";

const CLAUDE_DEFAULT_MODEL = "haiku";

export function createOneshotRouter({ claude, codex, antigravity, http, keys, getChoice }) {
  async function run(key, req) {
    const choice = getChoice(key);
    if (!choice || choice.id === "claude") {
      return claude({ ...req, model: choice?.model ?? CLAUDE_DEFAULT_MODEL });
    }
    const preset = presetById(choice.id);
    if (!preset) return { ok: false, error: `Unknown provider ${choice.id}` };
    if (preset.id === "codex") return codex({ ...req, model: choice.model });
    if (preset.id === "antigravity") return antigravity({ ...req, model: choice.model });
    if (preset.tier === "local") {
      return http({ ...req, baseUrl: preset.baseUrl, apiKey: null, model: choice.model });
    }
    const stored = await keys.get(preset.id);
    if (!stored?.key) {
      return {
        ok: false,
        error: `No API key for ${preset.label}. Enter one in Setup, Providers.`,
      };
    }
    const baseUrl = preset.baseUrl ?? stored.base_url;
    if (!baseUrl) return { ok: false, error: `No server URL for ${preset.label}.` };
    return http({ ...req, baseUrl, apiKey: stored.key, model: choice.model });
  }

  return {
    forKey: (key) => (req) =>
      run(key, req).then(
        (r) => r,
        (e) => ({ ok: false, error: String(e?.message ?? e) }),
      ),
    modelLabel(key) {
      const choice = getChoice(key);
      if (!choice || choice.id === "claude") return choice?.model ?? CLAUDE_DEFAULT_MODEL;
      return `${choice.id}:${choice.model ?? "default"}`;
    },
  };
}
