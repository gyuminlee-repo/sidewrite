// API keys of the key-tier providers (P14).
//
// A key lives in one of two places and nowhere else:
//   the environment   CC_OFFICE_KEY_<ID> (CC_OFFICE_KEY_OPENROUTER, ...),
//                     which wins over the file;
//   the key file      ~/.claude/office-addins/provider-keys.json, mode 0600,
//                     { "<id>": { "key": "...", "base_url": "..." } }.
// It is never sent to the task pane and never logged: the pane gets
// publicView() ({ has_key, source }) and the log lines name the provider only.
// get() is for the daemon's own requests to the provider.

import { chmod, mkdir, readFile, rename, unlink, writeFile } from "node:fs/promises";
import { randomUUID } from "node:crypto";
import { homedir } from "node:os";
import { basename, dirname, join } from "node:path";
import { presetById } from "./registry.mjs";

export const DEFAULT_KEYS_PATH = join(homedir(), ".claude", "office-addins", "provider-keys.json");

export const keyEnvName = (id) =>
  `CC_OFFICE_KEY_${String(id)
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "_")}`;

function checkBaseUrl(base_url) {
  let u;
  try {
    u = new URL(String(base_url));
  } catch {
    throw new Error("base_url must be an http or https URL");
  }
  if (u.protocol !== "http:" && u.protocol !== "https:") {
    throw new Error("base_url must be an http or https URL");
  }
  return String(base_url).trim().replace(/\/+$/, "");
}

export function createKeyStore({
  path = DEFAULT_KEYS_PATH,
  env = process.env,
  log = () => {},
} = {}) {
  let chain = Promise.resolve();

  async function readAll() {
    let raw;
    try {
      raw = await readFile(path, "utf8");
    } catch (e) {
      if (e?.code === "ENOENT") return {};
      throw e;
    }
    try {
      const data = JSON.parse(raw);
      return data && typeof data === "object" && !Array.isArray(data) ? data : {};
    } catch {
      // The text of the file is not put in the message: it may hold keys.
      throw new Error(`provider-keys file is not valid JSON: ${path}`);
    }
  }

  // Written under its final mode: the keys are never in a readable file.
  async function writeAll(data) {
    await mkdir(dirname(path), { recursive: true });
    const tmp = join(dirname(path), `.${basename(path)}.${process.pid}.${randomUUID()}.tmp`);
    await writeFile(tmp, JSON.stringify(data, null, 2) + "\n", { encoding: "utf8", mode: 0o600 });
    try {
      await chmod(tmp, 0o600);
      await rename(tmp, path);
    } catch (e) {
      await unlink(tmp).then(
        () => undefined,
        () => undefined,
      );
      throw e;
    }
  }

  async function get(id) {
    const row = (await readAll())[id] ?? {};
    const fromEnv = String(env[keyEnvName(id)] ?? "").trim();
    const fromFile = typeof row.key === "string" && row.key ? row.key : null;
    return {
      key: fromEnv || fromFile,
      base_url: typeof row.base_url === "string" && row.base_url ? row.base_url : null,
      source: fromEnv ? "env" : fromFile ? "file" : null,
    };
  }

  async function publicView(id) {
    const { key, source } = await get(id);
    return { has_key: !!key, source };
  }

  // key "": remove the stored key (an environment key is not touched).
  // Returns publicView(id), never the key.
  function set(id, { key, base_url } = {}) {
    const run = chain
      .then(
        () => undefined,
        () => undefined,
      )
      .then(async () => {
        const preset = presetById(id);
        if (!preset) throw new Error(`Unknown provider ${id}`);
        if (preset.tier !== "key") throw new Error(`${preset.label} takes no API key`);
        const k = String(key ?? "").trim();
        if (/[\r\n]/.test(k) || k.length > 1000) throw new Error("An API key is one line");
        const url = base_url == null || base_url === "" ? null : checkBaseUrl(base_url);
        const all = await readAll();
        const row = { ...(all[id] ?? {}) };
        if (k) row.key = k;
        else delete row.key;
        if (url) row.base_url = url;
        else if (base_url === "") delete row.base_url;
        if (Object.keys(row).length) all[id] = row;
        else delete all[id];
        await writeAll(all);
        log(`[providers] key ${k ? "stored" : "removed"} id=${id}`);
        return publicView(id);
      });
    chain = run;
    return run;
  }

  return { get, set, publicView, path };
}
