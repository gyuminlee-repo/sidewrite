import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createKeyStore, keyEnvName } from "../daemon/providers/keys.mjs";

const posixOnly = { skip: process.platform === "win32" };

async function withDir(fn) {
  const dir = await mkdtemp(join(tmpdir(), "keys-"));
  try {
    await fn(dir);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

test("env wins, file keeps 0600, public view hides the key", posixOnly, () =>
  withDir(async (dir) => {
    const store = createKeyStore({
      path: join(dir, "sub", "k.json"),
      env: { CC_OFFICE_KEY_GROQ: "envkey" },
    });
    await store.set("openrouter", { key: "sk-or-1" });
    await store.set("openai-compatible", { key: "x", base_url: "http://127.0.0.1:10100/v1" });
    assert.equal((await store.get("openrouter")).key, "sk-or-1");
    assert.equal((await store.get("groq")).key, "envkey");
    assert.equal((await store.get("openai-compatible")).base_url, "http://127.0.0.1:10100/v1");
    assert.equal((await stat(join(dir, "sub", "k.json"))).mode & 0o777, 0o600);
    assert.deepEqual(await store.publicView("openrouter"), { has_key: true, source: "file" });
    assert.deepEqual(await store.publicView("groq"), { has_key: true, source: "env" });
    assert.equal(JSON.stringify(await store.publicView("openrouter")).includes("sk-or"), false);
    await store.set("openrouter", { key: "" });
    assert.deepEqual(await store.publicView("openrouter"), { has_key: false, source: null });
    assert.deepEqual(await store.get("mistral"), { key: null, base_url: null, source: null });
  }),
);

test("base_url must be http(s); only key presets take a key; a key is one line", () =>
  withDir(async (dir) => {
    const store = createKeyStore({ path: join(dir, "k.json"), env: {} });
    await assert.rejects(
      store.set("openai-compatible", { key: "x", base_url: "file:///etc" }),
      /http/,
    );
    await assert.rejects(store.set("claude", { key: "x" }), /takes no API key/);
    await assert.rejects(store.set("nope", { key: "x" }), /Unknown provider/);
    await assert.rejects(store.set("groq", { key: "a\nb" }), /one line/);
    await assert.rejects(readFile(join(dir, "k.json"), "utf8"), /ENOENT/, "nothing was written");
  }));

test("no key material in the log lines, the return value of set, or any public view", () =>
  withDir(async (dir) => {
    const lines = [];
    const SECRET = "sk-or-v1-0123456789abcdef";
    const ENV_SECRET = "gsk_envsecret_zzzz";
    const store = createKeyStore({
      path: join(dir, "k.json"),
      env: { CC_OFFICE_KEY_GROQ: ENV_SECRET },
      log: (s) => lines.push(s),
    });
    const answer = await store.set("openrouter", { key: `  ${SECRET}  ` });
    assert.deepEqual(answer, { has_key: true, source: "file" });
    assert.equal((await store.get("openrouter")).key, SECRET, "stored trimmed");
    await store.set("openrouter", { key: "" });
    const seen = JSON.stringify([
      lines,
      answer,
      await store.publicView("openrouter"),
      await store.publicView("groq"),
    ]);
    for (const s of [SECRET, ENV_SECRET, SECRET.slice(0, 8), ENV_SECRET.slice(-4)]) {
      assert.equal(seen.includes(s), false, `leaked ${s}`);
    }
    assert.deepEqual(lines, [
      "[providers] key stored id=openrouter",
      "[providers] key removed id=openrouter",
    ]);
  }));

test("overlapping writes keep both keys; a broken file is an error, not an empty store", () =>
  withDir(async (dir) => {
    const path = join(dir, "k.json");
    const store = createKeyStore({ path, env: {} });
    await Promise.all([store.set("groq", { key: "g" }), store.set("mistral", { key: "m" })]);
    assert.deepEqual([(await store.get("groq")).key, (await store.get("mistral")).key], ["g", "m"]);
    const { writeFile } = await import("node:fs/promises");
    await writeFile(path, "{ not json");
    await assert.rejects(store.get("groq"), /provider-keys file is not valid JSON/);
  }));

test("environment variable names", () => {
  assert.equal(keyEnvName("openrouter"), "CC_OFFICE_KEY_OPENROUTER");
  assert.equal(keyEnvName("openai-compatible"), "CC_OFFICE_KEY_OPENAI_COMPATIBLE");
});
