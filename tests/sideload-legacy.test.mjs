import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join, dirname, win32 } from "node:path";
import { fileURLToPath } from "node:url";
import { EventEmitter } from "node:events";
import { runInNewContext } from "node:vm";
import { isLegacyManifest, manifestHash } from "../app/legacy-manifests.mjs";

const URL = new globalThis.URL("../app/sideload.mjs", import.meta.url);
const ROOT = join(dirname(fileURLToPath(URL)), "..");
const HOME = "/home/u";
const APPDATA = join(HOME, "AppData/Roaming");
const DEVELOPER = "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer";
const CATALOG =
  "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\TrustedCatalogs\\claude-code-office-trusted-catalog";
const oldPath = (file, appData = APPDATA) => join(appData, "Draftspect/manifests", file);
const newPath = (file, appData = APPDATA) => join(appData, "Sidewrite/manifests", file);
const macPath = (file) =>
  join(
    HOME,
    "Library/Containers",
    file === "word.xml" ? "com.microsoft.Word" : "com.microsoft.Excel",
    "Data/Documents/wef",
    file,
  );
const readFixture = (file) =>
  readFile(new globalThis.URL(`fixtures/legacy-manifests/${file}`, import.meta.url), "utf8");
const missing = () => Object.assign(new Error("missing"), { code: "ENOENT" });

// Run the actual sideload module with every filesystem/process boundary stubbed.
// No reg.exe, Office, system directories or user profile is touched on any OS.
async function sideload({ platform = "win32", appData = APPDATA } = {}) {
  const files = new Map();
  const dirs = new Set();
  const links = new Set();
  const registry = new Map();
  const calls = [];
  const warnings = [];
  const failures = { reg: () => false, copy: () => false };
  function directory(path) {
    for (let p = path; ; p = dirname(p)) {
      dirs.add(p);
      if (dirname(p) === p) break;
    }
  }
  function file(path, text) {
    directory(dirname(path));
    files.set(path, text);
  }
  function value(key, name, data = name, type = "REG_SZ") {
    if (!registry.has(key)) registry.set(key, new Map());
    registry.get(key).set(name, { data, type });
  }
  for (const name of ["word.xml", "excel.xml"]) {
    file(join(ROOT, "manifests", name), await readFile(join(ROOT, "manifests", name), "utf8"));
  }
  function spawn(command, args) {
    assert.equal(command, "reg.exe");
    const p = new EventEmitter();
    p.stdout = new EventEmitter();
    p.stderr = new EventEmitter();
    calls.push([command, ...args]);
    queueMicrotask(() => {
      const [verb, key] = args;
      const name = args.includes("/v") ? args[args.indexOf("/v") + 1] : undefined;
      const entries = registry.get(key);
      let code = 0;
      if (failures.reg(args)) code = 5;
      else if (verb === "add") value(key, name, args[args.indexOf("/d") + 1]);
      else if (verb === "query") {
        if (!entries || (name !== undefined && !entries.has(name))) code = 1;
        else {
          const rows = name === undefined ? [...entries] : [[name, entries.get(name)]];
          p.stdout.emit(
            "data",
            `${key}\r\n${rows.map(([n, v]) => `    ${n}    ${v.type}    ${v.data}`).join("\r\n")}\r\n`,
          );
        }
      } else if (verb === "delete") {
        if (name === undefined) registry.delete(key);
        else if (!entries?.delete(name)) code = 1;
      } else assert.fail(`Unexpected registry operation: ${args}`);
      if (code) p.stderr.emit("data", "stubbed registry failure");
      p.emit("exit", code);
    });
    return p;
  }
  const source = (await readFile(URL, "utf8"))
    .replace(/^import .*;\n/gm, "")
    .replace(/import\.meta\.url/g, JSON.stringify(URL.href))
    .replace(/^export /gm, "");
  const api = runInNewContext(`${source}\n;({ installAddin, uninstallAddin, isAddinInstalled });`, {
    join,
    dirname,
    win32,
    fileURLToPath,
    isLegacyManifest,
    manifestHash,
    spawn,
    homedir: () => HOME,
    process: { platform, env: { APPDATA: appData } },
    console: { warn: (message) => warnings.push(message) },
    readFile: async (path) => {
      if (!files.has(path)) throw missing();
      return files.get(path);
    },
    lstat: async (path) => {
      if (!files.has(path) && !dirs.has(path) && !links.has(path)) throw missing();
      return { isSymbolicLink: () => links.has(path), isFile: () => files.has(path) };
    },
    mkdir: async (path) => {
      calls.push(["mkdir", path]);
      directory(path);
    },
    copyFile: async (src, dst) => {
      calls.push(["copy", src, dst]);
      if (failures.copy(dst)) throw new Error("stubbed copy failure");
      if (!files.has(src)) throw missing();
      file(dst, files.get(src));
    },
    unlink: async (path) => {
      calls.push(["unlink", path]);
      if (!files.delete(path)) throw missing();
    },
    existsSync: (path) => files.has(path), // Also supports the unmodified regression baseline.
  });
  return { ...api, files, links, registry, calls, warnings, failures, file, value };
}

test("legacy manifests: verified pre-rename fixtures, BOM/CRLF, and custom content", async () => {
  for (const name of ["word.xml", "excel.xml"]) {
    const text = await readFixture(name);
    assert.equal(isLegacyManifest(name, text), true);
    assert.equal(isLegacyManifest(name, "\uFEFF" + text.replace(/\n/g, "\r\n")), true);
    assert.equal(isLegacyManifest("custom.xml", text), false);
    assert.equal(isLegacyManifest(name, text.replace("127.0.0.1:47834", "custom.example")), false);
    assert.equal(isLegacyManifest(name, text.replace("Draftspect for", "My custom")), false);
    assert.equal(isLegacyManifest(name, `<OfficeApp><!-- ${text} --></OfficeApp>`), false);
  }
});

test("Windows migration registers both replacements before removing only the owned old entries", async () => {
  const s = await sideload();
  for (const name of ["word.xml", "excel.xml"]) {
    s.file(oldPath(name), await readFixture(name));
    s.value(DEVELOPER, oldPath(name));
  }
  const unrelated = oldPath("custom.xml");
  s.file(unrelated, "keep custom manifest");
  s.file(join(APPDATA, "Draftspect/user-notes.txt"), "keep user files");
  s.value(DEVELOPER, unrelated);
  s.value(CATALOG, "Url", "\\\\server\\custom-catalog");
  await s.installAddin();
  for (const name of ["word.xml", "excel.xml"]) {
    assert.equal(s.files.has(oldPath(name)), false);
    assert.equal(s.registry.get(DEVELOPER).has(oldPath(name)), false);
    assert.equal(s.registry.get(DEVELOPER).has(newPath(name)), true);
    assert.ok(s.files.has(newPath(name)));
  }
  const lastAdd = s.calls.findLastIndex((c) => c[1] === "add");
  const firstDelete = s.calls.findIndex((c) => c[1] === "delete");
  assert.ok(firstDelete > lastAdd);
  assert.equal(s.files.get(unrelated), "keep custom manifest");
  assert.equal(s.files.get(join(APPDATA, "Draftspect/user-notes.txt")), "keep user files");
  assert.ok(s.registry.get(DEVELOPER).has(unrelated));
  assert.equal(s.registry.get(CATALOG).get("Url").data, "\\\\server\\custom-catalog");
  s.calls.length = 0;
  await s.installAddin();
  assert.ok(!s.calls.some((c) => c[1] === "delete"));
  await s.uninstallAddin();
  await s.uninstallAddin();
  assert.ok(s.files.has(unrelated));
});

test("Windows cleanup also removes owned old-only and stale registrations on uninstall", async () => {
  const s = await sideload();
  s.file(oldPath("word.xml"), await readFixture("word.xml"));
  s.value(DEVELOPER, oldPath("word.xml"));
  s.value(DEVELOPER, oldPath("excel.xml")); // File already missing.
  s.value(CATALOG, "Url", dirname(oldPath("word.xml")));
  await s.uninstallAddin();
  assert.equal(s.files.has(oldPath("word.xml")), false);
  assert.equal(s.registry.get(DEVELOPER).size, 0);
  assert.equal(s.registry.has(CATALOG), false);
  await s.uninstallAddin();
});

test("Windows cleanup preserves custom manifest content, registrations, and redirected paths", async () => {
  for (const kind of [
    "content",
    "data",
    "type",
    "query-failure",
    "file-link",
    "dir-link",
    "root-link",
  ]) {
    const s = await sideload();
    const path = oldPath("word.xml");
    const text = (await readFixture("word.xml")) + (kind === "content" ? "<!-- custom -->\n" : "");
    s.file(path, text);
    s.value(
      DEVELOPER,
      path,
      kind === "data" ? "C:\\custom.xml" : path,
      kind === "type" ? "REG_EXPAND_SZ" : "REG_SZ",
    );
    if (kind === "query-failure") s.failures.reg = (args) => args[0] === "query";
    if (kind === "file-link") s.links.add(path);
    if (kind === "dir-link") s.links.add(dirname(path));
    if (kind === "root-link") s.links.add(dirname(dirname(path)));
    await s.installAddin();
    assert.equal(s.files.get(path), text, kind);
    assert.ok(s.registry.get(DEVELOPER).has(path), kind);
    await s.uninstallAddin();
    assert.equal(s.files.get(path), text, kind);
    assert.ok(s.registry.get(DEVELOPER).has(path), kind);
  }
});

test("Windows migration leaves legacy usable when a replacement copy or registration fails", async () => {
  for (const kind of ["copy", "registration"]) {
    const s = await sideload();
    for (const name of ["word.xml", "excel.xml"]) {
      s.file(oldPath(name), await readFixture(name));
      s.value(DEVELOPER, oldPath(name));
    }
    if (kind === "copy") s.failures.copy = (dst) => dst === newPath("excel.xml");
    else s.failures.reg = (args) => args[0] === "add" && args.includes(newPath("excel.xml"));
    await assert.rejects(() => s.installAddin(), /stubbed/);
    for (const name of ["word.xml", "excel.xml"]) {
      assert.ok(s.files.has(oldPath(name)), kind);
      assert.ok(s.registry.get(DEVELOPER).has(oldPath(name)), kind);
    }
    assert.ok(!s.calls.some((c) => c[1] === "delete"));
    s.failures.copy = () => false;
    s.failures.reg = () => false;
    await s.installAddin();
    assert.equal(s.files.has(oldPath("word.xml")), false);
  }
});

test("Windows migration covers both redirected APPDATA and the historical tray location", async () => {
  const redirected = "/redirected/roaming";
  const s = await sideload({ appData: redirected });
  for (const root of [APPDATA, redirected]) {
    s.file(oldPath("word.xml", root), await readFixture("word.xml"));
    s.value(DEVELOPER, oldPath("word.xml", root));
  }
  await s.installAddin();
  assert.ok(s.files.has(newPath("word.xml", redirected)));
  for (const root of [APPDATA, redirected])
    assert.equal(s.files.has(oldPath("word.xml", root)), false);
});

test("registry parsing preserves custom values when APPDATA contains a registry type token", async () => {
  const root = "/redirected/Some REG_SZ Folder";
  for (const kind of ["type", "data"]) {
    const s = await sideload({ appData: root });
    const path = oldPath("word.xml", root);
    const text = await readFixture("word.xml");
    s.file(path, text);
    s.value(
      DEVELOPER,
      path,
      kind === "data" ? "C:\\custom.xml" : path,
      kind === "type" ? "REG_EXPAND_SZ" : "REG_SZ",
    );
    await s.installAddin();
    assert.equal(s.files.get(path), text);
    assert.ok(s.registry.get(DEVELOPER).has(path));
  }
});

test("Windows coexistence detection is read-only and offers migration once", async () => {
  const s = await sideload();
  await s.installAddin();
  assert.equal(await s.isAddinInstalled(), true);
  s.file(oldPath("word.xml"), await readFixture("word.xml"));
  s.value(DEVELOPER, oldPath("word.xml"));
  s.calls.length = 0;
  assert.equal(await s.isAddinInstalled(), false);
  assert.ok(s.calls.every((c) => c[1] === "query"));
  assert.ok(s.files.has(oldPath("word.xml")));
  await s.installAddin();
  assert.equal(await s.isAddinInstalled(), true);
});

test("macOS detects old GUIDs at the unchanged paths and replaces them on explicit install", async () => {
  const s = await sideload({ platform: "darwin" });
  for (const name of ["word.xml", "excel.xml"]) s.file(macPath(name), await readFixture(name));
  assert.equal(await s.isAddinInstalled(), false);
  assert.equal(s.calls.length, 0);
  await s.installAddin();
  assert.equal(await s.isAddinInstalled(), true);
  await s.installAddin();
  await s.uninstallAddin();
  await s.uninstallAddin();
  assert.equal(await s.isAddinInstalled(), false);
  assert.ok(!s.calls.some((c) => c[0] === "reg.exe"));
});

test("macOS and Windows preserve a custom destination manifest without partial install", async () => {
  for (const platform of ["darwin", "win32"]) {
    const s = await sideload({ platform });
    const path = platform === "darwin" ? macPath("excel.xml") : newPath("excel.xml");
    s.file(path, "<OfficeApp><Id>custom</Id></OfficeApp>");
    await assert.rejects(() => s.installAddin(), /custom manifest/);
    assert.equal(s.calls.length, 0);
    await assert.rejects(() => s.uninstallAddin(), /custom manifest/);
    assert.equal(s.files.get(path), "<OfficeApp><Id>custom</Id></OfficeApp>");
  }
});
