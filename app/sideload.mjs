// Auto-sideload the Sidewrite add-in manifests into Word and Excel.
//
// macOS: drop manifests into the per-host wef/ folder inside the Office
// app's container. Word/Excel scan that folder on launch and surface
// every manifest under Insert → Office Add-ins → SHARED FOLDER.
//
//   ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/word.xml
//   ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/excel.xml
//
// Windows: copy manifests to a folder under %APPDATA%, then register
// each manifest under the WEF\Developer registry key. Office scans that
// key on launch and sideloads each listed manifest directly — no admin,
// no network share. (A Trusted Catalog would require the folder to be a
// real SMB share with a UNC Url; Office silently ignores catalogs whose
// Url is a local path, which is why the catalog approach never showed
// the add-in. The Developer key is what the official Office dev tooling
// uses for local sideloading.)
//
//   %APPDATA%\Sidewrite\manifests\{word,excel}.xml
//   HKCU\Software\Microsoft\Office\16.0\WEF\Developer
//     <full manifest path> = <full manifest path>  (REG_SZ)
//
// All operations are idempotent: re-running install is safe, uninstall
// silently ignores already-removed pieces.

import { copyFile, mkdir, unlink, readFile, lstat } from "node:fs/promises";
import { join, dirname, win32 } from "node:path";
import { homedir } from "node:os";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import { isLegacyManifest, manifestHash } from "./legacy-manifests.mjs";

const __dirname = dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = join(__dirname, "..");
const MANIFESTS_DIR = join(PROJECT_ROOT, "manifests");

// One row per supported Office host. `mac_container` is the bundle ID whose
// sandbox we drop the manifest into on macOS. `guid` is purely informational
// here (the manifest carries its own Id element).
const HOSTS = [
  { host: "Word", file: "word.xml", mac_container: "com.microsoft.Word" },
  { host: "Excel", file: "excel.xml", mac_container: "com.microsoft.Excel" },
];

// null = missing, false = symlink/non-file (preserve), string = file content.
// Do not follow a redirected product/manifests directory during cleanup.
async function installedManifest(path) {
  for (const p of [dirname(dirname(path)), dirname(path), path]) {
    let stat;
    try {
      stat = await lstat(p);
    } catch (e) {
      if (e.code === "ENOENT") return null;
      throw e;
    }
    if (stat.isSymbolicLink() || (p === path && !stat.isFile())) return false;
  }
  return readFile(path, "utf8");
}

async function checkDestination(path, host) {
  const text = await installedManifest(path);
  if (text === null) return;
  const source = await readFile(join(MANIFESTS_DIR, host.file), "utf8");
  if (
    typeof text === "string" &&
    (manifestHash(text) === manifestHash(source) || isLegacyManifest(host.file, text))
  )
    return;
  throw new Error(`Preserved a custom manifest at ${path}; move it before installing Sidewrite.`);
}

// ---------------------------------------------------------------------------
// macOS
// ---------------------------------------------------------------------------
function macWefDir(container) {
  return join(homedir(), "Library", "Containers", container, "Data", "Documents", "wef");
}
function macWefPath(host) {
  return join(macWefDir(host.mac_container), host.file);
}

async function macInstall() {
  for (const host of HOSTS) await checkDestination(macWefPath(host), host);
  const installed = [];
  for (const host of HOSTS) {
    const wef = macWefDir(host.mac_container);
    await mkdir(wef, { recursive: true });
    const src = join(MANIFESTS_DIR, host.file);
    const dst = join(wef, host.file);
    await copyFile(src, dst);
    installed.push({ host: host.host, path: dst });
  }
  return { installed };
}

async function macUninstall() {
  const removed = [];
  for (const host of HOSTS) {
    const p = macWefPath(host);
    await checkDestination(p, host);
    try {
      await unlink(p);
      removed.push({ host: host.host, path: p });
    } catch (e) {
      if (e.code !== "ENOENT") throw e;
    }
  }
  return { removed };
}

async function macIsInstalled() {
  for (const host of HOSTS) {
    const text = await installedManifest(macWefPath(host));
    const source = await readFile(join(MANIFESTS_DIR, host.file), "utf8");
    if (typeof text !== "string" || manifestHash(text) !== manifestHash(source)) return false;
  }
  return true;
}

// ---------------------------------------------------------------------------
// Windows
// ---------------------------------------------------------------------------
// Office scans this key on launch and sideloads each listed manifest
// directly — the mechanism the official Office dev tooling uses. No
// admin, no SMB share. The value NAME is the manifest's full path (what
// Office reads); we set the data to the same path for legibility.
const WIN_DEVELOPER_KEY = "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer";

// Legacy: earlier builds (incorrectly) registered a Trusted Catalog with
// a LOCAL-path Url. Office silently ignores such catalogs — a catalog Url
// must be a UNC share — which is why the add-in never appeared. We delete
// this on install + uninstall so upgrading users don't keep a dead entry.
const WIN_LEGACY_CATALOG_KEY =
  "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\TrustedCatalogs\\claude-code-office-trusted-catalog";

function winCatalogDir() {
  if (process.env.APPDATA) return join(process.env.APPDATA, "Sidewrite", "manifests");
  return join(homedir(), "AppData", "Roaming", "Sidewrite", "manifests");
}

function winLegacyDirs() {
  // The tray used homedir, while the WSL installer used APPDATA (which can
  // be redirected). Inspect both exact former locations, never scan WEF.
  const roots = [join(homedir(), "AppData", "Roaming"), process.env.APPDATA].filter(Boolean);
  return [...new Set(roots.map((root) => join(root, "Draftspect", "manifests")))];
}

const sameWindowsPath = (a, b) =>
  typeof a === "string" && win32.normalize(a).toLowerCase() === win32.normalize(b).toLowerCase();

function regCommand(args) {
  return new Promise((resolve, reject) => {
    const p = spawn("reg.exe", args, { windowsHide: true });
    let stdout = "";
    let stderr = "";
    p.stdout.on("data", (d) => {
      stdout += d.toString();
    });
    p.stderr.on("data", (d) => {
      stderr += d.toString();
    });
    p.on("exit", (code) => {
      if (code === 0) resolve(stdout);
      else reject(new Error(`reg.exe ${args.join(" ")} exited ${code}: ${stderr.trim()}`));
    });
    p.on("error", reject);
  });
}

async function regAddString(key, name, value) {
  return regCommand(["add", key, "/v", name, "/t", "REG_SZ", "/d", value, "/f"]);
}
async function regDeleteKey(key) {
  return regCommand(["delete", key, "/f"]);
}
async function regDeleteValue(key, name) {
  return regCommand(["delete", key, "/v", name, "/f"]);
}
async function regReadString(key, name) {
  let output;
  try {
    // Read the key so a missing value is distinguishable from a failed query.
    output = await regCommand(["query", key]);
  } catch {
    return false; // Unreadable/missing key: preserve anything we cannot verify.
  }
  for (const line of output.split(/\r?\n/)) {
    const row = line.trimStart();
    if (row.slice(0, name.length).toLowerCase() !== name.toLowerCase()) continue;
    const rest = row.slice(name.length);
    if (!/^\s/.test(rest)) continue; // A different value with the same prefix.
    const match = rest.match(/^\s+(REG_\w+)\s+(.*)$/);
    return match?.[1] === "REG_SZ" ? match[2] : false;
  }
  return null;
}

async function winCleanupLegacy(hosts = HOSTS) {
  for (const dir of winLegacyDirs()) {
    for (const host of hosts) {
      const path = join(dir, host.file);
      const text = await installedManifest(path);
      if (text !== null && (typeof text !== "string" || !isLegacyManifest(host.file, text))) {
        console.warn(`legacy manifest: preserved custom file (${path})`);
        continue;
      }
      const value = await regReadString(WIN_DEVELOPER_KEY, path);
      if (value !== null && !sameWindowsPath(value, path)) {
        if (text !== null)
          console.warn(`legacy manifest: preserved unverified registration (${path})`);
        continue;
      }
      if (value !== null) await regDeleteValue(WIN_DEVELOPER_KEY, path);
      if (text !== null) await unlink(path);
    }
  }
  // A same-named custom catalog (for example an actual UNC share) is not ours.
  const url = await regReadString(WIN_LEGACY_CATALOG_KEY, "Url");
  if ([winCatalogDir(), ...winLegacyDirs()].some((dir) => sameWindowsPath(url, dir))) {
    await regDeleteKey(WIN_LEGACY_CATALOG_KEY);
  }
}
function regQueryValue(key, name) {
  return new Promise((resolve) => {
    const p = spawn("reg.exe", ["query", key, "/v", name], { windowsHide: true });
    p.on("exit", (code) => resolve(code === 0));
    p.on("error", () => resolve(false));
  });
}

async function winInstall() {
  const catalogDir = winCatalogDir();
  for (const host of HOSTS) await checkDestination(join(catalogDir, host.file), host);
  await mkdir(catalogDir, { recursive: true });
  const copied = [];
  for (const host of HOSTS) {
    const dst = join(catalogDir, host.file);
    await copyFile(join(MANIFESTS_DIR, host.file), dst);
    // Register the manifest under WEF\Developer (value name = path).
    await regAddString(WIN_DEVELOPER_KEY, dst, dst);
    copied.push({ host: host.host, path: dst });
  }
  // Migrate only after every replacement has been copied and registered.
  await winCleanupLegacy();
  return { installed: copied, catalog: catalogDir, registry: WIN_DEVELOPER_KEY };
}

async function winUninstall() {
  const removed = [];
  const catalogDir = winCatalogDir();
  for (const host of HOSTS) {
    const dst = join(catalogDir, host.file);
    await checkDestination(dst, host);
    try {
      await regDeleteValue(WIN_DEVELOPER_KEY, dst);
    } catch {
      /* value not present */
    }
    try {
      await unlink(dst);
      removed.push(dst);
    } catch {
      /* file gone — fine */
    }
  }
  await winCleanupLegacy();
  return { removed, registry: WIN_DEVELOPER_KEY };
}

async function winIsInstalled() {
  const catalogDir = winCatalogDir();
  // Installed iff every host's manifest is registered under the key.
  for (const host of HOSTS) {
    if (!(await regQueryValue(WIN_DEVELOPER_KEY, join(catalogDir, host.file)))) {
      return false;
    }
  }
  // A rename upgrade may already have both names registered. Offer the same
  // explicit reinstall flow instead of treating that partial migration as done.
  for (const dir of winLegacyDirs()) {
    for (const host of HOSTS) {
      const path = join(dir, host.file);
      const text = await installedManifest(path);
      if (text !== null && (typeof text !== "string" || !isLegacyManifest(host.file, text)))
        continue;
      const value = await regReadString(WIN_DEVELOPER_KEY, path);
      if (sameWindowsPath(value, path) || (text !== null && value === null)) return false;
    }
  }
  return true;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------
export async function isAddinInstalled() {
  if (process.platform === "darwin") return macIsInstalled();
  if (process.platform === "win32") return await winIsInstalled();
  return false;
}

export async function installAddin() {
  if (process.platform === "darwin") return macInstall();
  if (process.platform === "win32") return winInstall();
  throw new Error(
    `Auto-sideload not supported on ${process.platform}; see README for manual instructions.`,
  );
}

export async function uninstallAddin() {
  if (process.platform === "darwin") return macUninstall();
  if (process.platform === "win32") return winUninstall();
  throw new Error(`Auto-sideload not supported on ${process.platform}.`);
}
