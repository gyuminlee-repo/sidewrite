// Workspace-folder detection and marker management.
//
// "Workspace folder" = the folder Claude treats as cwd. The agent reads its
// CLAUDE.md, considers context files added to it, and uses it as the base
// for filesystem operations.
//
// Detection model (deliberately dead simple): the workspace is the folder
// that directly contains the open Word/Excel file. No walking up the tree,
// no marker search, no folder-name heuristics. For a document-centric tool
// the document *is* the subject, so its own folder is the natural cwd —
// and "no search" means no surprises (a stray ~/Desktop/.claude can't
// hijack the pick) and a result the user can always predict. The only
// thing that overrides it is the user explicitly choosing a folder in the
// picker; ensureWorkspaceMarker then drops a CLAUDE.md there.
//
// The guard: never auto-select $HOME itself, an OS-managed $HOME child
// (~/Library, …) or a filesystem root. A doc loose in ~ shouldn't make all
// of home the workspace; the user can pick a real folder instead. Under WSL
// the same applies to the Windows side: a drive root (<mount root>/d) and a
// Windows profile folder (<mount root>/c/Users/<name>) are never picked.

import { writeFile, stat } from "node:fs/promises";
import { homedir, release } from "node:os";
import { join, dirname, basename, resolve as resolvePath } from "node:path";
import { fileURLToPath } from "node:url";
import { isSystemHomeChild } from "./system-paths.mjs";

const MARKERS = ["CLAUDE.md", ".claude"];
const HOME = homedir();

// Office.js gives us paths in a few shapes depending on platform and where
// the doc came from:
//   - macOS local:   "/Users/x/Docs/spec.docx"
//   - macOS file://: "file:///Users/x/Docs/spec.docx"
//   - Windows local: "C:\\Users\\x\\Docs\\spec.docx"
//   - Windows URL:   "file:///C:/Users/x/Docs/spec.docx"
//   - SharePoint:    "https://contoso.sharepoint.com/…/spec.docx"
// Hand the URL forms to fileURLToPath so the Windows drive letter survives
// (a manual `replace(/^file:\/\//)` turns "file:///C:/x" into "/C:/x").
// Treat non-file:// URLs (SharePoint, OneDrive cloud) as unaddressable.
// WSL: the daemon runs inside WSL while Word runs on Windows, so Office
// hands us Windows paths. Translate them to the Linux view:
//   D:\a\b.docx, D:/a/b.docx, file:///D:/a/b.docx -> <mount root>/d/a/b.docx
//   \\wsl.localhost\<distro>\home\a, \\wsl$\<distro>\home\a -> /home/a
//     (null when WSL_DISTRO_NAME is set and names another distro)
//   any other UNC share -> null (not addressable from WSL)
// UNC forms accept "/" as well as "\" (//wsl.localhost/<distro>/home/a).
// CC_OFFICE_WSL_PATHS "1" forces on, "0" forces off, unset = auto (Linux
// kernel release mentions microsoft). CC_OFFICE_WSL_MOUNT_ROOT overrides
// the drive mount root (default /mnt). Read per call so tests can switch.
export function wslPathsEnabled(env = process.env) {
  if (env.CC_OFFICE_WSL_PATHS === "1") return true;
  if (env.CC_OFFICE_WSL_PATHS === "0") return false;
  return process.platform === "linux" && /microsoft/i.test(release());
}

function wslMountRoot(env = process.env) {
  return (env.CC_OFFICE_WSL_MOUNT_ROOT || "/mnt").replace(/\/+$/, "");
}

// undefined: not a Windows-form path, or translation is off (caller keeps
// its own handling). null: Windows-form but unreachable from WSL.
export function windowsPathToWsl(p, env = process.env) {
  if (typeof p !== "string" || !wslPathsEnabled(env)) return undefined;
  const root = wslMountRoot(env);
  let m = /^file:\/\/\/([a-zA-Z]):[\\/](.*)$/i.exec(p);
  if (m) {
    let rest;
    try {
      rest = decodeURIComponent(m[2]);
    } catch {
      return null;
    }
    return `${root}/${m[1].toLowerCase()}/${rest.replace(/\\/g, "/")}`;
  }
  m = /^([a-zA-Z]):[\\/](.*)$/.exec(p);
  if (m) return `${root}/${m[1].toLowerCase()}/${m[2].replace(/\\/g, "/")}`;
  m = /^[\\/]{2}(?:wsl\.localhost|wsl\$)[\\/]([^\\/]+)([\\/].*)?$/i.exec(p);
  if (m) {
    const distro = env.WSL_DISTRO_NAME;
    if (distro && distro.toLowerCase() !== m[1].toLowerCase()) return null;
    return (m[2] || "/").replace(/\\/g, "/");
  }
  if (/^[\\/]{2}[^\\/]/.test(p)) return null;
  return undefined;
}

// For user-typed paths (set_cwd, context entries): Windows-form input is
// translated under WSL, everything else is returned unchanged so relative
// context paths keep resolving against the workspace. null = unreachable.
export function toLocalPath(p, env = process.env) {
  const w = windowsPathToWsl(p, env);
  return w === undefined ? p : w;
}

function normalizeDocPath(docPath) {
  if (!docPath) return null;
  const wsl = windowsPathToWsl(docPath);
  if (wsl !== undefined) return wsl === null ? null : resolvePath(wsl);
  if (/^file:\/\//i.test(docPath)) {
    try {
      return resolvePath(fileURLToPath(docPath));
    } catch {
      return null;
    }
  }
  // A Windows drive path (C:\… or C:/…) or a UNC share (\\host\share)
  // is a real filesystem path, NOT a URL — and must be recognized BEFORE
  // the scheme check below, because a bare "C:" matches the URL-scheme
  // regex and would be wrongly discarded. That was the symptom on
  // Windows: the workspace always fell back to the daemon's launch dir
  // instead of the open document's folder. Don't decodeURIComponent
  // here — Office hands Windows local paths verbatim, so a literal "%"
  // in a filename must survive.
  if (/^[a-zA-Z]:[\\/]/.test(docPath) || /^\\\\[^\\]/.test(docPath)) {
    return resolvePath(docPath);
  }
  if (/^[a-z][a-z0-9+.-]*:/i.test(docPath)) {
    // Some non-file URL (https://…sharepoint…) — no filesystem path.
    return null;
  }
  return resolvePath(decodeURIComponent(docPath));
}

async function docDirOf(p) {
  try {
    const s = await stat(p);
    return s.isDirectory() ? p : dirname(p);
  } catch {
    return dirname(p);
  }
}

// The workspace for a document: the folder that directly contains it.
// Returns that folder, or null if the doc is unaddressable (cloud URL) or
// its folder isn't a sane cwd ($HOME itself / an OS-managed $HOME child /
// the filesystem root / under WSL a Windows drive root or profile folder).
export async function resolveWorkspaceRoot(docPath) {
  const p = normalizeDocPath(docPath);
  if (!p) return null;
  const docDir = await docDirOf(p);
  if (
    docDir === HOME ||
    isSystemHomeChild(docDir) ||
    isFilesystemRoot(docDir) ||
    isWindowsSideRoot(docDir)
  ) {
    return null;
  }
  return docDir;
}

// Under WSL: <mount root>/d (a drive) and <mount root>/c/Users/<name> (a
// Windows profile) are the Windows equivalents of "/" and $HOME.
function isWindowsSideRoot(p) {
  if (!wslPathsEnabled()) return false;
  const root = wslMountRoot();
  if (!p.startsWith(root + "/")) return false;
  const rest = p.slice(root.length);
  return /^\/[a-z]$/i.test(rest) || /^\/[a-z]\/users\/[^/]+$/i.test(rest);
}

// `dirname(p) === p` is the classic "p is a root" test, but on Windows it
// also fires for a UNC share root (`\\server\share\`) — and a share root
// IS a real folder a user can sanely treat as a workspace (the share is a
// folder, unlike `C:\` which is the drive itself). Block drive / POSIX
// roots, allow UNC share roots.
function isFilesystemRoot(p) {
  if (dirname(p) !== p) return false;
  return !/^\\\\/.test(p);
}

// Same answer, in the shape the taskpane expects. confidence is always
// "doc" (deterministic — the doc's own folder), so the pane can auto-apply
// it without a confirm banner. null when there's nothing sane to suggest.
export async function suggestWorkspaceRoot(docPath) {
  const cwd = await resolveWorkspaceRoot(docPath);
  return cwd ? { cwd, confidence: "doc" } : null;
}

// Drop an empty CLAUDE.md in the given folder if neither marker exists.
// Called on explicit user pick so the folder is set up for next time
// (context files, workspace instructions).
export async function ensureWorkspaceMarker(cwd) {
  if (!cwd) return false;
  for (const m of MARKERS) {
    try {
      await stat(join(cwd, m));
      return false;
    } catch {}
  }
  const seed =
    `# ${basename(cwd)}\n\n` +
    `_Workspace for the Office add-in. Add notes, glossary entries, ` +
    `style preferences, or instructions you want the assistant to follow ` +
    `here. The file is loaded automatically every time you open a document ` +
    `in this folder._\n`;
  await writeFile(join(cwd, "CLAUDE.md"), seed, { encoding: "utf8", flag: "wx" });
  return true;
}
