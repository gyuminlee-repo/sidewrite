// Unit tests for daemon/workspace.mjs.
//
// New model: the workspace is simply the folder that directly contains the
// open document — no walking up the tree, no marker search, no folder-name
// heuristic. The guard is "never $HOME / an OS-managed $HOME child / the
// filesystem root", plus under WSL "never a Windows drive root or profile
// folder". ensureWorkspaceMarker still drops a seed CLAUDE.md
// on explicit user pick. Tests use a fresh tmpdir per test.

import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdir, mkdtemp, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, dirname, resolve } from "node:path";

import {
  resolveWorkspaceRoot,
  suggestWorkspaceRoot,
  ensureWorkspaceMarker,
  windowsPathToWsl,
  toLocalPath,
} from "../daemon/workspace.mjs";

async function makeTmp() {
  return await mkdtemp(join(tmpdir(), "cc-office-ws-test-"));
}

test("resolveWorkspaceRoot returns the document's own folder", async () => {
  const root = await makeTmp();
  try {
    const dir = join(root, "a", "b");
    await mkdir(dir, { recursive: true });
    await writeFile(join(dir, "doc.docx"), "");
    assert.equal(await resolveWorkspaceRoot(join(dir, "doc.docx")), dir);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("resolveWorkspaceRoot does NOT walk up to ancestor markers (footgun fix)", async () => {
  // The ~/Desktop/.claude footgun: a stray marker in an ancestor used to
  // hijack the pick. Now the doc's own folder always wins.
  const root = await makeTmp();
  try {
    const sub = join(root, "Test Folder");
    await mkdir(sub, { recursive: true });
    await writeFile(join(root, "CLAUDE.md"), "# stray\n");
    await mkdir(join(root, ".claude"));
    await mkdir(join(root, ".git"));
    await writeFile(join(sub, "doc.docx"), "");
    // Returns the doc's folder, not `root` (where the markers are).
    assert.equal(await resolveWorkspaceRoot(join(sub, "doc.docx")), sub);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("resolveWorkspaceRoot returns the directory itself when handed a dir path", async () => {
  const root = await makeTmp();
  try {
    const dir = join(root, "ws");
    await mkdir(dir, { recursive: true });
    assert.equal(await resolveWorkspaceRoot(dir), dir);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("resolveWorkspaceRoot handles file:// URLs (POSIX)", async () => {
  const root = await makeTmp();
  try {
    const ws = join(root, "ws");
    await mkdir(ws, { recursive: true });
    const docPath = join(ws, "doc.docx");
    await writeFile(docPath, "");
    assert.equal(await resolveWorkspaceRoot("file://" + docPath), ws);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("resolveWorkspaceRoot returns null for non-file:// URLs (cloud docs)", async () => {
  const got = await resolveWorkspaceRoot("https://contoso.sharepoint.com/sites/x/spec.docx");
  assert.equal(got, null);
});

async function withEnv(vars, fn) {
  const prev = {};
  for (const [k, v] of Object.entries(vars)) {
    prev[k] = process.env[k];
    if (v === undefined) delete process.env[k];
    else process.env[k] = v;
  }
  try {
    return await fn();
  } finally {
    for (const [k, v] of Object.entries(prev)) {
      if (v === undefined) delete process.env[k];
      else process.env[k] = v;
    }
  }
}
const WSL_ON = {
  CC_OFFICE_WSL_PATHS: "1",
  CC_OFFICE_WSL_MOUNT_ROOT: undefined,
  WSL_DISTRO_NAME: undefined,
};
const WSL_OFF = { CC_OFFICE_WSL_PATHS: "0" };
const posixOnly = { skip: process.platform === "win32" };

test("non-WSL mode: Windows drive and UNC paths are NOT treated as URLs (regression)", () =>
  withEnv(WSL_OFF, async () => {
    // Upstream regression guard: a bare "C:" once matched the URL-scheme regex.
    assert.notEqual(await resolveWorkspaceRoot("C:\\Users\\me\\Docs\\spec.docx"), null);
    assert.notEqual(await resolveWorkspaceRoot("\\\\fileserver\\share\\spec.docx"), null);
  }));

test("non-WSL mode: result equals the untranslated path", () =>
  withEnv(WSL_OFF, async () => {
    assert.equal(await resolveWorkspaceRoot("D:\\a\\b.docx"), dirname(resolve("D:\\a\\b.docx")));
    assert.equal(toLocalPath("D:\\x"), "D:\\x");
    assert.equal(windowsPathToWsl("D:\\x"), undefined);
  }));

test("WSL mode: drive paths and file:// URLs map under the mount root", posixOnly, () =>
  withEnv(WSL_ON, async () => {
    assert.equal(await resolveWorkspaceRoot("D:\\a\\b.docx"), "/mnt/d/a");
    assert.equal(await resolveWorkspaceRoot("file:///D:/a/b.docx"), "/mnt/d/a");
    assert.equal(
      await resolveWorkspaceRoot("C:\\Users\\me\\Docs\\spec.docx"),
      "/mnt/c/Users/me/Docs",
    );
  }),
);

test("WSL mode: \\\\wsl.localhost and \\\\wsl$ shares map to Linux paths", posixOnly, () =>
  withEnv(WSL_ON, async () => {
    assert.equal(
      await resolveWorkspaceRoot("\\\\wsl.localhost\\Ubuntu\\home\\a\\b.docx"),
      "/home/a",
    );
    assert.equal(await resolveWorkspaceRoot("\\\\wsl$\\Ubuntu\\home\\a\\b.docx"), "/home/a");
  }),
);

test("WSL mode: a Windows drive root or profile folder is never the workspace", posixOnly, () =>
  withEnv(WSL_ON, async () => {
    assert.equal(await resolveWorkspaceRoot("D:\\x.docx"), null);
    assert.equal(await resolveWorkspaceRoot("C:\\Users\\me\\x.docx"), null);
    assert.equal(await resolveWorkspaceRoot("c:/users/me/x.docx"), null);
    assert.equal(await resolveWorkspaceRoot("D:\\proj\\x.docx"), "/mnt/d/proj");
  }),
);

test("WSL mode: drive-root guard follows the mount root override", posixOnly, () =>
  withEnv({ ...WSL_ON, CC_OFFICE_WSL_MOUNT_ROOT: "/x/" }, async () => {
    assert.equal(await resolveWorkspaceRoot("D:\\x.docx"), null);
    assert.equal(await resolveWorkspaceRoot("C:\\Users\\me\\x.docx"), null);
    assert.equal(await resolveWorkspaceRoot("D:\\proj\\x.docx"), "/x/d/proj");
  }),
);

test("non-WSL mode: no drive-root guard on /mnt paths", posixOnly, () =>
  withEnv(WSL_OFF, async () => {
    assert.equal(await resolveWorkspaceRoot("/mnt/d/x.docx"), "/mnt/d");
  }),
);

test("WSL mode: forward-slash UNC forms behave like backslash forms", posixOnly, () =>
  withEnv(WSL_ON, async () => {
    assert.equal(windowsPathToWsl("//wsl.localhost/Ubuntu/home/a/b.docx"), "/home/a/b.docx");
    assert.equal(windowsPathToWsl("//wsl$/Ubuntu/home/a"), "/home/a");
    assert.equal(windowsPathToWsl("//wsl.localhost/Ubuntu"), "/");
    assert.equal(windowsPathToWsl("//server/share/b.docx"), null);
    assert.equal(await resolveWorkspaceRoot("//wsl.localhost/Ubuntu/home/a/b.docx"), "/home/a");
    assert.equal(await resolveWorkspaceRoot("//server/share/b.docx"), null);
  }),
);

test("WSL mode: \\\\wsl.localhost paths of another distro are unaddressable", posixOnly, () =>
  withEnv({ ...WSL_ON, WSL_DISTRO_NAME: "Ubuntu" }, async () => {
    assert.equal(windowsPathToWsl("\\\\wsl.localhost\\ubuntu\\home\\a"), "/home/a");
    assert.equal(windowsPathToWsl("\\\\wsl$\\UBUNTU\\home\\a"), "/home/a");
    assert.equal(windowsPathToWsl("\\\\wsl.localhost\\Debian\\home\\a"), null);
    assert.equal(windowsPathToWsl("//wsl$/Debian/home/a"), null);
    assert.equal(await resolveWorkspaceRoot("\\\\wsl$\\Debian\\home\\a\\b.docx"), null);
  }),
);

test("WSL mode: other UNC shares and SharePoint are unaddressable", () =>
  withEnv(WSL_ON, async () => {
    assert.equal(await resolveWorkspaceRoot("\\\\server\\share\\b.docx"), null);
    assert.equal(await resolveWorkspaceRoot("\\\\fileserver\\share\\spec.docx"), null);
    assert.equal(
      await resolveWorkspaceRoot("https://contoso.sharepoint.com/sites/x/spec.docx"),
      null,
    );
  }));

test(
  "WSL mode: percent signs stay literal in drive paths and decode in file:// URLs",
  posixOnly,
  () =>
    withEnv(WSL_ON, async () => {
      assert.equal(await resolveWorkspaceRoot("D:\\a\\100%\\b.docx"), "/mnt/d/a/100%");
      assert.equal(await resolveWorkspaceRoot("file:///D:/a/100%25/b.docx"), "/mnt/d/a/100%");
    }),
);

test("WSL mode: mount root is overridable", posixOnly, () =>
  withEnv({ CC_OFFICE_WSL_PATHS: "1", CC_OFFICE_WSL_MOUNT_ROOT: "/x/" }, async () => {
    assert.equal(await resolveWorkspaceRoot("D:\\a\\b.docx"), "/x/d/a");
  }),
);

test("toLocalPath (set_cwd / context input) translates only Windows-form paths", () =>
  withEnv(WSL_ON, async () => {
    assert.equal(toLocalPath("D:\\x"), "/mnt/d/x");
    assert.equal(toLocalPath("/home/a"), "/home/a");
    assert.equal(toLocalPath("notes/ref.md"), "notes/ref.md");
    assert.equal(toLocalPath("\\\\server\\share"), null);
  }));

test("resolveWorkspaceRoot returns null when the doc sits in an OS-managed $HOME child", async () => {
  if (process.platform !== "darwin") return; // deny-list is macOS-specific
  const fakeHome = await mkdtemp(join(tmpdir(), "cc-office-ws-home-"));
  const prevHome = process.env.HOME;
  const prevUserProfile = process.env.USERPROFILE;
  process.env.HOME = fakeHome;
  process.env.USERPROFILE = fakeHome;
  try {
    const lib = join(fakeHome, "Library");
    await mkdir(lib, { recursive: true });
    await writeFile(join(lib, "doc.docx"), "");
    // Re-import with the faked HOME (workspace.mjs reads homedir() at load).
    const mod = await import("../daemon/workspace.mjs?home=" + Date.now() + Math.random());
    assert.equal(await mod.resolveWorkspaceRoot(join(lib, "doc.docx")), null);
  } finally {
    if (prevHome === undefined) delete process.env.HOME;
    else process.env.HOME = prevHome;
    if (prevUserProfile === undefined) delete process.env.USERPROFILE;
    else process.env.USERPROFILE = prevUserProfile;
    await rm(fakeHome, { recursive: true, force: true });
  }
});

test("suggestWorkspaceRoot returns the doc folder with confidence 'doc'", async () => {
  const root = await makeTmp();
  try {
    const ws = join(root, "ws");
    await mkdir(ws, { recursive: true });
    await writeFile(join(ws, "doc.docx"), "");
    assert.deepEqual(await suggestWorkspaceRoot(join(ws, "doc.docx")), {
      cwd: ws,
      confidence: "doc",
    });
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("suggestWorkspaceRoot returns null for cloud URLs", async () => {
  assert.equal(
    await suggestWorkspaceRoot("https://contoso.sharepoint.com/sites/x/spec.docx"),
    null,
  );
});

test("ensureWorkspaceMarker drops a CLAUDE.md when none exists, then no-ops", async () => {
  const root = await makeTmp();
  try {
    const ws = join(root, "ws");
    await mkdir(ws);
    assert.equal(await ensureWorkspaceMarker(ws), true);
    assert.equal(await ensureWorkspaceMarker(ws), false);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("ensureWorkspaceMarker does nothing when .claude already exists", async () => {
  const root = await makeTmp();
  try {
    const ws = join(root, "ws");
    await mkdir(ws);
    await mkdir(join(ws, ".claude"));
    assert.equal(await ensureWorkspaceMarker(ws), false);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});
