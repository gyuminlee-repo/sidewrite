// scripts/check-personal.sh: no personal value in a tracked file. The generic
// shapes are built in; real words come from an untracked local file, so this
// test (and the repository) holds none of them. Every sample below is built
// from parts, so this file passes the check it tests.
import { test } from "node:test";
import assert from "node:assert/strict";
import { execFileSync, spawnSync } from "node:child_process";
import { mkdtempSync, writeFileSync, copyFileSync, mkdirSync, readFileSync, rmSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const SCRIPT = fileURLToPath(new URL("../scripts/check-personal.sh", import.meta.url));

function repo(files, { local = null, allow = "" } = {}) {
  const d = mkdtempSync(join(tmpdir(), "pers-"));
  try {
    execFileSync("git", ["init", "-q"], { cwd: d });
    mkdirSync(join(d, "scripts"));
    copyFileSync(SCRIPT, join(d, "scripts/check-personal.sh"));
    writeFileSync(join(d, "scripts/personal-allow.txt"), allow);
    for (const [n, t] of Object.entries(files)) writeFileSync(join(d, n), t);
    if (local) writeFileSync(join(d, ".personal-patterns.local"), local);
    execFileSync("git", ["add", "-A", "--", ".", ":!.personal-patterns.local"], { cwd: d });
    return spawnSync("bash", ["scripts/check-personal.sh"], { cwd: d, encoding: "utf8" });
  } finally {
    rmSync(d, { recursive: true, force: true });
  }
}

// The check is a bash script on git grep: these tests are not run on Windows.
const BASH = { skip: process.platform === "win32" ? "needs bash" : false };

const HOME = "/ho" + "me/";
const WIN = "C:\\Us" + "ers\\";
const SAMPLES = [
  `${HOME}someone/x`,
  `${WIN}someone\\doc`,
  "/mnt/d/_work" + "space/y",
  "write to a@" + "b.com",
  "ABCD-D" + "-26-01234",
];

test("generic shapes fail anywhere in the tracked tree", BASH, () => {
  for (const t of SAMPLES) {
    const r = repo({ "notes.md": t });
    assert.equal(r.status, 1, `${t}\n${r.stdout}${r.stderr}`);
    assert.match(r.stdout, /notes\.md:1:/);
  }
});

test("placeholders pass", BASH, () => {
  const r = repo({
    "a.md":
      `${HOME}you/x user@` +
      `example.org noreply@host ${HOME}user/ ${WIN}you\\ git@` +
      "github.com:owner/repo",
  });
  assert.equal(r.status, 0, r.stdout + r.stderr);
  assert.match(r.stdout, /personal check: 0 hits/);
});

test("a user name of one or two letters is a placeholder, a longer one is not", BASH, () => {
  const ok = repo({ "a.md": `${HOME}a/x ${HOME}me/y ${WIN}u\\doc icon@` + "2x.png" });
  assert.equal(ok.status, 0, ok.stdout);
  for (const t of [`${HOME}bob/x`, `${WIN}bob\\doc`, "icon@" + "2x.org"])
    assert.equal(repo({ "a.md": t }).status, 1, t);
});

test("a hit on the same line as a placeholder is still a hit", BASH, () => {
  const r = repo({ "a.md": `${HOME}you/x and ${HOME}someone/y` });
  assert.equal(r.status, 1, r.stdout);
});

test("local patterns are added when the file exists; without it the script says so", BASH, () => {
  const without = repo({ "a.md": "the zebrafish paper" });
  assert.equal(without.status, 0, without.stdout + without.stderr);
  assert.match(without.stdout, /no \.personal-patterns\.local: generic shapes only/);
  const withFile = repo(
    { "a.md": "the zebrafish paper" },
    { local: "# my words\n\nzebrafish paper\n" },
  );
  assert.equal(withFile.status, 1);
  assert.match(withFile.stdout, /a\.md:1:/);
  assert.doesNotMatch(withFile.stdout, /generic shapes only/);
});

test("an allow line lets one known line through, and only that line", BASH, () => {
  const files = {
    LICENSE: `Copyright (c) someone <a@` + `b.com>\n`,
    "b.md": "mail a@" + "b.com\n",
  };
  const allow = "# upstream copyright line\nLICENSE:1:\n";
  const r = repo(files, { allow });
  assert.equal(r.status, 1);
  assert.match(r.stdout, /b\.md:1:/);
  assert.doesNotMatch(r.stdout, /LICENSE/);
});

test("the pattern file itself is ignored by git", () => {
  assert.match(
    readFileSync(new URL("../.gitignore", import.meta.url), "utf8"),
    /^\.personal-patterns\.local$/m,
  );
});

test("built-in profiles carry no response colour", async () => {
  for (const n of ["generic", "example-elsevier"]) {
    const p = JSON.parse(
      await readFile(new URL(`../profiles/journal_profiles/${n}.json`, import.meta.url), "utf8"),
    );
    for (const k of ["response_body", "response_revised", "response_kr"])
      assert.equal(p.colors?.[k] ?? null, null, `${n} ${k}`);
  }
});
