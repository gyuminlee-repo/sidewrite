// Invented comment labels and file names only. No Word runner is executed.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { runInNewContext } from "node:vm";
import {
  keepRefPattern,
  turnFromLog,
  rewritesFromLog,
  followFromLog,
  refusalsFromLog,
} from "../scripts/e2e/replay-lib.mjs";

for (const ref of [
  "R1.2.3",
  "R1.2+",
  "R1.2[bridge]",
  "R1.2(bridge)",
  "R1.2|R9.9",
  "R1.2\\draft",
  "R1.2^$?*{}",
]) {
  test(`replay patterns match only the literal reference ${ref}`, () => {
    const pattern = new RegExp(keepRefPattern([ref]));
    assert.equal(pattern.test(ref), true);
    assert.equal(pattern.test(ref.replaceAll(".", "x")), false);
    assert.deepEqual(turnFromLog(`[review] turn done ${ref}`, ref), {
      model: null,
      ended: true,
      ok: true,
      error: null,
    });
    assert.deepEqual(
      rewritesFromLog(`[review] rewrite ${ref} paragraph=3 parts=2 applied=1`, ref),
      [{ paragraph: 3, parts: 2, applied: 1, stopped: null }],
    );
    assert.deepEqual(
      followFromLog(`[review] follow ${ref} candidates=2 changes=1 questions=3`, ref),
      {
        candidates: 2,
        changes: 1,
        questions: 3,
      },
    );
    assert.deepEqual(refusalsFromLog(`[review] blocked ${ref} by=L1`, ref), { ForbidBlocked: 1 });
    const other = ref.replace(/\.(?=[^.]*$)/, "x");
    assert.equal(pattern.test(other), false);
    assert.equal(turnFromLog(`[review] turn done ${other}`, ref).ended, false);
    assert.deepEqual(
      rewritesFromLog(`[review] rewrite ${other} paragraph=3 parts=2 applied=1`, ref),
      [],
    );
    assert.equal(
      followFromLog(`[review] follow ${other} candidates=2 changes=1 questions=3`, ref),
      null,
    );
    assert.deepEqual(refusalsFromLog(`[review] blocked ${other} by=L1`, ref), {});
  });
}

test("runner and pane patterns escape complete references and file names without executing their entry points", () => {
  for (const file of ["run-acceptance.mjs", "judge.mjs", "pane-driver.cjs"]) {
    const source = readFileSync(new URL(`../scripts/e2e/${file}`, import.meta.url), "utf8");
    assert.doesNotMatch(source, /\.replace\(["']\.["'],\s*["']\\\\\.["']\)/, file);
    const helper = source.match(/^const escapeRe = .*;$/m)?.[0];
    assert.ok(helper, `${file} keeps a reusable literal escaper`);
    const escape = runInNewContext(`${helper}\nescapeRe`);
    for (const text of ["R1.2.3", "Library.notes[2](draft)+.docx", "R1.2|R9.9"]) {
      const re = new RegExp(`^${escape(text)}$`);
      assert.equal(re.test(text), true, file);
      assert.equal(re.test(text.replace(/\.(?=[^.]*$)/, "x")), false, file);
    }
  }
});
