// taskpane/shared/paths.js: where the open document is, against the workspace
// folder the daemon reports. The daemon runs under WSL and writes a Windows
// drive as /mnt/<drive>; Word reports the same folder as D:\...
import { test } from "node:test";
import assert from "node:assert/strict";
import { docDirFromActiveUrl, isInOrUnder } from "../taskpane/shared/paths.js";

test("a Windows document folder and the WSL workspace folder of it are one place", () => {
  const url = "D:\\_workspace\\cc\\draftspect-s2test\\e2e-20261002-004055\\review_manuscript.docx";
  const dir = docDirFromActiveUrl(url);
  assert.equal(dir, "D:/_workspace/cc/draftspect-s2test/e2e-20261002-004055");
  assert.equal(
    isInOrUnder(dir, "/mnt/d/_workspace/cc/draftspect-s2test/e2e-20261002-004055"),
    true,
  );
  assert.equal(isInOrUnder(dir, "/mnt/d/_workspace/cc/draftspect-s2test"), true);
  assert.equal(isInOrUnder(dir, "/mnt/d/_workspace"), true);
  assert.equal(isInOrUnder(dir, "/mnt/d"), true);
});

test("a folder that is not the document's is still a mismatch", () => {
  const dir = docDirFromActiveUrl("D:\\work\\paper\\m.docx");
  // another drive, a sibling folder, a folder whose name only starts the same
  assert.equal(isInOrUnder(dir, "/mnt/e/work/paper"), false);
  assert.equal(isInOrUnder(dir, "/mnt/d/work/other"), false);
  assert.equal(isInOrUnder(dir, "/mnt/d/work/pap"), false);
  assert.equal(isInOrUnder(dir, "/mnt/d/work/paper/sub"), false, "the document is above it");
  assert.equal(isInOrUnder(dir, "/home/me/work/paper"), false);
});

test("paths that did not meet WSL keep their old behaviour", () => {
  assert.equal(isInOrUnder("/Users/me/paper", "/Users/me"), true);
  assert.equal(isInOrUnder("/Users/me/paper", "/Users/m"), false);
  assert.equal(isInOrUnder("C:\\Users\\me\\paper", "c:/users/me"), true);
  assert.equal(isInOrUnder("/mnt/data/x", "/mnt/data"), true);
  assert.equal(isInOrUnder("/mnt/data/x", "/mnt/d"), false, "/mnt/data is not drive d");
});
