import { test } from "node:test";
import assert from "node:assert/strict";
import { execWithInteropRetry, isWslInteropLaunchFailure } from "../scripts/e2e/wsl-interop.mjs";

const VSOCK = "<3>WSL (708786 - ) ERROR: UtilAcceptVsock:281: accept4 failed 110\n";
// An error as execFileSync throws it: exit status 1, stderr in the message.
function failed(stderr, extra = {}) {
  return Object.assign(new Error(`Command failed: node.exe pane-driver.cjs propose\n${stderr}`), {
    status: 1,
    signal: null,
    stdout: "",
    stderr,
    ...extra,
  });
}

test("the WSL interop launch failure is retried", () => {
  assert.equal(isWslInteropLaunchFailure(failed(VSOCK)), true);
  // Only in the message (the runner kept e.message alone).
  assert.equal(
    isWslInteropLaunchFailure(failed("", { message: `Command failed: x\n${VSOCK}` })),
    true,
  );
});

test("other failures are not retried", () => {
  assert.equal(isWslInteropLaunchFailure(failed("Exception: Word is busy\n")), false);
  assert.equal(isWslInteropLaunchFailure(failed("")), false);
  assert.equal(isWslInteropLaunchFailure(new Error("spawnSync powershell.exe ENOENT")), false);
  assert.equal(isWslInteropLaunchFailure(null), false);
  // Half of the text is not the failure.
  assert.equal(isWslInteropLaunchFailure(failed("UtilAcceptVsock:281: something else\n")), false);
});

test("a timeout of our own is not retried, even with the interop text", () => {
  assert.equal(
    isWslInteropLaunchFailure(failed(VSOCK, { code: "ETIMEDOUT", signal: "SIGTERM" })),
    false,
  );
  assert.equal(isWslInteropLaunchFailure(failed(VSOCK, { signal: "SIGKILL" })), false);
});

test("a program that printed on stdout ran, and is not retried", () => {
  assert.equal(isWslInteropLaunchFailure(failed(VSOCK, { stdout: '{"ok":false}\n' })), false);
});

function fakeExec(...outcomes) {
  const calls = [];
  const exec = () => {
    const o = outcomes[calls.length];
    calls.push(o);
    if (o instanceof Error) throw o;
    return o;
  };
  return { exec, calls };
}
const recorder = () => {
  const lines = [];
  const sleeps = [];
  return { lines, sleeps, log: (l) => lines.push(l), sleep: (ms) => sleeps.push(ms) };
};

test("fails once with the interop text, then succeeds: the answer, with a NOTE", () => {
  const { exec, calls } = fakeExec(failed(VSOCK), "answer\n");
  const r = recorder();
  assert.equal(execWithInteropRetry(exec, { what: "pane-driver propose", ...r }), "answer\n");
  assert.equal(calls.length, 2);
  assert.deepEqual(r.sleeps, [3000]);
  assert.equal(r.lines.length, 1);
  assert.match(r.lines[0], /^NOTE {2}WSL could not start pane-driver propose .*once more/);
});

test("fails twice with the interop text: the second failure is thrown", () => {
  const second = failed(VSOCK);
  const { exec, calls } = fakeExec(failed(VSOCK), second);
  const r = recorder();
  assert.throws(
    () => execWithInteropRetry(exec, r),
    (e) => e === second,
  );
  assert.equal(calls.length, 2);
  assert.equal(r.sleeps.length, 1);
});

test("any other failure: thrown at once, no second call", () => {
  for (const err of [
    failed("Exception: Word is busy\n"),
    failed(VSOCK, { code: "ETIMEDOUT", signal: "SIGTERM" }),
    failed(VSOCK, { stdout: '{"ok":true}\n' }),
  ]) {
    const { exec, calls } = fakeExec(err, "never");
    const r = recorder();
    assert.throws(
      () => execWithInteropRetry(exec, r),
      (e) => e === err,
    );
    assert.equal(calls.length, 1);
    assert.deepEqual(r.sleeps, []);
    assert.deepEqual(r.lines, []);
  }
});

test("the default wait sleeps", () => {
  const { exec } = fakeExec(failed(VSOCK), "ok");
  const t0 = Date.now();
  assert.equal(execWithInteropRetry(exec, { waitMs: 50, log: () => {} }), "ok");
  assert.ok(Date.now() - t0 >= 45);
});
