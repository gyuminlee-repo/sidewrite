// A Windows program started from WSL can fail before it runs: the interop
// layer waits for the Windows side to connect back and gives up, printing
//   <3>WSL (<pid> - ) ERROR: UtilAcceptVsock:281: accept4 failed 110
// (110 is ETIMEDOUT) on stderr, with nothing on stdout. Seen under heavy load
// on the machine. Such a call is tried once more, after a short wait.
//
// Only that failure is retried, and only when the program printed nothing on
// stdout (the pane drivers and the PowerShell scripts always print a JSON
// line, so empty stdout means the program did not run, and a step that
// presses a button was not pressed). A timeout of our own (the exec killed
// the program) is never retried.

const text = (v) => (v == null ? "" : String(v));

export function isWslInteropLaunchFailure(err) {
  if (!err || typeof err !== "object") return false;
  if (err.code === "ETIMEDOUT" || err.signal) return false;
  if (text(err.stdout).trim()) return false;
  const said = `${text(err.stderr)}\n${text(err.message)}`;
  return said.includes("UtilAcceptVsock") && said.includes("accept4 failed");
}

const sleepSync = (ms) => Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);

// exec(): the call (synchronous). On the interop failure it is called once
// more after waitMs; any other failure, or a second one, is thrown as it is.
export function execWithInteropRetry(
  exec,
  { what = "a Windows program", waitMs = 3000, sleep = sleepSync, log = console.log } = {},
) {
  try {
    return exec();
  } catch (e) {
    if (!isWslInteropLaunchFailure(e)) throw e;
    log(
      `NOTE  WSL could not start ${what} (UtilAcceptVsock accept4 failed); trying once more in ${waitMs / 1000} s`,
    );
    sleep(waitMs);
    return exec();
  }
}
