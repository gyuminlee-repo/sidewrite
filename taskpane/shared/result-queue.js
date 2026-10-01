// Tool results produced while the WebSocket to the daemon is down.
//
// A tool call that Word finished after the connection dropped used to be
// thrown away with its result, so the daemon never learned that the edit is
// in the document. The results wait here and go out right after the hello of
// the next connection; the bridge hands a result for a call it gave up on to
// its late-result handler by call id. Bounded and with a time limit: an old
// result is not sent.

export function createResultQueue({ max = 20, ttlMs = 10 * 60_000, now = Date.now } = {}) {
  let items = [];
  return {
    push(obj) {
      items.push({ obj, at: now() });
      if (items.length > max) items = items.slice(-max);
    },
    // The results still in time, oldest first; the queue is empty afterwards.
    drain() {
      const t = now();
      const out = items.filter((x) => t - x.at <= ttlMs).map((x) => x.obj);
      items = [];
      return out;
    },
  };
}
