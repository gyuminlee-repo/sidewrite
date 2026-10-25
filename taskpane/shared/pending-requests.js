// Requests of the task pane waiting for their "_result" answer (taskpane.js
// sendRequest).
//
// A request waits until its answer comes, its own time runs out, or the
// connection closes. The time can be long (review file work on WSL drvfs
// waits up to 120 s, review-pane.js REVIEW_FILE_WAIT_MS), but a closed
// connection never brings an answer: every request still waiting is
// rejected at once then, so a button does not wait out the whole time with
// no error. Tables of other waits no request covers (the jobs of a pane,
// ended by a job_done event) listen for the same close through onClose.

export function createPendingRequests({ setTimer = setTimeout, clearTimer = clearTimeout } = {}) {
  const waiting = new Map();
  const closeListeners = [];
  return {
    // A promise of the answer to request `id`, rejected with
    // `Request "<type>" timed out` after timeoutMs.
    add(id, type, timeoutMs) {
      return new Promise((resolve, reject) => {
        const timer = setTimer(() => {
          if (!waiting.has(id)) return;
          waiting.delete(id);
          reject(new Error(`Request "${type}" timed out`));
        }, timeoutMs);
        waiting.set(id, { resolve, reject, timer });
      });
    },
    // The answer came: true when a request was waiting for it.
    resolve(id, msg) {
      const w = waiting.get(id);
      if (!w) return false;
      waiting.delete(id);
      clearTimer(w.timer);
      w.resolve(msg);
      return true;
    },
    // The connection closed: every waiting request ends with `error`, then
    // every close listener is called with it. Returns how many requests
    // were waiting.
    rejectAll(error) {
      const all = [...waiting.values()];
      waiting.clear();
      for (const w of all) {
        clearTimer(w.timer);
        w.reject(error);
      }
      for (const fn of closeListeners) {
        try {
          fn(error);
        } catch (e) {
          console.warn("[pending] close listener failed:", e);
        }
      }
      return all.length;
    },
    // fn(error) is called on every close (rejectAll).
    onClose: (fn) => {
      closeListeners.push(fn);
    },
    get size() {
      return waiting.size;
    },
  };
}
