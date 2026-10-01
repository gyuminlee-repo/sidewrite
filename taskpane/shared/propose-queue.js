// "Propose all" (P15): review turns for a list of comments, one after
// another. proposeOne(ref) resolves when the turn of that comment has ended,
// with { ok: true } or { ok: false, error }. The first failure ends the run:
// a usage limit or a lost connection would fail every comment after it too.
//
// run(refs) resolves with { done } and, when it ended early, either
// { failed, error } or { stopped: true, next }. stop() lets the running turn
// finish and starts no other.

export function createProposeQueue({ proposeOne, onProgress = () => {} }) {
  let stopRequested = false;
  let running = false;
  return {
    get running() {
      return running;
    },
    stop() {
      stopRequested = true;
    },
    async run(refs) {
      if (running) throw new Error("Propose all is already running");
      running = true;
      stopRequested = false;
      const done = [];
      try {
        for (let i = 0; i < refs.length; i++) {
          if (stopRequested) return { done, stopped: true, next: refs[i] };
          onProgress({ index: i, total: refs.length, ref: refs[i] });
          let r;
          try {
            r = await proposeOne(refs[i]);
          } catch (e) {
            return { done, failed: refs[i], error: e?.message ?? String(e) };
          }
          done.push({ ref: refs[i], ...r });
          if (!r.ok) return { done, failed: refs[i], error: r.error };
        }
        return { done };
      } finally {
        running = false;
      }
    },
  };
}
