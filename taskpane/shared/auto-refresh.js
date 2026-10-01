// When the Review tab refreshes its status by itself (P15): after any turn
// of the pane ends and when the pane gets the focus back. One refresh runs at
// a time.
//
// trigger(reason) answers what it did:
//   "busy"     isBusy() is true (a turn or a Propose all run, or no comments)
//   "started"  a refresh was started
//   "queued"   one is running and reason is "turn": one more runs after it,
//              once, however many turn ends come meanwhile
//   "running"  one is running and reason is not "turn": that one covers it
//   "too_soon" the last one started less than minIntervalMs ago (a turn end
//              is never too soon: the document has just changed)
// A run that fails is reported to onError and does not stop later ones.

export function createAutoRefresh({
  run,
  isBusy = () => false,
  minIntervalMs = 10_000,
  now = Date.now,
  onError = () => {},
}) {
  let running = false;
  let queued = false;
  let lastStart = -Infinity;
  function start() {
    running = true;
    lastStart = now();
    // run() is called here and now; a throw becomes a rejection.
    new Promise((resolve) => resolve(run())).then(
      () => ended(),
      (e) => {
        onError(e);
        ended();
      },
    );
  }
  function ended() {
    running = false;
    if (!queued) return;
    queued = false;
    if (!isBusy()) start();
  }
  return {
    trigger(reason) {
      if (isBusy()) return "busy";
      if (running) {
        if (reason !== "turn") return "running";
        queued = true;
        return "queued";
      }
      if (reason !== "turn" && now() - lastStart < minIntervalMs) return "too_soon";
      start();
      return "started";
    },
  };
}
