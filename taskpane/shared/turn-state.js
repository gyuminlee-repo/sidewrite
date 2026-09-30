// Whether a turn of this pane is running (a chat message or a Review-tab
// turn). taskpane.js starts and ends every turn here, so the composer and
// the Review tab buttons change together; the Review tab does not wait for
// its own turn_done, which a chat turn never sends.

// assistant_event events that end the running turn.
export const TURN_END_EVENTS = new Set(["turn_complete", "error", "auth_error"]);

// getReviewPane: the Review tab controller (initReviewPane) or null.
export function createTurnState({
  setComposerDisabled = () => {},
  getReviewPane = () => null,
} = {}) {
  let inFlight = false;
  const set = (v) => {
    inFlight = v;
    setComposerDisabled(v);
    getReviewPane()?.setTurnBusy?.(v);
  };
  return {
    get inFlight() {
      return inFlight;
    },
    begin: () => set(true),
    end: () => set(false),
    onAssistantEvent(msg) {
      if (TURN_END_EVENTS.has(msg?.event)) set(false);
    },
  };
}
