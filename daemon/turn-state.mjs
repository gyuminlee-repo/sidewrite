// Which turn of a pane is running, and who runs it. A pane has one turn at a
// time. Its owner is the pane's Claude session (a chat message, or a review
// turn Claude runs with the office tools) or the proposal runner (a review
// turn of a provider that proposes, proposal-turns.mjs).
//
// A turn is ended by its owner only. The Claude session ends its turn from
// many places (a result, a Stop, a stream that ended, a restart for a new
// model, context or workspace); none of them may end a proposal turn, which
// does not run in that session: the task pane would be told that the review
// turn is over (turn_done) while the runner still writes to Word.
//
// A turn holds a review turn ({ ref, ... }) or none (a chat message). The
// review turn is shown to its owner only, so the office tools of the Claude
// session never take a proposal turn for their own.

export const CLAUDE = "claude";
export const PROPOSAL = "proposal";

// onReviewTurnEnded(key, turn, outcome): a review turn ended; outcome is
// { ok: true } or { ok: false, error } (review-service.mjs turnEnded).
export function createTurnState({ onReviewTurnEnded = () => {} } = {}) {
  const turns = new Map(); // pane key -> { owner, turn }
  return {
    busy: (key) => turns.has(key),
    owner: (key) => turns.get(key)?.owner ?? null,
    // A turn begins. A pane that has one keeps it: false, nothing changes.
    begin(key, owner, turn = null) {
      if (turns.has(key)) return false;
      turns.set(key, { owner, turn });
      return true;
    },
    reviewTurn(key, owner) {
      const t = turns.get(key);
      return t && t.owner === owner ? t.turn : null;
    },
    // The turn of `owner` is over. false when the pane has no turn or the
    // turn is another owner's: then nothing is ended and nobody is told.
    end(key, owner, outcome = { ok: true }) {
      const t = turns.get(key);
      if (!t || t.owner !== owner) return false;
      turns.delete(key);
      if (t.turn) onReviewTurnEnded(key, t.turn, outcome);
      return true;
    },
  };
}
