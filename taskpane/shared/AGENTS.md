# taskpane/shared/

## OVERVIEW
Host-aware pane runtime: dispatcher, Office.js tool implementations, per-tab controllers and their pure helpers. Earned its file: score ~22 (26 files, ~16k LOC, `review-view.js` imported by panes, Word tools, daemon-side tests and e2e).

## WHERE TO LOOK
| Task | File | Notes |
|------|------|-------|
| Route a daemon tool call | `taskpane.js` (2338) | entry; `office_*` -> `tools-word.js`, `excel_*` -> `tools-excel.js` |
| Word operation (tracked edit, comments, snapshot) | `tools-word.js` (2879) | exports `PANE_CAPS`, `toolTrackedEdit`, `collectReviewSnapshot`, `documentSaved` |
| Excel operation | `tools-excel.js` (594) | A1 addresses, optionally sheet-qualified |
| Review tab | `review-pane.js` (2353) glue, `review-view.js` (1338) pure | largest controller; ~90 pure helpers in view |
| Response tab | `response-pane.js` / `response-view.js` | view duplicates a daemon mapping |
| Tools tab (read-only) | `tools-pane.js` / `tools-view.js` | daemon writes a separate copy |
| Provider and model picker | `providers-pane.js` / `providers-view.js` | model rows via `model-menu.js`, `model-line.js` |
| Citation fields (Zotero, EndNote) | `field-text.js` | parse, verdicts, safe splice, unread detection |
| Follow-through candidates | `follow-places.js`, `residue.js`, `named-terms.js`, `paragraph-label.js` | bounded search |
| OOXML run formatting | `format-runs.js` | uses `RUN_CHAR`, `xmlText` from review-view |
| Busy state across chat + Review | `turn-state.js` | `TURN_END_EVENTS`: `turn_complete`, `error`, `auth_error` |
| Queues | `propose-queue.js`, `result-queue.js`, `auto-refresh.js`, `go-to.js` | |
| Path helpers | `paths.js` | URL, Windows, POSIX, `/mnt/<drive>` |
| Shared styling | `styles.css` (1761) | `.msg.event` never hidden by diagnostics toggle (l.447) |

## CONVENTIONS
- Split rule: `*-pane.js` = DOM glue + bridge requests; `*-view.js` = pure functions (no DOM, no Office, no Node). Tests import views directly under Node, so keep views free of `window`, `document`, `Word`.
- Every Office.js call goes through `Word.run` / `Excel.run`.
- Review actions are serialized by `guard()` in `review-pane.js`; long jobs that outlive the bridge 10 s default are handled outside the ordinary guard.
- `createTurnState` checks turn identity on end; chat turns emit no `turn_done`, so never wait for one from Review.

## ANTI-PATTERNS
- Never write the Word document from the Tools tab.
- Never rewrite a paragraph holding fields as plain text; never edit inside or across a citation field. Fail closed when field text is unread.
- Never let cleanup alter user-written comments (`review-view.js`).
- In `follow-places.js`: a smaller change context must not hide an answer; checked answers must not use up the bounded batch before unseen ones.
- `propose-queue.js` refuses concurrent runs; do not bypass.
- No `require`, no `node:` imports here; CI `node --check` passes but the browser fails.

## TESTS
One test file per module, usually same name: `tests/<module>.test.mjs`, DOM glue in `tests/<tab>-pane-dom.test.mjs` / `review-pane-dom.test.mjs` (3229 lines, dynamic import of `review-pane.js`). Word-tool behavior: `tests/tracked-edit-word.test.mjs`, `tracked-edit-fields.test.mjs` (fake Word model with failure injection). `scripts/e2e/make_tracked_special.py` mirrors `review-view.js` `trackedRuns`; change both.
