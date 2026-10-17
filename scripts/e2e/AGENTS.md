# scripts/e2e/

## OVERVIEW
Hands-free Word acceptance runner and replay, driven from WSL into Windows Word + the real pane over CDP. Earned its file: score ~24 (33 files, ~15k LOC, three largest files in the repo, imported by 12+ tests). Full procedure: `README.md` (1252 lines; read by section).

## WHERE TO LOOK
| Task | File | Notes |
|------|------|-------|
| Run an acceptance stage | `run-acceptance.mjs` (3175) | stages `s2` (default), `s25`, `s3`, `s4`, `cite`, `parts`, `resident` |
| Judge a row / evidence | `judge.mjs` (3364) | 80+ pure named exports; imports 6 daemon modules |
| Pane UI steps | `pane-driver.cjs` (2051) | one step per process, one JSON line out; `replay-driver.cjs` extends `steps` |
| Replay a real review | `replay.mjs` (862), `replay-lib.mjs` (751) | lib = pure parsing, turns, reports |
| Compare replay docx | `replay-compare.py` | `--before --after --tool --out` |
| Word COM ownership | `word-own.ps1`, `word-com.ps1`, `word-rot.ps1`, `word-launch.ps1`, `word-back.ps1` | |
| WSL interop retry | `wsl-interop.mjs` | |
| Pause resident daemon | `resident-unit.mjs` | uses `scripts/install/unit.mjs` |
| Fixture builders | `make_*_fixture.py`, `make_tracked_special.py`, `embed_taskpane.py` | argparse, `--help` |
| Marked-diff pairs in Word | `marked-diff-word-*.mjs`, `marked-diff-word.ps1`, `marked_diff_pairs.py` | `PAIRS` is shared data |
| Equation probe | `word-math-probe*.{mjs,ps1}` | creates unique `wmp-*` folder |

## CONVENTIONS
- Role suffixes: `*-lib` pure, `*-driver` OS-facing step runner, `*-probe`, `*-judge`. Keep decisions in libs/judge so tests run without Word.
- Drivers are CommonJS `.cjs` (Windows Node, `require`); everything else ESM.
- Every run gets an isolated run folder on a Windows drive; scripts are copied there before execution. Reusing `--run-dir` deletes stale markers/evidence first.
- Word: dedicated `/n` process, run-folder anchor, full-path ROT lookup, containment check against the run folder.
- Evidence is saved as JSON; a verdict comes from observed pane text, Word COM state, daemon log and verify-command exit codes.
- Python writers refuse `--out` equal to `--src`.

## ANTI-PATTERNS
- Never automate a pre-existing Word or close a document not proven to be under the run folder.
- Never force-kill (`Stop-Process`, `.Kill()`) when ownership or save state is unknown; use `CloseMainWindow`.
- Never minimize or foreground the run Word window; send it behind.
- Never run Windows scripts from a UNC path.
- Never judge a Word document that was never saved; never count an ambiguous undo as applied.
- Retry only the WSL `UtilAcceptVsock` launch failure with empty stdout, once; never timeouts or calls with output.
- Judge must not default to Claude when another provider is named; two panes must not share a session.
- `replay-open.ps1` never writes its input; repair goes to `-RepairTo`.
- `term-family-count.mjs` prints counts only, never manuscript text.
- Screenshots are opt-in and must never affect a verdict.

## TESTS
`tests/e2e-*.test.mjs`, `replay-lib.test.mjs` (2033), `marked-diff-word-check.test.mjs`, `word-math-probe.test.mjs`, `test_replay_compare.py`, `test_make_*.py`, `test_embed_taskpane.py`, `word-cleanup.test.ps1` (parses PowerShell source, runs against fakes; source-level only). A daemon rename breaks `judge.mjs` silently until these run.
