# PROJECT KNOWLEDGE BASE

**Generated:** 2026-10-17
**Commit:** 8eadc37
**Branch:** main

## OVERVIEW
Sidewrite: Word + Excel Office.js add-ins over a local Claude Agent SDK daemon (Node ESM, Electron tray, Python CLI tools). Architecture, branding, auth, file layout: see `CLAUDE.md`. Not repeated here.

## STRUCTURE
```
sidewrite/
├── daemon/      # Node ESM brain: SDK loop, services, revision state, providers  -> daemon/AGENTS.md
├── tools/       # Python CLIs run by the daemon (JSON last line + exit code)   -> tools/AGENTS.md
├── taskpane/    # browser runtime: host pages, figures page, CSP              -> taskpane/AGENTS.md
│   └── shared/  # dispatcher, Office.js tools, *-pane glue / *-view pure     -> taskpane/shared/AGENTS.md
├── app/         # Electron tray, daemon lifecycle, sideload (no child file; invariants below)
├── profiles/    # journal profiles; lookup + overrides in profiles/README.md
├── scripts/     # check-personal, pdf-line-map, fixture makers               -> scripts/AGENTS.md
│   ├── e2e/     # WSL -> Windows Word acceptance runner                      -> scripts/e2e/AGENTS.md
│   └── install/ # systemd user unit, HKCU registration, doctor              -> scripts/install/AGENTS.md
├── tests/       # node:test + unittest + one .ps1                            -> tests/AGENTS.md
└── manifests/   # word.xml, excel.xml
```

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| Daemon entry, sessions, permission guard | `daemon/index.mjs` | `npm run dev` |
| Tray app, restart cap, IPC | `app/main.mjs` | `npm start`; package.json main |
| Add-in install/uninstall | `app/sideload.mjs`, `app/legacy-manifests.mjs` | hash-based, see NOTES |
| Pane tool dispatch | `taskpane/shared/taskpane.js` | host via `Office.context.host` |
| Spawn a Python tool | `daemon/py-tool.mjs` (runPy) | `CC_OFFICE_PYTHON` override |
| Repo root for tools | `daemon/doc-tools.mjs` REPO_ROOT | |
| Journal profile lookup | `daemon/journal-profile.mjs` | only reader of `profiles/journal_profiles/*.json` |
| LLM provider routing | `daemon/providers/` | section in daemon/AGENTS.md |
| PDF line numbers | `daemon/response-lines.mjs` -> `scripts/pdf-line-map.py` | |
| Word acceptance | `scripts/e2e/run-acceptance.mjs` | stages s2 s25 s3 s4 cite parts |
| Resident daemon install | `scripts/install/install.sh` | |
| CI | `.github/workflows/ci.yml` | |

## CODE MAP
| Symbol | Type | Location | Refs | Role |
|--------|------|----------|------|------|
| loadRevision | fn | `daemon/revision-store.mjs` | 27 files | revision.json state, widest fan-in |
| runPy | fn | `daemon/py-tool.mjs` | daemon services | spawns `python3 tools/<name>.py` |
| REPO_ROOT | const | `daemon/doc-tools.mjs` | daemon | tool path resolution |
| ledger-check, response-quotes, edit-facts | modules | `daemon/` | wide | change with care |
| *-view.js | pure modules | `taskpane/shared/` | Node tests, `scripts/e2e/judge.mjs` | testable pane logic |
| judge.mjs | script | `scripts/e2e/` | imports 6 daemon modules | daemon renames break e2e silently |

Largest files: `tools/build_marked_diff.py` ~2830, `daemon/review-service.mjs` ~2000, `daemon/index.mjs` ~1780, `daemon/revision-store.mjs` ~1690 lines.

## CONVENTIONS
- Ports: WS 47833, HTTP 47834 (`daemon/index.mjs` l.71-72; WS = HTTP - 1). WS literal also in `taskpane.js` and host-page CSP; change all together.
- Daemon serves `taskpane/` statically (under-root check) and rewrites `/shared/taskpane.js|styles.css` with `?v=<bridge token>` for cache busting.
- Python tools print one JSON line last. Exit 0 ok, 1 refused / check failed, 2 bad input, 3 optional dep missing.
- Per-paper values go in `profile_overrides` (revision.json / submission.json), never in a profile file.
- Prettier: printWidth 100, double quotes, trailingComma all. No lint script.
- Env vars beyond the CLAUDE.md list: `CC_OFFICE_PYTHON`, `CC_OFFICE_CLAUDE_PATH`, `CC_OFFICE_CODEX_BIN`, `CC_OFFICE_AGY_BIN`, `CC_OFFICE_KEY_<ID>`, `CC_OFFICE_PROFILE`, `CC_OFFICE_PROFILE_DIRS`, `CC_OFFICE_CONFIG_DIR`, `CC_OFFICE_RATIONALE_LANG`, `CC_OFFICE_RESPONSE_MODEL`, `CC_OFFICE_MODELS_TIMEOUT_MS`, `CC_OFFICE_WSL_PATHS`, `CC_OFFICE_WSL_MOUNT_ROOT`, `CC_OFFICE_REVIEW_PROMPT_FORBID`.

## ANTI-PATTERNS (THIS PROJECT)
- No fallback between LLM providers. Only Claude runs review turns (office tools). API keys never reach the pane or logs.
- Prompt and refusal strings in the daemon are model-facing behavior, not copy; edit as code.
- Pane code: no `require`, no `node:` imports, no daemon imports (`node --check` passes, browser fails).
- No innerHTML with daemon or document text; never drop the pane CSP.
- Never automate or force-kill a Word process the run does not own.
- Never remove legacy services, tasks or manifests by name; full-content match only.
- No hardcoded home / Windows-user paths or real names in tracked files (`scripts/check-personal.sh`).
- No fixed sleeps or poll loops in tests (existing debt listed in tests/AGENTS.md).
- Mirrored logic, change both sides: anonymize categories (`tools/anonymize_docx.py` + `scripts/verify-anonymized.py`); citation field rules (`tools/citation_fields_check.py` + `taskpane/shared/field-text.js`).

## UNIQUE STYLES
- `*-pane.js` = Office.js glue; `*-view.js` = pure helpers shared with Node tests and e2e judge.
- Optional Python deps (lxml, python-docx, PyMuPDF/fitz, citeproc-py) imported lazily; tests skip when absent. No manifest declares them (unverified beyond grep).

## COMMANDS
```bash
npm start                         # Electron tray (daemon + UI)
npm run dev                       # daemon only
npm test                          # node --test tests/*.test.mjs (Node only)
python3 -m unittest tests/test_<name>.py   # Python tests, not in npm test or CI
npm run format:check              # prettier
bash scripts/check-personal.sh    # leak gate: 0 clean, 1 hits, 2 outside git
python3 tools/response_verify.py --selftest
python3 tools/figure_render.py --probe
python3 tools/csl_render.py --check
node scripts/e2e/run-acceptance.mjs --stage s2   # Windows Word via WSL
```
CI: `node --check` on every .js/.mjs under daemon app taskpane scripts tests (Node 20.x, 22.x), `npm test`, `npm audit --audit-level=high`, `npm run format:check`. Wider than the 3-file `node --check` in CLAUDE.md.

## NOTES
- On WSL, "the changed PowerShell scripts parse" (`tests/replay-lib.test.mjs`) fails 1 of 2654; `npm test` otherwise passes.
- CI installs no optional Python packages, so dependent tests never run in CI. `tests/word-cleanup.test.ps1` is also outside CI.
- `app/main.mjs`: MAX_RESTART 3; counter resets only after STABLE_AFTER_MS 45000 uptime, never on first "running". Daemon IPC on fd 3; pickers via `pick_path` / `pick_path_result`. Windows tray wants `tray-icon-win.ico` (absent today, PNG fallback active).
- `app/sideload.mjs`: install/uninstall compare content hash, never filename/GUID. Custom manifests raise, never silently removed. Legacy Draftspect manifests removed only on hash match and only after the new copy and registration succeed. Never follow symlinked or redirected dirs during cleanup.
- Profiles lookup order: manuscript folder, `CC_OFFICE_PROFILE_DIRS`, `CC_OFFICE_CONFIG_DIR`, repo `profiles/`.
