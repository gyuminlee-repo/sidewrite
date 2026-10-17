# daemon/

## OVERVIEW
Node ESM daemon: Agent SDK sessions, pane bridge, Review/Response/figure/aux services, LLM provider routing. Earned its file: score 12 (70 files, high code ratio, symbol density, fan-in) plus distinct domain.

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| Permission guard (Office file writes) | `index.mjs` `canUseTool` path, `OFFICE_FILE_EXT`, `REDIR_TO_OFFICE` | read-only .docx inspection stays allowed |
| SDK options, env flags | `query-options.mjs` | `buildQueryOptions`, `resolveSettingSources` |
| Review tab | `review-service.mjs` (factory), `review-tools.mjs` (tracked edit/rewrite/ask-author), `review-prompt.mjs`, `review-split.mjs` | |
| Revision state on disk | `revision-store.mjs` | zod schemas, `<stem>.revision.json`, `REVISION_VERSION` |
| Change state of tracked edits | `revision-classify.mjs`, `revision-parts.mjs` (main vs supplement parts) | |
| Response letter | `response-service.mjs`, `response-quotes.mjs`, `response-check.mjs`, `response-lines.mjs`, `response-model.mjs`, `response-prompt.mjs` | |
| Ledger checks (regex, llm, count) | `ledger-check.mjs`, `count-check.mjs`, `ledger-suggest.mjs` | 8 importers; stable surface |
| Submission checklist, guidelines | `submission-check.mjs`, `guideline-extract.mjs`, `format-check.mjs` | |
| Aux jobs (anonymize, marked diff, translate, references) | `aux-service.mjs`, `anonymize.mjs`, `marked-diff.mjs`, `deepl.mjs`, `crossref.mjs`, `csl-style.mjs`, `languagetool.mjs` | |
| Figure lane | `figure-service.mjs`, `figure-http.mjs` (`FIGURES_PREFIX`) | |
| Python tool calls | `py-tool.mjs` `runPy`, `doc-tools.mjs` `pyJson`/`REPO_ROOT` | scripts under `../tools/` |
| One-shot judgments, provider choice | `oneshot.mjs`, `providers/` | see Providers below |
| Journal profile lookup | `journal-profile.mjs` | only reader of `../profiles/`; rules in `profiles/README.md` |
| Fetch deadlines | `fetch-timeout.mjs` | every outbound fetch uses `deadline` |

## PROVIDERS (`providers/`, covered here, score 6)
- `registry.mjs` PRESETS in tiers: subscription (claude, codex, agy), local (Ollama), key (API key, off by default).
- Only Claude gets role `turn`: review turns need the office tools, so no other provider edits the manuscript.
- `selection.mjs` per-pane role state restoring/ready/refused; `ready()` waits up to 45 s, then throws.
- `oneshot-router.mjs` `pin()` resolves the choice once per job; every batch of that job uses the same provider and model.
- NO fallback to another provider anywhere: an unresolved choice is an error and nothing is sent.
- `keys.mjs`: env `CC_OFFICE_KEY_<ID>` wins over `~/.claude/office-addins/provider-keys.json` (mode 0600). Keys never reach the pane or logs; pane gets `publicView()`.
- `probe.mjs` decides availability; a refusal carries the probe row reason.

## CONVENTIONS
- Relative imports with explicit `.mjs`, named exports only.
- Services are `createXService({ ...deps })` factories; inject `runPy` and clocks for tests. Handlers go through `handler-wrap.mjs` `wrapHandlers`.
- Failures become error replies, never silent defaults. `runPy` never throws; read `{ ok, code, json, error }`.
- Tools print one JSON line last; `runPy` parses the last stdout line. Python binary from `CC_OFFICE_PYTHON`, default `python3`.
- State files: temp file in the same folder, then rename; writes serialized per file (`revision-store.mjs` header).
- Counts (words, lines) are computed in code, never asked of the model.
- Prompt rules are string arrays; `POLICY_PACKS` refer to rules by index, so append new rules, never insert.
- Model-facing refusals start with `Not applied:`; a forbid rule that fails to compile blocks the edit.
- `locateParagraph` returns index null on ambiguous identical paragraphs; never guess.
- Workspace = folder that directly holds the document (`workspace.mjs`); never auto-pick `$HOME`, an OS-managed `$HOME` child, a filesystem root, a WSL drive root or Windows profile folder.

## ANTI-PATTERNS
- Adding a route string in `index.mjs` for figures (`tests/daemon-wiring.test.mjs` fails).
- Changing exports of `ledger-check.mjs`, `response-quotes.mjs`, `edit-facts.mjs`, `revision-store.mjs` without grepping importers (wide fan-in).
- Rewording prompt or refusal strings as cosmetic edits: they are model-facing behavior.
- Clearing stale marks or rewriting letter line numbers on a citation refresh (`response-quotes.mjs` `markStale`).
- Letting an older ledger check lower a mark a newer check set.
- Documents over `DOC_MAX_BYTES` (100 MiB, `doc-bytes.mjs`): pane must save first (`SAVE_FIRST`).
- Removing the NEVER rule for `office_replace_text` in `system-prompt-word.md`.

## HOTSPOTS
`review-service.mjs` ~2000, `index.mjs` ~1780, `revision-store.mjs` ~1690, `review-tools.mjs` ~1500, `review-split.mjs` ~1360, `response-service.mjs` ~1020 lines. Check import cycles (response-quotes -> review-split -> revision-store) before splitting any.

## TESTS
`node --test tests/<name>.test.mjs`; provider tests are `tests/providers-*.test.mjs`, `tests/cli-exec.test.mjs`.
