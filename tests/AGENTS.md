# tests/

## OVERVIEW
Flat test directory for every layer: Node `node:test` suites, Python `unittest` files for the Python tools, one PowerShell source test. Earned its file: score ~27 (191 entries, ~100k LOC, cross-layer contract tests).

## STRUCTURE
```
tests/
├── <module>.test.mjs   # node:test + node:assert/strict, run by npm test
├── test_<tool>.py      # unittest for tools/*.py and scripts/*.py, NOT run by npm test
├── word-cleanup.test.ps1  # parses e2e PowerShell, runs against fakes
├── marked_diff_corpus.py, anon_docx_fixture.py   # Python fixture generators
├── helpers/            # plain-pane, figure-world, figures-page, submission-world/-fixture, stalled-body
└── fixtures/           # docx/json/jsonl/csl/txt samples, fake-bin.mjs, fake-systemctl.sh, py-packages.mjs, legacy-manifests/, review_formats/
```

## WHERE TO LOOK
| Change in | Tests |
|-----------|-------|
| `daemon/X.mjs` | `X.test.mjs`, plus `*-bridge-integration.test.mjs` (real bridge over `ws`) |
| `taskpane/shared/X.js` | `X.test.mjs`, DOM glue in `*-pane-dom.test.mjs` |
| Route/tool wiring | `daemon-wiring.test.mjs` (asserts figure routes live in `figure-http.mjs`, not `index.mjs`) |
| `scripts/e2e/judge.mjs` | `e2e-judge*.test.mjs` |
| `tools/*.py`, `scripts/*.py` | `test_<name>.py` |
| Revision state schema | `revision-store.test.mjs`, `revision-roundtrip.test.mjs`, and ~15 other importers |

## CONVENTIONS
- Test names are behavior sentences ("legacy ownership: only the exact generated unit, ..."); `submission-check` titles carry `row N:` matching the checklist row.
- Temp state per test: `mkdtemp(join(tmpdir(), "<prefix>-"))`, removed with `rm`. Fixed clocks, invented manuscript text.
- Optional Python packages: `{ skip: missing("lxml") }` from `fixtures/py-packages.mjs` (probes `CC_OFFICE_PYTHON` or `python3`); Python side uses `unittest.skipUnless`. CI installs no Python packages, so these tests skip there.
- Python tests load the tool by file path via `importlib.util`; run line is in each file header.
- Fakes over Office: fake Word model with failure injection, fake DOM, `plainPane` helper, stub `sendRequest`, fake `runPy`.
- Cache-busting dynamic import (`import(url + "?x")`) to reload module state (`transcript`, `workspace`).
- Prompt-text assertions (`review-prompt`, `review-requirements`, `response-service`) match prompt wording; edit the regex together with the prompt.

## ANTI-PATTERNS
- Do not add fixed sleeps or poll loops. Existing debt: `review-service.test.mjs` (~l.1380-2858), `supplement-link*.test.mjs`, `selection-oneshot`, `pane-caps` `until()`, `marked-build` l.43. New async tests subscribe to the event, then await with a bounded timeout.
- A skip is not a pass: a green CI run proves nothing for lxml/python-docx/fitz/citeproc tests; run them locally before claiming coverage.
- `review-turn-record.test.mjs` skips when SDK internal `McpServer._registeredTools` disappears; check after SDK bumps.
- Source-contract tests (`word-cleanup.test.ps1`, some e2e runner checks) read code text; behavior may break while they pass.
- Never put real manuscripts, names, or home paths in fixtures (`check-personal.test.mjs` guards). Literal `/mnt/d/...` strings in `paths.test.mjs` are pure data.
- No `.only(`.

## COMMANDS
```bash
npm test                                    # node --test tests/*.test.mjs
node --test tests/<name>.test.mjs
for f in tests/test_*.py; do python3 -m unittest "$f"; done
```
