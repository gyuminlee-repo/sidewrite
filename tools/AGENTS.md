# tools/

## OVERVIEW
Python 3 CLI tools the daemon and the e2e runner call as subprocesses: letter write/read/verify, marked diff, anonymize, citation field check, figure and CSL render. Earned its file: score 5, but distinct domain (separate runtime, CLI plus JSON contract instead of imports).

## WHERE TO LOOK
| Task | File | Notes |
|------|------|-------|
| Write response letter docx | `response_docx.py` | `--model --out --overwrite`; deterministic bytes (fixed zip timestamps) |
| Read letter, propose profile overrides | `response_read.py` | `--docx --profile` |
| Verify letter vs model and manuscript | `response_verify.py` | imports `response_docx`, `response_read`; `--selftest` |
| Marked (track changes) docx from baseline + clean | `build_marked_diff.py` | ~2830 lines, needs lxml |
| Strip author/identifying metadata | `anonymize_docx.py` | `--keep --also` |
| Zotero/EndNote fields survive an edit | `citation_fields_check.py` | `--before --after --fld-data judge\|note` |
| First page of PDF/AI/SVG/raster to PNG | `figure_render.py` | PyMuPDF; `--probe` |
| CSL reference render | `csl_render.py` | citeproc-py optional; `--check` |

## CONTRACT
- Invoked as `python3 tools/<name>.py ...` via `daemon/py-tool.mjs` `runPy` (interpreter from `CC_OFFICE_PYTHON`) or from `scripts/e2e/`. Path found through `daemon/doc-tools.mjs` `REPO_ROOT`.
- Print ONE JSON object as the last stdout line; the daemon parses only that line. Extra stdout after it breaks the caller.
- Exit codes: 0 ok; 1 check failed or refused (nothing written); 2 usage or bad input; 3 optional dependency missing. `response_docx.py` also returns 4 (verify mapping in source before relying on it).
- Refusals are exceptions (`Refused`, `Locked`, `Exists`, `NotADocx`, `BadInput`) turned into an exit code plus JSON, not tracebacks.

## CONVENTIONS
- Flat layout: no package, no `__init__.py`; siblings import by bare name (script dir on `sys.path`).
- Every file: `#!/usr/bin/env python3`, `main(argv=None)`, `sys.exit(main())` guard.
- Stdlib first; lxml only in `build_marked_diff.py`; fitz and citeproc imported lazily inside functions, so a missing package fails only on that path.
- Output never overwrites by default; write to a temp file then `place(tmp, out, overwrite)` with a `same_file` guard.
- `build_marked_diff.py` uses single quotes; others double. Do not normalize in a focused fix.

## ANTI-PATTERNS
- Writing partial output on a failed check (`build_marked_diff.py` asserts paragraph text adds up, else writes nothing).
- Editing anonymize categories here only: same ids and defaults live in `scripts/verify-anonymized.py`; citation rules mirror `taskpane/shared/field-text.js`. Change all copies together.
- Diverging the duplicated `Refused`/`same_file`/`place` helpers between `anonymize_docx.py` and `build_marked_diff.py`.
- Calling `python`: this machine has only `python3`.

## TESTS
- `python3 -m unittest tests/test_<tool>.py` (one file per tool; not wired into `npm test`).
- Node side: `node --test tests/marked-diff.test.mjs tests/e2e-judge-cite.test.mjs`.
- `tests/helpers/figure-world.mjs` stands in for `figure_render.py` in Node tests.
