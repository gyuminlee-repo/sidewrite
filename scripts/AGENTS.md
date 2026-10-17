# scripts/

## OVERVIEW
Developer and operator tooling outside the daemon: acceptance automation (`e2e/`), resident install (`install/`), and a handful of standalone helpers. Earned its file: distinct domain (mixed Bash/Python/PowerShell/Node run by hand or by the daemon, not by the pane).

## STRUCTURE
```
scripts/
├── e2e/        # hands-free Word acceptance + replay (own AGENTS.md, README.md 1252 lines)
├── install/    # systemd user unit, Windows registration, doctor (own AGENTS.md)
├── check-personal.sh + personal-allow.txt
├── pdf-line-map.py          # daemon runs it: margin line numbers from rendered PDF
├── verify-anonymized.py     # known-answer check of an anonymized docx
├── make-test-manuscript.py, make-test-response.py   # public fake fixtures
├── probe-hi.mjs             # Word-free token probe against a running daemon
└── build-icons.py, build-tray-icon.py
```

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| Personal-value leak gate | `check-personal.sh` | exit 0 "personal check: 0 hits", 1 hits, 2 outside git; real names go in untracked `.personal-patterns.local` |
| Allow a known hit | `personal-allow.txt` | fixed text `path:line:` with reason above |
| Line numbers for response letter | `pdf-line-map.py` | called from `daemon/response-lines.mjs` (`LINE_MAP`); tested by `tests/test_pdf_line_map.py` |
| Measure prompt/hook token cost | `probe-hi.mjs` | `--jsonl <file>` to summarize an existing session; tested by `tests/probe.test.mjs` |
| Regenerate fake manuscript | `make-test-manuscript.py` | Word line numbering on (`w:lnNumType`) |

## CONVENTIONS
- Python is always `python3` (daemon honors `CC_OFFICE_PYTHON`). Scripts are stdlib unless the header names a package; tests skip when the package is missing.
- Each script opens with a header comment: purpose, usage lines, side effects. Keep it current; it is the only doc for top-level helpers.
- All fixture content is invented; never commit a real manuscript, name, or path.
- `check-personal.sh` builds its own patterns in split pieces (`'/ho''me'`) so it does not hit itself; keep that trick when adding patterns.
- `__pycache__/` here is ignored build output, not source.

## ANTI-PATTERNS
- Never hardcode a home folder, Windows user folder, mount path or e-mail in a tracked file; `check-personal.sh` (and `tests/check-personal.test.mjs`) fails on it.
- `probe-hi.mjs` leaves its temp workspace in `~/.claude/office-addins/sessions.json`; do not run it against a production daemon expecting no trace.
- Do not move a script the daemon shells out to (`pdf-line-map.py`) without updating `daemon/response-lines.mjs` and the matching `tests/test_*.py`.

## VERIFY
```bash
bash scripts/check-personal.sh
python3 -m py_compile scripts/*.py
node --check scripts/probe-hi.mjs
```
