# Word acceptance run without hands

`run-acceptance.mjs` runs a Review tab acceptance procedure end to end:
daemon, a separate Word process, the task pane clicks, the author's Word
steps, and the feature_list verify commands. It has four flows: stage 2
(`S2-18..S2-34`, the default), stage 2.5 (`--stage s25`, `S25-30..S25-41`),
stage 3 (`--stage s3`, `S3-30..S3-41`, the response letter) and stage 4
(`--stage s4`, `S4-30..S4-33,S4-37,S4-38`, the Tools tab).

Stage 4 opens the fake manuscript `tests/fixtures/aux_manuscript.docx`
(`make_aux_fixture.py`, every sentence invented) with the journal profile
the same script prints, put under `journal_profiles/` in the run folder. Its
steps: translate a selected paragraph and ask about it (the pane's snapshot
of the document is the same before and after, and the daemon log shows no
agent turn and no tool call); paste `tests/fixtures/fake_guideline.txt`,
keep the abstract word limit of the candidates and add it (it is in the
ledger of `revision.json` with the sentence it cites); run the checks (the
count rule reports the planted abstract, the format rules report the five
planted findings, the document is unchanged); make the marked copy in both
modes (each file equals what `tools/build_marked_diff.py` writes from the
same two files); make the anonymized copy (the runner runs
`scripts/verify-anonymized.py` on it again: total 0, same text). The
selection and guideline steps ask the judgment provider chosen in the pane
(`--judgment-provider` picks one). `--features` takes a list of ranges for
this stage (`S4-30..S4-33,S4-37`).

## Prerequisites

- Windows with Word (Office 16), driven from WSL. `powershell.exe`,
  `wslpath`, `ss`, `jq` and `python3` on the WSL side.
- The add-in manifest `manifests/word.xml` sideloaded (registered under
  `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`).
- Node on Windows and an existing `playwright-core` package on Windows (no
  browser download is needed: the runner only attaches over CDP).
- Ports 47833/47834 free (no other Draftspect daemon) and a free CDP port
  (default 9333).
- The run folder must be on a Windows drive (for example under `/mnt/d`).
  Windows scripts are copied there and run from there, never from a UNC path.

## Use

```sh
node scripts/e2e/run-acceptance.mjs \
  --base-dir /mnt/d/path/to/test-root \
  --playwright-core 'D:\path\to\node_modules\playwright-core' \
  --feature-list /path/to/feature_list.json \
  --plan-dir <folder the feature_list commands name> \
  --plan-repo <repo path the feature_list commands cd into> \
  --conditional S2-27
```

The exit code is 0 only when every step passed, every verify command exited
0 and the Word started by the run was closed. `--conditional <id>` (repeat
or comma-separate) names a feature that depends on an event the run cannot
force. Such a feature is reported `SKIPPED-CONDITION-NOT-MET` and does not
fail the run only when its verify command failed and the event is provably
absent from this run's daemon logs (the logs show finished turns and no
line of the event). Known conditions (`CONDITIONS` in `judge.mjs`): `S2-27`,
a turn that ended on a usage limit (`[agent] turn complete (usage_limit)`).
An id without a known condition counts like any other feature.

A new folder `e2e-<UTC time>` is made under `--base-dir`. It holds the tagged
copy of the test manuscript, its revision.json and comment backup, the Windows
scripts (`tools/`), the machine-judged marker files (`ok/`, each with an
`.evidence.json`) and `result.json` (every step with the pane texts read, the
Word COM outcomes, the verify command exit codes and the cleanup record). The
daemon log goes to `--log` (default `/tmp/daemon-<stage>-e2e.log`; a second
daemon started for the forbid-list knob writes `-knob.log`).

Verify commands are rewritten before they run (`rewriteVerify` in
`judge.mjs`): `--plan-dir` (default `/tmp/<stage>-run`) becomes the run
folder, `cd <--plan-repo> ` becomes this checkout, `--plan-log` (default
`/tmp/daemon-<stage>.log`) and its `-knob.log` become this run's logs,
`/tmp/<stage>-` markers and evidence files become `ok/<stage>-`. No marker
path of a rewritten command is in `/tmp`, so a file left there by an earlier
run or by hand cannot make it pass. When a run folder is reused
(`--run-dir`), the `ok/<stage>-*.ok` and `ok/<stage>-*.evidence.json` files
in it are removed at the start (`stale_markers_removed` in `result.json`).

The stage is the prefix of the feature ids and of the markers, and it
chooses the flow: `s2`, `s25` or `s3`. `--stage s25` sets it (the features
then default to `S25-30..S25-41`); without it the stage is read from
`--features` (default `S2-18..S2-34`). A range stays in one stage.

Markers are written by the runner only (`marker(name, evidence)`): the
evidence file always, the `.ok` file only when `evidence.pass` is true, and
`pass` is computed by a pure function of `judge.mjs` from machine data (the
texts read in the pane, revision.json, the daemon log, the answers of Word
over COM, `pdftotext`). No marker is made by a person.

## Stage 2.5

```sh
node scripts/e2e/run-acceptance.mjs --stage s25 \
  --base-dir <folder on a Windows drive, as a WSL path> \
  --playwright-core 'D:\path\to\node_modules\playwright-core' \
  --feature-list <stage 2.5 feature_list.json> \
  --plan-repo <repo path the feature_list commands cd into> \
  --conditional S25-33,S25-34,S25-35,S25-36
```

The copy of the manuscript is `paper.docx`, so the files the feature list
reads are `paper.decisions.jsonl` and `paper.ledger.json` in the run folder;
`--plan-dir` and `--plan-log` default to `/tmp/s25-run` and
`/tmp/daemon-s25.log`, the paths that list writes. Ollama must be running
with a chat model (`--judgment-provider` names another provider for the
judgments). The steps and what each feature is judged from:

| Feature         | Steps                                                                                   | Evidence (`ok/`), marker only when it passes                                                                               |
| --------------- | --------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------- |
| S25-30          | `models`                                                                                | `s25-models`: a model outside sonnet, opus, haiku in the menu and `[models] source=sdk` in the log                         |
| S25-31          | `providers`                                                                             | `s25-providers`: one entry per provider (`enabled`, `note` of the Judgments menu)                                          |
| S25-32          | `select-provider oneshot=ollama`, `load-review-file`, `split`, `edit-draft`             | `s25-ollama-split`: Ollama selected, the split done, a receipt with `"model":"ollama:..."`                                 |
| (no feature id) | daemon stopped and started again, `split` as soon as the pane is back                   | `s25-restart-split`: the stored provider restored, the split's judgments on it, no judgment or receipt of another provider |
| S25-39          | `import-ledger` (fixture), `export-ledger`, `import-ledger` (the exported file)         | the step rows; the feature reads `paper.ledger.json` and the `[review] ledger import` log line                             |
| S25-40          | `suggest-ledger` on `tests/fixtures/fake_review_rules.txt`, ticking "cell-free extract" | `s25-ledger-suggest`: the candidate ticked, added, and in the ledger                                                       |
| S25-33          | `select-provider turn=codex`, `propose R1.1`                                            | `s25-codex-turn`: the turn started and ended with code 0, the endpoint was called, a change with its comment               |
| S25-34          | `propose R1.3` (asks for a word rule L2 forbids)                                        | `s25-codex-guard`: `document_changes` = recorded changes that hold the word (0 to pass)                                    |
| S25-36          | `propose-stop R1.4`                                                                     | `s25-cli-stop`: Stop pressed mid-turn, pane Ready, no `codex exec --json` process left                                     |
| S25-35          | `select-provider turn=antigravity`, `propose R1.2`                                      | `s25-agy-turn`: as S25-33                                                                                                  |
| S25-37          | `propose-all` (with Claude)                                                             | `s25-propose-all`: `overlaps`, `proposed`, `open_before`                                                                   |
| S25-38          | Word COM deletes a sentence, `focus-refresh`                                            | `s25-focus-refresh`: the auto-refresh line changed and an outside change is listed                                         |
| S25-41          | cleanup (Word quit, daemon stopped), then the verify commands                           | `s25-cleanup`: own Word gone, earlier Word processes alive, ports closed                                                   |

When the pane does not offer Codex or Antigravity for review turns, the
`select-provider turn=...` row is `SKIP`, no CLI turn is run, and the four
evidence files carry `pass: false` with the pane's reason. With
`--conditional` those features are then `SKIPPED-CONDITION-NOT-MET` (the log
shows finished turns and no `[cli-turn] start`); without it they fail the
run. The pane outputs these judgments were unit-tested with are written by
hand in the form the driver returns: none was recorded from Word before the
first stage 2.5 run.

## Stage 3

```sh
node scripts/e2e/run-acceptance.mjs --stage s3 \
  --base-dir <folder on a Windows drive, as a WSL path> \
  --playwright-core 'D:\path\to\node_modules\playwright-core' \
  --feature-list <stage 3 feature_list.json> \
  --plan-repo <repo path the feature_list commands cd into>
```

The response letter (Response tab). More is needed than for stage 2:
`pdftotext` (poppler-utils) and PyMuPDF (`python3 -c "import fitz"`) on the
WSL side, the first as a second reader of the PDF Word renders, the second
for the daemon's own line reader. The copy of the manuscript is
`review_manuscript.docx` as in stage 2; `--plan-dir` and `--plan-log` default
to `/tmp/s3-run` and `/tmp/daemon-s3.log`. The features default to
`S3-30..S3-40`, checked while Word and the daemon are still up, and
`--after-cleanup` (default `S3-41`) names the features whose verify command
runs after the cleanup. `--letter` is the existing response letter that is
read for its colours (default `tests/fixtures/prior_response.docx`).

The review part is the stage 2 procedure up to the author's decisions in
Word, with `S3_DECISIONS` (`judge.mjs`): the deletion asked for in R1.4 is
accepted, so the letter cites a removal. The author's own pending insertion,
the knob run, the comment cleanup, the mode 14 copy and the special
character document are left out: a tracked change left pending is a warning
of the letter check, and S3-34 asks for a check without one.

The drafts are written by the judgment provider of the pane.
`--judgment-provider <id>` chooses it (step `select-provider oneshot=<id>`);
without the option the pane's own choice stays, which is Claude on a first
run. `--provider` chooses the provider of the review turns as in stage 2.

| Feature | Steps                                                                                                                                        | Evidence (`ok/`), marker only when it passes                                                                             |
| ------- | -------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------ |
| S3-30   | `open-response`                                                                                                                              | `s3-open`: `cards`, `verbatim` (each card's comment equals revision.json), `profile`                                     |
| S3-31   | `load-letter`, then revision.json is read, `apply-proposal`                                                                                  | `s3-proposal`: the proposal rows, `before` (overrides while only proposed) and `after`                                   |
| S3-32   | `response-decline` (R1.5, and every request no accepted change answers), `response-draft-all`                                                | none: the verify command reads revision.json and the daemon log                                                          |
| S3-33   | `response-locate`, `pdftotext -layout` on the rendered PDF                                                                                   | `s3-lines`: `secondReader`, one row per cited sentence with the text of its start line                                   |
| S3-34   | `response-check`                                                                                                                             | `s3-check`: `summary`, `checks`                                                                                          |
| S3-35   | `response-export`, `tools/response_read.py` on the written letter                                                                            | `s3-export`: `body`, `revised` (the colours the reader finds), `path`                                                    |
| S3-36   | Word COM `insert-paragraph`, `refresh` (Review tab), `open-response`, `response-locate`, `pdftotext`                                         | `s3-line-shift`: `lineShift` (`moved`, `kept`, `wrong`, `stale_shown`), `inserted_at`, `second`                          |
| S3-37   | Word COM `replace-text` in a quoted sentence, `refresh`, `open-response`, `response-check`, `response-edit` (kept), `response-locate`, check | `s3-stale`: `stale_shown`, `failed_before`, `passed_after`, the daemon's `[response] stale` line                         |
| S3-38   | `ledger-add` (response scope), `response-edit` (forbidden word), `response-export`, `response-edit` (put back), `response-check`             | `s3-ledger`: `ledger_fail`, `exported_while_failing` (the letter files and their times before and after), `passed_after` |
| S3-40   | `response-export` (text), Word COM `open-check` on the letter, `response-export`, Word COM `close-other`                                     | `s3-open-check`: `opened`, `repaired`, `paragraphs_match` (Word's count and model.json), `first_path`, `second_path`     |
| S3-39   | `response-set-lang` (German), `response-draft` R1.1, `preview` R1.1 (Review tab), `response-set-lang` (empty)                                | `s3-lang`: the language in the daemon's draft line and in the review prompt, `korean_found`                              |
| S3-41   | cleanup (Word quit, daemon stopped), then its verify command                                                                                 | `s3-cleanup`: own Word gone, earlier Word processes alive, `ports_free`                                                  |

What differs from a reading of the feature descriptions alone:

- S3-37: the stale answer is cleared by saving it again (the author read it
  and keeps it) and updating the line numbers, not by a refresh. A refresh
  computes the citations anew, and the mark stays until the answer is written
  again.
- S3-38: the export is asked for right after the edit, while the button
  still reads "Export letter (.docx)". The daemon finds the failure, writes
  nothing, and the button then reads "Export anyway", which the run does not
  press.
- S3-40 runs before S3-39, so the letter that is opened in Word and the
  export next to it are in the language of the checks that passed.
- "R1.4 cites the removal" is a row of the table, `SKIP` when the model
  answered R1.4 with another kind of change. No verify command reads it.

A step that only reads or computes again (`open-response`, `response-locate`,
`response-check`) is run a second time when it fails once (`retries` in
`result.json`); a step that writes or asks a model is not. The flow has no
fixed wait: every Response step waits for the message its own action ends
with.

`--provider <id>` (claude, codex, antigravity; stage 2 flow) chooses the
provider of the review turns in the pane's Providers section right after the
Review tab opened (step `select-provider turn=<id>`). A provider the pane
does not offer for review turns is not chosen: the step fails with the text
of its menu row, and the turns run with Claude. Known conditions for
`--conditional` in stage 2.5: `S25-33`, `S25-34` (a Codex review turn
started), `S25-35` (an Antigravity review turn started), `S25-36` (any CLI
review turn started), each read from `[cli-turn] start provider=...`.

Other options: `--cdp-port`, `--win-node`, `--fixture`, `--review`,
`--ledger`, `--features S2-18..S2-34`, `--skip-knob`, `--skip-mode14`,
`--keep-on-fail` (leave Word and the daemon up after a failed step).

## Pane steps

`pane-driver.cjs <step> <args.json>` runs one step and prints one JSON line.
Stage 2: `open-review`, `split`, `edit-draft`, `ledger-add`, `check-all`,
`preview`, `propose`, `refresh`, `clean-tool`, `delete-all`, `view`.
`split` and `suggest-ledger` wait for the message their own action ends with
(`Split by ...`, `... candidate rule(s) by ...`) or for an error. The line
can show other text first: the notice of a click held until the daemon has
confirmed the stored provider, or the late result of an earlier action.
Stage 2.5:

| Step               | Args                                  | Does and returns                                                                                                                                |
| ------------------ | ------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| `models`           |                                       | the rows of the chat box's model menu once the daemon's list is in (`title`, `options`); a list that is late is asked for once more (`retried`) |
| `propose-stop`     | `ref`, `after_ms`                     | Propose changes, then Stop while the turn runs; `stopped`, the pane view after it                                                               |
| `providers`        |                                       | Setup tab; both menus (`id`, `text`, `enabled`, `note`), every provider row (state, status, policy, key note), message                          |
| `select-provider`  | `role` (turn, oneshot), `id`, `model` | chooses it; `selected` is what the daemon accepted, `refused` the menu text of a provider that cannot be chosen                                 |
| `load-review-file` | `path` (Windows)                      | review text from a .txt or .md file; `message`, `chars` in the paste box                                                                        |
| `export-ledger`    |                                       | Export ledger; the message names `<stem>.ledger.json`                                                                                           |
| `import-ledger`    | `path` (Windows)                      | Import ledger; `Added a ledger rule(s), skipped s ...`                                                                                          |
| `suggest-ledger`   | `review_file`, `pick`, `add`          | candidates from the review text; ticks the ones holding a `pick` text, adds the ticked ones unless `add` is false                               |
| `propose-all`      |                                       | Propose all until its Stop button is gone; `order` of the comments tried, the summary message                                                   |
| `focus-refresh`    | `wait_ms`                             | dispatches `focus` on the pane window; `changed` is true when the auto-refresh line changed with no Refresh click                               |

Stage 3 (Response tab). Each step returns what the tab shows: `message`,
`next`, `profile`, `summary`, `checks`, `export_label`, `output`,
`proposal_rows` and `cards` (`ref`, `comment`, `text`, `declined`, `stale`,
`undecided`, `citations` with `line` and `quote`, `open_reqs`).

| Step                 | Args                                                        | Does, and the message it waits for                                                       |
| -------------------- | ----------------------------------------------------------- | ---------------------------------------------------------------------------------------- |
| `open-response`      |                                                             | shows the tab, presses Refresh from the manuscript; `Refreshed from the manuscript: ...` |
| `response-view`      |                                                             | reads the tab, presses nothing                                                           |
| `response-refresh`   |                                                             | Refresh from the manuscript                                                              |
| `load-letter`        | `path` (Windows)                                            | reads an existing letter; `Read the letter: ...`; `rows` are the proposal lines          |
| `apply-proposal`     |                                                             | Apply to this paper; `Applied to this paper: ...`                                        |
| `response-set-lang`  | `lang`, `rationale_lang`                                    | the language form, Save for this paper; `Saved for this paper: profile ...`              |
| `response-decline`   | `comments` (ref, reason), `open_requests`, `request_reason` | declines the comments, then every open request; `Declined ...`                           |
| `response-draft`     | `ref`                                                       | Draft on one card; `Drafted <ref>.`                                                      |
| `response-draft-all` |                                                             | Draft all missing; `Drafted n responses.`                                                |
| `response-edit`      | `ref`, `text` or `prepend`                                  | fills the answer (or leaves it), Save; `Saved <ref>.`; `previous_text`                   |
| `response-locate`    |                                                             | Update line numbers; `Line numbers updated: ...`                                         |
| `response-check`     |                                                             | Run checks; `Checked: ...`                                                               |
| `response-export`    | `format` (`text` for the plain-text files)                  | Export; `Exported...` or, while a check fails, `Not exported: ...`                       |

`judge.mjs` holds the expected texts of these steps (`JUDGES`) and the
evidence builders of the stage 2.5 markers (`modelsEvidence`,
`providersEvidence`, `ollamaSplitEvidence`, `cliTurnEvidence`,
`guardEvidence`, `stopEvidence`, `proposeAllEvidence`, `focusEvidence`,
`suggestEvidence`, `cleanupEvidence`) and of the stage 3 markers
(`openEvidence`, `proposalEvidence`, `secondReader`, `exportEvidence`,
`lineShift`, `staleEvidence`, `ledgerEvidence`, `langEvidence`,
`openCheckEvidence`).
There is no step that registers the office tools with agy: the daemon
registers them for the length of one agy review turn only.

## Resolved rationale comments

- After the author's Word steps and the Refresh, the runner reads
  `Comment.Done` of every comment over COM and requires that the rationale
  comment of each accepted or rejected change is resolved (row
  `word: rationale comments resolved`). The test manuscript is built in
  compatibility mode 15 for this; in mode 14 Word refuses to resolve.
- A change next to another tracked change (row `word: R1.4 next to a pending
revision`): before the author's steps the runner puts a tracked insertion
  of the author (`Checked: `) at the start of the paragraph that holds the
  R1.4 deletion and leaves it pending, then rejects R1.4. After the Refresh
  R1.4 must read Rejected (not Pending) in revision.json and on its badge,
  with its rationale comment resolved, while Word still lists the author's
  insertion as a revision. The paragraph's revision mark total used to
  decide, which kept such a change Pending.
- Tracked special characters (row `special: tracked runs found in Word's
paragraph text`, files under `special/`): `make_tracked_special.py` builds
  a document whose tracked insertions and deletions hold a no-break hyphen,
  a soft hyphen, a tab, a line break, a carriage return and symbols. It is
  opened in a third Word of the run, the pane's own snapshot is read, and
  every tracked run must be found in the paragraph text Word reports. This
  is the measurement behind the character table in `trackedRuns`
  (`taskpane/shared/review-view.js`). No model turn; `--skip-special` leaves
  it out.
- Mode-14 check (`mode-14:` rows, files under `mode14/` in the run folder,
  daemon log `-mode14.log`): after the main Word is closed, the runner
  builds a compatibility mode 14 copy with
  `scripts/make-test-manuscript.py --compat 14`, opens it in a second Word
  of its own, keeps comment R1.1, proposes its change, accepts it over COM
  and refreshes. Expected: the refresh message carries the counts and
  `n rationale comment(s) not marked resolved (Word: ...)`, and the comment
  is still open in Word. It costs one more model turn.

## How Word is started

- `word-launch.ps1` sets `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=<port>`
  and `WEBVIEW2_USER_DATA_FOLDER=<run>/webview2` in its own process only and
  starts `WINWORD.EXE /n /q <doc>`. `/n` starts a new Word instance even when
  Word is already open, and the child inherits the two variables. No user or
  machine environment variable is set, so other WebView2 apps get no port.
  The separate user data folder keeps the pane out of the WebView2 browser
  process of a Word that is already running.
- `New-Object -ComObject Word.Application` is not used to start Word: measured
  on Office 16, it starts `WINWORD.EXE /Automation -Embedding` under
  `svchost.exe` (DCOM), so the caller's `$env:` does not reach it.
- `embed_taskpane.py` tags a copy of the manuscript with a webextension part
  (`Office.AutoShowTaskpaneWithDocument=true`, `storeType="Registry"`,
  `store="developer"`, id and version from the manifest) and a taskpane part
  with `visibility="1"`, so the pane opens with the document.
- `word-com.ps1` reaches the document through the Running Object Table by its
  full path, so documents in any other Word are not touched. `quit` closes
  without saving and only when every open document is in the run folder; the
  runner then checks that the started process id is gone and that every Word
  process that was running before is still there.
- When `quit` refuses (a document outside the run folder is open in the Word
  the runner started) or cannot say which documents are open, the runner does
  not stop that Word: it stays up, `result.json` lists it under
  `cleanup.left_running`, and the run exits non-zero. The process is stopped
  by force only when `quit` ran on run-folder documents and the process is
  still there.

## Files

| File                 | Runs on      | Does                                                                                               |
| -------------------- | ------------ | -------------------------------------------------------------------------------------------------- |
| `run-acceptance.mjs` | WSL          | the whole run                                                                                      |
| `judge.mjs`          | WSL          | COM plan from revision.json, expected pane texts, path rewrite (tests: `tests/e2e-judge.test.mjs`) |
| `pane-driver.cjs`    | Windows node | one pane step per call over CDP                                                                    |
| `word-launch.ps1`    | Windows      | Word with a CDP port on its pane                                                                   |
| `word-com.ps1`       | Windows      | accept/reject, hand edits, comments, quit                                                          |
| `embed_taskpane.py`  | WSL          | auto-open task pane parts in a copy                                                                |

## Word as the judge of the marked copy

`tools/build_marked_diff.py` checks its output with its own model of "accept
all" and "reject all". `marked-diff-word-check.mjs` asks Word instead. It needs
no daemon, no port and no task pane.

```
node scripts/e2e/marked-diff-word-check.mjs --base-dir /mnt/d/some/folder
```

- For 98 fixed pairs of `tests/marked_diff_corpus.py` (the list and what each
  pair covers: `PAIRS` in `marked-diff-word-lib.mjs`) it writes the baseline and
  the clean document under `<base-dir>/mdw-<time>/pairs/<nn>/` and runs the tool
  in both modes.
- `marked-diff-word.ps1` starts a Word process of its own on `anchor.docx` of
  the run folder (`/n`), finds it through the Running Object Table by that
  path, and in that Word opens each file read-only, without a repair and
  without dialogs. The track-changes output is opened twice: once for
  `Revisions.AcceptAll()`, once for `Revisions.RejectAll()`. The structure is
  read each time (the text of every paragraph, how many pictures and fields
  it holds, the code and result of each field and the text of each equation, every table with its rows,
  columns, cell texts, whether cells are merged (`Table.Uniform`) and where
  each cell stands, every bookmark with its text, and the number of
  comments, footnotes and endnotes) and the document is
  closed without saving. The clean and the baseline document are read the
  same way; the colour output is opened and its paragraphs and tables counted.
- A pair passes when Word opens every file without an error and without
  changing it on the way in, no revision is left, accept-all equals the clean
  document, reject-all equals the baseline, and the colour output has the
  paragraphs and tables of the track-changes output. Comments are the
  exception on the reject side: a comment is no tracked change, so after
  reject-all the document must still have the comments of the clean document.
- Word is quit only when every document open in it is under the run folder;
  the cleanup rules are those of the acceptance run above. A Word that was
  already running is never touched.
- `--prepare-only` writes the pairs, the tool's output and `jobs.json` and
  stops before Word. `--only "table,break"` keeps the pairs whose name has one
  of these texts.
- Output: one line per pair with the first difference, `result.json` in the run
  folder, exit code 0 only when every pair passed and the Word of the run is
  gone.

| File                         | Runs on | Does                                                                                             |
| ---------------------------- | ------- | ------------------------------------------------------------------------------------------------ |
| `marked-diff-word-check.mjs` | WSL     | the whole check                                                                                  |
| `marked-diff-word-lib.mjs`   | WSL     | the pairs, the comparison, the verdict (tests: `tests/marked-diff-word-check.test.mjs`)          |
| `marked_diff_pairs.py`       | WSL     | the docx pairs from the corpus                                                                   |
| `marked-diff-word.ps1`       | Windows | Word: open, accept or reject all, read the structure, close without saving, quit                 |
| `word-own.ps1`               | Windows | the Word of the run's own: start, ownership check, quit (dot-sourced by the two below and above) |

### How Word writes a tracked change of an equation

The tool marks an equation that is new or gone the way Word does. That
notation was read from Word with `word-math-probe.mjs`:

```
node scripts/e2e/word-math-probe.mjs --base-dir /mnt/d/some/folder
```

In a Word process of its own (the rules of `word-own.ps1`),
`word-math-probe.ps1` makes `base.docx` under `<base-dir>/wmp-<time>/` (five
paragraphs, four with an equation), then `revised.docx` with Track Changes on:
an equation deleted and another inserted, a word changed next to an equation,
a paragraph with an equation deleted, a paragraph with an equation added, one
character changed inside an equation. The `word/document.xml` of both files is
written next to them and summarised: for every paragraph with an equation or a
change mark, what stands in what, and where the change marks stand (around
`m:oMath`, inside it, in a `w:rPr`, in `m:ctrlPr`, in the paragraph mark).
Exit code 0 when Word made both files, all five changes ran and the Word of
the run is gone. The summary is `mathSummary` and `markRelations` in
`word-math-probe-lib.mjs` (tests: `tests/word-math-probe.test.mjs`).
