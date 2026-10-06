# Word acceptance run without hands

`run-acceptance.mjs` runs a Review tab acceptance procedure end to end:
daemon, a separate Word process, the task pane clicks, the author's Word
steps, and the feature_list verify commands. It has four flows: stage 2
(`S2-18..S2-34`, the default), stage 2.5 (`--stage s25`, `S25-30..S25-41`),
stage 3 (`--stage s3`, `S3-30..S3-41`, the response letter) and stage 4
(`--stage s4`, `S4-30..S4-33,S4-37,S4-38`, the Tools tab). A fifth flow,
`--stage cite`, has no feature list: it asks whether tracked edits leave the
fields of Zotero and EndNote whole (section "Stage cite" below). Stage parts
(`--stage parts`) has none either, nor has `--resident`, a short stage on
the installed daemon (section "Stage resident"). Stage 3 takes two steps more
with `--check-now` and `--figures-button`.

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
- Ports 47833/47834 free (no other Sidewrite daemon) and a free CDP port
  (default 9333). The installed daemon (`scripts/install`, a systemd user
  unit) is the one exception: when it is active the run stops it at the
  start and starts it again when the run ends, also after a failure or
  Ctrl-C (`resident-unit.mjs`). With `--keep-on-fail` the run's own daemon
  stays on the ports, so the unit cannot come back until that daemon is
  stopped; then `bash scripts/install/daemon-service.sh install`. With
  `--resident` it is the other way round: the unit must be running, and it
  stays running.
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

Stage 2 shows the forbid guard without a model (S2-23, marker
`s2-forbid-block`). After the review turns the run puts a blocked edit on
file for R1.3 that holds the forbidden word of L2 and names L1 as the rule it
waits for (`FORBID_GUARD_EDIT` in `judge.mjs`, written with the store's own
`addBlocked` while no request is running), and presses "Allow exception" in
the pane. The daemon sends the stored edit through the same guarded write as
an edit of a turn; the exception lifts L1 only, so L2 must refuse it:
`[review] blocked R1.3 by=L2` in the log, the refusal on the message line, a
blocked edit recorded by the daemon, no change and no new tracked change in
Word. The run then takes its own item off the file. The two steps that show
whether the model avoids the word (`propose R1.3` and the knob run) stay, and
S2-23 no longer depends on them.

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
  --plan-repo <repo path the feature_list commands cd into>
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
| S25-37          | `propose-all` (with Claude)                                                             | `s25-propose-all`: `overlaps`, `proposed`, `open_before`                                                                   |
| S25-38          | Word COM deletes a sentence, `focus-refresh`                                            | `s25-focus-refresh`: the auto-refresh line changed and an outside change is listed                                         |
| S25-41          | cleanup (Word quit, daemon stopped), then the verify commands                           | `s25-cleanup`: own Word gone, earlier Word processes alive, ports closed                                                   |

Review turns run with Claude. Codex is a provider for judgments only (the
Review turns menu does not list it), so the flow has no step that chooses it
for a review turn and the feature ids S25-33 to S25-36, which were about such
turns, have no steps and no evidence. The pane outputs these judgments were unit-tested with are written by
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

### Stage 3 with `--check-now` and `--figures-button`

```sh
node scripts/e2e/run-acceptance.mjs --stage s3 --check-now --figures-button \
  --base-dir <folder on a Windows drive, as a WSL path> \
  --playwright-core 'D:\path\to\node_modules\playwright-core' \
  --feature-list <stage 3 feature_list.json> \
  --plan-repo <repo path the feature_list commands cd into>
```

Two buttons of the Review tab that had only been tested without Word. Both
flags are off by default; without them stage 3 runs as above. Stage 3 carries
them because it is the cheapest flow that exports a response letter (stage 4
has no review loaded, so "Before submission" is hidden, and stage parts runs
the turns of two documents). Each judged row is printed as
`PASS|FAIL  check now (<press>): <row id>  <state>: <the text the pane shows>`
and kept in `result.json` (`check_now`, `figures_button`).

`--check-now` presses "Check now" (pane step `submission-check`) three times
(`submissionCheck` in `run-acceptance.mjs`, judged by `submissionEvidence` in
`judge.mjs`):

1. `unsaved`: right after the export of S3-40, while the manuscript holds the
   run's Word edits that were never saved.
2. `saved`: after `word-com.ps1 save` of the manuscript (in the run folder).
3. `after S3-39`: at the end of the stage, after the German draft of R1.1.

| Row                                              | Expected, from what the run reads itself                                                                                                                                                                                                                                                                                                                          |
| ------------------------------------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `job`                                            | rows drawn, no error on the message line, the daemon logged `[tools] submission_check rows=...`                                                                                                                                                                                                                                                                   |
| `saved`                                          | the `saved=` of that line is `Document.Saved` as `word-com.ps1 state` reads it just before the press                                                                                                                                                                                                                                                              |
| `changes_decided`, `tool_comments`, `file_clean` | ready exactly when Word over COM lists no revision; no comment that starts `[R<n>.<m>]` or `[E.<m>]` (`TOOL_COMMENT_PREFIX`); no revision and no comment                                                                                                                                                                                                          |
| `marked_copy`, `anonymized_copy`                 | Word says unsaved: not ready with the text "This document has changes that are not saved. ...". Saved: no copy is on record in this stage, so not ready with "No marked copy is on record for this file. ..." and not the unsaved text. The anonymized row may instead be "not needed" when the profile asks for no copy                                          |
| `letter`                                         | the run hashes the file `response_export.path` names. When the record says `fail` 0 and the hash is `docx_sha`: ready and naming that file at presses 1 and 2; at press 3 not ready "... changed after <file> was exported ..." when the answers, a decline or `profile_overrides` changed since press 2 (`letterInputsChanged`), else ready. Otherwise not ready |
| `figures`                                        | not needed: no figure change is on record in this stage                                                                                                                                                                                                                                                                                                           |
| `as_of`                                          | the line under the rows says the document has unsaved changes, or "Press Check now again after a change."                                                                                                                                                                                                                                                         |
| `answers`, `requests`, `format_counts`           | `NOTE`: printed with their text, not judged (the run has no reading of its own for them)                                                                                                                                                                                                                                                                          |

Press 1 is followed by the row `check now: the manuscript saved through
Word`. When Word already says `Saved` before it, a `NOTE` says that the rows
of an unsaved document were not exercised. The save changes the manuscript
on disk before S3-39; nothing in the flow reads that file afterwards, but a
verify command of the feature list that does would see it.

`--figures-button` presses "Figures and other files" under More (pane step
`figures-open`, `figuresButton`, judged by `figuresOpenEvidence`), after the
first two presses of `--check-now`. The button is the one of the whole
manuscript under More (`#review-figures-all`, `figuresLink()` without a
comment), so the page is asked for with `doc=` and `ref=null`; a comment card
has a button of its own only when the comment has figure work
(`review-pane.js` `hasFigureWork`), which stage 3 has not. The step records
whether the pane's own click handler prevented the link
(`default_prevented`: `openInBrowser` returned true, so the API was called
and did not throw), wraps `Office.context.ui.openBrowserWindow` to see the
address it is called with (`wrapped` only when a fresh read of
`Office.context.ui.openBrowserWindow` returns the wrapper; `installed_on`
says whether it went on the object or on its prototype), lists the pages of
the WebView2 before and after, and clicks. The daemon logs each load of the
page itself (`figure-http.mjs`: `[figures] page GET doc="..." ref=...
ua="..."`, quoted as JSON). Afterwards the run shows the Response tab again
(`show-tab`, nothing pressed in it): S3-39 fills fields of that tab without
opening it, and the first run with these flags (2026-10-03) left the Review
tab shown, so those fields were hidden and the fills timed out. The rows:

| Row                | Passes when                                                                                                                                                                                                                                                                                              |
| ------------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `clicked`          | the link was there and its address names the document (`doc=`)                                                                                                                                                                                                                                           |
| `route`            | the handler prevented the link (`route=api`) with no error, and, when the wrapper held, one call with that address; or the API is missing and the link was followed (`route=fallback`). Fails on `fallback-despite-api` (the API is there, the link was followed) and `unknown` (the click was not seen) |
| `served`           | within `--figures-timeout` seconds (default 30) the daemon logged a page load with that `doc`                                                                                                                                                                                                            |
| `outside the pane` | the pane stayed on its page and no `/figures/` page appeared among the pages of Word's WebView2                                                                                                                                                                                                          |
| `user agent`       | `NOTE`: the browser's User-Agent; it is not used to tell a browser from the WebView2                                                                                                                                                                                                                     |

The browser window the button opened is closed only when it is surely the
run's own (`browserCloseDecision`, then `browser-windows.ps1 close`): a
process that started after the click, of a browser (msedge, chrome, firefox,
brave, opera, vivaldi) no process of which ran before the click, the only such
new window, its title "Sidewrite figures and other files", and exactly one
tab in it by UI Automation. Then `CloseMainWindow`, never a forced stop.
Otherwise the window stays and the row `figures button: the browser window it
opened` says `left open: <why>`. With a browser of the user already running
(Edge keeps background processes for startup boost) the page opens as a tab
of the user's own window, and it is always left open.

What these two steps assume, and what the run shows when it is wrong:

| Assumed                                                                                                                                                                | When it is wrong                                                                                                                                                                             |
| ---------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| a paragraph holds `w_ins` or `w_del` in the pane's snapshot exactly when Word over COM lists a revision there (`RevList` of `word-com.ps1`, every kind of `Revisions`) | `changes_decided` and `file_clean` fail with `ready: ...` while COM counts n revisions: Word has a revision the snapshot does not see (a formatting change, a paragraph mark)                |
| the snapshot lists every comment, resolved ones too, as COM does (`collectReviewSnapshot` loads `items/resolved` and keeps them)                                       | `tool_comments` reads ready while COM counts n comments of the tool                                                                                                                          |
| `Word.Document.saved` in the pane agrees with `Document.Saved` over COM                                                                                                | `check now (...): saved` fails with `pane sent saved=true, COM Saved=false` (or `null`), and `marked_copy` then shows the record's text where the unsaved one was expected                   |
| COM edits without a save leave `Document.Saved` false (press 1)                                                                                                        | a `NOTE` "the rows of an unsaved document were not exercised"; the press is judged as a saved one                                                                                            |
| the letter model computed by Check now, on the record refreshed in memory, hashes like the one the export wrote                                                        | `letter` fails at press 1 or 2 with `not_ready: The answers, the accepted changes or the line numbers changed after ... was exported`. That is a finding about the hash, not about the judge |
| `word-com.ps1 save` works on the manuscript while its pane is open                                                                                                     | `check now: the manuscript saved through Word` fails with Word's error; press 2 is then judged on an unsaved document                                                                        |
| `Office.context.ui.openBrowserWindow` can be wrapped where the pane reads it (the first run read `wrapped=true` with no call; the old check was wrong)                 | `wrapped=no (address not watched)` in the `route` row: the row is then judged on `default_prevented` alone; `served` still checks the address the browser asked for                          |
| a `click` listener on the document runs after the link's own handler and sees its `preventDefault`                                                                     | `route=unknown default_prevented=-`: the row fails                                                                                                                                           |
| Office opens the address in the default browser of Windows (OpenBrowserWindowApi 1.1)                                                                                  | `served` fails with `no page load logged`. When the link fallback opened the page inside Word instead, `outside the pane` fails and names the page                                           |
| a Word on Windows reaches the daemon at 127.0.0.1:47834 from the browser (WSL forwards localhost)                                                                      | `served` fails; the browser shows a page that cannot be reached                                                                                                                              |
| the browser sets the window title to the page's `<title>` within 10 s of the page load                                                                                 | the window is left open: `left open: no new browser window shows "Sidewrite figures and other files" ...`                                                                                    |
| UI Automation lists the tabs of Edge and Chrome as `TabItem` elements of the window                                                                                    | the window is left open: `left open: the window has 0 tabs` (or more)                                                                                                                        |

## Stage cite

```sh
node scripts/e2e/run-acceptance.mjs --stage cite \
  --base-dir <folder on a Windows drive, as a WSL path> \
  --playwright-core 'D:\path\to\node_modules\playwright-core'
```

Do the two edit tools leave reference-manager fields whole? The document is
`tests/fixtures/citation_manuscript.docx` (`make_citation_fixture.py`, every
sentence and reference invented): eleven Zotero citations
(`ADDIN ZOTERO_ITEM CSL_CITATION {json}`, some instructions split over
several runs, two with several items, one before a full stop, one next to
italic and superscript runs), the Zotero bibliography (`ADDIN ZOTERO_BIBL`,
one field over six paragraphs), `ZOTERO_PREF_1` and `_2` in
`docProps/custom.xml`, two EndNote citations (`ADDIN EN.CITE`, one with the
record in the instruction, one with it in `w:fldData` and the nested
`ADDIN EN.CITE.DATA` field), the EndNote list (`ADDIN EN.REFLIST`), the
`EN.*` document variables, and as controls a `REF` field and a hyperlink.
The head of the script says where each shape comes from and which parts are
not confirmed from a source. The EndNote shapes are among those: they were
written from secondary sources and were not compared with a document a real
EndNote wrote, and no EndNote plugin is installed on the machine the stage
ran on (2026-10-02; the row of the plugins says what the Word of a run has).
The review is `tests/fixtures/cite_review.txt` (four invented comments). No
feature list, no ledger, no `--conditional`.

The rule the tools are held to: an edit either leaves every field exactly as
it was, or removes a field as a whole (begin through end inside the tracked
deletion, so accepting removes the field and rejecting restores it). Nothing
in between. What is judged is `tools/citation_fields_check.py` on the files
Word saved, and for the edits of the run itself also the answer of the tool
(the column `expect` of `CASES`). The answers of a model turn are recorded
and not expected.

```sh
python3 tools/citation_fields_check.py --before A.docx --after B.docx [--expect-removed N] [--fld-data judge|note] [--detail]
```

It reads both files without Word: every field of every story (complex
fields with `w:instrText` and, inside a tracked deletion, `w:delInstrText`;
`w:fldSimple`; `w:hyperlink`). B passes when every story is balanced (no
`end` or `separate` without a `begin`, no `begin` left open), no field is
split by a tracked change (its markers and its instruction lie all outside a
change, all in deletions, or all in insertions), no field is hollowed, no
result of a field differs from the result the same field has in A, the
instructions are those of A as a multiset of exact texts, the `w:fldData`
are those of A, and the `ZOTERO_*` properties and the `EN.*` document
variables are the same. With `--expect-removed N`, B lacks exactly N
instructions of A and has no other.

A field is hollowed when it is still there (its markers are not all inside
tracked deletions) and nothing of its result is left once the changes are
accepted, although it had text: the result lies whole in a tracked deletion,
or, after accept-all, the field has no result run at all. That is what three
separate deletions of "the words before the citation", "the visible text of
the citation" and "the full stop after it" leave (the first real run,
2026-10-02): the citation looks deleted, the field stays, and the plugin
writes the text again at its next refresh and keeps the reference in the
bibliography. By the counts, the instructions and the balance such a file is
whole, so it has a failure of its own, `hollowed`. A field whose markers and
result all lie inside tracked deletions is deleted as a whole (`in
deletions`) and passes. A result that is partly deleted, or has other words
than in A, fails as `result-changed`: the text of a field is written by the
reference manager, and a tool that edits it leaves a citation that no longer
says what its instruction records. Both are read for complex fields only,
not for `w:fldSimple` and `w:hyperlink`.

`--fld-data note` prints a difference of `w:fldData` as a `NOTE` line and
does not fail on it. The run passes it when Word's own save of the untouched
fixture already changed `w:fldData` (the first row of the stage): Word
writes the invented EndNote data of the fixture anew on every save, with
another value each time, so the difference says nothing about a tool. When
the first row finds `w:fldData` kept, it is judged in every later row.
For each failure it prints one line with the story and paragraph, the first
40 characters of the paragraph, the kind of the field and its markers before
and after, for example

```
FAIL split-field | word/document.xml p3 "Shaken flasks remain the usual first ves" | kind=ZOTERO_ITEM | before: begin@- instr@- separate@- result@- end@- | after: begin@del instr@del separate@- result@- end@-
```

(`@-` outside a tracked change, `@del` in a deletion, `@ins` in an
insertion), for a hollowed or changed result also `| result: "before" ->
"after"`, then one JSON line. Exit code 0 only when everything holds. A
live citation whose visible text is no longer the text its instruction
records (`plainCitation`, `DisplayText`) is listed under `text_changed` and
does not fail on that alone: A is not looked at for it, so a citation the
author had changed by hand is listed and passes, while one that differs
from A fails as `result-changed`. Nor does a field under `flattened` fail: it lies whole in a tracked
deletion while its visible text stands again in the same paragraph as plain
inserted text, which is what Word makes of an edit whose range held the
field. Nothing is half, but once that change is accepted the citation is
plain text and the plugin no longer knows it. With `--detail` (the run uses
it) the JSON line also lists every field with its state and the markers of
every paragraph that holds one.

The steps, each row of the table of the run:

| Row                                                                                                                                                           | What is done                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            | Judged                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| ------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `cite: the Zotero and EndNote plugins of this Word`                                                                                                           | `word-com.ps1 addins`: the COM add-ins and global templates of the Word of the run, read only                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           | no (`NOTE`); `cite.plugins` in `result.json`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     |
| `cite: Word's own save keeps the fields of the fixture`                                                                                                       | `word-com.ps1 save` before any edit, the file copied to `cite/baseline.docx`, the checker on the fixture against it                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     | no (`NOTE` when it differs): every later file is compared with this baseline, so what Word itself changes in the fixture is not taken for the work of a tool. When `w:fldData` differs here, the row says so and every later check runs with `--fld-data note` (`cite.fld_data` in `result.json`)                                                                                                                                                                                                                                                                                                                                                |
| `cite: next-to`, `spans`, `straddles`, `cite-text-zotero`, `cite-text-endnote`, `cite-text-replace`, `delete-sentence`, `delete-sentence-zotero`, `formatted` | one tracked edit each (`CASES` in `make_citation_fixture.py`), saved to `cite/case-<name>.docx`. Expected `applied`: a phrase next to a Zotero citation; a phrase next to a field in the paragraph with italic and superscript runs. Expected `refused`: a replacement whose find holds the whole visible text of a citation; a find that ends inside it; the visible text of a Zotero citation alone, deleted; the same for the EndNote citation of the first real run; the visible text of a citation replaced by other text. Expected `removed`: a whole sentence with an EndNote citation, and one with a Zotero citation, each deleted as ONE edit | the checker on the copy before the edit against the copy after it (no split field, no hollowed field, no changed result; the row gives the count of each), and the answer of the tool: `applied` with no field newly in a deletion; `refused FieldNotEditable` with Word unchanged; `applied` with exactly one field newly whole in a deletion. The row reads `tool: <answer> [why=<reason> by=<daemon or pane> ...] \| expected <...>: met or NOT MET (fields newly in deletions: n) \| fields ...`, and the message of a refusal is printed under it. Also fails: no readable answer, or a refusal after which Word has another tracked change |
| `cite: rewrite`                                                                                                                                               | the first part of a whole-paragraph rewrite of a paragraph with two citations, sent with `plain_only`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   | the one expected answer: refused `ParagraphNotPlain`, no tracked change added, the checker passes                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| `propose R1.1`, `R1.2`, `R1.3` and `cite: fields after the turn of ...`                                                                                       | one model turn per comment (reword a sentence that cites; delete a sentence that cites; rewrite a paragraph with three citations), each saved to `cite/turn-<ref>.docx`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 | the turn ran; the checker on the copy before the turn against the copy after it. The row tells the changes recorded, the rewrites tried (`stopped=ParagraphNotPlain`) and the refusals of the turn, from the daemon log                                                                                                                                                                                                                                                                                                                                                                                                                          |
| `cite: saved with the changes tracked, every field whole`                                                                                                     | `cite/tracked.docx` against the baseline                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                | the checker; marker `ok/cite-tracked`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            |
| `cite: citations whose text differs from what their field records`                                                                                            | the `text_changed` list of that check                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   | no (`NOTE`)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |
| `cite: citations an edit replaced by plain text`                                                                                                              | the `flattened` list of that check; the row of the edit or turn that did it says so as well                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             | no (`NOTE`): the fields are whole, and the citation is lost once the change is accepted                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          |
| `cite: Word opens the saved file as it is, accepts and rejects`                                                                                               | after the Word of the run is closed and the daemon stopped: `marked-diff-word.ps1` in a Word of its own opens `cite/tracked.docx` read-only without a repair, runs `Revisions.AcceptAll()` and `RejectAll()` and writes `cite/accepted.docx` and `cite/rejected.docx` (`save_as` of a job)                                                                                                                                                                                                                                                                                                                                                              | every file opened, no revision left                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| `cite: after reject-all the fields are the original ones`                                                                                                     | the checker, baseline against `rejected.docx`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           | equal, every field live                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          |
| `cite: after accept-all only the fields that lay in deletions`                                                                                                | the checker, baseline against `accepted.docx` with `--expect-removed <fields that lay whole in a deletion of tracked.docx>`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             | passes, every field live, and the instructions that are gone are exactly those of the deleted fields. The row names them: a citation an edit replaced by plain text is whole by this rule and is listed here                                                                                                                                                                                                                                                                                                                                                                                                                                     |
| `cite: Word lists the original fields after reject-all`                                                                                                       | Word's own `Field.Code` and `Field.Result` of every paragraph, baseline against reject-all                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              | equal; marker `ok/cite-accept-reject` holds all of the last four rows                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            |

How the edits of the run reach the tools. A phrase edit goes the way of the
forbid guard of stage 2: it is put on file as a blocked edit of R1.4 (the
store's own `addBlocked`) that names a rule the ledger does not have, and the
pane step `apply-seeded` presses "Allow exception", so the daemon sends it
through `applyTrackedEdit` to `office_tracked_edit` as it sends an edit of a
turn. The rationale comment of such an edit ends in the note of an approved
exception. `office_tracked_rewrite` has no such way in without a model: the
call it makes first is built with the daemon's own functions
(`citeRewriteCall` in `judge.mjs`: `rewriteSegments`, `narrowEdit`,
`paragraphHash`) and handed to the pane's `toolTrackedEdit` by the step
`tracked-edit-direct`, which is where a paragraph that is not plain text is
refused. The daemon's half of a rewrite runs only when a model turn chooses
the tool; the row of that turn shows it.

What the stage does not do: it starts no Zotero and no EndNote, presses no
Refresh of a plugin and reads no library. So it cannot tell whether a plugin
still updates the document afterwards; it tells whether the fields the
plugin reads are still there, whole and unchanged. Each failing check prints
the checker's lines under its row, and `result.json` keeps them under
`cite.checks.<name>` with the full result. A field one edit left split or
hollowed is still so in every later copy while the change is pending, so
the rows after it fail as well: the paragraph named in the line tells which
edit it was. An edit that is expected to go in has a paragraph of its own.
The edits that must be refused share one with a later edit (the two in the
baffles paragraph with the turn of R1.3, `cite-text-endnote` with the turn
of R1.2, `delete-sentence-zotero` with the turn of R1.1): when one of them
goes in against the expectation, its own row fails first.
`--screenshots` works as in the other stages.

How the tools keep the rule (`taskpane/shared/field-text.js`, used by the
daemon and the task pane alike). Before Word is called the daemon places the
edit in the paragraph as the last status refresh gave it, which shows a field
as U+0013 instruction U+0014 result U+0015: a range inside the text of a
field, across its edge, or a replacement over a field is refused as
`FieldNotEditable` (log line `[review] field edit refused <ref> why=<inside,
partial or replace-across> by=daemon paragraph=<n>`). The task pane does the
same on its own reviewed text. A deletion that holds a field whole with
words on both sides goes to Word as `Range.delete()` on the range the search
found. After every edit in a paragraph with fields the pane reads the fields
again from the paragraph XML and compares them with those read before; a
field that is hollowed, split, changed or gone makes it reject the tracked
changes on the range of the edit and refuse (`why=after-edit by=pane
found=<what> undone=1`). When Word cannot take the edit back the tool
answers `FieldEditNotUndone`, the change stays on record without a comment
and the model is told to leave it to the author.

What the undo came to is read from the document (`undoOutcome` in
`field-text.js`): the fields of the paragraph, and its reviewed text before
the edit, after the edit and after the undo. Fields as before and the text
as before the edit: a clean undo, refused as above. A field still not as it
was, or the text still as after the edit: the edit is in the document
(`FieldEditNotUndone`). Fields as before and a text that is neither: the
edit is gone and so is something else, because Word had merged the edit with
a pending change of the same author next to it and rejected both. Then the
edit is refused and not recorded (`... undone=1 took_more=1`, then
`[review] undo took more <ref> paragraph=<n>: refresh asked`), the pane
refreshes the status, which reads the change that vanished from the
document, and the card of the comment tells the author to check the
paragraph. A text that could not be read counts as "the edit is in the
document".

A paragraph whose fields nobody has seen is not edited. The status refresh
records which paragraphs came without reviewed text and whether the document
holds a field anywhere (the pane looks for one in the XML of every
paragraph, `has_fields`). An edit or a rewrite in a paragraph without
reviewed text, or after one (a field over several paragraphs would begin
there), is refused as `FieldsUnread` (`[review] fields unread <ref>
paragraph=<n> by=daemon`) unless the document is known to hold no field; a
snapshot written before this was recorded is refused the same way until the
next refresh (every Propose refreshes first). A rewrite of a paragraph that
lies in a field over several paragraphs is refused by the daemon
(`stopped=ParagraphNotPlain`), which knows the carry from the paragraphs
before it; the task pane sees no marker in a middle paragraph.

What that assumes of Word, and what a row shows when it does not hold. None
of the Word side could be run when the tool was written; the first two lines
were seen in the run of 2026-10-02.

| Assumed                                                                                                                                                              | The row when it is wrong                                                                                                                                                                                                                                                                                                                                                               |
| -------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Range.getReviewedText` shows a field as U+0013 instruction U+0014 result U+0015 (seen)                                                                              | the cases expected `refused` are not refused beforehand: their rows read `why=after-edit by=pane` (taken back) and not `why=inside by=daemon`, or `applied` with `hollowed 1`                                                                                                                                                                                                          |
| a search finds a sentence across a field (seen: the pane answered with its paragraph), and the range it finds runs from before the begin marker to after the end one | `delete-sentence`: `refused NotFound`, or `refused FieldNotEditable [why=after-edit by=pane ... found=split undone=1]` with `NOT MET`, the fields unchanged                                                                                                                                                                                                                            |
| `Range.delete()` with change tracking on puts the whole range, field markers included, into a tracked deletion                                                       | the same row: `why=after-edit` with `found=split`, `hollowed` or `count_a->b`, `NOT MET`                                                                                                                                                                                                                                                                                               |
| `Paragraph.getOoxml()` shows the tracked deletions and the field markers of the paragraph as the saved file does                                                     | a good deletion is taken back (`found=count_a->b`, `NOT MET`), or a hollowed field is not seen: `applied`, and the checker fails the row with `hollowed`                                                                                                                                                                                                                               |
| `Range.getTrackedChanges().rejectAll()` on the range of the edit takes back that edit and nothing else (WordApi 1.6)                                                 | took back less: `applied (delete, no comment: FieldEditNotUndone)`, the daemon log has `[review] field edit not undone`, and the checker names the field. Took back more (a pending change next to it): `refused FieldNotEditable [why=after-edit by=pane ... undone=1 took_more=1]`, the log has `[review] undo took more`, and Word lists fewer tracked changes than before the edit |
| `Office.context.requirements.isSetSupported("WordApi", "1.6")` is true in this Word (the paragraph ids of the run are the 1.6 ones)                                  | `delete-sentence`: `refused FieldNotEditable [why=no-undo by=pane ...]`, `NOT MET`, nothing changed                                                                                                                                                                                                                                                                                    |
| after such a deletion the reviewed text of the paragraph no longer shows the field                                                                                   | the next edit in that paragraph is refused once as `StaleTarget` and the pane refreshes: the turn of R1.1 (the paragraph of `delete-sentence-zotero`) reads `stale 1` or more, and the refresh may list the paragraph as changed outside the tool                                                                                                                                      |
| the tracked-deleted sentence can still be found for the rationale comment                                                                                            | `applied (delete, no comment: CommentAnchorMissing)`: the change is recorded, its comment is missing                                                                                                                                                                                                                                                                                   |

Not covered: a paragraph inside a field that runs over several paragraphs (a
bibliography) is refused by the daemon, which sees the whole document; the
task pane alone sees it only in the first and the last paragraph of the
field. `w:fldSimple` and `w:hyperlink` are guarded through the reviewed text
and are not read by the check after the edit. A Word that gives no reviewed
text: an edit goes in as before only when the pane found no field in the XML
of any paragraph; otherwise it is refused (`FieldsUnread`).

The Word side of the stage (`word-com.ps1 save` and `addins`, `save_as` of
`marked-diff-word.ps1`, the pane steps `apply-seeded` and
`tracked-edit-direct`) had not run in Word when it was written; its judges
were tested with inputs written by hand (`tests/e2e-judge-cite.test.mjs`).

`--provider <id>` (stage 2 flow) chooses the provider of the review turns
in the pane's Providers section right after the Review tab opened (step
`select-provider turn=<id>`). Only `claude` is offered for review turns; any
other id is not chosen: the step fails with what the menu shows, and the
turns run with Claude. Stage 2.5 has no condition for `--conditional`.

Other options: `--cdp-port`, `--win-node`, `--fixture`, `--review`,
`--ledger`, `--features S2-18..S2-34`, `--skip-knob`, `--skip-mode14`,
`--keep-on-fail` (leave Word and the daemon up after a failed step).

`--screenshots` (off by default) saves two PNGs of the task pane after every
pane step in `shots/` of the run folder: `NN-<step>[-<comment or tab>].png`
is the window of the pane as the user sees it, with what the step was for
scrolled into view (the card of its comment, else the message line), and
`...-full.png` is the whole tab. The pane scrolls inside an element, so a
full-page picture alone is the window again; for the whole tab the elements
it scrolls in are let grow for the moment of the picture and every style and
scroll position is put back. A step that takes that long is photographed 4 s
and 20 s in as well (`-wait1`, `-wait2`, both files), with nothing scrolled:
that is what the click left the user with. A step that fails is photographed
too.

`shots/timings.json` has a row per step: the seconds from the start of the
step to its answer, the seconds of the turn as the step measured them, the
size of the pane's window in CSS pixels and its pixel ratio, what the pane
said during the wait (message, progress line, status, the models line),
whether the progress line was in view then (`progress_in_view`), and whether
the step waits for a model.

Stage 2 adds pictures that are not steps of the run and are not judged: every
tab before anything is pasted, the chat and Setup after the turns, the answer
of "Check revised paragraphs", the Word window of the run once the first
tracked change is in, and after the verify commands the pane with the daemon
stopped ("Refresh status" pressed), after which the daemon is started again
for the cleanup. The Word window is drawn by the window itself
(`word-shot.ps1`, PrintWindow at the real pixel size), so no other window is
in the picture and it is not moved or focused; when more than a tenth of it
comes out black (the WebView2 pane can), no file is written and
`word_window_shot` in `result.json` says `dropped` with the share.

## Stage parts

```sh
node scripts/e2e/run-acceptance.mjs --stage parts \
  --base-dir <folder on a Windows drive, as a WSL path> \
  --playwright-core 'D:\path\to\node_modules\playwright-core'
```

A manuscript and its supplementary Word file in one revision (README
"Supplementary Word documents", steps A1 to A3): the link, the copied
comment list, a turn in each file, the forbid guard in the supplementary
file, where a request stands in the other file, a turn in both panes at the
same time, the response letter of the manuscript with a change of the
supplementary file (accepted, then changed behind Sidewrite's back, then
rejected), and leaving. Both documents are open in the Word of the run, each
with its pane, on one daemon; rows k to m close and reopen `SI.docx`.

The files are invented (`make_parts_fixture.py`, `--spec` prints what the
stage reads): `paper.docx` (`tests/fixtures/parts_paper.docx`), `SI.docx`
(`parts_si.docx`: supplementary methods and one table), the review
`parts_review.txt` of four comments and the ledger `parts_ledger.json` with
one forbidden word, "novel" (rule L1).

| Comment | Asks for a change              | The stage uses it for  |
| ------- | ------------------------------ | ---------------------- |
| R1.1    | only in the supplementary file | c, d, e, f, h, k, l, m |
| R1.2    | in both files                  | g                      |
| R1.3    | only in the manuscript         | (context)              |
| R1.4    | in a figure file               | (context)              |

No feature list, no `--conditional`. `--supplement`, `--fixture`, `--review`
and `--ledger` replace the files; `--provider` chooses the provider of the
review turns in both panes. The requests of each comment are typed by the
run (`edit-draft`), so their ids do not depend on a model.

The rows of the table of the run, in order. A judged row writes
`ok/parts-<letter>.evidence.json` (and `.ok` when it passed) and prints, for
each check that failed, which pane or file, what was expected and what was
found (`judge.mjs` `partsVerdict`); the table at the end cuts that line, the
line printed when the row was judged and `result.json` have it whole.

| Row | What is done                                                                                                                                                                    | Judged (`judge.mjs`)                                                                                                                                                                                                                                                                                                                          |
| --- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| a   | the manuscript gets its review, list and rule; `SI.docx` is opened in the same Word (`word-com.ps1 open-edit`)                                                                  | `partsBeforeJoin`: the pane of `SI.docx` shows the join box with "Join the revision of paper.docx, 4 comments" and no comment row; neither review file holds a link                                                                                                                                                                           |
| b   | the candidate button is pressed in the pane of `SI.docx`; the pane of `paper.docx` is only watched (`wait-linked`)                                                              | `partsJoined`: four rows in `SI.docx`, the rule shown with no Edit or Remove button and the note "These rules are edited in paper.docx.", Propose changes in every card, `part_of` and `parts` in the two files, the same request ids, "Linked files: SI.docx" in the manuscript's pane                                                       |
| c   | Propose changes for R1.1 in the pane of `paper.docx`                                                                                                                            | `partsMainOutside`: every request of R1.1 is in `requirements_outside` of `paper.revision.json`, no change was recorded, Word lists no new tracked change in `paper.docx`                                                                                                                                                                     |
| d   | Propose changes for R1.1 in the pane of `SI.docx`                                                                                                                               | `partsSuppChange`: a new change under R1.1 in `SI.revision.json`, every one with its comment; Word lists more tracked changes in `SI.docx` and a comment that begins `[R1.1]`; none in `paper.revision.json`; the paragraph texts of `paper.docx` (pane step `snapshot`) are the same before and after                                        |
| e   | an edit that writes "novel" is put on file for R1.1 of `SI.docx` under a rule the ledger does not have (L9) and sent with "Allow exception", as `forbidGuard()` of stage 2 does | `forbidBlockEvidence` with the review file and the Word state of `SI.docx`: the pane shows the refusal by L1, the daemon logged `[review] blocked R1.1 by=L1`, recorded the blocked edit, Word has no new tracked change                                                                                                                      |
| f   | nothing is pressed in the pane of `paper.docx`: the step waits for its row of R1.1 (`wait-state`)                                                                               | `partsCrossState`: the row says "n change(s) to review in SI.docx" and the request the change names shows "Waiting for your decision in SI.docx"                                                                                                                                                                                              |
| g   | Propose changes for R1.2 in both panes, started together (two driver processes)                                                                                                 | `partsBoth`: both steps ended; both `[review] prompt R1.2` lines come before the first `[review] turn done R1.2`; a new change in each review file; in each document more tracked changes and between one and that many new `[R1.2]` comments; no log line with `outcome unknown`, `from another pane`, `unknown id` or `late result dropped` |
| g2  | nothing is done: the log of the whole run is read                                                                                                                               | `partsSessions`: each document has at least one `[agent] turn doc=<path> session=<id> asked=<id or -> (started, resumed or continued)` line, and no session id of one document (run on or asked for) is one of the other                                                                                                                      |
| h   | in the pane of `paper.docx` the review is split again and saved without R1.1 (`save-without`)                                                                                   | `partsKeep`: the message line shows "Comment R1.1 has recorded change(s) in SI.docx; keep it" as an error; the comment list in `paper.revision.json` is as it was                                                                                                                                                                             |
| i   | `SI.docx` is saved through Word (`word-com.ps1 save`), then its pane refreshes                                                                                                  | not judged, see below                                                                                                                                                                                                                                                                                                                         |
| j   | "Leave this revision" in the pane of `SI.docx`                                                                                                                                  | `partsLeft`: the join box is shown again with the manuscript as a candidate, `part_of` and `parts` are gone, every change `SI.revision.json` held is still in it                                                                                                                                                                              |

Rows k, l and m run between row i and row j (the leave stays last). Each is
one judged row; a step that throws fails its row with `the step threw: ...`
and the stage goes on. `ok/parts-k.evidence.json` and the two others hold
what every step answered.

- **k, the letter counts and cites the accepted change** (`partsCited`). In
  the pane of `SI.docx` the file is named "Supplementary Information"
  (`link-label`). `SI.docx` is saved and copied to `parts/SI.pending.docx`
  (the file with the change still tracked, for row l). The changes of R1.1
  are accepted through Word (`word-com.ps1 apply` with
  `parts/accept-plan.json`, the accept step of stage 3 limited to one
  comment, `partsPlan`), the file is saved, its pane refreshes, and the file
  is closed without saving (`close-other`). In the pane of `paper.docx`:
  Response tab, R1.2 to R1.4 declined with a reason, Draft for R1.1 (a
  model), Run checks, and, only when no check fails, Export. Judged: the
  link carries the name; Word applied every op; the review file of `SI.docx`
  has every change of R1.1 accepted, `doc_state.saved` true with the time
  and size the file has after the close, and citations for R1.1 without line
  numbers; the card of R1.1 has an answer and a location line that starts
  with "Supplementary Information ["; no `FAIL R1.1 ...` and no
  `FAIL part_state` row; the message "Exported..."; the letter model
  (`paper.response.model.json`) has a `location` paragraph of R1.1 that
  starts with "Supplementary Information [Revised in " and a `quote`
  paragraph right under it whose text is a quote the review file of
  `paper.docx` records for `SI.docx`; the daemon logged
  `[response] export path=... verify=ok` (the verifier ran on the paragraphs
  of both files).
- **l, a file changed without a refresh is not trusted** (`partsGuard`).
  Word cannot reject a change it has accepted, so `parts/SI.pending.docx` is
  copied over the closed `SI.docx`, the file is opened again (`open-edit`),
  the changes of R1.1 are rejected through Word (`parts/reject-plan.json`),
  the file is saved and closed. Its pane is not driven. In the pane of
  `paper.docx` the Response tab is refreshed (the button then reads "Export
  letter (.docx)") and Export is pressed once. Judged: every Word step
  answered ok; the file on disk has another time or size than
  `doc_state.file`; the button did not offer to force; exactly one row
  `FAIL part_state: ... SI.docx ... Open it, save, press Refresh.`; the
  message "Not exported: ...", the letter files of the run folder and
  `response_export` as they were; the daemon logged
  `[response] export held fail=n`.
- **m, refreshed again** (`partsReopened`). `SI.docx` is opened, its pane
  refreshes. In the pane of `paper.docx`: Response tab, Run checks. Judged:
  every change of R1.1 is `rejected` in the review file of `SI.docx`, whose
  `doc_state` is saved and of the file on disk; no `FAIL part_state` row; a
  row `FAIL R1.1 requirements: ...` and the request shown open on the card;
  a stale line on the card of R1.1.

One row is told and not judged (`NOTE`):

- **Document.saved and the file on disk** (row i). The pane step `doc-saved`
  reads `Word.Document.saved` in the pane of `SI.docx` at four points: right
  after it opened, after the tracked edit of row d, before the save and
  after the save through Word. With each reading: the modification time and
  the size of `SI.docx` as the daemon's side sees them (`fs.statSync` from
  WSL), then `doc_state` as the status refresh after the save wrote it into
  `SI.revision.json`. The row prints the raw property with its type, what
  the pane would send (`true`, `false` or `null`) and the error when Office
  gave none. This is the experiment for the confirmed-state guard of the
  next step (does the property turn false on an edit and true on a save; do
  the time and the size change with a save and only then).

Row g2, the session each pane's turns ran on. The daemon prints one line per
turn without a debug flag (`daemon/sessions.mjs` `turnSessionLine`). The
resume hint of a document is kept per host and folder; a file of a revision
with parts (the manuscript and each supplementary file) has a hint of its
own, by its path, and never reads the one of its folder. So with both files
in one folder each pane starts its own session, and the row fails when one
pane ran on, or was asked to resume, a session of the other. `asked` is read
as well: should the SDK give a resumed session a new id, the row still fails.

What the stage assumes about Word and Office.js, and what the log shows when
an assumption does not hold:

| Assumption                                                                                                                                                                       | When it is wrong                                                                                                                                                                                                                                                                                                                                                                                                                                                             |
| -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Documents.Open` over COM opens a second document writable in the Word of the run                                                                                                | the row "SI.docx opened in the Word of the run" fails with Word's error, or with `read_only=true`, and the stage stops there                                                                                                                                                                                                                                                                                                                                                 |
| `word-com.ps1` finds that document by its full path afterwards (`BindToMoniker`), as it finds the first one                                                                      | the rows that read `SI.docx` in Word (d, e, g, the NOTE of i) fail and print the binding error of PowerShell where the count of tracked changes would be                                                                                                                                                                                                                                                                                                                     |
| a document opened that way shows its task pane by itself (the `Office.AutoShowTaskpaneWithDocument` part `embed_taskpane.py` writes), as the first document does                 | "SI.docx: open-review" fails after 120 s with `pane not reachable on port 9333: no pane page of SI.docx yet (pane pages: ...)`; the list names the documents of the panes that are there                                                                                                                                                                                                                                                                                     |
| the panes of two documents of one Word are pages of the same WebView2 debugging endpoint                                                                                         | the same message, with only the manuscript in the list                                                                                                                                                                                                                                                                                                                                                                                                                       |
| a pane page can be told by `Office.context.document.url` (it ends in the file name), or once it has loaded by its "Review file:" line                                            | the same message with `(a pane that names no document yet)` for a page whose `Office.context.document.url` is empty and whose Review tab has not loaded                                                                                                                                                                                                                                                                                                                      |
| a pane whose document is not the window in front still runs: it answers the daemon and applies tracked edits                                                                     | rows d, e or g fail: the step of the pane in the back ends with `turn ... still running`, or the log has `outcome unknown` lines (row g names them)                                                                                                                                                                                                                                                                                                                          |
| two driver processes can be connected to the endpoint at once                                                                                                                    | row g: "both propose steps ended without an error" fails with the error of the second connection                                                                                                                                                                                                                                                                                                                                                                             |
| the two turns of row g overlap (each takes far longer than the second click)                                                                                                     | row g: "the two turns ran at the same time" fails with the positions of the log lines; the other checks of the row still say whether each change landed in its own files                                                                                                                                                                                                                                                                                                     |
| the turn of the manuscript notes the request of R1.1 as another file's (a model decides it)                                                                                      | row c fails with what the turn did instead: `outside [...], questions n, changes n`                                                                                                                                                                                                                                                                                                                                                                                          |
| the turn of the supplementary file names the request in `rationale.requirement_ids` (a model does)                                                                               | row f: "a pending change of R1.1 names a request" fails with the ids of each change                                                                                                                                                                                                                                                                                                                                                                                          |
| the CLI sends a `system/init` message at the start of every turn, as it did in the runs so far (the line of a turn is printed from it)                                           | row g2: "a turn ran on a session" fails for a document with `no turn`, although its rows c, d or g passed                                                                                                                                                                                                                                                                                                                                                                    |
| `Word.Document.saved` (WordApi 1.1) can be read in the pane                                                                                                                      | the NOTE of row i shows `Document.saved=null` with Office's error; the pane then sends `null` and `doc_state.saved` is `null`                                                                                                                                                                                                                                                                                                                                                |
| a file Word holds open can be read for its time and size from WSL                                                                                                                | the NOTE of row i shows `file not read`, and `doc_state.file` is `null`                                                                                                                                                                                                                                                                                                                                                                                                      |
| k: `Word.Document.saved` is true in the pane right after a save through COM (the NOTE of row i shows what it is)                                                                 | row k: "the refresh was taken on a saved document" fails with `saved false` or `saved null`, and its check row is `FAIL part_state: SI.docx had changes that were not saved ...`                                                                                                                                                                                                                                                                                             |
| k: closing a saved document without saving leaves the time and the size of its file as they were                                                                                 | row k: the same check fails with the two times and sizes                                                                                                                                                                                                                                                                                                                                                                                                                     |
| k: a file Word holds open can be copied from WSL right after a save (as stage cite copies its document)                                                                          | row l: "reopened as it was before the accept" fails with `put back: ENOENT ...`; row k itself is not affected                                                                                                                                                                                                                                                                                                                                                                |
| k: Word keeps the rationale comment of a change when the change is accepted, so the accept step finds each change by its comment                                                 | row k: "the change(s) of R1.1 accepted" fails with `ERR rationale comment not found` or `ERR no revision at the comment` per change                                                                                                                                                                                                                                                                                                                                          |
| k: the turn of row d named the request in `rationale.requirement_ids` (the assumption of row f) and the draft of the answer passes the checks of the profile (a model writes it) | row k: "no failing check for R1.1" fails and prints the rows, for example `FAIL R1.1 requirements: q1 is neither cited nor declined` or `FAIL R1.1 rule_opening: ...`                                                                                                                                                                                                                                                                                                        |
| k: the model file of the letter is `paper.response.model.json` next to the letter                                                                                                | row k: "a location line of R1.1 names Supplementary Information" fails with `[]` although the export passed                                                                                                                                                                                                                                                                                                                                                                  |
| l: a closed `SI.docx` can be overwritten from WSL, and `Documents.Open` opens it again in the Word of the run with its pane                                                      | row l: the first check fails and names the step: `put back: EBUSY ...` (Word still held the file: the close of row k failed), `open: <Word's error>`                                                                                                                                                                                                                                                                                                                         |
| l: the pane of `SI.docx`, which loads by itself when the file is opened, does not run a status refresh between the save and the close (it refreshes when it gets the focus)      | a refresh before the save does no harm: row l passes and its line ends with "the pane of SI.docx refreshed by itself before the save", the check row may then be the "had changes that were not saved" one. A refresh after the save makes the record current: "on disk is no longer the file its review file records" fails with "(the pane of SI.docx refreshed by itself: doc_state.at ... -> ...)", and the part_state check fails with the rows that were shown instead |
| l: the Response tab offers "Export letter (.docx)" after "Refresh from the manuscript", whatever the last check said                                                             | row l: "the export button did not offer to force" fails with the label found; the export of the row was then a forced one and the other checks of the row say what it did                                                                                                                                                                                                                                                                                                    |
| m: a document opened from disk and not edited has `Word.Document.saved` true                                                                                                     | row m: "the refresh was taken on a saved document" and "no failing check for the state of SI.docx" fail, the latter with the `FAIL part_state` row                                                                                                                                                                                                                                                                                                                           |
| l, m: the pane page of a closed document goes away, so the driver finds the pane of the reopened `SI.docx` and not a dead one                                                    | row m: "opened again and refreshed" fails with the driver's error for the refresh (`pane not reachable ...` or a timeout)                                                                                                                                                                                                                                                                                                                                                    |
| l: nothing is asked of `SI.docx` through `word-com.ps1` while it is closed (`BindToMoniker` would load a closed file)                                                            | by construction; `tests/e2e-judge-parts.test.mjs` reads the runner for it                                                                                                                                                                                                                                                                                                                                                                                                    |

Pane steps of the stage (all take the pane of the document named by the
environment variable `E2E_PANE_DOC`, which the runner sets to `paper` or
`SI`):

| Step           | Args                                  | Does and returns                                                                                                               |
| -------------- | ------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------ |
| `link-join`    | `main` (file name)                    | presses "Join the revision of `main`, ..."; the view once the line of the manuscript shows, or with the error message          |
| `link-leave`   |                                       | presses "Leave this revision"; the view after "This file left the revision of ..."                                             |
| `link-label`   | `label`                               | types the name in "Name in the response letter", presses "Save name"; the view after "The response letter calls this file ..." |
| `wait-state`   | `ref`, `until` (RegExp), `timeout_ms` | presses nothing; `matched` is true once the row of `ref` says `until`; the view                                                |
| `wait-linked`  | `until` (RegExp), `timeout_ms`        | presses nothing; `matched` is true once the line of linked files says `until`; the view                                        |
| `save-without` | `review_file`, `drop`, `requirements` | splits the review again, removes the row `drop`, Save; the view after the refusal or after the save                            |
| `doc-saved`    |                                       | `saved` (as the pane sends it), `raw` and `raw_type` (the property as Office.js gave it), `error`                              |

The view of every Review step now also has, per comment, `state` (the line
of its row), `requests` (`id`, `state`, `text`) and `can_propose`, and for
the tab `linked` (the line that names the manuscript or the linked files),
`link_box` (`shown`, `candidates`), `ledger_note` and `ledger_edit_buttons`.

The Word side of the stage (`word-com.ps1 open-edit`, the choice of a pane
by its document, two driver processes at once, the six pane steps above)
had not run in Word when it was written, and neither had rows k to m with
their step `link-label`. Its judges were tested with inputs
written by hand (`tests/e2e-judge-parts.test.mjs`), the choice of the pane
with document names written by hand (the same file), and the wait conditions
and reads of the new steps on an empty page (`tests/e2e-pane-driver.test.mjs`).

## Stage resident

```sh
node scripts/e2e/run-acceptance.mjs --resident \
  --base-dir <folder on a Windows drive, as a WSL path> \
  --playwright-core 'D:\path\to\node_modules\playwright-core'
```

Every other stage stops the installed daemon (`sidewrite-daemon.service`)
and starts one of its own. `--resident` is a short stage that does neither:
the Word of the run, with the add-in registered and the pane opened by the
document (as in every stage), connects to the daemon of the unit on
47833/47834, the Review tab opens (`open-review`: the line "Review file:
<run folder>/paper.revision.json" comes from the daemon's answer) and is read
again (`view`). No model, no write to Word. Nothing of the unit is started,
stopped or restarted; the run ends with "unit is not active" when it is not
running (start it with `systemctl --user start sidewrite-daemon.service`).

The unit runs the checkout it was installed from (its `WorkingDirectory`),
not necessarily the one the runner is started from: the pane and the daemon
the stage exercises are those of that checkout. To test a branch, install the
unit from it (`bash scripts/install/daemon-service.sh install`) or restart it
after that checkout has the code.

The log of the unit is the file its `StandardOutput=append:` line names
(`systemctl --user cat sidewrite-daemon.service`; the install writes
`$HOME/.claude/office-addins/daemon.log`), read from the size it had when the
run started. The daemon's lines are not in the journal, which holds only the
messages of systemd itself; `journalctl --user -u sidewrite-daemon.service
--since @<start>` is read instead only when the unit names no file. The rows:

| Row                                                                | Passes when                                                                                                                      |
| ------------------------------------------------------------------ | -------------------------------------------------------------------------------------------------------------------------------- |
| `open-review`, `view (the Review tab read again)`                  | the pane answered with the review file of the run folder, and read the tab again without an error                                |
| `no tool call on other documents`                                  | no line of the unit's log since the start names a `.docx` outside the run folder (hello lines aside)                             |
| `resident: serving`                                                | before Word started: the unit is active and its `MainPID` holds both ports (`ss -ltnp`)                                          |
| `resident: not restarted`                                          | after Word was closed: the same `MainPID`, and `uptime_s` of `GET /health` grew by the time between the two reads (5 s of slack) |
| `resident: still serving`                                          | after Word was closed: active, both ports held by `MainPID`                                                                      |
| `resident: the Word of the run closed, other Word processes alive` | the cleanup of every stage                                                                                                       |

| Assumed                                                                                                                     | When it is wrong                                                                                                                            |
| --------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------- |
| nobody else uses the unit during the run (it also serves the user's own Word panes)                                         | `no tool call on other documents` fails and names the line of the other document. The run cannot tell such a line from one it caused itself |
| the unit writes its output to the file of `StandardOutput=append:` (as `scripts/install/unit.mjs` renders it)               | `result.json` `resident.log` shows `{ since }` and the row reads the journal, where the daemon's lines are not: it passes on an empty log   |
| `ss -ltnp` shows the pid of a process of the same user                                                                      | `resident: serving` fails with `47833:[] 47834:[]` while the unit is active                                                                 |
| `systemctl --user show -p MainPID --value` gives the pid of node (`Type=simple`)                                            | `resident: serving` fails with the pids `ss` shows next to another MainPID                                                                  |
| the pane of a Word started with the WebView2 variables of `word-launch.ps1` loads from the unit as from a daemon of the run | `open-review` fails with `pane not reachable on port 9333` or no "Review file:" line within 60 s                                            |

Opening the pane from the Word ribbon instead of by the document is not
automated. `manifests/word.xml` declares no add-in commands (no
`VersionOverrides`), so Word has no Sidewrite button on the ribbon: by hand the
pane is opened through the Office Add-ins dialog (Home or Insert, Add-ins, My
Add-ins, Developer Add-ins), a dialog of web content with labels in the
language of Office and in a window of its own. Driving it with UI
Automation, even scoped to the Word the run started, would rest on localized
names and on the inner structure of that dialog, and none of it could be
tried when the stage was written.

## Pane steps

`pane-driver.cjs <step> <args.json>` runs one step and prints one JSON line.
Stage 2: `open-review`, `split`, `edit-draft`, `ledger-add`, `check-all`,
`preview`, `propose`, `refresh`, `clean-tool`, `delete-all`, `view`.
`split` and `suggest-ledger` wait for the message their own action ends with
(`Split into ...`, `... candidate rule(s) by ...`) or for an error. The line
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

Review tab, the two buttons of `--check-now` and `--figures-button` (stage 3):

| Step               | Args        | Does and returns                                                                                                                                                                                                                                                                                                       |
| ------------------ | ----------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `submission-check` |             | opens "Before submission", presses Check now, waits until the button is enabled again and the message line is neither "Reading the document..." nor "Checking..."; `rows` (`id`, `state`, `label`, `text`, `go`), `asof`, `count`, `message`, `message_error`                                                          |
| `figures-open`     | `settle_ms` | opens More, wraps `openBrowserWindow`, presses "Figures and other files", waits `settle_ms` (default 3000); `href`, `spy` (`supported`, `wrapped`, `installed_on`, `calls`, `errors`, `default_prevented`), `pane_url_before`, `pane_url_after`, `figures_pages_before`, `figures_pages_after` (pages of the WebView2) |

`judge.mjs` holds the expected texts of these steps (`JUDGES`) and the
evidence builders of the stage 2.5 markers (`modelsEvidence`,
`providersEvidence`, `ollamaSplitEvidence`, `proposeAllEvidence`, `focusEvidence`,
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
  without saving and only when its current Application PID matches a positive
  `ExpectPid` and every open document is in the run folder; the
  runner then checks that the started process id is gone and that every Word
  process that was running before is still there.
- When `quit` refuses (a document outside the run folder is open in the Word
  the runner started) or cannot say which documents are open, the runner does
  not stop that Word: it stays up, `result.json` lists it under
  `cleanup.left_running`, and the run exits non-zero. The process is stopped
  by force only when `quit` ran on run-folder documents and the process is
  still there.

## Files

| File                  | Runs on      | Does                                                                                               |
| --------------------- | ------------ | -------------------------------------------------------------------------------------------------- |
| `run-acceptance.mjs`  | WSL          | the whole run                                                                                      |
| `judge.mjs`           | WSL          | COM plan from revision.json, expected pane texts, path rewrite (tests: `tests/e2e-judge.test.mjs`) |
| `pane-driver.cjs`     | Windows node | one pane step per call over CDP                                                                    |
| `word-launch.ps1`     | Windows      | Word with a CDP port on its pane                                                                   |
| `word-com.ps1`        | Windows      | accept/reject, hand edits, comments, a second document, save, quit                                 |
| `browser-windows.ps1` | Windows      | the browser processes before and after the figures button; closes the run's own window only        |
| `embed_taskpane.py`   | WSL          | auto-open task pane parts in a copy                                                                |

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

## Replay of a real revision

`replay.mjs` runs the Review tab on a copy of a real manuscript with its real
review, and `replay-compare.py` puts the result next to what the authors did.
Neither is part of the acceptance run and neither judges anything.

```sh
node scripts/e2e/replay.mjs \
  --manuscript /path/to/before.docx --review /path/to/review.txt \
  --base-dir /mnt/d/path/to/test-root \
  --playwright-core 'D:\path\to\node_modules\playwright-core' \
  [--model opus] [--writing-model opus] [--judgment-model haiku] \
  [--only R1.1,R2.3] [--instructions rules.txt] [--forbid forbid.txt] \
  [--draft-responses] [--turn-minutes 10] [--repair] [--rewrite-stacked] \
  [--follow-through]

python3 scripts/e2e/replay-compare.py \
  --before /path/to/before.docx --after /path/to/submitted.docx \
  --tool /mnt/d/path/to/test-root/replay-<time>/changes.json --out /path/to/out
```

The manuscript and the review are only read. The run folder
`replay-<UTC time>` holds the copy (`manuscript.docx`, saved with its tracked
changes and rationale comments), `manuscript.revision.json`,
`manuscript.decisions.jsonl`, `changes.json` (per comment: the requests with
their state `changed`, `needs_author`, `outside_this_file` or `open`, the
questions for the author, the requests declared outside this file, the terms
a change took out that other paragraphs still hold, the terms the reviewer
named with every paragraph that holds them now and where the turn itself
said the same term is left, the paragraphs written anew with the changes of
each rewrite, what the paragraph check found on its changes, the edits a
guard refused by code, the changes with the paragraph before and after,
rationale, state, the id of their rewrite and the words that come from the
review, how the turn ended, its time and model), `paragraph_check.json` (the
result of "Check revised paragraphs", run once after the last turn),
`responses.json` with `--draft-responses`, `summary.txt`, `result.json` and
`daemon.log`. The screen shows step names and counts, never text of the
manuscript. One review turn per comment, in order; a turn that fails or runs
longer than `--turn-minutes` is recorded and the next comment goes on. Word
and the resident daemon are handled as in the acceptance run. Each Word step
of the run (the open checks, the Word launch, the save of the copy) lists the
WINWORD processes right before and after it and records a new one that is
not the step's own Word in `word.strays` (and `open_checks[].strays`). The final
sweep reports these processes in `cleanup.sweep` without closing them. A PID
delta, start time or blank window title cannot prove ownership: a user may
have opened Word during the step. Replay never stops Word by process ID after
a failed, refused or timed-out COM quit. A remaining owned process is recorded
in `cleanup.left_running` (also `open_checks[].left_running` for an open check)
and makes cleanup fail. `word-own.ps1` refuses unknown or mismatched ownership
both before use and again before quit. Normal cleanup still quits through COM
when the current PID matches and all documents are in the run folder.
The Word scripts find a document only in the Running Object Table
(`word-rot.ps1`): a bind by path would start a Word of its own when the
document is not registered yet.

`node scripts/e2e/term-family-count.mjs <run folder>` counts again, with the
code of this checkout, where the terms the reviewers named stand in the
text the run left (`snapshot.paragraphs` of `manuscript.revision.json`):
per term the places matched as written and those matched only by a word
family, counts only. Run it from two checkouts to compare a change of
`taskpane/shared/named-terms.js` on a real run.

Before the run a Word of its own is asked whether it opens the copy the run
will use (the add-in tagged on) as it is, with the same call as the Word judge
of the marked copies (`OpenReadOnly` in `word-own.ps1`). A file Word wants to repair ends the run with "run again with --repair";
with `--repair` Word opens it with Open and Repair, saves a repaired copy in
the run folder and the run uses that copy (`summary.txt` says so). The given
file is copied once and never opened in Word itself.

`--model`, `--writing-model` and `--judgment-model` are chosen in Setup
(Providers: Edits, Writing, Judgments; Claude). `summary.txt` and each turn
of `changes.json` hold what the pane's own line says the work ran on. The
run fails (FAILED in `summary.txt`, exit code 1) when that line, at the end,
names another model than the one asked for and gives no reason. A document
that stays on another model because the chosen one declines it says so in
the line and is no failure; `summary.txt` then has a line "edits stayed on
... because ... declines this manuscript (...)". Nor are the edits a failure
when the last turn of the run was the one the model was replaced in after a
refusal. Replacements and restorations are counted.

With `--rewrite-stacked`, after the paragraph check every paragraph that holds
three or more changes gets its "Rewrite this paragraph" turn (the buttons in
the order of the check, each once), then the status is refreshed; the turns
are in `changes.json` under `rewrite_turns` and the new changes carry a
`rewrite_id`. Off by default.

With `--follow-through`, after the turns of the comments and before the
paragraph check every comment whose card offers "Check other places" gets
that turn: a second turn of the comment for the places its changes did not
touch (the heading above a changed paragraph, the legends of the figures it
cites, the methods, the rest of its section). The turns are in `changes.json`
under `follow_turns`, and `follow` of a comment holds the places the turn was
given (paragraph indexes), the changes it added and the number of its
questions; the changes themselves stand among the changes of the comment.
`summary.txt` counts them in the two "follow-through" lines. Off by default.

`summary.txt` lists the batches of the request extraction as the daemon
logged them (labelled, unlabelled refs, or FAILED with the error).

The refused edits are counted from the daemon log of each turn:
`NumberNeedsAuthor` (a number changed without the author), `ForbidBlocked`
(a forbid rule; the daemon gives this one no code), `ForbidRuleBroken`,
`TaskPaneOutdated`, `RewriteBlocked`, `RewriteStopped` (a rewrite that
stopped after one or more parts), the code of what stopped a rewrite before
its first part (`ParagraphNotPlain` and others) and `Refused` (a refusal the
replay does not know). One log line is one refused edit or rewrite.
`RewriteNotExpressible` leaves no log line and is not counted. When no
request was extracted in the whole run, `summary.txt` says "not measured" for
the requirements no change names.

`--instructions` and `--forbid` are text files with one rule per line (`#`
starts a note); each line becomes a ledger rule the model checks.

`replay-compare.py` uses the standard library only. It writes
`reference_changes.json` (the paragraphs that differ between the two clean
files: modified, rewritten, deleted, inserted; "rewritten" is a removed and
an added paragraph that share enough words) and `compare.md` (counts, then the
paragraphs changed by both, by the authors only, by the tool only, and the
tool changes it could not place on a paragraph, and per comment what the
tool said instead of changing: its questions for the author and the requests
it declared outside this file, so silence and a question can be told apart;
the terms the reviewer named and where they still stand, the paragraphs
written anew, the changes with words from the review, and the findings of
the paragraph check).
Without `--tool` it writes the authors' side only.
