# Word acceptance run without hands

`run-acceptance.mjs` runs a Review tab acceptance procedure end to end:
daemon, a separate Word process, the task pane clicks, the author's Word
steps, and the feature_list verify commands. It has two flows, stage 2
(`S2-18..S2-34`, the default) and stage 2.5 (`--stage s25`, `S25-30..S25-45`).

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
chooses the flow: `s2` or `s25`. `--stage s25` sets it (the features then
default to `S25-30..S25-45`); without it the stage is read from `--features`
(default `S2-18..S2-34`). A range stays in one stage.

## Stage 2.5

```sh
node scripts/e2e/run-acceptance.mjs --stage s25 \
  --base-dir <folder on a Windows drive, as a WSL path> \
  --playwright-core 'D:\path\to\node_modules\playwright-core' \
  --feature-list <stage 2.5 feature_list.json> \
  --plan-repo <repo path the feature_list commands cd into> \
  --conditional S25-42,S25-43,S25-44,S25-45
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
| S25-42          | `select-provider turn=codex` (agrees to the first-use notice), `propose R1.1`           | `s25-proposal-turn`: the turn ended `outcome=ok` with a change applied, each change has its comment, Word holds them       |
| S25-45          | the same turn                                                                           | `s25-proposal-receipt`: one `proposal_batch` line for R1.1, model `codex:...`, applied ids on file, no text of a change    |
| S25-44          | `propose-stop R1.4` (Stop after 6 s)                                                    | `s25-proposal-stop`: Stop pressed, no change and no attempt of R1.4 on file, no `codex exec --json` process left           |
| S25-43          | `propose R1.3` (asks for a word a ledger rule forbids), again with the forbid list off  | `s25-proposal-guard`: `[review] blocked R1.3`, the turn stopped there, the word in no change and no tracked change of Word |
| S25-37          | `propose-all` (with Claude)                                                             | `s25-propose-all`: `overlaps`, `proposed`, `open_before`                                                                   |
| S25-38          | Word COM deletes a sentence, `focus-refresh`                                            | `s25-focus-refresh`: the auto-refresh line changed and an outside change is listed                                         |
| S25-41          | cleanup (Word quit, daemon stopped), then the verify commands                           | `s25-cleanup`: own Word gone, earlier Word processes alive, ports closed                                                   |

Codex is given no tools in a review turn: it returns a list of proposed
changes and the daemon applies each one as a tracked change
(`daemon/proposal-turns.mjs`). The S25-43 turn is run a second time with
`CC_OFFICE_REVIEW_PROMPT_FORBID=0` (its log is `daemon-knob.log` next to the
daemon log) when the first try proposed nothing a rule forbids; `--skip-knob`
leaves that out. Antigravity is offered for nothing and has no step.

`--proposal-fixture <file.json>` (stage 2.5) runs the stretch from the
proposal runner to Word without a model: the daemon is started with
`CC_OFFICE_E2E_PROPOSAL_FIXTURE` and takes the reply of every proposal turn
from that file (`daemon/proposal-fixture.mjs`; the daemon log says so in its
first lines, and receipts carry the model `fixture:<file>`). The run selects
Codex for review turns as usual (the pane must offer it; nothing is sent to
it), proposes each comment of the file in its order, writes
`ok/s25-fixture-<ref>.evidence.json` (the done line, proposals skipped,
rejected or stopped with their codes, changes on file, comments and tracked
changes in Word) and ends there. The verify commands of the other features
do not pass in such a run. `tests/fixtures/proposal_fixture_deleted_text.json`
measures whether Word's search counts text that a tracked change deleted:
R1.2 replaces "cell-free extract", R1.5 (after the status refresh of its own
turn) looks for "cell-free" in that paragraph (`stopped` with `Ambiguous`
when Word counts the deleted text, applied when not), and R1.4 does both in
one turn, where the daemon passes the second over (`FoundInDeletedText`).

When the pane does not offer Codex for review turns (not installed, not
signed in), the `select-provider turn=codex` row is `SKIP`, no proposal turn
is run, and the four evidence files carry `pass: false` with the pane's
reason. With `--conditional` those features are then
`SKIPPED-CONDITION-NOT-MET` in the feature table (the log shows finished
turns and no `[proposal] prompt ... provider=codex:` line), which is not
`PASS`; without it they fail the run. The pane outputs these judgments were
unit-tested with are written by hand in the form the driver returns: none
was recorded from Word before the first run with proposal turns.

`--provider <id>` (claude, codex, antigravity; stage 2 flow) chooses the
provider of the review turns in the pane's Providers section right after the
Review tab opened (step `select-provider turn=<id>`). A provider the pane
does not offer for review turns is not chosen: the step fails with the text
of its menu row, and the turns run with Claude. Known conditions for
`--conditional` in stage 2.5: `S25-42` to `S25-45` (a proposal turn of Codex
started), read from `[proposal] prompt ref=... provider=codex:...`.

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

`judge.mjs` holds the expected texts of these steps (`JUDGES`) and the
evidence builders of the stage 2.5 markers (`modelsEvidence`,
`providersEvidence`, `ollamaSplitEvidence`, `proposalTurnEvidence`,
`proposalReceiptEvidence`, `proposalStopEvidence`, `proposalGuardEvidence`,
`proposeAllEvidence`, `focusEvidence`,
`suggestEvidence`, `cleanupEvidence`).
No step hands a CLI the office tools: a provider other than Claude never
calls them.

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
