# Word acceptance run without hands

`run-acceptance.mjs` runs the stage-2 Review tab acceptance procedure end to
end: daemon, a separate Word process, the task pane clicks, the author's Word
steps, and the feature_list verify commands.

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
daemon log goes to `--log` (default `/tmp/daemon-s2-e2e.log`; a second
daemon started for the forbid-list knob writes `-knob.log`).

Verify commands are rewritten before they run: `--plan-dir` becomes the run
folder, `cd <--plan-repo> ` becomes this checkout, `--plan-log` (default
`/tmp/daemon-s2.log`) and its `-knob.log` become this run's logs, `/tmp/s2-`
markers become `ok/s2-`. Files of earlier runs are never read.

Other options: `--cdp-port`, `--win-node`, `--fixture`, `--review`,
`--ledger`, `--features S2-18..S2-34`, `--skip-knob`, `--skip-mode14`,
`--keep-on-fail` (leave Word and the daemon up after a failed step).

## Resolved rationale comments

- After the author's Word steps and the Refresh, the runner reads
  `Comment.Done` of every comment over COM and requires that the rationale
  comment of each accepted or rejected change is resolved (row
  `word: rationale comments resolved`). The test manuscript is built in
  compatibility mode 15 for this; in mode 14 Word refuses to resolve.
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
