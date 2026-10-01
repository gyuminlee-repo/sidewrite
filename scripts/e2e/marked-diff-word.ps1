# Word as the judge of marked-diff output (marked-diff-word-check.mjs).
#
# Runs in a Word process of its own (word-own.ps1: started with /n on one
# anchor document of the run folder; a Word that was already open, and its
# documents, are never touched; every document this script closes is one it
# opened itself; at the end only the Word of this run is quit).
#
# A job is { id, path, action }: the file is opened read-only, without a
# repair (OpenAndRepair is not used) and without a dialog (DisplayAlerts off),
# then for action "accept" Revisions.AcceptAll() and for "reject"
# Revisions.RejectAll() run, the structure is read, and the document is
# closed without saving. A file Word refuses to open is a job with ok = false
# and the error text.
#
# Writes -Out (JSON: { jobs: [...] }) after every job and prints one JSON
# line at the end: { ok, error, pid, preexisting, jobs, quit }.
param(
  [Parameter(Mandatory = $true)][string]$Jobs,
  [Parameter(Mandatory = $true)][string]$Out,
  [Parameter(Mandatory = $true)][string]$Anchor,
  [Parameter(Mandatory = $true)][string]$AllowDir,
  [Parameter(Mandatory = $true)][string]$PidFile,
  [int]$WaitSeconds = 90
)
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$utf8 = New-Object System.Text.UTF8Encoding($false)

. (Join-Path $PSScriptRoot "word-own.ps1")

# What the comparison needs of an open document, in Word's own terms.
function Shape($doc) {
  $paras = @()
  $inline = @()
  $fields = @()
  $codes = @()
  foreach ($x in $doc.Paragraphs) {
    $paras += [string]$x.Range.Text
    $inline += [int]$x.Range.InlineShapes.Count
    $fields += [int]$x.Range.Fields.Count
    # code, U+001F, result, U+001E for every field of the paragraph
    $code = ""
    foreach ($f in $x.Range.Fields) { $code += [string]$f.Code.Text + [char]31 + [string]$f.Result.Text + [char]30 }
    $codes += $code
  }
  $tables = @()
  foreach ($t in $doc.Tables) {
    # Rows and columns are only counted, cells are read through the range of the table:
    # a single row or column of a table with merged cells cannot be reached.
    $cells = @()
    $at = @()
    foreach ($c in $t.Range.Cells) {
      $cells += [string]$c.Range.Text
      $at += ([string]$c.RowIndex + "," + [string]$c.ColumnIndex)
    }
    $tables += [ordered]@{ rows = [int]$t.Rows.Count; columns = [int]$t.Columns.Count; uniform = [bool]$t.Uniform; cells = $cells; at = $at }
  }
  $marks = @()
  foreach ($b in $doc.Bookmarks) { $marks += [ordered]@{ name = [string]$b.Name; text = [string]$b.Range.Text } }
  return [ordered]@{
    paragraphs = $paras; inline = $inline; fields = $fields; codes = $codes; tables = $tables; bookmarks = $marks
    comments = [int]$doc.Comments.Count; footnotes = [int]$doc.Footnotes.Count; endnotes = [int]$doc.Endnotes.Count
  }
}

function RunJob($app, $job) {
  $res = [ordered]@{ id = $job.id; action = $job.action; ok = $false; error = $null }
  if (-not (Under $job.path)) { $res.error = "the file is outside the run folder"; return $res }
  if ((OpenNames $app) -contains $job.path) { $res.error = "the file is still open from an earlier job"; return $res }
  $doc = $null
  try {
    # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
    $doc = $app.Documents.Open($job.path, $false, $true, $false)
  } catch {
    $res.error = "open: " + $_.Exception.Message
    return $res
  }
  try {
    $res.opened = [string]$doc.FullName
    # A document that is not saved right after it was opened was changed by Word on the way in.
    $res.dirty_on_open = -not [bool]$doc.Saved
    $res.revisions = [int]$doc.Revisions.Count
    $res.raw = [ordered]@{ paragraphs = [int]$doc.Paragraphs.Count; tables = [int]$doc.Tables.Count }
    if ($job.action -eq "accept") { $doc.Revisions.AcceptAll() | Out-Null }
    elseif ($job.action -eq "reject") { $doc.Revisions.RejectAll() | Out-Null }
    $res.revisions_left = [int]$doc.Revisions.Count
    $res.shape = Shape $doc
    $res.ok = $true
  } catch {
    $res.error = "read: " + $_.Exception.Message
  } finally {
    # Nothing is saved: the document is marked as saved, then closed without saving (0 = wdDoNotSaveChanges).
    try { $doc.Saved = $true } catch { $res.close_note = "Saved could not be set: " + $_.Exception.Message }
    try { $doc.Close(0) | Out-Null } catch { $res.close_note = "Close(0): " + $_.Exception.Message }
  }
  if ((OpenNames $app) -contains $job.path) {
    $res.ok = $false
    $res.error = "the document could not be closed: " + $res.close_note
  }
  return $res
}

$w = StartOwnWord $Anchor $PidFile $WaitSeconds
$before = $w.before
$p = $w.process
$d = $w.document
$app = $w.app

$results = @()
$failure = $null
$quit = $null
$ours = $false
try {
  AssertOwnWord $app $p
  $ours = $true
  $app.DisplayAlerts = 0
  $list = Get-Content -Raw -Encoding UTF8 $Jobs | ConvertFrom-Json
  foreach ($job in $list) {
    $results += (RunJob $app $job)
    [System.IO.File]::WriteAllText($Out, (Line ([ordered]@{ jobs = $results })), $utf8)
  }
} catch {
  $failure = $_.Exception.Message
} finally {
  $quit = QuitOwnWord $app $d $p $ours
}
[System.IO.File]::WriteAllText($Out, (Line ([ordered]@{ jobs = $results })), $utf8)
Line ([ordered]@{ ok = ($null -eq $failure); error = $failure; pid = $p.Id; preexisting = $before; jobs = $results.Count; quit = $quit })
