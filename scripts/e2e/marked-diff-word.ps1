# Word as the judge of marked-diff output (marked-diff-word-check.mjs).
#
# Starts a Word process of its own (/n: a new instance even when Word is
# already open) on one anchor document of the run folder, finds that document
# through the Running Object Table by its full path (as word-com.ps1 does),
# and uses its Application for every job. A Word that was already open, and
# its documents, are never touched: the Application is used only when it is
# the process this script started and every document open in it is under
# -AllowDir, every document this script closes is one it opened itself, and
# at the end Word is quit only under the same two conditions.
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

function Line($o) { $o | ConvertTo-Json -Depth 10 -Compress }
function Under([string]$path) {
  $allow = $AllowDir.TrimEnd("\") + "\"
  return $path.StartsWith($allow, [System.StringComparison]::OrdinalIgnoreCase)
}
function OpenNames($app) {
  $names = @()
  foreach ($x in $app.Documents) { $names += [string]$x.FullName }
  return , $names
}

# The process that owns a Word Application, from its window. 0 when it cannot be read.
Add-Type -Namespace MdwNative -Name User32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);'
function AppProcessId($app) {
  try {
    [uint32]$owner = 0
    $hwnd = [System.IntPtr][int]$app.ActiveWindow.Hwnd
    [MdwNative.User32]::GetWindowThreadProcessId($hwnd, [ref]$owner) | Out-Null
    return [int]$owner
  } catch {
    return 0
  }
}

# What the comparison needs of an open document, in Word's own terms.
function Shape($doc) {
  $paras = @()
  $inline = @()
  $fields = @()
  foreach ($x in $doc.Paragraphs) {
    $paras += [string]$x.Range.Text
    $inline += [int]$x.Range.InlineShapes.Count
    $fields += [int]$x.Range.Fields.Count
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
  return [ordered]@{ paragraphs = $paras; inline = $inline; fields = $fields; tables = $tables; bookmarks = $marks }
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

$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
if (-not (Under $Anchor)) { throw "the anchor document is outside the run folder" }
$exe = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe")."(default)"
if (-not $exe -or -not (Test-Path $exe)) { throw "WINWORD.EXE not found in App Paths" }
$p = Start-Process -FilePath $exe -ArgumentList @("/n", "/q", ('"' + $Anchor + '"')) -PassThru
[System.IO.File]::WriteAllText($PidFile, (Line ([ordered]@{ pid = $p.Id; preexisting = $before })), $utf8)

$deadline = (Get-Date).AddSeconds($WaitSeconds)
while ((Get-Date) -lt $deadline -and -not $p.HasExited) {
  if ($p.MainWindowHandle -ne 0) { break }
  Start-Sleep -Milliseconds 500
  $p.Refresh()
}
if ($p.HasExited -or $before -contains $p.Id) { throw "the Word process for this run did not start" }
# The window is there before the document is registered: give it a moment.
Start-Sleep -Seconds 3

$d = $null
while (-not $d -and (Get-Date) -lt $deadline) {
  try {
    $d = [System.Runtime.InteropServices.Marshal]::BindToMoniker($Anchor)
    if ([string]$d.FullName -ne $Anchor) { $d = $null }
  } catch { $d = $null }
  if (-not $d) { Start-Sleep -Milliseconds 500 }
}
if (-not $d) { throw "the anchor document was not found in the new Word" }
$app = $d.Application

$results = @()
$failure = $null
$quit = $null
$ours = $false
try {
  $owner = AppProcessId $app
  if ($owner -ne 0 -and $owner -ne $p.Id) { throw "the anchor document is open in another Word (process $owner), not in the one this run started ($($p.Id))" }
  $outside = @((OpenNames $app) | Where-Object { -not (Under $_) })
  if ($outside.Count -gt 0) { throw "the Word this run reached has other documents open" }
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
  # Quit, as word-com.ps1 does: only the Word of this run, and only when every open document is under the run folder.
  try {
    $names = OpenNames $app
    $outside = @($names | Where-Object { -not (Under $_) })
    if (-not $ours -or $outside.Count -gt 0) {
      $why = if ($ours) { "foreign_documents" } else { "not_the_word_of_this_run" }
      $quit = [ordered]@{ ok = $false; refused = $why; documents = $names; outside = $outside; pid = $p.Id }
    } else {
      foreach ($x in $app.Documents) { $x.Saved = $true }
      $app.Quit() | Out-Null
      [System.Runtime.InteropServices.Marshal]::ReleaseComObject($d) | Out-Null
      $gone = $false
      for ($i = 0; $i -lt 60; $i++) {
        if (-not (Get-Process -Id $p.Id -ErrorAction SilentlyContinue)) { $gone = $true; break }
        Start-Sleep -Milliseconds 500
      }
      $quit = [ordered]@{ ok = $true; documents = $names; pid = $p.Id; exited = $gone }
    }
  } catch {
    $quit = [ordered]@{ ok = $false; error = $_.Exception.Message; pid = $p.Id }
  }
}
[System.IO.File]::WriteAllText($Out, (Line ([ordered]@{ jobs = $results })), $utf8)
Line ([ordered]@{ ok = ($null -eq $failure); error = $failure; pid = $p.Id; preexisting = $before; jobs = $results.Count; quit = $quit })
