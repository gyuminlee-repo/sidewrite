# Does Word open this file as it is? Asked before replay.mjs starts its run.
#
# Runs in a Word process of its own (word-own.ps1: started with /n on one
# anchor document of the run folder; a Word that was already open, and its
# documents, are never touched; at the end only the Word of this run is quit).
# -Doc is a copy in the run folder. It is opened read-only, without a dialog
# (DisplayAlerts off) and without a repair. When Word refuses it and
# -RepairTo is given, it is opened with Open and Repair and saved under that
# name (a new file in the run folder); -Doc itself is never written.
#
# Prints one JSON line:
#   { ok, error, pid, preexisting, opens, open_error, repaired, repair_error,
#     paragraphs, quit }
param(
  [Parameter(Mandatory = $true)][string]$Doc,
  [Parameter(Mandatory = $true)][string]$Anchor,
  [Parameter(Mandatory = $true)][string]$AllowDir,
  [Parameter(Mandatory = $true)][string]$PidFile,
  [string]$RepairTo = "",
  [int]$WaitSeconds = 90
)
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

. (Join-Path $PSScriptRoot "word-own.ps1")

# Close without saving (0 = wdDoNotSaveChanges). A failure is returned as text.
function CloseQuiet($doc) {
  try { $doc.Saved = $true; $doc.Close(0) | Out-Null; return $null } catch { return $_.Exception.Message }
}

$w = StartOwnWord $Anchor $PidFile $WaitSeconds
$before = $w.before
$p = $w.process
$d = $w.document
$app = $w.app

$res = [ordered]@{ opens = $false; open_error = $null; repaired = $false; repair_error = $null; paragraphs = $null; close_note = $null }
$failure = $null
$quit = $null
$ours = $false
try {
  AssertOwnWord $app $p
  $ours = $true
  $app.DisplayAlerts = 0
  if (-not (Under $Doc)) { throw "the file is outside the run folder" }
  if ($RepairTo -and -not (Under $RepairTo)) { throw "the repaired copy would be outside the run folder" }
  $doc = $null
  try {
    # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
    $doc = $app.Documents.Open($Doc, $false, $true, $false)
    $res.opens = $true
    $res.paragraphs = [int]$doc.Paragraphs.Count
  } catch {
    $res.open_error = $_.Exception.Message
  }
  if ($doc) { $res.close_note = CloseQuiet $doc; $doc = $null }
  if (-not $res.opens -and $RepairTo) {
    try {
      $m = [Type]::Missing
      # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument,
      #      PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format,
      #      Encoding, Visible, OpenAndRepair)
      $doc = $app.Documents.Open($Doc, $false, $true, $false, $m, $m, $m, $m, $m, $m, $m, $m, $true)
      # 12 = wdFormatXMLDocument. Some Word versions want the arguments by reference.
      try { $doc.SaveAs2($RepairTo, 12) } catch { $doc.SaveAs2([ref]$RepairTo, [ref]12) }
      $res.repaired = $true
      $res.paragraphs = [int]$doc.Paragraphs.Count
    } catch {
      $res.repair_error = $_.Exception.Message
    }
    if ($doc) { $res.close_note = CloseQuiet $doc; $doc = $null }
  }
} catch {
  $failure = $_.Exception.Message
} finally {
  $quit = QuitOwnWord $app $d $p $ours
}
$res.ok = ($null -eq $failure)
$res.error = $failure
$res.pid = $p.Id
$res.preexisting = $before
$res.quit = $quit
Line $res
