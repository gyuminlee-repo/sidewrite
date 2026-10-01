# Does Word open this file as it is? Asked by replay.mjs about the very file
# its run is going to open, before the run starts.
#
# Runs in a Word process of its own (word-own.ps1: started with /n on one
# anchor document of the run folder; a Word that was already open, and its
# documents, are never touched; at the end only the Word of this run is quit).
#
# Without -RepairTo: -File is opened with OpenReadOnly of word-own.ps1, the
# call marked-diff-word.ps1 uses (read-only, no repair, no dialog), and closed
# without saving. It counts as opened only when Word gives a document with
# that path and at least one paragraph.
# With -RepairTo: -File is opened with Open and Repair and saved under that
# name (a new file in the run folder). -File itself is never written.
#
# PowerShell variable names ignore case (a local written like a parameter in
# small letters IS the parameter): every variable here has one spelling, and
# none is named like a parameter.
#
# Prints one JSON line:
#   { ok, error, pid, preexisting, opens, open_error, repaired, repair_error,
#     paragraphs, quit }
param(
  [Parameter(Mandatory = $true)][string]$File,
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
function CloseQuiet($x) {
  try { $x.Saved = $true; $x.Close(0) | Out-Null; return $null } catch { return $_.Exception.Message }
}

$w = StartOwnWord $Anchor $PidFile $WaitSeconds
$before = $w.before
$p = $w.process
$anchorDoc = $w.document
$app = $w.app

$res = [ordered]@{ opens = $false; open_error = $null; repaired = $false; repair_error = $null; paragraphs = $null; close_note = $null }
$failure = $null
$quit = $null
$ours = $false
try {
  AssertOwnWord $app $p
  $ours = $true
  $app.DisplayAlerts = 0
  if (-not (Under $File)) { throw "the file is outside the run folder" }
  $opened = $null
  if (-not $RepairTo) {
    try {
      $opened = OpenReadOnly $app $File
      $count = [int]$opened.Paragraphs.Count
      $res.paragraphs = $count
      if ([string]$opened.FullName -ne $File) { $res.open_error = "Word gave another document than the one asked for" }
      elseif ($count -lt 1) { $res.open_error = "Word gave a document without a paragraph" }
      else { $res.opens = $true }
    } catch {
      $res.open_error = $_.Exception.Message
    }
  } else {
    if (-not (Under $RepairTo)) { throw "the repaired copy would be outside the run folder" }
    try {
      $m = [Type]::Missing
      # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument,
      #      PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format,
      #      Encoding, Visible, OpenAndRepair)
      $opened = $app.Documents.Open($File, $false, $true, $false, $m, $m, $m, $m, $m, $m, $m, $m, $true)
      # 12 = wdFormatXMLDocument. Some Word versions want the arguments by reference.
      try { $opened.SaveAs2($RepairTo, 12) } catch { $opened.SaveAs2([ref]$RepairTo, [ref]12) }
      $res.repaired = $true
      $res.paragraphs = [int]$opened.Paragraphs.Count
    } catch {
      $res.repair_error = $_.Exception.Message
    }
  }
  if ($opened) { $res.close_note = CloseQuiet $opened }
} catch {
  $failure = $_.Exception.Message
} finally {
  $quit = QuitOwnWord $app $anchorDoc $p $ours
}
$res.ok = ($null -eq $failure)
$res.error = $failure
$res.pid = $p.Id
$res.preexisting = $before
$res.quit = $quit
Line $res
