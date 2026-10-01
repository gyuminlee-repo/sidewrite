# How Word writes tracked changes of equations (word-math-probe.mjs).
#
# Runs in a Word process of its own (word-own.ps1: started with /n on one
# anchor document of the run folder; a Word that was already open, and its
# documents, are never touched; at the end only the Word of this run is quit).
#
# 1. Makes -Base: five paragraphs, four of them with an equation (typed as
#    linear text, turned into an equation with OMaths.Add and built up).
#    Saved and closed.
# 2. Copies it to -Revised, opens the copy, turns TrackRevisions on and makes
#    five changes, each a step of its own (a step that fails is reported and
#    the next one still runs):
#      1  paragraph 1: the equation is deleted, another one is inserted after it
#      2  paragraph 2: one word is replaced, the equation stays
#      3  paragraph 3: the whole paragraph, with its equation, is deleted
#      4  a new paragraph with an equation is added at the end
#      5  paragraph 5: one character inside the equation is replaced
#    Saved and closed.
#
# Writes -Out and prints the same JSON as one line at the end:
# { ok, error, pid, preexisting, steps: [{ name, ok, error }], info, quit }.
param(
  [Parameter(Mandatory = $true)][string]$Base,
  [Parameter(Mandatory = $true)][string]$Revised,
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

$Paras = @(
  "Text before x^2 text after",
  "Second paragraph holds y+1 and some words here",
  "Third paragraph holds z-1 and goes away",
  "Fourth paragraph is plain text",
  "Fifth paragraph holds a+b inside"
)
$Maths = @("x^2", "y+1", "z-1", "", "a+b")
$Added = "New paragraph holds c+d here"

# Turns $length characters of a paragraph, from $offset on, into an equation and builds it up.
function MakeMath($doc, [int]$para, [int]$offset, [int]$length) {
  $start = $doc.Paragraphs.Item($para).Range.Start + $offset
  $rng = $doc.Range($start, $start + $length)
  $rng = $doc.OMaths.Add($rng)
  $rng.OMaths.Item(1).BuildUp()
}

# One change: its name, whether it ran, and the error text when it did not.
function Step([string]$name, [scriptblock]$body) {
  $res = [ordered]@{ name = $name; ok = $false; error = $null }
  try {
    & $body | Out-Null
    $res.ok = $true
  } catch {
    $res.error = $_.Exception.Message
  }
  return $res
}

$w = StartOwnWord $Anchor $PidFile $WaitSeconds
$before = $w.before
$p = $w.process
$d = $w.document
$app = $w.app

$steps = @()
$info = $null
$failure = $null
$quit = $null
$ours = $false
$doc = $null
$closeNote = $null
try {
  AssertOwnWord $app $p
  $ours = $true
  $app.DisplayAlerts = 0
  if (-not (Under $Base) -or -not (Under $Revised)) { throw "the documents must be written under the run folder" }

  # 1. The baseline. 12 = wdFormatXMLDocument.
  $doc = $app.Documents.Add()
  $doc.TrackRevisions = $false
  $doc.Content.Text = ($Paras -join "`r")
  for ($i = 0; $i -lt $Paras.Count; $i++) {
    $linear = $Maths[$i]
    if ($linear -ne "") {
      $at = $Paras[$i].IndexOf($linear)
      $len = $linear.Length
      MakeMath $doc ($i + 1) $at $len
    }
  }
  $doc.SaveAs2($Base, 12)
  $doc.Close(0) | Out-Null
  $doc = $null

  # 2. The revised copy. Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
  Copy-Item -LiteralPath $Base -Destination $Revised
  $doc = $app.Documents.Open($Revised, $false, $false, $false)
  $doc.TrackRevisions = $true

  $steps += (Step "1 delete the equation of paragraph 1 and insert another after it" {
    $mr = $doc.Paragraphs.Item(1).Range.OMaths.Item(1).Range
    $end = $mr.End
    $mr.Delete() | Out-Null
    # one character further: behind the space that follows the deleted equation
    $ins = $doc.Range($end + 1, $end + 1)
    $ins.InsertAfter("y^2")
    $ins = $doc.OMaths.Add($ins)
    $ins.OMaths.Item(1).BuildUp()
  })
  $steps += (Step "2 replace a word of paragraph 2, the equation stays" {
    $r = $doc.Paragraphs.Item(2).Range
    if (-not $r.Find.Execute("words")) { throw "the word was not found in paragraph 2" }
    $r.Text = "terms"
  })
  $steps += (Step "3 delete paragraph 3 with its equation" {
    $doc.Paragraphs.Item(3).Range.Delete() | Out-Null
  })
  $steps += (Step "5 replace one character inside the equation of paragraph 5" {
    $mr = $doc.Paragraphs.Item(5).Range.OMaths.Item(1).Range
    $ch = $mr.Characters.Item($mr.Characters.Count)
    $ch.Text = "c"
  })
  $steps += (Step "4 add a paragraph with an equation at the end" {
    $doc.Content.InsertParagraphAfter()
    $doc.Content.InsertAfter($Added)
    $last = $doc.Paragraphs.Item($doc.Paragraphs.Count).Range
    $at = $last.Start + $Added.IndexOf("c+d")
    $rng = $doc.Range($at, $at + 3)
    $rng = $doc.OMaths.Add($rng)
    $rng.OMaths.Item(1).BuildUp()
  })

  $info = [ordered]@{
    paragraphs = [int]$doc.Paragraphs.Count
    equations = [int]$doc.OMaths.Count
    revisions = [int]$doc.Revisions.Count
  }
  $doc.Save()
  $doc.Close(0) | Out-Null
  $doc = $null
} catch {
  $failure = $_.Exception.Message
} finally {
  # A document of this script that is still open is closed without saving, so that Word can be quit.
  if ($null -ne $doc) {
    try {
      $doc.Saved = $true
      $doc.Close(0) | Out-Null
    } catch {
      $closeNote = $_.Exception.Message
    }
  }
  $quit = QuitOwnWord $app $d $p $ours
}
$result = [ordered]@{
  ok = ($null -eq $failure); error = $failure; pid = $p.Id; preexisting = $before
  steps = $steps; info = $info; close_note = $closeNote; quit = $quit
}
[System.IO.File]::WriteAllText($Out, (Line $result), $utf8)
Line $result
