# Word COM steps of the acceptance run, on one document opened in a Word
# process that the runner started (word-launch.ps1). The document is found
# through the Running Object Table by its full path (BindToMoniker), so a
# Word that was already open with other documents is never touched.
# Every action prints one JSON object.
#
# Actions
#   state           comments (with done = resolved) and revisions of the document
#   apply           -PlanJson <file>: accept or reject the tracked changes of
#                   each rationale comment, optionally edit one inserted word
#   hand-insert     -PlanJson <file> -ChangeId <id> -Text <s>: a tracked
#                   insertion by the author at the start of the paragraph
#                   that holds this change, left pending
#   delete-text     -Text <s>: delete its first occurrence, track changes off
#   add-comment     -Text <s> [-Paragraph <n>]: a comment by the Word user
#   quit            -AllowDir <dir> -ExpectPid <n>: close without saving when
#                   every open document is under AllowDir; else refused =
#                   "foreign_documents" and nothing is closed
param(
  [Parameter(Mandatory = $true)][string]$Doc,
  [Parameter(Mandatory = $true)][string]$Action,
  [string]$PlanJson,
  [string]$Text,
  [string]$ChangeId,
  [int]$Paragraph = 2,
  [string]$AllowDir,
  [int]$ExpectPid = 0
)
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Norm([string]$s) {
  if ($null -eq $s) { return "" }
  return (($s -replace "[\r\n\u000b\u0007]+", " ") -replace "\s+", " ").Trim()
}
function Out($o) { $o | ConvertTo-Json -Depth 8 -Compress }

$d = [System.Runtime.InteropServices.Marshal]::BindToMoniker($Doc)
if ((Norm $d.FullName) -ne (Norm $Doc)) { throw "bound to $($d.FullName), expected $Doc" }

function RevList() {
  $list = @()
  foreach ($r in $d.Revisions) {
    $list += [pscustomobject]@{ rev = $r; type = [int]$r.Type; start = $r.Range.Start; end = $r.Range.End; text = $r.Range.Text }
  }
  return $list
}
function CommentList() {
  $list = @()
  foreach ($c in $d.Comments) {
    # Done: the comment is resolved. Unreadable in some documents: null then.
    $done = $null
    try { $done = [bool]$c.Done } catch { $done = $null }
    $list += [pscustomobject]@{ com = $c; text = (Norm $c.Range.Text); start = $c.Scope.Start; end = $c.Scope.End; author = $c.Author; done = $done }
  }
  return $list
}

# Revisions of one change, found from its rationale comment: an insertion is
# the inserted range itself; a replace or a deletion anchors the comment on
# the deleted text, and a replace adds the insertion that touches it.
function RevisionsFor($c, [string]$kind) {
  $hits = @()
  foreach ($r in (RevList)) {
    $overlap = ($r.start -lt $c.end) -and ($r.end -gt $c.start)
    $touch = ($r.start -eq $c.end) -or ($r.end -eq $c.start)
    if ($kind -eq "insert" -and $r.type -eq 1 -and $overlap) { $hits += $r }
    elseif ($kind -eq "delete" -and $r.type -eq 2 -and $overlap) { $hits += $r }
    elseif ($kind -eq "replace" -and (($r.type -eq 2 -and $overlap) -or ($r.type -eq 1 -and ($overlap -or $touch)))) { $hits += $r }
  }
  return $hits
}

function WithTrackingOff([scriptblock]$body) {
  $was = $d.TrackRevisions
  $d.TrackRevisions = $false
  try { & $body } finally { $d.TrackRevisions = $was }
}

switch ($Action) {
  "state" {
    Out ([ordered]@{
      ok = $true
      track_revisions = $d.TrackRevisions
      comments = @(CommentList | ForEach-Object { [ordered]@{ text = $_.text; author = $_.author; start = $_.start; end = $_.end; done = $_.done } })
      revisions = @(RevList | ForEach-Object { [ordered]@{ type = $_.type; text = $_.text; start = $_.start; end = $_.end } })
    })
  }
  "apply" {
    $plan = Get-Content -Raw -Encoding UTF8 $PlanJson | ConvertFrom-Json
    $results = @()
    foreach ($op in $plan.ops) {
      $want = Norm $op.comment_text
      $c = @(CommentList | Where-Object { $_.text -eq $want }) | Select-Object -First 1
      $res = [ordered]@{ change_id = $op.change_id; op = $op.op; comment_found = [bool]$c; revisions = @(); edited = $null; error = $null }
      if (-not $c) { $res.error = "rationale comment not found"; $results += $res; continue }
      $revs = @(RevisionsFor $c $op.kind)
      $res.revisions = @($revs | ForEach-Object { [ordered]@{ type = $_.type; text = $_.text } })
      if ($revs.Count -eq 0) { $res.error = "no revision at the comment"; $results += $res; continue }
      $ins = @($revs | Where-Object { $_.type -eq 1 }) | Select-Object -First 1
      $insStart = if ($ins) { $ins.start } else { -1 }
      $insLen = if ($ins) { $ins.text.Length } else { 0 }
      foreach ($r in $revs) { if ($op.op -eq "reject") { $r.rev.Reject() } else { $r.rev.Accept() } }
      if ($op.op -eq "accept_edit") {
        if (-not $ins) { $res.error = "no inserted text to edit" }
        else {
          # The insertion keeps its start when the deletion before it goes.
          $del = @($revs | Where-Object { $_.type -eq 2 -and $_.start -lt $insStart })
          $shift = 0; foreach ($x in $del) { $shift += $x.text.Length }
          $span = $d.Range($insStart - $shift, $insStart - $shift + $insLen)
          # Word may merge neighbouring insertions into one revision: keep
          # to this change's own text when it is inside.
          $own = if ($op.text) { $span.Text.IndexOf([string]$op.text) } else { -1 }
          if ($own -ge 0) { $span = $d.Range($span.Start + $own, $span.Start + $own + ([string]$op.text).Length) }
          # A word the plan names first, else the first word of 4+ letters.
          $m = $null
          if ($op.replacements) {
            foreach ($k in $op.replacements.PSObject.Properties.Name) {
              $hit = [regex]::Match($span.Text, "\b" + [regex]::Escape($k) + "\b")
              if ($hit.Success) { $m = $hit; break }
            }
          }
          if (-not $m) { $m = [regex]::Match($span.Text, "[A-Za-z]{4,}") }
          if (-not $m.Success) { $res.error = "no word of 4+ letters in the inserted text" }
          else {
            $from = $m.Value
            $to = if ($op.replacements -and $op.replacements.$from) { $op.replacements.$from } else { $from + "ed" }
            WithTrackingOff {
              $w = $d.Range($span.Start + $m.Index, $span.Start + $m.Index + $from.Length)
              $w.Text = $to
            }
            $res.edited = [ordered]@{ from = $from; to = $to; inserted_text = $span.Text }
          }
        }
      }
      $results += $res
    }
    Out ([ordered]@{ ok = (@($results | Where-Object { $_.error }).Count -eq 0); results = $results })
  }
  "hand-insert" {
    $plan = Get-Content -Raw -Encoding UTF8 $PlanJson | ConvertFrom-Json
    $op = @($plan.ops | Where-Object { $_.change_id -eq $ChangeId }) | Select-Object -First 1
    if (-not $op) { Out ([ordered]@{ ok = $false; error = "no plan op $ChangeId" }); break }
    $want = Norm $op.comment_text
    $c = @(CommentList | Where-Object { $_.text -eq $want }) | Select-Object -First 1
    if (-not $c) { Out ([ordered]@{ ok = $false; error = "rationale comment not found" }); break }
    $para = $c.com.Scope.Paragraphs.Item(1).Range
    $was = $d.TrackRevisions
    $d.TrackRevisions = $true
    try { $d.Range($para.Start, $para.Start).InsertBefore($Text) } finally { $d.TrackRevisions = $was }
    $inPara = @(RevList | Where-Object { $_.start -ge $para.Start -and $_.end -le $c.com.Scope.Paragraphs.Item(1).Range.End })
    Out ([ordered]@{ ok = $true; change_id = $ChangeId; inserted = $Text; revisions_in_paragraph = $inPara.Count })
  }
  "delete-text" {
    $r = $d.Content
    $f = $r.Find
    $f.ClearFormatting()
    $f.MatchCase = $true
    $f.Forward = $true
    $f.Wrap = 0
    if (-not $f.Execute($Text)) { Out ([ordered]@{ ok = $false; error = "text not found" }); break }
    if ($r.Start -gt 0 -and $d.Range($r.Start - 1, $r.Start).Text -eq " ") { $r.SetRange($r.Start - 1, $r.End) }
    $removed = $r.Text
    WithTrackingOff { $r.Delete() | Out-Null }
    Out ([ordered]@{ ok = $true; removed = $removed })
  }
  "add-comment" {
    $p = $d.Paragraphs.Item($Paragraph).Range
    $words = $p.Words
    $anchor = $words.Item(1)
    $c = $d.Comments.Add($anchor, $Text)
    Out ([ordered]@{ ok = $true; anchor = $anchor.Text; author = $c.Author; comments = $d.Comments.Count })
  }
  "quit" {
    $app = $d.Application
    $names = @()
    foreach ($x in $app.Documents) { $names += $x.FullName }
    $allow = (Norm $AllowDir).TrimEnd("\") + "\"
    $outside = @($names | Where-Object { -not (Norm $_).StartsWith($allow, [System.StringComparison]::OrdinalIgnoreCase) })
    # Refused: Word stays up with every document untouched. The runner must not
    # stop the process either (cleanupDecision in judge.mjs).
    if ($outside.Count -gt 0) { Out ([ordered]@{ ok = $false; refused = "foreign_documents"; error = "documents outside the test folder are open"; documents = $names; outside = $outside; pid = $ExpectPid }); break }
    foreach ($x in $app.Documents) { $x.Saved = $true }
    $app.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($d) | Out-Null
    $gone = $false
    if ($ExpectPid -gt 0) {
      for ($i = 0; $i -lt 60; $i++) {
        if (-not (Get-Process -Id $ExpectPid -ErrorAction SilentlyContinue)) { $gone = $true; break }
        Start-Sleep -Milliseconds 500
      }
    }
    Out ([ordered]@{ ok = $true; documents = $names; pid = $ExpectPid; exited = $gone })
  }
  default { throw "unknown action $Action" }
}
