# Save the copy of the manuscript that replay.mjs opened, with its tracked
# changes and comments as they are. The document is found through the Running
# Object Table by its full path (BindToMoniker), as word-com.ps1 does, so a
# Word that is open with other documents is never touched. Nothing is saved
# unless the document lies under -AllowDir (the run folder).
# Prints one JSON object: { ok, saved, path, revisions, comments }.
param(
  [Parameter(Mandatory = $true)][string]$Doc,
  [Parameter(Mandatory = $true)][string]$AllowDir
)
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Out($o) { $o | ConvertTo-Json -Depth 4 -Compress }

$d = [System.Runtime.InteropServices.Marshal]::BindToMoniker($Doc)
$full = [string]$d.FullName
if ($full -ne $Doc) { throw "bound to $full, expected $Doc" }
$allow = $AllowDir.TrimEnd("\") + "\"
if (-not $full.StartsWith($allow, [System.StringComparison]::OrdinalIgnoreCase)) {
  Out ([ordered]@{ ok = $false; saved = $false; error = "the document is not in the run folder"; path = $full })
  exit 0
}
$d.Save()
Out ([ordered]@{ ok = $true; saved = [bool]$d.Saved; path = $full; revisions = $d.Revisions.Count; comments = $d.Comments.Count })
