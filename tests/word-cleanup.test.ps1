# Pure ownership tests: parse production code, execute with fake applications.
# No Word, ROT lookup, native window call or real process operation is executed.
$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
function Parse($relative) {
  $tokens = $null; $errors = $null
  $tree = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root $relative), [ref]$tokens, [ref]$errors)
  if ($errors.Count) { throw ($errors | Out-String) }
  return $tree
}
$own = Parse "scripts/e2e/word-own.ps1"
$com = Parse "scripts/e2e/word-com.ps1"
function WithoutComRelease([string]$text) {
  return $text.Replace('[System.Runtime.InteropServices.Marshal]::ReleaseComObject($d) | Out-Null', '$null = $d')
}
foreach ($name in @("Under", "OpenNames", "AssertOwnWord", "QuitOwnWord")) {
  $fn = $own.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
  . ([scriptblock]::Create((WithoutComRelease $fn.Extent.Text)))
}
$switch = $com.Find({ param($n) $n -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)
$quit = @($switch.Clauses | Where-Object { $_.Item1.Value -eq "quit" })[0].Item2.Extent.Text
# Imports are not run: AppProcessId below supplies the measured ownership.
$quit = $quit.Replace('. (Join-Path $PSScriptRoot "word-own.ps1")', '')
$quit = [scriptblock]::Create('switch ("quit") { "quit" ' + (WithoutComRelease $quit) + ' }')
function AppProcessId($app) { return $script:owner }
function Get-Process { return $null }
function Start-Sleep { throw "unexpected wait" }
function Norm([string]$s) { return $s.Trim() }
function Out($o) { return $o }
function Check($condition, $label) {
  if (-not $condition) { throw "FAIL: $label" }
  Write-Output "PASS: $label"
}
$AllowDir = "C:\run"
$p = [pscustomobject]@{ Id = 42 }
foreach ($case in @(
  @{ name = "unknown owner"; owner = 0; expected = 42; foreign = $false; ours = $true; allowed = $false },
  @{ name = "mismatched owner"; owner = 99; expected = 42; foreign = $false; ours = $true; allowed = $false },
  @{ name = "unknown expected PID"; owner = 42; expected = 0; foreign = $false; ours = $true; allowed = $false },
  @{ name = "foreign document"; owner = 42; expected = 42; foreign = $true; ours = $true; allowed = $false },
  @{ name = "owned documents"; owner = 42; expected = 42; foreign = $false; ours = $true; allowed = $true }
)) {
  $script:owner = $case.owner
  $ExpectPid = $case.expected
  $p.Id = $case.expected
  foreach ($entry in @("own", "com")) {
    $doc = [pscustomobject]@{ FullName = "C:\run\anchor.docx"; Saved = $false }
    $foreign = [pscustomobject]@{ FullName = "C:\user\notes.docx"; Saved = $false }
    $app = [pscustomobject]@{ Documents = @($doc); QuitCalled = $false }
    $app | Add-Member ScriptMethod Quit { $this.QuitCalled = $true }
    if ($case.foreign) { $app.Documents += $foreign }
    $d = [pscustomobject]@{ Application = $app }
    if ($entry -eq "own") {
      $asserted = $true
      try { AssertOwnWord $app $p } catch { $asserted = $false }
      Check ($asserted -eq $case.allowed) "$($case.name): AssertOwnWord"
      $answer = QuitOwnWord $app $d $p $case.ours
    } else { $answer = & $quit }
    Check ($answer.ok -eq $case.allowed) "$($case.name): $entry quit decision"
    Check ($app.QuitCalled -eq $case.allowed) "$($case.name): $entry Quit call"
    Check ($doc.Saved -eq $case.allowed) "$($case.name): $entry Saved mutation"
    Check (-not $foreign.Saved) "$($case.name): $entry foreign unchanged"
    if (-not $case.allowed) {
      $reason = if ($case.foreign) { "foreign_documents" } else { "ownership_unproven" }
      Check ($answer.refused -eq $reason) "$($case.name): $entry refusal preserved"
    }
  }
}
Write-Output "PowerShell ownership checks: 0 failures; changed PowerShell files: 0 parse errors"
