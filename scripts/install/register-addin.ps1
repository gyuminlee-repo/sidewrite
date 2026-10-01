# Registers the Word add-in for the current Windows user, the way the tray
# app does (app/sideload.mjs winInstall): the manifest is copied to
# %APPDATA%\Draftspect\manifests and a value named after its full path is
# written under the WEF\Developer key. Word reads that key when it starts
# and loads the add-in. Current user only; no administrator rights.
#
#   register-addin.ps1 -Manifest <path of word.xml>
#   register-addin.ps1 -Manifest <path of word.xml> -Unregister
#
# Run again, it leaves what is in place alone. The names here are the ones
# of scripts/install/win-args.mjs (registerPlan); a test holds them together.
param(
  [Parameter(Mandatory = $true)][string]$Manifest,
  [switch]$Unregister
)
$ErrorActionPreference = "Stop"

$keyPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
$dir = Join-Path $env:APPDATA "Draftspect\manifests"
$dst = Join-Path $dir (Split-Path -Leaf $Manifest)

function Test-Registered {
  if (-not (Test-Path -LiteralPath $keyPath)) { return $false }
  return $null -ne (Get-Item -LiteralPath $keyPath).GetValue($dst, $null)
}

if ($Unregister) {
  if (Test-Registered) {
    Remove-ItemProperty -LiteralPath $keyPath -Name $dst
    Write-Output "registry: removed"
  } else {
    Write-Output "registry: not there"
  }
  if (Test-Path -LiteralPath $dst) {
    Remove-Item -LiteralPath $dst -Force
    Write-Output "manifest: removed ($dst)"
  } else {
    Write-Output "manifest: not there"
  }
  exit 0
}

if (-not (Test-Path -LiteralPath $Manifest)) {
  throw "No manifest at $Manifest"
}
$same = $false
if (Test-Path -LiteralPath $dst) {
  $same = (Get-FileHash -LiteralPath $dst).Hash -eq (Get-FileHash -LiteralPath $Manifest).Hash
}
if ($same) {
  Write-Output "manifest: in place ($dst)"
} else {
  New-Item -ItemType Directory -Force -Path $dir | Out-Null
  Copy-Item -LiteralPath $Manifest -Destination $dst -Force
  Write-Output "manifest: copied ($dst)"
}
if (Test-Registered) {
  Write-Output "registry: in place"
} else {
  if (-not (Test-Path -LiteralPath $keyPath)) {
    New-Item -Path $keyPath -Force | Out-Null
  }
  Set-ItemProperty -LiteralPath $keyPath -Name $dst -Value $dst -Type String
  Write-Output "registry: written (restart Word to see the add-in)"
}
