# Registers the Word add-in for the current Windows user, the way the tray
# app does (app/sideload.mjs winInstall): the manifest is copied to
# %APPDATA%\Sidewrite\manifests and a value named after its full path is
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
$dir = Join-Path $env:APPDATA "Sidewrite\manifests"
$dst = Join-Path $dir (Split-Path -Leaf $Manifest)

function Remove-LegacyManifest {
  # Exact shipped files before PR #15; keep custom manifests, even when they
  # use the old filename or GUID. These hashes also live in legacy-manifests.mjs.
  $hashes = @{
    "word.xml" = "e4c72462952e21aad24517c5fe474dfb370d05ff99b279f81bb3a44b9fbc834d"
    "excel.xml" = "4573041c6fcd00e7de42cf5ec7c39424cd963bceb936afde98ea298e8a91fd2f"
  }
  $file = Split-Path -Leaf $Manifest
  if (-not $hashes.ContainsKey($file)) { return }
  $legacyRoot = Join-Path $env:APPDATA "Draftspect"
  $legacyDir = Join-Path $legacyRoot "manifests"
  $legacy = Join-Path $legacyDir $file
  foreach ($path in @($legacyRoot, $legacyDir, $legacy)) {
    $item = Get-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
    if ($item -and ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
      Write-Warning "legacy manifest: preserved redirected path ($path)"
      return
    }
  }
  $exists = Test-Path -LiteralPath $legacy
  if ($exists) {
    if (-not (Test-Path -LiteralPath $legacy -PathType Leaf)) { return }
    $text = [IO.File]::ReadAllText($legacy).TrimStart([char]0xFEFF).Replace("`r`n", "`n")
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
      $hash = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($text))).Replace("-", "").ToLowerInvariant()
    } finally {
      $sha.Dispose()
    }
    if ($hash -ne $hashes[$file]) {
      Write-Warning "legacy manifest: preserved custom file ($legacy)"
      return
    }
  }
  if (Test-Path -LiteralPath $keyPath) {
    $key = Get-Item -LiteralPath $keyPath
    $value = $key.GetValue($legacy, $null)
    if ($null -ne $value) {
      if (($key.GetValueKind($legacy) -ne [Microsoft.Win32.RegistryValueKind]::String) -or ($value -ne $legacy)) {
        Write-Warning "legacy manifest: preserved custom registration ($legacy)"
        return
      }
      Remove-ItemProperty -LiteralPath $keyPath -Name $legacy
    }
  }
  if ($exists) { Remove-Item -LiteralPath $legacy -Force }
}

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
  Remove-LegacyManifest
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
# Register the replacement first; a failed copy/registration leaves the old
# add-in usable. Only this host's exact old manifest and value are removed.
Remove-LegacyManifest
