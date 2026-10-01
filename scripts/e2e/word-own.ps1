# A Word process of the run's own. Dot-sourced by marked-diff-word.ps1 and
# word-math-probe.ps1; the caller has $AllowDir (the run folder).
#
# StartOwnWord starts Word (/n: a new instance even when Word is already
# open) on one anchor document of the run folder and finds that document
# through the Running Object Table by its full path (as word-com.ps1 does).
# A Word that was already open, and its documents, are never touched:
# AssertOwnWord lets the Application be used only when it is the process
# this run started and every document open in it is under $AllowDir, and
# QuitOwnWord quits Word only under the same two conditions.

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

# Starts the Word of this run on the anchor document. Writes -PidFile
# ({ pid, preexisting }) as soon as the process is there. Returns
# @{ before (the Word processes that ran already); process; document (the
# anchor); app }.
function StartOwnWord([string]$Anchor, [string]$PidFile, [int]$WaitSeconds) {
  $utf8 = New-Object System.Text.UTF8Encoding($false)
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
  return @{ before = $before; process = $p; document = $d; app = $d.Application }
}

# Throws unless the Application is the Word this run started and holds only documents of the run folder.
function AssertOwnWord($app, $p) {
  $owner = AppProcessId $app
  if ($owner -ne 0 -and $owner -ne $p.Id) { throw "the anchor document is open in another Word (process $owner), not in the one this run started ($($p.Id))" }
  $outside = @((OpenNames $app) | Where-Object { -not (Under $_) })
  if ($outside.Count -gt 0) { throw "the Word this run reached has other documents open" }
}

# Quit, as word-com.ps1 does: only the Word of this run ($ours: AssertOwnWord
# passed), and only when every open document is under the run folder.
# Returns what happened: { ok, documents, pid, exited } or { ok = false, ... }.
function QuitOwnWord($app, $d, $p, [bool]$ours) {
  try {
    $names = OpenNames $app
    $outside = @($names | Where-Object { -not (Under $_) })
    if (-not $ours -or $outside.Count -gt 0) {
      $why = if ($ours) { "foreign_documents" } else { "not_the_word_of_this_run" }
      return [ordered]@{ ok = $false; refused = $why; documents = $names; outside = $outside; pid = $p.Id }
    }
    foreach ($x in $app.Documents) { $x.Saved = $true }
    $app.Quit() | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($d) | Out-Null
    $gone = $false
    for ($i = 0; $i -lt 60; $i++) {
      if (-not (Get-Process -Id $p.Id -ErrorAction SilentlyContinue)) { $gone = $true; break }
      Start-Sleep -Milliseconds 500
    }
    return [ordered]@{ ok = $true; documents = $names; pid = $p.Id; exited = $gone }
  } catch {
    return [ordered]@{ ok = $false; error = $_.Exception.Message; pid = $p.Id }
  }
}
