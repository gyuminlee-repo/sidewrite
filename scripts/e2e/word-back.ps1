# Keep the Word window of a run behind the user's windows, without taking
# the focus and without minimizing it.
#
# A Word the runners start comes to the front and would sit over what the
# user is doing. It is sent to the bottom of the z-order (SetWindowPos
# HWND_BOTTOM, SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE). It is never
# minimized: with Word minimized the tracked edits of a turn went in without
# their rationale comments (acceptance run of 2026-10-05, stage 3), and with
# Word sent to the back the same run passed. Only the process the run
# started is touched, by its id; no other Word.
#
# Dot-sourced (no -WatchProcessId): defines SendWordToBack and StartWordBackWatch.
# Run with -WatchProcessId: the watch. Word brings its window to the front again
# as it opens a document and loads the pane, so the window of that process
# is sent to the back every 2 s until the process ends (or -WatchSeconds pass).
# A window the user minimized or restored is left as it is shown otherwise:
# a minimized one is not touched.
param(
  [int]$WatchProcessId = 0,
  [int]$WatchSeconds = 21600
)

# This file, for the watch StartWordBackWatch starts.
$SwBackPath = $MyInvocation.MyCommand.Path

if (-not ("SwBack.User32" -as [type])) {
  Add-Type -Namespace SwBack -Name User32 -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetWindowPos(System.IntPtr hWnd, System.IntPtr after, int x, int y, int cx, int cy, uint flags);
[DllImport("user32.dll")] public static extern bool IsIconic(System.IntPtr hWnd);
'@
}

# Sends the main window of process $Id to the bottom of the z-order. True
# when it did; false when the process has no window yet, is gone, or its
# window is minimized (left alone).
function SendWordToBack([int]$Id) {
  $p = Get-Process -Id $Id -ErrorAction SilentlyContinue
  if (-not $p -or $p.ProcessName -ne "WINWORD") { return $false }
  $h = $p.MainWindowHandle
  if ($h -eq [System.IntPtr]::Zero) { return $false }
  if ([SwBack.User32]::IsIconic($h)) { return $false }
  # HWND_BOTTOM = 1; SWP_NOSIZE 0x1 | SWP_NOMOVE 0x2 | SWP_NOACTIVATE 0x10.
  return [SwBack.User32]::SetWindowPos($h, [System.IntPtr]1, 0, 0, 0, 0, 0x13)
}

# Starts the watch for process $Id in a hidden PowerShell of its own, which
# ends with that process. Returns its process id, or 0 when it could not
# start (the run goes on; the window may then stay in front).
function StartWordBackWatch([int]$Id) {
  try {
    $w = Start-Process -FilePath "powershell.exe" -WindowStyle Hidden -PassThru -ArgumentList @(
      "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ('"' + $SwBackPath + '"'), "-WatchProcessId", $Id
    )
    return $w.Id
  } catch {
    return 0
  }
}

if ($WatchProcessId -gt 0) {
  $first = Get-Process -Id $WatchProcessId -ErrorAction SilentlyContinue
  if (-not $first -or $first.ProcessName -ne "WINWORD") { exit 0 }
  $started = $first.StartTime
  $end = (Get-Date).AddSeconds($WatchSeconds)
  while ((Get-Date) -lt $end) {
    $p = Get-Process -Id $WatchProcessId -ErrorAction SilentlyContinue
    # Gone, or the id now belongs to another process: the watch ends.
    if (-not $p -or $p.StartTime -ne $started) { break }
    [void](SendWordToBack $WatchProcessId)
    Start-Sleep -Seconds 2
  }
}
