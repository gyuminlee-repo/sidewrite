# Start a separate Word process whose WebView2 add-in panes expose a Chrome
# DevTools Protocol port, and open one document in it.
#
# The debug port is passed only to this Word process: the WebView2 variables
# are set in this PowerShell process and inherited by the child (Microsoft
# WebView2 docs: WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS with
# --remote-debugging-port, WEBVIEW2_USER_DATA_FOLDER). No user or machine
# environment variable is written. A separate user data folder keeps the pane
# out of the WebView2 browser process of any Word that is already running.
# /n starts a new Word instance even when Word is already open.
# Once its window is there, the window of this Word (only this process) is
# sent behind the user's windows without taking the focus, and kept there
# by a hidden watch until it exits (word-back.ps1). It is never minimized.
# Prints one JSON line: { pid, preexisting, exe, port, user_data_folder,
# back (sent to the back once), back_watch (pid of the watch, 0 if none) }.
param(
  [Parameter(Mandatory = $true)][string]$DocPath,
  [Parameter(Mandatory = $true)][int]$Port,
  [Parameter(Mandatory = $true)][string]$UserDataFolder,
  [int]$WaitSeconds = 60
)
$ErrorActionPreference = "Stop"
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$busy = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
if ($busy) { throw "port $Port is already listening (pid $($busy[0].OwningProcess))" }
$exe = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe")."(default)"
if (-not $exe -or -not (Test-Path $exe)) { throw "WINWORD.EXE not found in App Paths" }
New-Item -ItemType Directory -Force -Path $UserDataFolder | Out-Null
$env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS = "--remote-debugging-port=$Port"
$env:WEBVIEW2_USER_DATA_FOLDER = $UserDataFolder
$p = Start-Process -FilePath $exe -ArgumentList @("/n", "/q", ('"' + $DocPath + '"')) -PassThru
$deadline = (Get-Date).AddSeconds($WaitSeconds)
while ((Get-Date) -lt $deadline -and -not $p.HasExited) {
  if ($p.MainWindowHandle -ne 0) { break }
  Start-Sleep -Milliseconds 500
  $p.Refresh()
}
$back = $false
$watch = 0
if (-not $p.HasExited -and $before -notcontains $p.Id) {
  . (Join-Path $PSScriptRoot "word-back.ps1")
  $back = SendWordToBack $p.Id
  $watch = StartWordBackWatch $p.Id
}
$out = [ordered]@{
  pid = $p.Id
  exited = $p.HasExited
  preexisting = $before
  exe = $exe
  port = $Port
  user_data_folder = $UserDataFolder
  back = [bool]$back
  back_watch = $watch
}
$out | ConvertTo-Json -Compress
