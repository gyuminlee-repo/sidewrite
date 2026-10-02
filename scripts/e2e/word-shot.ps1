# A picture of the main window of one Word process, as that window draws
# itself (PrintWindow): the task pane beside the document. Nothing of another
# window is in it, the window is not moved, shown, focused or changed, and
# nothing is installed (System.Drawing and user32 are part of Windows).
# Used by run-acceptance.mjs --screenshots on the Word that run started.
# Prints one JSON line: { ok, width, height, minimized, drawn }.
param(
  [Parameter(Mandatory = $true)][int]$ProcessId,
  [Parameter(Mandatory = $true)][string]$Out
)
$ErrorActionPreference = "Stop"
function Answer($o) { $o | ConvertTo-Json -Compress }
try {
  Add-Type -AssemblyName System.Drawing
  Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class WordShot {
  [StructLayout(LayoutKind.Sequential)]
  public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdc, uint flags);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
}
"@
  $proc = Get-Process -Id $ProcessId
  if ($proc.ProcessName -ne "WINWORD") { throw "process $ProcessId is $($proc.ProcessName), not WINWORD" }
  $handle = $proc.MainWindowHandle
  if ($handle -eq [IntPtr]::Zero) { throw "process $ProcessId has no main window" }
  $minimized = [WordShot]::IsIconic($handle)
  $rect = New-Object WordShot+RECT
  if (-not [WordShot]::GetWindowRect($handle, [ref]$rect)) { throw "GetWindowRect failed" }
  $width = $rect.Right - $rect.Left
  $height = $rect.Bottom - $rect.Top
  if ($width -lt 1 -or $height -lt 1) { throw "the window has no size ($width x $height)" }
  $bitmap = New-Object System.Drawing.Bitmap $width, $height
  $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
  $hdc = $graphics.GetHdc()
  # 2: PW_RENDERFULLCONTENT, so content drawn by the compositor (the pane) is in it.
  $drawn = [WordShot]::PrintWindow($handle, $hdc, 2)
  $graphics.ReleaseHdc($hdc)
  $graphics.Dispose()
  $bitmap.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png)
  $bitmap.Dispose()
  Answer ([ordered]@{ ok = [bool]$drawn; width = $width; height = $height; minimized = [bool]$minimized; drawn = [bool]$drawn })
} catch {
  Answer ([ordered]@{ ok = $false; error = $_.Exception.Message })
}
