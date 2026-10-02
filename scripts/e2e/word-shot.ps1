# A picture of the main window of one Word process, as that window draws
# itself (PrintWindow): the task pane beside the document. Nothing of another
# window is in it, the window is not moved, shown, focused or changed, and
# nothing is installed (System.Drawing and user32 are part of Windows).
# Used by run-acceptance.mjs --screenshots on the Word that run started.
#
# This process is made DPI aware first, so the size it reads is the size of
# the window in real pixels (without that, a scaled display gives a smaller
# size and the picture is the top left corner only).
# The task pane is drawn by WebView2 and can come out black in such a
# picture. A picture with a black area that large is not written: no file is
# better than one that shows a pane that was never black on the screen.
# Prints one JSON line: { ok, width, height, minimized, drawn, dpi_aware,
# black_share, dropped }.
param(
  [Parameter(Mandatory = $true)][int]$ProcessId,
  [Parameter(Mandatory = $true)][string]$Out,
  [double]$MaxBlackShare = 0.10
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
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdc, uint flags);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
}
"@
  $aware = [WordShot]::SetProcessDPIAware()
  $proc = Get-Process -Id $ProcessId
  if ($proc.ProcessName -ne "WINWORD") { throw "process $ProcessId is $($proc.ProcessName), not WINWORD" }
  $handle = $proc.MainWindowHandle
  if ($handle -eq [IntPtr]::Zero) { throw "process $ProcessId has no main window" }
  $minimized = [WordShot]::IsIconic($handle)
  if ($minimized) { throw "the window is minimized; it is left as it is and not photographed" }
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

  # The share of black pixels, read from every fourth pixel of every fourth row.
  $area = New-Object System.Drawing.Rectangle 0, 0, $width, $height
  $data = $bitmap.LockBits($area, [System.Drawing.Imaging.ImageLockMode]::ReadOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $bytes = New-Object byte[] ($data.Stride * $height)
  [System.Runtime.InteropServices.Marshal]::Copy($data.Scan0, $bytes, 0, $bytes.Length)
  $bitmap.UnlockBits($data)
  $seen = 0
  $black = 0
  for ($y = 0; $y -lt $height; $y += 4) {
    $row = $y * $data.Stride
    for ($x = 0; $x -lt $width; $x += 4) {
      $i = $row + $x * 4
      $seen++
      if ($bytes[$i] -lt 8 -and $bytes[$i + 1] -lt 8 -and $bytes[$i + 2] -lt 8) { $black++ }
    }
  }
  $share = if ($seen) { [math]::Round($black / $seen, 4) } else { 1 }
  $dropped = (-not $drawn) -or ($share -gt $MaxBlackShare)
  if (-not $dropped) { $bitmap.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png) }
  $bitmap.Dispose()
  Answer ([ordered]@{
    ok = (-not $dropped)
    width = $width
    height = $height
    minimized = [bool]$minimized
    drawn = [bool]$drawn
    dpi_aware = [bool]$aware
    black_share = $share
    dropped = [bool]$dropped
  })
} catch {
  Answer ([ordered]@{ ok = $false; error = $_.Exception.Message })
}
