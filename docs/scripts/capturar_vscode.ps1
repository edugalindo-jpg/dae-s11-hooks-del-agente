param(
  [string]$Folder = "C:\Users\User\Documents\web\_repo_s11",
  [string]$File = "src\hooks\usePlans.js",
  [string]$Out = "C:\Users\User\Documents\web\_repo_s11\docs\capturas\16_vscode_estructura.png",
  [int]$ZoomSteps = 3
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WinCap {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
}
"@

$leaf = Split-Path -Leaf $Folder
$existing = Get-Process Code -ErrorAction SilentlyContinue |
  Where-Object { $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -like "*$leaf*" } |
  Select-Object -First 1

$launched = $false
if (-not $existing) {
  $filePath = Join-Path $Folder $File
  Start-Process "code.cmd" -ArgumentList @("-n", "`"$Folder`"", "-g", "`"$filePath`:1`"") | Out-Null
  $launched = $true
  $deadline = (Get-Date).AddSeconds(60)
  do {
    Start-Sleep -Milliseconds 700
    $existing = Get-Process Code -ErrorAction SilentlyContinue |
      Where-Object { $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -like "*$leaf*" } |
      Select-Object -First 1
  } until ($existing -or (Get-Date) -gt $deadline)
}

if (-not $existing) { throw "No se encontro la ventana de VS Code para '$leaf'." }

$handle = $existing.MainWindowHandle
[WinCap]::ShowWindow($handle, 3) | Out-Null
[WinCap]::SetForegroundWindow($handle) | Out-Null
Start-Sleep -Seconds 2

if ($ZoomSteps -gt 0) {
  [System.Windows.Forms.SendKeys]::SendWait(("{=}" * 0 + ("^{=}" * $ZoomSteps)))
  Start-Sleep -Seconds 2
}

[WinCap]::SetForegroundWindow($handle) | Out-Null
Start-Sleep -Milliseconds 700

$rect = New-Object WinCap+RECT
[WinCap]::GetWindowRect($handle, [ref]$rect) | Out-Null
$width = $rect.Right - $rect.Left
$height = $rect.Bottom - $rect.Top
Write-Output "title: $((Get-Process -Id $existing.Id).MainWindowTitle)"
Write-Output "rect: $($rect.Left),$($rect.Top) ${width}x${height}"

function Is-Blank([System.Drawing.Bitmap]$bmp) {
  $colors = @{}
  for ($x = 0; $x -lt $bmp.Width; $x += [Math]::Max(1, [int]($bmp.Width / 40))) {
    for ($y = 0; $y -lt $bmp.Height; $y += [Math]::Max(1, [int]($bmp.Height / 40))) {
      $colors[$bmp.GetPixel($x, $y).ToArgb()] = 1
    }
  }
  return ($colors.Count -lt 8)
}

$bitmap = New-Object System.Drawing.Bitmap($width, $height)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$hdc = $graphics.GetHdc()
$ok = [WinCap]::PrintWindow($handle, $hdc, 2)
$graphics.ReleaseHdc($hdc)
$graphics.Dispose()
$mode = "PrintWindow=$ok"

if (Is-Blank $bitmap) {
  $bitmap.Dispose()
  [WinCap]::SetForegroundWindow($handle) | Out-Null
  Start-Sleep -Seconds 1
  $bitmap = New-Object System.Drawing.Bitmap($width, $height)
  $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
  $graphics.CopyFromScreen($rect.Left, $rect.Top, 0, 0, $bitmap.Size)
  $graphics.Dispose()
  $mode = "ScreenCopy"
}

$dir = Split-Path -Parent $Out
if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
$bitmap.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png)
$bitmap.Dispose()
Write-Output "captured: $Out ($width x $height) $mode"

if ($launched) {
  # Close only the window we opened (WM_CLOSE), then let VS Code exit on its own.
  [WinCap]::PostMessage($handle, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
}
