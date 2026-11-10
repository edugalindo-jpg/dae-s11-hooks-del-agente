param(
  [Parameter(Mandatory = $true)][string]$Folder,
  [Parameter(Mandatory = $true)][string]$Command,
  [Parameter(Mandatory = $true)][string]$Title,
  [Parameter(Mandatory = $true)][string]$Out,
  [int]$WaitSeconds = 15,
  [string]$OccludeWith = ""
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

Add-Type @"
using System;
using System.Runtime.InteropServices;
public class CaptureWin32 {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern bool SetWindowText(IntPtr hWnd, string text);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
}
"@

# The inner script: sets a readable console font (28 px, Consolas), then runs the command.
$inner = @"
`$Host.UI.RawUI.WindowTitle = '$Title'
Set-Location -LiteralPath '$Folder'
`$env:NO_COLOR = '1'
`$env:FORCE_COLOR = '0'
`$env:CI = '1'

Add-Type @'
using System;
using System.Runtime.InteropServices;
[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct CONSOLE_FONT_INFOEX {
  public uint cbSize;
  public uint nFont;
  public short dwFontSizeX;
  public short dwFontSizeY;
  public uint FontFamily;
  public uint FontWeight;
  [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public string FaceName;
}
public class ConsoleFont {
  [DllImport("kernel32.dll", SetLastError = true)] public static extern IntPtr GetStdHandle(int nStdHandle);
  [DllImport("kernel32.dll", SetLastError = true)] public static extern bool SetCurrentConsoleFontEx(IntPtr hConsoleOutput, bool bMaximumWindow, ref CONSOLE_FONT_INFOEX lpConsoleCurrentFontEx);
  public static void Set(string face, short height) {
    CONSOLE_FONT_INFOEX info = new CONSOLE_FONT_INFOEX();
    info.cbSize = (uint)Marshal.SizeOf(info);
    info.FaceName = face;
    info.dwFontSizeX = 0;
    info.dwFontSizeY = height;
    info.FontFamily = 54;
    info.FontWeight = 400;
    SetCurrentConsoleFontEx(GetStdHandle(-11), false, ref info);
  }
}
'@
[ConsoleFont]::Set('Consolas', 28) | Out-Null
Clear-Host
$Command
"@

$temp = Join-Path $env:TEMP ("capture_cmd_" + [guid]::NewGuid().ToString("N") + ".ps1")
Set-Content -Path $temp -Value $inner -Encoding UTF8

$proc = Start-Process powershell.exe -ArgumentList @(
  "-NoExit", "-ExecutionPolicy", "Bypass", "-File", $temp
) -PassThru

# Wait until the console window exists and has the title we set.
$deadline = (Get-Date).AddSeconds(15)
$window = $null
do {
  Start-Sleep -Milliseconds 300
  $window = Get-Process -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -like "*$Title*" } |
    Select-Object -First 1
} until ($window -or (Get-Date) -gt $deadline)

if (-not $window) {
  throw "No se encontro la ventana con titulo '$Title'."
}

$handle = $window.MainWindowHandle
[CaptureWin32]::ShowWindow($handle, 3) | Out-Null   # SW_MAXIMIZE

# Let the command run and the console finish scrolling.
Start-Sleep -Seconds $WaitSeconds
[CaptureWin32]::SetWindowText($handle, $Title) | Out-Null

# Optional self-test: put another window in front before capturing.
if ($OccludeWith) {
  $other = Get-Process -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -like "*$OccludeWith*" } |
    Select-Object -First 1
  if ($other) {
    [CaptureWin32]::SetForegroundWindow($other.MainWindowHandle) | Out-Null
    Start-Sleep -Seconds 1
    Write-Output "occluded with: $($other.MainWindowTitle)"
  }
}

$rect = New-Object CaptureWin32+RECT
[CaptureWin32]::GetWindowRect($handle, [ref]$rect) | Out-Null
$width = $rect.Right - $rect.Left
$height = $rect.Bottom - $rect.Top
if ($width -le 0 -or $height -le 0) { throw "La ventana no tiene un tamano valido." }

$bitmap = New-Object System.Drawing.Bitmap($width, $height)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$hdc = $graphics.GetHdc()
$ok = [CaptureWin32]::PrintWindow($handle, $hdc, 2)   # PW_RENDERFULLCONTENT
$graphics.ReleaseHdc($hdc)
$graphics.Dispose()

$dir = Split-Path -Parent $Out
if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
$bitmap.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png)
$bitmap.Dispose()

Write-Output "captured: $Out ($width x $height) PrintWindow=$ok"

# Close the console window we opened.
if (-not $proc.HasExited) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
