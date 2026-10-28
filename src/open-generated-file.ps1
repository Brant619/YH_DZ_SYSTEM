param(
  [string]$TargetPath,
  [string]$TargetPathFile
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($TargetPath)) {
  $TargetPath = $env:YH_OPEN_TARGET_PATH
}
if ([string]::IsNullOrWhiteSpace($TargetPath) -and -not [string]::IsNullOrWhiteSpace($TargetPathFile)) {
  $TargetPath = Get-Content -LiteralPath $TargetPathFile -Raw -Encoding UTF8
}
if ([string]::IsNullOrWhiteSpace($TargetPath)) {
  throw "Target path is required."
}
$TargetPath = $TargetPath.Trim()
$resolvedPath = (Resolve-Path -LiteralPath $TargetPath).Path
$item = Get-Item -LiteralPath $resolvedPath

if (-not $item.PSIsContainer) {
  Start-Process -FilePath $resolvedPath
  $resolvedPath
  return
}

$openPath = if ($item.PSIsContainer) { $resolvedPath } else { $item.DirectoryName }
$normalizedTarget = [System.IO.Path]::GetFullPath($openPath).TrimEnd("\")

Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class WindowActivation {
  [DllImport("user32.dll")]
  public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

  [DllImport("user32.dll")]
  public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@

function Find-ExplorerWindow {
  param($Shell, [string]$LocationPath)

  return @($Shell.Windows()) |
    Where-Object {
      if (-not $_.FullName -or [System.IO.Path]::GetFileName($_.FullName) -ine "explorer.exe") {
        return $false
      }
      try {
        $windowPath = [System.IO.Path]::GetFullPath(([System.Uri]$_.LocationURL).LocalPath).TrimEnd("\")
        return $windowPath -ieq $LocationPath
      } catch {
        return $false
      }
    } |
    Select-Object -First 1
}

$shell = New-Object -ComObject Shell.Application
$window = Find-ExplorerWindow $shell $normalizedTarget

if (-not $window) {
  if ($item.PSIsContainer) {
    $shell.Explore($resolvedPath)
  } else {
    Start-Process -FilePath explorer.exe -ArgumentList @("/select,", $resolvedPath)
  }
  for ($attempt = 0; $attempt -lt 20 -and -not $window; $attempt++) {
    Start-Sleep -Milliseconds 150
    $window = Find-ExplorerWindow $shell $normalizedTarget
  }
}

if (-not $window) {
  throw "Explorer window did not open for: $resolvedPath"
}

$handle = [IntPtr]::new([int64]$window.HWND)
[WindowActivation]::ShowWindowAsync($handle, 9) | Out-Null
Start-Sleep -Milliseconds 100
[WindowActivation]::SetForegroundWindow($handle) | Out-Null

$resolvedPath
