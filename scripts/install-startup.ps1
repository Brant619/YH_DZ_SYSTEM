param(
  [int]$Port = 8099
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $PSScriptRoot
$StartScript = Join-Path $PSScriptRoot "start-service.ps1"
$StartupDirectory = [Environment]::GetFolderPath("Startup")
$ShortcutPath = Join-Path $StartupDirectory "YH-SMS-Reconciliation-Service.lnk"
$PowerShellExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($ShortcutPath)
$shortcut.TargetPath = $PowerShellExe
$shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$StartScript`" -Port $Port"
$shortcut.WorkingDirectory = $ProjectRoot
$shortcut.IconLocation = "$PowerShellExe,0"
$shortcut.Description = "Start the YH SMS reconciliation service on port $Port"
$shortcut.Save()

Write-Output "Startup shortcut installed: $ShortcutPath"
