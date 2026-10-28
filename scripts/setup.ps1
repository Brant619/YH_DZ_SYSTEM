param(
  [string]$DataDirectory = "D:\YH\永辉短信对账",
  [string]$EdgeProfileDirectory = "",
  [string]$LegacyConfigPath = "D:\YH\永辉短信对账\ai\config.json",
  [switch]$Force
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $PSScriptRoot
$ConfigDirectory = Join-Path $ProjectRoot "config"
$ConfigPath = Join-Path $ConfigDirectory "local.json"
$ExamplePath = Join-Path $ConfigDirectory "config.example.json"

New-Item -ItemType Directory -Force -Path $ConfigDirectory | Out-Null

if ((Test-Path -LiteralPath $ConfigPath) -and -not $Force) {
  Write-Output "Config already exists: $ConfigPath"
  Write-Output "Use -Force to regenerate it."
  exit 0
}

if (Test-Path -LiteralPath $LegacyConfigPath) {
  $legacyJson = [IO.File]::ReadAllText($LegacyConfigPath, [Text.Encoding]::UTF8)
  $config = $legacyJson | ConvertFrom-Json
} else {
  $exampleJson = [IO.File]::ReadAllText($ExamplePath, [Text.Encoding]::UTF8)
  $config = $exampleJson | ConvertFrom-Json
}

$resolvedDataDirectory = [IO.Path]::GetFullPath($DataDirectory)
$config | Add-Member -NotePropertyName dataDirectory -NotePropertyValue $resolvedDataDirectory -Force
$config | Add-Member -NotePropertyName appDirectory -NotePropertyValue (Join-Path $ProjectRoot "src") -Force

if (-not $config.seastar) {
  $config | Add-Member -NotePropertyName seastar -NotePropertyValue ([pscustomobject]@{}) -Force
}

$profilePath = if ($EdgeProfileDirectory) {
  [IO.Path]::GetFullPath($EdgeProfileDirectory)
} elseif ($config.seastar.browserProfileDirectory -and (Test-Path -LiteralPath $config.seastar.browserProfileDirectory)) {
  [IO.Path]::GetFullPath($config.seastar.browserProfileDirectory)
} elseif (Test-Path -LiteralPath (Join-Path $resolvedDataDirectory "ai\edge-profile")) {
  Join-Path $resolvedDataDirectory "ai\edge-profile"
} else {
  Join-Path $resolvedDataDirectory ".edge-profile"
}
$config.seastar | Add-Member -NotePropertyName browserProfileDirectory -NotePropertyValue $profilePath -Force

$config | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $ConfigPath -Encoding utf8

Write-Output "Created local config: $ConfigPath"
Write-Output "Data directory: $resolvedDataDirectory"
Write-Output "Edge profile: $profilePath"

if (-not (Test-Path -LiteralPath $resolvedDataDirectory)) {
  Write-Warning "The data directory does not exist yet. Restore it from the USB drive before starting a workflow."
}
