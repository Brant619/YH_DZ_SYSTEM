param(
  [int]$Port = 8099,
  [switch]$Open
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $PSScriptRoot
$ServerScript = Join-Path $ProjectRoot "src\server.js"
$ConfigPath = Join-Path $ProjectRoot "config\local.json"
$StateRoot = Join-Path $env:LOCALAPPDATA "YH-SMS-Reconciliation"
$LogDirectory = Join-Path $StateRoot "logs"
$OutLog = Join-Path $LogDirectory "server.out.log"
$ErrLog = Join-Path $LogDirectory "server.err.log"
$Entry = "http://127.0.0.1:$Port/"

if (-not (Test-Path -LiteralPath $ConfigPath)) {
  throw "Missing config/local.json. Run scripts\setup.ps1 first."
}

New-Item -ItemType Directory -Force -Path $LogDirectory | Out-Null

function Test-Service {
  try {
    $health = Invoke-RestMethod -Uri "http://127.0.0.1:$Port/api/health" -TimeoutSec 2
    return $health.service -eq "yh-sms-reconciliation"
  } catch {
    return $false
  }
}

if (Test-Service) {
  Write-Output "Service already running: $Entry"
  if ($Open) { Start-Process $Entry }
  exit 0
}

$connection = Get-NetTCPConnection -LocalAddress "127.0.0.1" -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
if ($connection) {
  throw "Port $Port is already used by another process."
}

$node = (Get-Command node.exe -ErrorAction SilentlyContinue).Source
if (-not $node) {
  throw "node.exe was not found in PATH. Install Node.js 20 or newer first."
}

$previousPort = $env:PORT
$previousConfigPath = $env:YH_CONFIG_PATH
try {
  $env:PORT = [string]$Port
  $env:YH_CONFIG_PATH = $ConfigPath
  Start-Process `
    -FilePath $node `
    -ArgumentList "`"$ServerScript`"" `
    -WorkingDirectory $ProjectRoot `
    -WindowStyle Hidden `
    -RedirectStandardOutput $OutLog `
    -RedirectStandardError $ErrLog
} finally {
  $env:PORT = $previousPort
  $env:YH_CONFIG_PATH = $previousConfigPath
}

for ($attempt = 0; $attempt -lt 30; $attempt++) {
  Start-Sleep -Milliseconds 500
  if (Test-Service) {
    Write-Output "Service started: $Entry"
    if ($Open) { Start-Process $Entry }
    exit 0
  }
}

throw "Service failed to start. Check $ErrLog"
