param(
  [int]$Port = 8099
)

$ErrorActionPreference = "Stop"

try {
  $health = Invoke-RestMethod -Uri "http://127.0.0.1:$Port/api/health" -TimeoutSec 2
} catch {
  Write-Output "Service is not running on port $Port."
  exit 0
}

if ($health.service -ne "yh-sms-reconciliation" -or -not $health.pid) {
  throw "The listener on port $Port is not the YH reconciliation service."
}

Stop-Process -Id ([int]$health.pid) -Force
Write-Output "Stopped YH reconciliation service (PID $($health.pid))."
