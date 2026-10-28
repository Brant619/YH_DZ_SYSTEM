param(
  [string]$DshConfigPath = (Join-Path $env:USERPROFILE ".dsh\profiles\desktop\cordis.patch.yml")
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $PSScriptRoot
$McpServer = Join-Path $ProjectRoot "mcp\server.js"
$ConfigPath = Join-Path $ProjectRoot "config\local.json"
$NodePath = (Get-Command node.exe -ErrorAction SilentlyContinue).Source

if (-not $NodePath) {
  throw "node.exe was not found in PATH."
}
if (-not (Test-Path -LiteralPath $DshConfigPath)) {
  throw "DSH config was not found: $DshConfigPath"
}

$existing = Get-Content -LiteralPath $DshConfigPath -Raw
if ($existing -match "id:\s*mcp-yh-reconciliation\b") {
  Write-Output "DSH MCP entry already exists: $DshConfigPath"
  exit 0
}

function ConvertTo-YamlSingleQuoted([string]$Value) {
  return "'" + $Value.Replace("'", "''") + "'"
}

$nodeYaml = ConvertTo-YamlSingleQuoted $NodePath
$serverYaml = ConvertTo-YamlSingleQuoted $McpServer
$rootYaml = ConvertTo-YamlSingleQuoted $ProjectRoot
$configYaml = ConvertTo-YamlSingleQuoted $ConfigPath

$block = @"

- insert:
    - id: mcp-yh-reconciliation
      name: "@deepseek-ai/dsh-mcp-client"
      config:
        serverName: yh-reconciliation
        transport: stdio
        command: $nodeYaml
        args:
          - $serverYaml
        env:
          YH_PROJECT_ROOT: $rootYaml
          YH_CONFIG_PATH: $configYaml
        toolCallTimeoutMs: 900000
        failOnStartupError: false
"@

Copy-Item -LiteralPath $DshConfigPath -Destination "$DshConfigPath.bak-yh" -Force
Add-Content -LiteralPath $DshConfigPath -Value $block -Encoding utf8
Write-Output "Installed DSH MCP entry in: $DshConfigPath"
Write-Output "Restart DeepSeek Harness to load mcp__yh-reconciliation__* tools."
