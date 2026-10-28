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

$existing = [IO.File]::ReadAllText($DshConfigPath, [Text.Encoding]::UTF8)

function ConvertTo-YamlSingleQuoted([string]$Value) {
  return "'" + $Value.Replace("'", "''") + "'"
}

$nodeYaml = ConvertTo-YamlSingleQuoted $NodePath
$serverYaml = ConvertTo-YamlSingleQuoted $McpServer
$rootYaml = ConvertTo-YamlSingleQuoted $ProjectRoot
$configYaml = ConvertTo-YamlSingleQuoted $ConfigPath

$block = @"

# BEGIN YH-RECONCILIATION MCP
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
# END YH-RECONCILIATION MCP
"@.Trim()

$markerPattern = "(?s)# BEGIN YH-RECONCILIATION MCP.*?# END YH-RECONCILIATION MCP"
if ($existing -match $markerPattern) {
  $updated = [regex]::Replace($existing, $markerPattern, $block)
} elseif ($existing -match "id:\s*mcp-yh-reconciliation\b") {
  $legacyPattern = "(?ms)^- insert:\s*\r?\n\s+- id: mcp-yh-reconciliation\b.*?(?=^- |\z)"
  $updated = [regex]::Replace($existing, $legacyPattern, "")
  $updated = $updated.TrimEnd() + "`r`n`r`n" + $block + "`r`n"
} else {
  $updated = $existing.TrimEnd() + "`r`n`r`n" + $block + "`r`n"
}

Copy-Item -LiteralPath $DshConfigPath -Destination "$DshConfigPath.bak-yh" -Force
[IO.File]::WriteAllText($DshConfigPath, $updated, [Text.UTF8Encoding]::new($false))
Write-Output "Installed DSH MCP entry in: $DshConfigPath"
Write-Output "Restart DeepSeek Harness to load mcp__yh-reconciliation__* tools."
