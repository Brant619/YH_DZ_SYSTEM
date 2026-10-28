param(
  [Parameter(Mandatory=$true)][string]$OutputXlsx,
  [Parameter(Mandatory=$true)][string]$SourcesJson
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

function ConvertTo-JsonLine($obj) {
  $obj | ConvertTo-Json -Depth 6 -Compress
}

function Safe-SheetName([string]$name) {
  $safe = $name -replace '[:\\/\?\*\[\]]', '_'
  if ($safe.Length -gt 31) {
    return $safe.Substring(0, 31)
  }
  return $safe
}

$excel = $null
$outputWorkbook = $null
$openedWorkbooks = New-Object System.Collections.Generic.List[object]

try {
  $sources = $SourcesJson | ConvertFrom-Json
  if (-not $sources -or $sources.Count -eq 0) {
    throw 'No source workbooks were provided.'
  }

  foreach ($source in $sources) {
    if (-not (Test-Path -LiteralPath $source.filePath)) {
      throw "Source workbook not found: $($source.filePath)"
    }
  }

  $outDir = Split-Path -Parent $OutputXlsx
  if (-not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
  }
  if (Test-Path -LiteralPath $OutputXlsx) {
    try {
      Remove-Item -LiteralPath $OutputXlsx -Force
    } catch {
      $stamp = Get-Date -Format 'yyyyMMddHHmmss'
      $base = [System.IO.Path]::GetFileNameWithoutExtension($OutputXlsx)
      $OutputXlsx = Join-Path $outDir "$base`_$stamp.xlsx"
    }
  }

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false

  $outputWorkbook = $excel.Workbooks.Add()
  while ($outputWorkbook.Worksheets.Count -gt 1) {
    $outputWorkbook.Worksheets.Item($outputWorkbook.Worksheets.Count).Delete()
  }
  $placeholder = $outputWorkbook.Worksheets.Item(1)
  $placeholder.Name = 'placeholder'

  $copied = 0
  foreach ($source in $sources) {
    $sourceWorkbook = $excel.Workbooks.Open($source.filePath, 3, $true)
    $openedWorkbooks.Add($sourceWorkbook) | Out-Null
    $sheet = $sourceWorkbook.Worksheets.Item(1)
    $sheet.Copy([Type]::Missing, $outputWorkbook.Worksheets.Item($outputWorkbook.Worksheets.Count))
    $newSheet = $outputWorkbook.Worksheets.Item($outputWorkbook.Worksheets.Count)
    $newSheet.Name = Safe-SheetName ([string]$source.sheetName)
    $copied++
  }

  if ($copied -gt 0) {
    $placeholder.Delete()
  }
  $outputWorkbook.Worksheets.Item(1).Activate() | Out-Null
  $outputWorkbook.SaveAs($OutputXlsx, 51)

  ConvertTo-JsonLine @{
    ok = $true
    outputFile = $OutputXlsx
    sheetCount = $copied
    sheets = @($sources | ForEach-Object { $_.sheetName })
  }
}
catch {
  ConvertTo-JsonLine @{
    ok = $false
    reason = $_.Exception.Message
    outputFile = $OutputXlsx
  }
  exit 1
}
finally {
  foreach ($workbook in $openedWorkbooks) {
    try { $workbook.Close($false) | Out-Null } catch {}
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch {}
  }
  if ($outputWorkbook -ne $null) {
    try { $outputWorkbook.Close($false) | Out-Null } catch {}
    [Runtime.InteropServices.Marshal]::ReleaseComObject($outputWorkbook) | Out-Null
  }
  if ($excel -ne $null) {
    try { $excel.Quit() | Out-Null } catch {}
    [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
