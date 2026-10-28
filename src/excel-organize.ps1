param(
  [Parameter(Mandatory=$true)][string]$SourceCsv,
  [Parameter(Mandatory=$true)][string]$OutputXlsx,
  [int]$TotalColumn = 12
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

function ConvertTo-JsonLine($obj) {
  $obj | ConvertTo-Json -Depth 6 -Compress
}

function Set-CellFormat($cell, $format) {
  try {
    $cell.NumberFormat = $format
  } catch {
    try { $cell.NumberFormatLocal = $format } catch {}
  }
}

$excel = $null
$workbook = $null

try {
  if (-not (Test-Path -LiteralPath $SourceCsv)) {
    throw "Source CSV not found: $SourceCsv"
  }

  $firstLine = Get-Content -LiteralPath $SourceCsv -TotalCount 1 -Encoding UTF8
  $columnCount = [Math]::Max(1, ($firstLine -split ',').Count)
  $fieldInfo = New-Object 'object[,]' $columnCount, 2
  for ($i = 0; $i -lt $columnCount; $i++) {
    $fieldInfo[$i, 0] = $i + 1
    $fieldInfo[$i, 1] = 2
  }

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false

  # OpenText imports every CSV column as text, avoiding Excel scientific notation conversion.
  $excel.Workbooks.OpenText(
    $SourceCsv,
    65001,
    1,
    1,
    1,
    $false,
    $false,
    $false,
    $true,
    $false,
    $false,
    [Type]::Missing,
    $fieldInfo
  )

  $workbook = $excel.ActiveWorkbook
  $sheet = $workbook.Worksheets.Item(1)

  $sheet.Columns.Item(1).Insert() | Out-Null
  $sheet.Columns.Item(1).Insert() | Out-Null
  $sheet.Cells.Item(1, 1).Value2 = "$([char]0x8D26)$([char]0x5355)"
  $sheet.Cells.Item(1, 2).Value2 = "$([char]0x53D1)$([char]0x7968)"

  $used = $sheet.UsedRange
  $rowCount = $used.Rows.Count
  $colCount = $used.Columns.Count
  $totalColumn = $TotalColumn + 2
  $issues = New-Object System.Collections.Generic.List[string]

  for ($r = 1; $r -le $rowCount; $r++) {
    for ($c = 1; $c -le $colCount; $c++) {
      $cell = $sheet.Cells.Item($r, $c)
      $text = [string]$cell.Text
      if ($text -match '^="(.*)"$') {
        $text = $Matches[1].Replace('""', '"')
        $cell.Value2 = $text
      }
      if ($text -match '^[+-]?\d+(\.\d+)?E[+-]?\d+$') {
        $issues.Add("R$r C$c appears as scientific notation: $text")
      }
      if ($r -eq 1) {
        Set-CellFormat $cell '@'
      } elseif ($c -eq $totalColumn) {
        $raw = ([string]$cell.Text).Trim().Replace(',', '')
        if ($raw -match '^-?\d+(\.\d+)?$') {
          $cell.Value2 = [double]$raw
          Set-CellFormat $cell '0.00'
        } else {
          Set-CellFormat $cell '@'
        }
      } elseif ($text -match '^[+-]?\d{12,}$') {
        Set-CellFormat $cell '@'
      } elseif ($text -match '^-?\d+(\.\d+)?$') {
        $cell.Value2 = [double]$text
        Set-CellFormat $cell 'General'
      } else {
        Set-CellFormat $cell '@'
      }
    }
  }

  $mergedGroups = 0
  $row = 2
  while ($row -le $rowCount) {
    $value = ([string]$sheet.Cells.Item($row, $totalColumn).Text).Trim()
    if ([string]::IsNullOrWhiteSpace($value)) {
      $row++
      continue
    }

    $end = $row
    while ($end + 1 -le $rowCount) {
      $nextValue = ([string]$sheet.Cells.Item($end + 1, $totalColumn).Text).Trim()
      if ($nextValue -ne $value) { break }
      $end++
    }

    if ($end -gt $row) {
      $range = $sheet.Range($sheet.Cells.Item($row, $totalColumn), $sheet.Cells.Item($end, $totalColumn))
      $range.Merge()
      $range.HorizontalAlignment = -4108
      $range.VerticalAlignment = -4108
      $mergedGroups++
    }

    $row = $end + 1
  }

  $filterRange = $sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item($rowCount, $colCount))
  $filterRange.AutoFilter() | Out-Null
  $sheet.UsedRange.Columns.AutoFit() | Out-Null

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

  $workbook.SaveAs($OutputXlsx, 51)
  ConvertTo-JsonLine @{
    ok = $true
    sourceFile = $SourceCsv
    outputFile = $OutputXlsx
    mergedGroups = $mergedGroups
    issues = @($issues)
  }
}
catch {
  ConvertTo-JsonLine @{
    ok = $false
    reason = $_.Exception.Message
    sourceFile = $SourceCsv
    outputFile = $OutputXlsx
  }
  exit 1
}
finally {
  if ($workbook -ne $null) {
    try { $workbook.Close($false) | Out-Null } catch {}
    [Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
  }
  if ($excel -ne $null) {
    try { $excel.Quit() | Out-Null } catch {}
    [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
