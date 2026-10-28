param(
  [Parameter(Mandatory=$true)][string]$HmtFile,
  [Parameter(Mandatory=$true)][string]$TargetXlsx
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

function U([int[]]$codes) {
  -join ($codes | ForEach-Object { [char]$_ })
}

function ConvertTo-JsonLine($obj) {
  $obj | ConvertTo-Json -Depth 6 -Compress
}

function Normalize($value) {
  $s = ([string]$value).Trim()
  $s = $s.Replace(" ", "")
  $s = $s.Replace((U @(0x3000)), "")
  return $s
}

function Build-HeaderMap($sheet) {
  $map = @{}
  $colCount = $sheet.UsedRange.Columns.Count
  for ($c = 1; $c -le $colCount; $c++) {
    $name = Normalize $sheet.Cells.Item(1, $c).Text
    if ($name -and -not $map.ContainsKey($name)) {
      $map[$name] = $c
    }
  }
  return $map
}

function Get-Col($map, [string[]]$names) {
  foreach ($name in $names) {
    $key = Normalize $name
    if ($map.ContainsKey($key)) { return $map[$key] }
  }
  return $null
}

function Set-TextCell($cell, $value) {
  try { $cell.NumberFormat = "@" } catch {}
  $cell.Value2 = [string]$value
}

function Set-AmountCell($cell, $value) {
  $numericText = ([string]$value).Trim().Replace(",", "")
  $cell.NumberFormat = "0.00"
  if ($numericText -match '^-?\d+(\.\d+)?$') {
    $cell.Value = [double]$numericText
    return [double]$numericText
  }
  $cell.Value2 = [string]$value
  return $null
}

function Remove-ExistingHmtRows($sheet, $headers, $jiangxiJumeng, $gang, $ao, $tai, $channelHeader, $supplierHeader, $signatureHeader) {
  $channelCol = Get-Col $headers @($channelHeader, $supplierHeader)
  $signatureCol = Get-Col $headers @($signatureHeader)
  if (-not $signatureCol) { return 0 }

  $removed = 0
  $lastRow = $sheet.UsedRange.Rows.Count
  for ($r = $lastRow; $r -ge 2; $r--) {
    $signature = ([string]$sheet.Cells.Item($r, $signatureCol).Text).Trim()
    $isHmt = $signature.Contains($gang) -or $signature.Contains($ao) -or $signature.Contains($tai)
    if (-not $isHmt) { continue }

    $isJumeng = $true
    if ($channelCol) {
      $channel = ([string]$sheet.Cells.Item($r, $channelCol).Text).Trim()
      $isJumeng = ($channel -eq $jiangxiJumeng)
    }

    if ($isJumeng) {
      $sheet.Rows.Item($r).Delete() | Out-Null
      $removed++
    }
  }
  return $removed
}

function Merge-TotalGroups($sheet, $totalCol, $startRow, $endRow) {
  if (-not $totalCol -or $startRow -gt $endRow) { return 0 }
  $merged = 0
  $row = $startRow
  while ($row -le $endRow) {
    $value = ([string]$sheet.Cells.Item($row, $totalCol).Text).Trim()
    if ([string]::IsNullOrWhiteSpace($value)) {
      $row++
      continue
    }

    $groupEnd = $row
    while ($groupEnd + 1 -le $endRow) {
      $nextValue = ([string]$sheet.Cells.Item($groupEnd + 1, $totalCol).Text).Trim()
      if ($nextValue -ne $value) { break }
      $groupEnd++
    }

    if ($groupEnd -gt $row) {
      $range = $sheet.Range($sheet.Cells.Item($row, $totalCol), $sheet.Cells.Item($groupEnd, $totalCol))
      $range.Merge()
      $range.HorizontalAlignment = -4108
      $range.VerticalAlignment = -4108
      $merged++
    }
    $row = $groupEnd + 1
  }
  return $merged
}

function Get-Col($map, [string[]]$names) {
  foreach ($name in $names) {
    $key = Normalize $name
    if ($map.ContainsKey($key)) { return $map[$key] }
  }
  return $null
}

function Highlight-KeyCells($sheet, $headers, $startRow, $endRow, $color) {
  if ($startRow -gt $endRow) { return }
  $cols = @(
    (Get-Col $headers @((U @(0x901A,0x9053,0x540D)))),
    (Get-Col $headers @((U @(0x8D26,0x53F7,0x7F16,0x7801)))),
    (Get-Col $headers @((U @(0x8D26,0x53F7,0x540D)))),
    (Get-Col $headers @((U @(0x77ED,0x4FE1,0x7C7B,0x578B)))),
    (Get-Col $headers @((U @(0x7B7E,0x540D)))),
    (Get-Col $headers @((U @(0x5408,0x8BA1))))
  ) | Where-Object { $_ }

  foreach ($col in $cols) {
    $range = $sheet.Range($sheet.Cells.Item($startRow, $col), $sheet.Cells.Item($endRow, $col))
    $range.Interior.Color = $color
  }
}

$excel = $null
$hmtWorkbook = $null
$targetWorkbook = $null

try {
  if (-not (Test-Path -LiteralPath $HmtFile)) {
    throw "HMT bill file not found: $HmtFile"
  }
  if (-not (Test-Path -LiteralPath $TargetXlsx)) {
    throw "Organized workbook not found: $TargetXlsx"
  }

  $billHeader = U @(0x8D26,0x5355)
  $hmtLabel = U @(0x6E2F,0x6FB3,0x53F0,0x8D26,0x5355)
  $jiangxiJumeng = U @(0x6C5F,0x897F,0x805A,0x68A6)
  $signatureHeader = U @(0x7B7E,0x540D)
  $channelHeader = U @(0x901A,0x9053,0x540D)
  $supplierHeader = U @(0x4F9B,0x5E94,0x5546)
  $amountHeader = U @(0x91D1,0x989D)
  $totalHeader = U @(0x5408,0x8BA1)
  $contactHeader = U @(0x5BF9,0x63A5,0x4EBA)
  $guJiawen = U @(0x987E,0x4F73,0x96EF)
  $gang = U @(0x6E2F)
  $ao = U @(0x6FB3)
  $tai = U @(0x53F0)

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false

  $hmtWorkbook = $excel.Workbooks.Open($HmtFile)
  $targetWorkbook = $excel.Workbooks.Open($TargetXlsx)
  $dst = $targetWorkbook.Worksheets.Item(1)
  $dstHeaders = Build-HeaderMap $dst
  $removedRows = Remove-ExistingHmtRows $dst $dstHeaders $jiangxiJumeng $gang $ao $tai $channelHeader $supplierHeader $signatureHeader
  $dstHeaders = Build-HeaderMap $dst
  $dstLastRow = $dst.UsedRange.Rows.Count
  $appended = 0
  $scannedRows = 0
  $firstAppendRow = $null
  $lastAppendRow = $null
  $appendedAmountSum = 0.0
  $appendedRowsWithTotal = 0
  $amountDstCol = Get-Col $dstHeaders @($amountHeader)
  $totalDstCol = Get-Col $dstHeaders @($totalHeader)
  $contactDstCol = Get-Col $dstHeaders @($contactHeader)

  foreach ($src in @($hmtWorkbook.Worksheets)) {
    $srcHeaders = Build-HeaderMap $src
    $signatureCol = Get-Col $srcHeaders @($signatureHeader)
    if (-not $signatureCol) { continue }

    $channelCol = Get-Col $srcHeaders @($channelHeader, $supplierHeader)
    $sheetName = [string]$src.Name
    $sheetIsJumeng = $sheetName.Contains($jiangxiJumeng)
    $srcRows = $src.UsedRange.Rows.Count

    for ($r = 2; $r -le $srcRows; $r++) {
      $scannedRows++
      $signature = ([string]$src.Cells.Item($r, $signatureCol).Text).Trim()
      $isHmt = $signature.Contains($gang) -or $signature.Contains($ao) -or $signature.Contains($tai)
      if (-not $isHmt) { continue }

      if ($channelCol) {
        $channel = ([string]$src.Cells.Item($r, $channelCol).Text).Trim()
        $rowIsJumeng = ($channel -eq $jiangxiJumeng)
      } else {
        $rowIsJumeng = $sheetIsJumeng
      }
      if (-not $rowIsJumeng) { continue }

      $dstLastRow++
      $appended++
      if ($firstAppendRow -eq $null) { $firstAppendRow = $dstLastRow }
      $lastAppendRow = $dstLastRow

      foreach ($targetName in $dstHeaders.Keys) {
        $targetCol = $dstHeaders[$targetName]
        if ($targetCol -le 2) { continue }

        $sourceCol = $null
        if ($srcHeaders.ContainsKey($targetName)) {
          $sourceCol = $srcHeaders[$targetName]
        } elseif ($targetName -eq (Normalize $supplierHeader) -and $srcHeaders.ContainsKey((Normalize $channelHeader))) {
          $sourceCol = $srcHeaders[(Normalize $channelHeader)]
        } elseif ($targetName -eq (Normalize $channelHeader) -and $srcHeaders.ContainsKey((Normalize $supplierHeader))) {
          $sourceCol = $srcHeaders[(Normalize $supplierHeader)]
        }

        if ($sourceCol) {
          $value = $src.Cells.Item($r, $sourceCol).Text
          if ($targetName -eq (Normalize $amountHeader) -or $targetName -eq (Normalize $totalHeader)) {
            $number = Set-AmountCell $dst.Cells.Item($dstLastRow, $targetCol) $value
            if ($targetName -eq (Normalize $amountHeader) -and $number -ne $null) {
              $appendedAmountSum += [double]$number
            }
            if ($targetName -eq (Normalize $totalHeader) -and $number -ne $null) {
              $appendedRowsWithTotal++
            }
          } else {
            Set-TextCell $dst.Cells.Item($dstLastRow, $targetCol) $value
          }
        }
      }
      if ($contactDstCol) {
        Set-TextCell $dst.Cells.Item($dstLastRow, $contactDstCol) $guJiawen
      }
    }
  }

  if ($appended -gt 0) {
    if ($totalDstCol -and $appendedRowsWithTotal -eq 0) {
      for ($r = $firstAppendRow; $r -le $lastAppendRow; $r++) {
        Set-AmountCell $dst.Cells.Item($r, $totalDstCol) $appendedAmountSum | Out-Null
      }
    }
    $mergedGroups = Merge-TotalGroups $dst $totalDstCol $firstAppendRow $lastAppendRow
    Highlight-KeyCells $dst $dstHeaders $firstAppendRow $lastAppendRow 10092543
    $dst.UsedRange.Columns.AutoFit() | Out-Null
    $targetWorkbook.Save()
  } else {
    $mergedGroups = 0
  }

  ConvertTo-JsonLine @{
    ok = $true
    hmtFile = $HmtFile
    targetFile = $TargetXlsx
    scannedRows = $scannedRows
    removedRows = $removedRows
    appendedRows = $appended
    mergedGroups = $mergedGroups
  }
}
catch {
  ConvertTo-JsonLine @{
    ok = $false
    reason = $_.Exception.Message
    hmtFile = $HmtFile
    targetFile = $TargetXlsx
  }
  exit 1
}
finally {
  if ($hmtWorkbook -ne $null) {
    try { $hmtWorkbook.Close($false) | Out-Null } catch {}
    [Runtime.InteropServices.Marshal]::ReleaseComObject($hmtWorkbook) | Out-Null
  }
  if ($targetWorkbook -ne $null) {
    try { $targetWorkbook.Close($true) | Out-Null } catch {}
    [Runtime.InteropServices.Marshal]::ReleaseComObject($targetWorkbook) | Out-Null
  }
  if ($excel -ne $null) {
    try { $excel.Quit() | Out-Null } catch {}
    [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
