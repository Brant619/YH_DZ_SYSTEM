param(
  [Parameter(Mandatory=$true)][string]$AiShortFile,
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
  foreach ($ch in @("(", ")", (U @(0xFF08)), (U @(0xFF09)))) {
    $s = $s.Replace($ch, "")
  }
  return $s
}

function Build-HeaderMap($sheet, [int]$headerRow) {
  $map = @{}
  $colCount = $sheet.UsedRange.Columns.Count
  for ($c = 1; $c -le $colCount; $c++) {
    $name = Normalize $sheet.Cells.Item($headerRow, $c).Text
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

function Find-HeaderRow($sheet, [string]$taskHeader) {
  $rows = [Math]::Min(20, $sheet.UsedRange.Rows.Count)
  $cols = $sheet.UsedRange.Columns.Count
  for ($r = 1; $r -le $rows; $r++) {
    for ($c = 1; $c -le $cols; $c++) {
      if ((Normalize $sheet.Cells.Item($r, $c).Text) -eq (Normalize $taskHeader)) {
        return $r
      }
    }
  }
  return $null
}

function Remove-ExistingRows($sheet, $headers, [string]$billLabel, [string]$taskName, [string]$accountCode, [string]$shortType, [string]$voiceType) {
  $billCol = Get-Col $headers @((U @(0x8D26,0x5355)))
  $channelCol = Get-Col $headers @((U @(0x901A,0x9053,0x540D)), (U @(0x8D26,0x53F7,0x540D)))
  $accountCodeCol = Get-Col $headers @((U @(0x8D26,0x53F7,0x7F16,0x7801)))
  $typeCol = Get-Col $headers @((U @(0x77ED,0x4FE1,0x7C7B,0x578B)))
  if (-not $billCol -and -not $channelCol -and -not ($accountCodeCol -and $typeCol)) { return 0 }

  $removed = 0
  $lastRow = $sheet.UsedRange.Rows.Count
  for ($r = $lastRow; $r -ge 2; $r--) {
    $bill = if ($billCol) { ([string]$sheet.Cells.Item($r, $billCol).Text).Trim() } else { "" }
    $channel = if ($channelCol) { ([string]$sheet.Cells.Item($r, $channelCol).Text).Trim() } else { "" }
    $code = if ($accountCodeCol) { ([string]$sheet.Cells.Item($r, $accountCodeCol).Text).Trim() } else { "" }
    $type = if ($typeCol) { ([string]$sheet.Cells.Item($r, $typeCol).Text).Trim() } else { "" }
    $isNewAiRow = ($code -eq $accountCode -and ($type -eq $shortType -or $type -eq $voiceType))
    if ($bill -eq $billLabel -or $channel -eq $taskName -or $isNewAiRow) {
      $sheet.Rows.Item($r).Delete() | Out-Null
      $removed++
    }
  }
  return $removed
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
$sourceWorkbook = $null
$targetWorkbook = $null

try {
  if (-not (Test-Path -LiteralPath $AiShortFile)) {
    throw "AI short bill file not found: $AiShortFile"
  }
  if (-not (Test-Path -LiteralPath $TargetXlsx)) {
    throw "Organized workbook not found: $TargetXlsx"
  }

  $billLabel = "AI$(U @(0x6302,0x77ED,0x8D26,0x5355))"
  $taskName = "$(U @(0x6C38,0x8F89,0x8BED,0x97F3))AI"
  $ttsName = "$(U @(0x6C38,0x8F89,0x8BED,0x97F3))TTS"
  $taskHeader = U @(0x4EFB,0x52A1,0x540D,0x79F0)
  $voiceLabel = U @(0x8BED,0x97F3,0x8D39,0x7528)
  $shortLabel = U @(0x6302,0x77ED,0x8D39,0x7528)
  $voiceHeader = "$voiceLabel$(U @(0x5143))"
  $shortHeader = "$shortLabel$(U @(0x5143))"
  $totalLabel = U @(0x8D39,0x7528,0x603B,0x8BA1)
  $totalHeader = "$totalLabel$(U @(0x5143))"
  $durationHeader = "$(U @(0x8BA1,0x8D39,0x65F6,0x957F))$(U @(0x5206,0x949F))"
  $smsCountHeader = U @(0x77ED,0x4FE1,0x8BA1,0x8D39,0x6761,0x6570)

  $fixedChannel = U @(0x6C5F,0x897F,0x805A,0x68A6)
  $fixedAccountCode = "s_ywtj"
  $fixedAccountName = U @(0x4E0A,0x6D77,0x6C38,0x8F89)
  $fixedContact = U @(0x80E1,0x653F)
  $fixedShortType = "AI$(U @(0x6302,0x77ED))"
  $fixedVoiceType = "AI$(U @(0x8BED,0x97F3))"
  $fixedShortSignature = U @(0x6302,0x77ED)
  $fixedVoiceSignature = "$(U @(0x6C38,0x8F89,0x4F1A,0x5458,0x8425,0x9500))AI$(U @(0x4EA4,0x4E92))"
  $fixedShortUnit = "0.03"
  $fixedVoiceUnit = "0.1$(U @(0x5143))/60$(U @(0x79D2))"
  $fixedInvoiceTitle = U @(0x798F,0x5EFA,0x6C38,0x8F89,0x4E91,0x521B,0x79D1,0x6280,0x6709,0x9650,0x516C,0x53F8)
  $fixedTaxId = "91350100MA2XTAHM40"

  $dstBillHeader = U @(0x8D26,0x5355)
  $dstMonthHeader = U @(0x6708,0x4EFD)
  $dstChannelHeader = U @(0x901A,0x9053,0x540D)
  $dstAccountCodeHeader = U @(0x8D26,0x53F7,0x7F16,0x7801)
  $dstAccountNameHeader = U @(0x8D26,0x53F7,0x540D)
  $dstTypeHeader = U @(0x77ED,0x4FE1,0x7C7B,0x578B)
  $dstSignatureHeader = U @(0x7B7E,0x540D)
  $dstSuccessHeader = U @(0x6210,0x529F,0x6761,0x6570)
  $dstContactHeader = U @(0x5BF9,0x63A5,0x4EBA)
  $dstUnitHeader = U @(0x5355,0x4EF7)
  $dstAmountHeader = U @(0x91D1,0x989D)
  $dstTotalHeader = U @(0x5408,0x8BA1)
  $dstInvoiceTitleHeader = U @(0x5F00,0x7968,0x62AC,0x5934)
  $dstTaxIdHeader = U @(0x7EB3,0x7A0E,0x4EBA,0x6807,0x8BC6,0x53F7)

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false

  $sourceWorkbook = $excel.Workbooks.Open($AiShortFile)
  $targetWorkbook = $excel.Workbooks.Open($TargetXlsx)
  $dst = $targetWorkbook.Worksheets.Item(1)
  $dstHeaders = Build-HeaderMap $dst 1
  $removedRows = Remove-ExistingRows $dst $dstHeaders $billLabel $taskName $fixedAccountCode $fixedShortType $fixedVoiceType
  $dstHeaders = Build-HeaderMap $dst 1

  $matchedSheet = $null
  $matchedRow = $null
  $voiceFee = $null
  $shortFee = $null
  $voiceCount = $null
  $shortCount = $null
  $sourceTotal = $null
  $scannedRows = 0

  foreach ($src in @($sourceWorkbook.Worksheets)) {
    $headerRow = Find-HeaderRow $src $taskHeader
    if (-not $headerRow) { continue }

    $srcHeaders = Build-HeaderMap $src $headerRow
    $taskCol = Get-Col $srcHeaders @($taskHeader)
    $voiceCol = Get-Col $srcHeaders @($voiceHeader, $voiceLabel)
    $shortCol = Get-Col $srcHeaders @($shortHeader, $shortLabel)
    $totalCol = Get-Col $srcHeaders @($totalHeader, $totalLabel)
    $durationCol = Get-Col $srcHeaders @($durationHeader, (U @(0x8BA1,0x8D39,0x65F6,0x957F)))
    $smsCountCol = Get-Col $srcHeaders @($smsCountHeader)
    if (-not $taskCol -or -not $voiceCol -or -not $shortCol) { continue }

    $lastRow = $src.UsedRange.Rows.Count
    for ($r = $headerRow + 1; $r -le $lastRow; $r++) {
      $scannedRows++
      $name = ([string]$src.Cells.Item($r, $taskCol).Text).Trim()
      if ($name -eq $ttsName) { continue }
      if ($name -ne $taskName) { continue }

      $voiceFee = ([string]$src.Cells.Item($r, $voiceCol).Text).Trim()
      $shortFee = ([string]$src.Cells.Item($r, $shortCol).Text).Trim()
      if ($durationCol) {
        $voiceCount = ([string]$src.Cells.Item($r, $durationCol).Text).Trim()
      }
      if ($smsCountCol) {
        $shortCount = ([string]$src.Cells.Item($r, $smsCountCol).Text).Trim()
      }
      if ($totalCol) {
        $sourceTotal = ([string]$src.Cells.Item($r, $totalCol).Text).Trim()
      }
      $matchedSheet = [string]$src.Name
      $matchedRow = $r
      break
    }
    if ($matchedRow) { break }
  }

  if (-not $matchedRow) {
    throw "AI short bill row not found: Yonghui voice AI."
  }
  if ([string]::IsNullOrWhiteSpace($voiceFee) -or [string]::IsNullOrWhiteSpace($shortFee)) {
    throw "Yonghui voice AI row is missing voice fee or short-call fee."
  }

  $voiceNumber = ([string]$voiceFee).Replace(",", "")
  $shortNumber = ([string]$shortFee).Replace(",", "")
  if ($voiceNumber -notmatch '^-?\d+(\.\d+)?$' -or $shortNumber -notmatch '^-?\d+(\.\d+)?$') {
    throw "Voice fee or short-call fee is not a valid number."
  }

  $total = [double]$voiceNumber + [double]$shortNumber
  if (-not [string]::IsNullOrWhiteSpace($sourceTotal)) {
    $sourceTotalNumber = ([string]$sourceTotal).Replace(",", "")
    if ($sourceTotalNumber -match '^-?\d+(\.\d+)?$') {
      $total = [double]$sourceTotalNumber
    }
  }

  $monthCol = Get-Col $dstHeaders @($dstMonthHeader)
  $channelCol = Get-Col $dstHeaders @($dstChannelHeader)
  $accountCodeCol = Get-Col $dstHeaders @($dstAccountCodeHeader)
  $accountNameCol = Get-Col $dstHeaders @($dstAccountNameHeader)
  $typeCol = Get-Col $dstHeaders @($dstTypeHeader)
  $signatureCol = Get-Col $dstHeaders @($dstSignatureHeader)
  $successCol = Get-Col $dstHeaders @($dstSuccessHeader)
  $contactCol = Get-Col $dstHeaders @($dstContactHeader)
  $unitCol = Get-Col $dstHeaders @($dstUnitHeader)
  $amountCol = Get-Col $dstHeaders @($dstAmountHeader)
  $totalColDst = Get-Col $dstHeaders @($dstTotalHeader)
  $invoiceTitleCol = Get-Col $dstHeaders @($dstInvoiceTitleHeader)
  $taxIdCol = Get-Col $dstHeaders @($dstTaxIdHeader)
  if (-not $amountCol -or -not $totalColDst) {
    throw "Organized workbook is missing amount or total column."
  }

  $dstLastRow = $dst.UsedRange.Rows.Count
  $rows = @(
    @{ type = $fixedShortType; signature = $fixedShortSignature; count = $shortCount; unit = $fixedShortUnit; amount = $shortFee },
    @{ type = $fixedVoiceType; signature = $fixedVoiceSignature; count = $voiceCount; unit = $fixedVoiceUnit; amount = $voiceFee }
  )
  $firstAppendRow = $dstLastRow + 1
  $appended = 0

  foreach ($item in $rows) {
    $dstLastRow++
    $appended++
    if ($monthCol) { Set-TextCell $dst.Cells.Item($dstLastRow, $monthCol) ([string]$dst.Cells.Item(2, $monthCol).Text) }
    if ($channelCol) { Set-TextCell $dst.Cells.Item($dstLastRow, $channelCol) $fixedChannel }
    if ($accountCodeCol) { Set-TextCell $dst.Cells.Item($dstLastRow, $accountCodeCol) $fixedAccountCode }
    if ($accountNameCol) { Set-TextCell $dst.Cells.Item($dstLastRow, $accountNameCol) $fixedAccountName }
    if ($typeCol) { Set-TextCell $dst.Cells.Item($dstLastRow, $typeCol) $item.type }
    if ($signatureCol) { Set-TextCell $dst.Cells.Item($dstLastRow, $signatureCol) $item.signature }
    if ($successCol -and -not [string]::IsNullOrWhiteSpace($item.count)) { Set-AmountCell $dst.Cells.Item($dstLastRow, $successCol) $item.count | Out-Null }
    if ($contactCol) { Set-TextCell $dst.Cells.Item($dstLastRow, $contactCol) $fixedContact }
    if ($unitCol) { Set-TextCell $dst.Cells.Item($dstLastRow, $unitCol) $item.unit }
    Set-AmountCell $dst.Cells.Item($dstLastRow, $amountCol) $item.amount | Out-Null
    Set-AmountCell $dst.Cells.Item($dstLastRow, $totalColDst) $total | Out-Null
    if ($invoiceTitleCol) { Set-TextCell $dst.Cells.Item($dstLastRow, $invoiceTitleCol) $fixedInvoiceTitle }
    if ($taxIdCol) { Set-TextCell $dst.Cells.Item($dstLastRow, $taxIdCol) $fixedTaxId }
  }

  $lastAppendRow = $dstLastRow
  if ($lastAppendRow -gt $firstAppendRow) {
    $range = $dst.Range($dst.Cells.Item($firstAppendRow, $totalColDst), $dst.Cells.Item($lastAppendRow, $totalColDst))
    $range.Merge()
    $range.HorizontalAlignment = -4108
    $range.VerticalAlignment = -4108
  }

  Highlight-KeyCells $dst $dstHeaders $firstAppendRow $lastAppendRow 15128749
  $dst.UsedRange.Columns.AutoFit() | Out-Null
  $targetWorkbook.Save()

  ConvertTo-JsonLine @{
    ok = $true
    aiShortFile = $AiShortFile
    targetFile = $TargetXlsx
    matchedSheet = $matchedSheet
    matchedRow = $matchedRow
    scannedRows = $scannedRows
    removedRows = $removedRows
    appendedRows = $appended
    voiceFee = $voiceFee
    shortFee = $shortFee
    voiceCount = $voiceCount
    shortCount = $shortCount
    totalFee = "{0:F2}" -f $total
  }
}
catch {
  ConvertTo-JsonLine @{
    ok = $false
    reason = $_.Exception.Message
    aiShortFile = $AiShortFile
    targetFile = $TargetXlsx
  }
  exit 1
}
finally {
  if ($sourceWorkbook -ne $null) {
    try { $sourceWorkbook.Close($false) | Out-Null } catch {}
    [Runtime.InteropServices.Marshal]::ReleaseComObject($sourceWorkbook) | Out-Null
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
