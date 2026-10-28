param(
  [Parameter(Mandatory=$true)][string]$YhFile,
  [Parameter(Mandatory=$true)][string]$JmFile,
  [Parameter(Mandatory=$true)][string]$JmAiFile,
  [Parameter(Mandatory=$true)][string]$BeijingFile,
  [Parameter(Mandatory=$true)][string]$JdFile,
  [Parameter(Mandatory=$true)][string]$TengyuFile,
  [Parameter(Mandatory=$true)][string]$OutputFile
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

function U([int[]]$codes) {
  -join ($codes | ForEach-Object { [char]$_ })
}

function ConvertTo-JsonLine($obj) {
  $obj | ConvertTo-Json -Depth 8 -Compress
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

function To-Number($value) {
  $s = ([string]$value).Trim().Replace(",", "")
  $s = [regex]::Replace($s, '[^\d\.\-]', '')
  if ($s -match '^-?\d+(\.\d+)?$') { return [double]$s }
  return 0.0
}

function Build-HeaderMap($sheet, [int]$headerRow) {
  $map = @{}
  if (-not $headerRow) { return $map }
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

function Find-HeaderRow($sheet, [string]$header) {
  $needle = Normalize $header
  $rows = [Math]::Min(30, $sheet.UsedRange.Rows.Count)
  $cols = $sheet.UsedRange.Columns.Count
  for ($r = 1; $r -le $rows; $r++) {
    for ($c = 1; $c -le $cols; $c++) {
      if ((Normalize $sheet.Cells.Item($r, $c).Text) -eq $needle) { return $r }
    }
  }
  return $null
}

function Set-Money($cell, [double]$value) {
  $cell.NumberFormat = "0.00"
  $cell.Value = [Math]::Round($value, 2)
}

function Set-Text($cell, [string]$value) {
  $cell.Value2 = $value
}

function Set-Formula($cell, [string]$formula) {
  $cell.Formula = $formula
  $cell.NumberFormat = "0.00"
}

$excel = $null
$books = @()

try {
  foreach ($p in @($YhFile, $JmFile, $JmAiFile, $BeijingFile, $JdFile, $TengyuFile)) {
    if (-not (Test-Path -LiteralPath $p)) { throw "File not found: $p" }
  }

  $outDir = Split-Path -Parent $OutputFile
  if (-not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
  }

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false

  $channelHeader = U @(0x901A,0x9053,0x540D)
  $typeHeader = U @(0x77ED,0x4FE1,0x7C7B,0x578B)
  $signatureHeader = U @(0x7B7E,0x540D)
  $amountHeader = U @(0x91D1,0x989D)
  $jmName = U @(0x6C5F,0x897F,0x805A,0x68A6)
  $jdName = U @(0x4EAC,0x4E1C)
  $bjName = U @(0x5317,0x4EAC,0x51CC,0x6E21)
  $tyName = U @(0x817E,0x57DF)
  $aiShortType = "AI$(U @(0x6302,0x77ED))"
  $aiVoiceType = "AI$(U @(0x8BED,0x97F3))"
  $hmtPattern = "$(U @(0x9999,0x6E2F))|$(U @(0x6FB3,0x95E8))|$(U @(0x53F0,0x6E7E))|$(U @(0x6E2F,0x6FB3,0x53F0))"

  $yhBook = $excel.Workbooks.Open($YhFile)
  $books += $yhBook
  $yhSheet = $yhBook.Worksheets.Item(1)
  $yhHeaders = Build-HeaderMap $yhSheet 1
  $yhChannelCol = Get-Col $yhHeaders @($channelHeader)
  $yhTypeCol = Get-Col $yhHeaders @($typeHeader)
  $yhSignatureCol = Get-Col $yhHeaders @($signatureHeader)
  $yhAmountCol = Get-Col $yhHeaders @($amountHeader)
  if (-not ($yhChannelCol -and $yhTypeCol -and $yhSignatureCol -and $yhAmountCol)) {
    throw "YH workbook is missing required columns."
  }

  $yhJmSms = 0.0
  $yhJmHmt = 0.0
  $yhJmShort = 0.0
  $yhJmAi = 0.0
  $yhJd = 0.0
  $yhBj = 0.0
  $yhTengyu = 0.0

  for ($r = 2; $r -le $yhSheet.UsedRange.Rows.Count; $r++) {
    $channel = ([string]$yhSheet.Cells.Item($r, $yhChannelCol).Text).Trim()
    $type = ([string]$yhSheet.Cells.Item($r, $yhTypeCol).Text).Trim()
    $signature = ([string]$yhSheet.Cells.Item($r, $yhSignatureCol).Text).Trim()
    $amount = To-Number $yhSheet.Cells.Item($r, $yhAmountCol).Text

    if ($channel -eq $jmName) {
      if ($type -eq $aiShortType) { $yhJmShort += $amount }
      elseif ($type -eq $aiVoiceType) { $yhJmAi += $amount }
      elseif ($signature -match $hmtPattern) { $yhJmHmt += $amount }
      else { $yhJmSms += $amount }
    } elseif ($channel -eq $jdName) {
      $yhJd += $amount
    } elseif ($channel -eq $bjName) {
      $yhBj += $amount
    } elseif ($channel -eq $tyName) {
      $yhTengyu += $amount
    }
  }

  $moneyYuanHeader = "$(U @(0x91D1,0x989D))$(U @(0xFF08,0x5143,0xFF09))"
  $totalLabel = U @(0x5408,0x8BA1)
  $smsSheetName = U @(0x77ED,0x4FE1)
  $intlSheetName = U @(0x56FD,0x9645)

  $jmBook = $excel.Workbooks.Open($JmFile)
  $books += $jmBook
  $jmSms = 0.0
  $jmHmt = 0.0
  foreach ($sheet in @($jmBook.Worksheets)) {
    if ($sheet.Name -eq $smsSheetName -or $sheet.Name -eq $intlSheetName) {
      $hr = Find-HeaderRow $sheet $moneyYuanHeader
      $headers = Build-HeaderMap $sheet $hr
      $amountCol = Get-Col $headers @($moneyYuanHeader)
      if ($amountCol) {
        for ($r = $hr + 1; $r -le $sheet.UsedRange.Rows.Count; $r++) {
          if (([string]$sheet.Cells.Item($r, 1).Text).Trim() -eq $totalLabel) {
            if ($sheet.Name -eq $smsSheetName) { $jmSms = To-Number $sheet.Cells.Item($r, $amountCol).Text }
            else { $jmHmt = To-Number $sheet.Cells.Item($r, $amountCol).Text }
            break
          }
        }
      }
    }
  }

  $taskHeader = U @(0x4EFB,0x52A1,0x540D,0x79F0)
  $voiceFeeHeader = "$(U @(0x8BED,0x97F3,0x8D39,0x7528))$(U @(0xFF08,0x5143,0xFF09))"
  $shortFeeHeader = "$(U @(0x6302,0x77ED,0x8D39,0x7528))$(U @(0xFF08,0x5143,0xFF09))"

  $jmAiBook = $excel.Workbooks.Open($JmAiFile)
  $books += $jmAiBook
  $jmTts = 0.0
  $jmShort = 0.0
  $jmVoiceAi = 0.0
  foreach ($sheet in @($jmAiBook.Worksheets)) {
    $hr = Find-HeaderRow $sheet $taskHeader
    if ($hr) {
      $headers = Build-HeaderMap $sheet $hr
      $taskCol = Get-Col $headers @($taskHeader)
      $voiceCol = Get-Col $headers @($voiceFeeHeader, (U @(0x8BED,0x97F3,0x8D39,0x7528)))
      $shortCol = Get-Col $headers @($shortFeeHeader, (U @(0x6302,0x77ED,0x8D39,0x7528)))
      if ($taskCol -and $voiceCol) {
        for ($r = $hr + 1; $r -le $sheet.UsedRange.Rows.Count; $r++) {
          $task = ([string]$sheet.Cells.Item($r, $taskCol).Text).Trim()
          $taskKey = (Normalize $task).ToUpperInvariant()
          if ($taskKey -match 'TTS') {
            $jmTts += To-Number $sheet.Cells.Item($r, $voiceCol).Text
          } elseif ($taskKey -match 'AI') {
            $jmVoiceAi = To-Number $sheet.Cells.Item($r, $voiceCol).Text
            if ($shortCol) { $jmShort = To-Number $sheet.Cells.Item($r, $shortCol).Text }
          }
        }
      }
    }
  }

  $billingAmountHeader = "$(U @(0x8BA1,0x8D39,0x91D1,0x989D))$(U @(0xFF08,0x5143,0xFF09))"
  $subtotalLabel = U @(0x8D39,0x7528,0x5C0F,0x8BA1)
  $bjBook = $excel.Workbooks.Open($BeijingFile)
  $books += $bjBook
  $bjSheet = $bjBook.Worksheets.Item(1)
  $hr = Find-HeaderRow $bjSheet $billingAmountHeader
  $headers = Build-HeaderMap $bjSheet $hr
  $bjAmountCol = Get-Col $headers @($billingAmountHeader)
  $bjAmount = 0.0
  if ($bjAmountCol) {
    for ($r = $hr + 1; $r -le $bjSheet.UsedRange.Rows.Count; $r++) {
      $label = ([string]$bjSheet.Cells.Item($r, 2).Text).Trim()
      if ($label -match $subtotalLabel) {
        $bjAmount = To-Number $bjSheet.Cells.Item($r, $bjAmountCol).Text
        break
      }
    }
  }

  $settlementSheetName = U @(0x7ED3,0x7B97,0x5355)
  $totalAmountHeader = "$(U @(0x603B,0x91D1,0x989D))$(U @(0xFF08,0x5143,0xFF09))"
  $jdBook = $excel.Workbooks.Open($JdFile)
  $books += $jdBook
  $jdAmount = 0.0
  foreach ($sheet in @($jdBook.Worksheets)) {
    if ($sheet.Name -eq $settlementSheetName) {
      $hr = Find-HeaderRow $sheet $totalAmountHeader
      $headers = Build-HeaderMap $sheet $hr
      $totalCol = Get-Col $headers @($totalAmountHeader)
      if ($totalCol) { $jdAmount = To-Number $sheet.Cells.Item($hr + 1, $totalCol).Text }
      break
    }
  }

  $tyBook = $excel.Workbooks.Open($TengyuFile); $books += $tyBook; $tyAmount = 0.0; $tySheet = $tyBook.Worksheets.Item(1)
  $tyUsed = $tySheet.UsedRange
  # 腾域账单金额在第 15 列，取该列最后一个正数（合计行）。
  for ($tr = 1; $tr -le $tyUsed.Rows.Count; $tr++) {
    $candidate = To-Number $tySheet.Cells.Item($tr, 15).Text
    if ($candidate -gt 0) { $tyAmount = $candidate }
  }
  $outBook = $excel.Workbooks.Add()
  $books += $outBook
  $outSheet = $outBook.Worksheets.Item(1)
  $outSheet.Name = U @(0x8D39,0x7528,0x5BF9,0x6BD4)

  $rows = @(
    @((U @(0x5185,0x5BB9,0x7C7B,0x578B)), "YH$(U @(0x5E10,0x5355))", "$jmName$(U @(0x5E10,0x5355))", (U @(0x5DEE,0x5F02)), (U @(0x6700,0x7EC8,0x5BF9,0x8D26,0x91D1,0x989D))),
    @((U @(0x77ED,0x4FE1)), $yhJmSms, $jmSms, $null, $null),
    @((U @(0x8BED,0x97F3)) + "TTS", $null, $jmTts, $null, $null),
    @((U @(0x5F69,0x4FE1)), $null, 0.0, $null, $null),
    @((U @(0x6E2F,0x6FB3,0x53F0)), $yhJmHmt, $jmHmt, $null, $null),
    @((U @(0x5C0F,0x8BA1)) + ":(" + (U @(0x6263,0x6302,0x77ED,0x548C,0x8BED,0x97F3)) + ")", $null, $null, $null, $null),
    @((U @(0x6302,0x77ED)), $null, $jmShort, $null, $null),
    @((U @(0x8BED,0x97F3)) + "AI", $null, $jmVoiceAi, $null, $null),
    @($totalLabel, $null, $null, $null, $null),
    @($jdName, "YH$(U @(0x5E10,0x5355))", "$jdName$(U @(0x5E10,0x5355))", (U @(0x5DEE,0x5F02)), (U @(0x6700,0x7EC8,0x5BF9,0x8D26,0x91D1,0x989D))),
    @("", $yhJd, $jdAmount, $null, $null),
    @($bjName, "YH$(U @(0x5E10,0x5355))", "$bjName$(U @(0x5E10,0x5355))", (U @(0x5DEE,0x5F02)), (U @(0x6700,0x7EC8,0x5BF9,0x8D26,0x91D1,0x989D))),
    @("", $yhBj, $bjAmount, $null, $null),
    @("", $null, $null, $null, $null)
  )

  for ($r = 1; $r -le $rows.Count; $r++) {
    for ($c = 1; $c -le 5; $c++) {
      $v = $rows[$r - 1][$c - 1]
      if ($v -is [double] -or $v -is [int]) { Set-Money $outSheet.Cells.Item($r, $c) ([double]$v) }
      elseif ($v -ne $null) { Set-Text $outSheet.Cells.Item($r, $c) ([string]$v) }
    }
  }

  $finalAmountHeader = U @(0x6700,0x7EC8,0x5BF9,0x8D26,0x91D1,0x989D)
  $diffHeader = U @(0x5DEE,0x5F02)
  $contentTypeHeader = U @(0x5185,0x5BB9,0x7C7B,0x578B)
  $yhBillHeader = "YH$(U @(0x5E10,0x5355))"
  $jmBillHeader = "$jmName$(U @(0x5E10,0x5355))"
  $jdBillHeader = "$jdName$(U @(0x5E10,0x5355))"
  $bjBillHeader = "$bjName$(U @(0x5E10,0x5355))"
  $smsLabel = U @(0x77ED,0x4FE1)
  $ttsLabel = "$(U @(0x8BED,0x97F3))TTS"
  $mmsLabel = U @(0x5F69,0x4FE1)
  $hmtLabel = U @(0x6E2F,0x6FB3,0x53F0)
  $shortLabel = U @(0x6302,0x77ED)
  $voiceAiLabel = "$(U @(0x8BED,0x97F3))AI"
  $subTotalLabel = "$(U @(0x5C0F,0x8BA1)):($(U @(0x6263,0x6302,0x77ED,0x548C,0x8BED,0x97F3)))"

  $outSheet.Range("A1:E16").ClearContents() | Out-Null
  Set-Text $outSheet.Cells.Item(1, 1) $contentTypeHeader
  Set-Text $outSheet.Cells.Item(1, 2) $yhBillHeader
  Set-Text $outSheet.Cells.Item(1, 3) $jmBillHeader
  Set-Text $outSheet.Cells.Item(1, 4) $diffHeader
  Set-Text $outSheet.Cells.Item(1, 5) $finalAmountHeader
  Set-Text $outSheet.Cells.Item(2, 1) $smsLabel
  Set-Money $outSheet.Cells.Item(2, 2) $yhJmSms
  Set-Money $outSheet.Cells.Item(2, 3) $jmSms
  Set-Text $outSheet.Cells.Item(3, 1) $ttsLabel
  Set-Money $outSheet.Cells.Item(3, 3) $jmTts
  Set-Text $outSheet.Cells.Item(4, 1) $mmsLabel
  Set-Money $outSheet.Cells.Item(4, 3) 0.0
  Set-Text $outSheet.Cells.Item(5, 1) $hmtLabel
  Set-Money $outSheet.Cells.Item(5, 2) $yhJmHmt
  Set-Money $outSheet.Cells.Item(5, 3) $jmHmt
  Set-Text $outSheet.Cells.Item(6, 1) $subTotalLabel
  Set-Text $outSheet.Cells.Item(7, 1) $shortLabel
  Set-Money $outSheet.Cells.Item(7, 3) $jmShort
  Set-Text $outSheet.Cells.Item(8, 1) $voiceAiLabel
  Set-Money $outSheet.Cells.Item(8, 3) $jmVoiceAi
  Set-Text $outSheet.Cells.Item(9, 1) $totalLabel
  Set-Text $outSheet.Cells.Item(10, 1) $jdName
  Set-Text $outSheet.Cells.Item(10, 2) $yhBillHeader
  Set-Text $outSheet.Cells.Item(10, 3) $jdBillHeader
  Set-Text $outSheet.Cells.Item(10, 4) $diffHeader
  Set-Text $outSheet.Cells.Item(10, 5) $finalAmountHeader
  Set-Money $outSheet.Cells.Item(11, 2) $yhJd
  Set-Money $outSheet.Cells.Item(11, 3) $jdAmount
  Set-Text $outSheet.Cells.Item(12, 1) $bjName
  Set-Text $outSheet.Cells.Item(12, 2) $yhBillHeader
  Set-Text $outSheet.Cells.Item(12, 3) $bjBillHeader
  Set-Text $outSheet.Cells.Item(12, 4) $diffHeader
  Set-Text $outSheet.Cells.Item(12, 5) $finalAmountHeader
   Set-Money $outSheet.Cells.Item(13, 2) $yhBj
   Set-Money $outSheet.Cells.Item(13, 3) $bjAmount

  # 鑵惧煙璐﹀崟锛堝崟鐙垪鍑猴紝渚涜垂鐢ㄥ姣斿拰缇よ瘽鏈鍙栵級
  Set-Text $outSheet.Cells.Item(14, 1) $tyName
  Set-Text $outSheet.Cells.Item(14, 2) $yhBillHeader
  Set-Text $outSheet.Cells.Item(14, 3) "$tyName$(U @(0x5E10,0x5355))"
  Set-Text $outSheet.Cells.Item(14, 4) $diffHeader
  Set-Text $outSheet.Cells.Item(14, 5) $finalAmountHeader
  Set-Money $outSheet.Cells.Item(15, 2) $yhTengyu
  Set-Money $outSheet.Cells.Item(15, 3) $tyAmount

  Set-Formula $outSheet.Cells.Item(2, 4) "=B2+B5-C2-C3-C4-C5"
  Set-Formula $outSheet.Cells.Item(6, 2) "=B2+B5"
  Set-Formula $outSheet.Cells.Item(6, 3) "=C2+C3+C4+C5"
  Set-Formula $outSheet.Cells.Item(6, 4) "=B6-C6"
  Set-Formula $outSheet.Cells.Item(7, 2) "=C7"
  Set-Formula $outSheet.Cells.Item(8, 2) "=C8"
  Set-Formula $outSheet.Cells.Item(9, 2) "=B2+B5+C7+C8"
  Set-Formula $outSheet.Cells.Item(9, 3) "=C2+C3+C4+C5+C7+C8"
  Set-Formula $outSheet.Cells.Item(9, 4) "=B9-C9"
  Set-Formula $outSheet.Cells.Item(11, 4) "=B11-C11"
  Set-Formula $outSheet.Cells.Item(13, 4) "=B13-C13"
  Set-Formula $outSheet.Cells.Item(15, 4) "=B15-C15"
  Set-Formula $outSheet.Cells.Item(16, 2) "=SUM(B9,B11,B13,B15)"
  Set-Formula $outSheet.Cells.Item(16, 3) "=SUM(C9,C11,C13,C15)"
  Set-Formula $outSheet.Cells.Item(16, 4) "=SUM(D9,D11,D13,D15)"
  Set-Formula $outSheet.Cells.Item(16, 5) "=SUM(E9,E11,E13,E15)"

  $outSheet.Range("B2:B4").Merge() | Out-Null
  $outSheet.Range("B2:B4").VerticalAlignment = -4108
  $outSheet.Range("D2:D5").Merge() | Out-Null
  $outSheet.Range("E2:E5").Merge() | Out-Null
  $outSheet.Range("A10:A11").Merge() | Out-Null
  $outSheet.Range("A12:A13").Merge() | Out-Null
  $outSheet.Range("A14:A15").Merge() | Out-Null
  $outSheet.Range("A1:E16").Borders.LineStyle = 1
  $outSheet.Range("A1:E16").HorizontalAlignment = -4108
  $outSheet.Range("A1:E16").VerticalAlignment = -4108
  $outSheet.Range("A1:E1").Font.Bold = $true
  $outSheet.Range("A10:E10").Font.Bold = $true
  $outSheet.Range("A12:E12").Font.Bold = $true
  $outSheet.Range("A11:E11").Font.Bold = $true
  $outSheet.Range("A13:E13").Font.Bold = $true
  $outSheet.Range("A14:E14").Font.Bold = $true
  $outSheet.Range("A15:E15").Font.Bold = $true
  $outSheet.Range("A9:E9").Font.Bold = $true
  $outSheet.Range("A16:E16").Font.Bold = $false
  $outSheet.Range("A1:E1").Interior.Color = 15189684
  $outSheet.Range("A10:E10").Interior.Color = 13434879
  $outSheet.Range("A12:E12").Interior.Color = 13434828
  $outSheet.Range("A2:A9").Interior.Color = 15189684
  $outSheet.Range("A11:E11").Interior.Color = 15921906
  $outSheet.Range("A13:E13").Interior.Color = 15138810
  $outSheet.Range("A14:E14").Interior.Color = 14277081
  $outSheet.Range("A15:E15").Interior.Color = 14277081
  $outSheet.Range("A16:E16").Interior.Color = 14277081
  $outSheet.Columns.Item(1).ColumnWidth = 22
  $outSheet.Columns.Item(2).ColumnWidth = 16
  $outSheet.Columns.Item(3).ColumnWidth = 18
  $outSheet.Columns.Item(4).ColumnWidth = 14
  $outSheet.Columns.Item(5).ColumnWidth = 18
  $outSheet.Rows.Item(1).RowHeight = 24
  $outSheet.Range("A2:A9").EntireRow.RowHeight = 30
  $outSheet.Rows.Item(10).RowHeight = 24
  $outSheet.Rows.Item(11).RowHeight = 30
  $outSheet.Rows.Item(12).RowHeight = 24
  $outSheet.Rows.Item(13).RowHeight = 30
  $outSheet.Rows.Item(14).RowHeight = 24
  $outSheet.Rows.Item(15).RowHeight = 30
  $outSheet.Rows.Item(16).RowHeight = 30

  if (Test-Path -LiteralPath $OutputFile) {
    Remove-Item -LiteralPath $OutputFile -Force
  }
  $outBook.SaveAs($OutputFile, 51)

  ConvertTo-JsonLine @{
    ok = $true
    outputFile = $OutputFile
    yhJmSms = [Math]::Round($yhJmSms, 2)
    yhJmHmt = [Math]::Round($yhJmHmt, 2)
    yhJmShort = [Math]::Round($yhJmShort, 2)
    yhJmAi = [Math]::Round($yhJmAi, 2)
    jmSms = [Math]::Round($jmSms, 2)
    jmHmt = [Math]::Round($jmHmt, 2)
    jmTts = [Math]::Round($jmTts, 2)
    jmShort = [Math]::Round($jmShort, 2)
    jmVoiceAi = [Math]::Round($jmVoiceAi, 2)
    yhJd = [Math]::Round($yhJd, 2)
    jd = [Math]::Round($jdAmount, 2)
    tengyu = [Math]::Round($tyAmount, 2)
    yhBeijing = [Math]::Round($yhBj, 2)
    yhTengyu = [Math]::Round($yhTengyu, 2)
    beijing = [Math]::Round($bjAmount, 2)
  }
}
catch {
  ConvertTo-JsonLine @{ ok = $false; reason = $_.Exception.Message; outputFile = $OutputFile }
  exit 1
}
finally {
  for ($i = $books.Count - 1; $i -ge 0; $i--) {
    try { $books[$i].Close($false) | Out-Null } catch {}
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($books[$i]) | Out-Null } catch {}
  }
  if ($excel) {
    try { $excel.Quit() | Out-Null } catch {}
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
