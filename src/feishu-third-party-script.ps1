param(
  [Parameter(Mandatory=$true)][string]$InputFile,
  [Parameter(Mandatory=$true)][string]$OutputFile,
  [Parameter(Mandatory=$true)][int]$Year,
  [Parameter(Mandatory=$true)][int]$Month
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

function To-Number($value) {
  if ($null -eq $value -or $value -eq '') { return 0.0 }
  return [double]$value
}

function Money([double]$value) {
  return $value.ToString('0.00', [Globalization.CultureInfo]::InvariantCulture)
}

function Difference-Message([string]$internalLabel, [double]$internal, [string]$providerLabel, [double]$provider, [double]$difference) {
  if ([Math]::Abs($difference) -lt 0.005) {
    return "$internalLabel $(Money $internal) 元，$providerLabel $(Money $provider) 元，金额一致。本月按照永辉结算的金额 $(Money $internal) 元进行结算，请确认"
  }
  return "$internalLabel $(Money $internal) 元，$providerLabel $(Money $provider) 元，差额 $(Money $difference) 元，是否可接受呢？"
}

$excel = $null
$book = $null
$sheet = $null

try {
  if (-not (Test-Path -LiteralPath $InputFile)) { throw "未找到费用对比表：$InputFile" }

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $book = $excel.Workbooks.Open($InputFile, 0, $true)
  $sheet = $book.Worksheets.Item(1)

  # 费用对比表固定布局：江西聚梦在 6-9 行，京东在 11 行，北京凌渡在 13 行，腾域在 15 行。
  $jmInternal = To-Number $sheet.Cells.Item(6, 2).Value2
  $jmProvider = To-Number $sheet.Cells.Item(6, 3).Value2
  $jmDifference = To-Number $sheet.Cells.Item(6, 4).Value2
  $jmShort = To-Number $sheet.Cells.Item(7, 3).Value2
  $jmAi = To-Number $sheet.Cells.Item(8, 3).Value2
  $jdInternal = To-Number $sheet.Cells.Item(11, 2).Value2
  $jdProvider = To-Number $sheet.Cells.Item(11, 3).Value2
  $jdDifference = To-Number $sheet.Cells.Item(11, 4).Value2
  $bjInternal = To-Number $sheet.Cells.Item(13, 2).Value2
  $bjProvider = To-Number $sheet.Cells.Item(13, 3).Value2
  $bjDifference = To-Number $sheet.Cells.Item(13, 4).Value2
  $tyInternal = To-Number $sheet.Cells.Item(15, 2).Value2
  $tyProvider = To-Number $sheet.Cells.Item(15, 3).Value2
  $tyDifference = To-Number $sheet.Cells.Item(15, 4).Value2

  $monthText = "${Month}月份"
  $jingdong = "@熊龙彬 ，@王安钦Eric 你好，${monthText}短信账单：`r`n" +
    (Difference-Message '永辉内部核算' $jdInternal '京东账单' $jdProvider $jdDifference)
  $jiangxi = "@邹国香 你好，${monthText}短信账单：`r`n" +
    (Difference-Message '去除“挂短”和“永辉会员营销AI交互”，永辉内部结算' $jmInternal '江西聚梦账单为' $jmProvider $jmDifference) +
    "`r`n@胡政 ${monthText}账单`“挂短`”产生费用 $(Money $jmShort) 元，`“永辉会员营销AI交互`”产生费用 $(Money $jmAi) 元。请知悉。"
  $beijing = "@李艳超 你好，${monthText}的短信账单：`r`n" +
    (Difference-Message '永辉内部结算' $bjInternal '北京凌渡账单' $bjProvider $bjDifference)

  $tengyu = "@杨莹莹 你好，${monthText}短信账单：`r`n" +
    (Difference-Message '永辉内部结算' $tyInternal '腾域账单' $tyProvider $tyDifference)

  $content = "京东：`r`n$jingdong`r`n`r`n江西聚梦：`r`n$jiangxi`r`n`r`n北京凌渡：`r`n$beijing`r`n`r`n腾域：`r`n$tengyu`r`n"
  $outputDir = Split-Path -Parent $OutputFile
  if (-not (Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
  }
  [IO.File]::WriteAllText($OutputFile, $content, (New-Object Text.UTF8Encoding($false)))

  @{
    ok = $true
    outputFile = $OutputFile
    messages = @{
      jingdong = $jingdong
      jiangxiJumeng = $jiangxi
      beijingLingdu = $beijing
      tengyu = $tengyu
    }
  } | ConvertTo-Json -Depth 5 -Compress
}
catch {
  @{ ok = $false; reason = $_.Exception.Message; outputFile = $OutputFile } | ConvertTo-Json -Depth 4 -Compress
  exit 1
}
finally {
  if ($book) { try { $book.Close($false) | Out-Null } catch {} }
  if ($sheet) { try { [Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null } catch {} }
  if ($book) { try { [Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null } catch {} }
  if ($excel) {
    try { $excel.Quit() | Out-Null } catch {}
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
