param(
  [Parameter(Mandatory=$true)][string]$SourceXlsx,
  [Parameter(Mandatory=$true)][string]$OutputXlsx,
  [Parameter(Mandatory=$true)][string]$ColumnName,
  [Parameter(Mandatory=$true)][string]$MatchValue
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Add-Type -AssemblyName System.IO.Compression.FileSystem

function ConvertTo-JsonLine($obj) {
  $obj | ConvertTo-Json -Depth 6 -Compress
}

function Get-ColumnName([string]$cellRef) {
  return ($cellRef -replace '\d+', '')
}

function Column-Number-To-Name([int]$number) {
  $name = ''
  while ($number -gt 0) {
    $number--
    $name = [char](65 + ($number % 26)) + $name
    $number = [Math]::Floor($number / 26)
  }
  return $name
}

function Column-Name-To-Number([string]$name) {
  $number = 0
  foreach ($ch in $name.ToUpperInvariant().ToCharArray()) {
    $number = ($number * 26) + ([int][char]$ch - [int][char]'A' + 1)
  }
  return $number
}

function Read-Number([string]$value) {
  $text = $value.Trim().Replace(',', '')
  if ($text -match '^-?\d+(\.\d+)?$') {
    return [double]$text
  }
  return 0
}

function Get-CellValue($cell, $sharedStrings, $ns) {
  $type = [string]$cell.GetAttribute('t')
  $valueNode = $cell.GetElementsByTagName('v', $ns) | Select-Object -First 1
  if ($type -eq 's') {
    $indexText = if ($valueNode) { [string]$valueNode.InnerText } else { '' }
    if ($indexText -match '^\d+$') {
      $index = [int]$indexText
      if ($index -ge 0 -and $index -lt $sharedStrings.Count) {
        return $sharedStrings[$index]
      }
    }
    return ''
  }
  if ($type -eq 'inlineStr') {
    $textNodes = $cell.GetElementsByTagName('t', $ns)
    return (($textNodes | ForEach-Object { $_.InnerText }) -join '')
  }
  if ($valueNode) {
    return [string]$valueNode.InnerText
  }
  return ''
}

function Get-Cell($row, [string]$column, $ns) {
  foreach ($cell in $row.GetElementsByTagName('c', $ns)) {
    if ((Get-ColumnName ([string]$cell.GetAttribute('r'))) -eq $column) {
      return $cell
    }
  }
  return $null
}

function Ensure-Cell($doc, $row, [string]$column, [int]$rowNumber, $ns) {
  $cell = Get-Cell $row $column $ns
  if ($cell) { return $cell }

  $cell = $doc.CreateElement('c', $ns)
  $cell.SetAttribute('r', "$column$rowNumber")
  $targetColumnNumber = Column-Name-To-Number $column
  $inserted = $false
  foreach ($existing in @($row.GetElementsByTagName('c', $ns))) {
    $existingColumnNumber = Column-Name-To-Number (Get-ColumnName ([string]$existing.GetAttribute('r')))
    if ($existingColumnNumber -gt $targetColumnNumber) {
      $row.InsertBefore($cell, $existing) | Out-Null
      $inserted = $true
      break
    }
  }
  if (-not $inserted) {
    $row.AppendChild($cell) | Out-Null
  }
  return $cell
}

function Clear-CellValue($cell, $ns) {
  if (-not $cell) { return }
  if ($cell.HasAttribute('t')) {
    $cell.RemoveAttribute('t')
  }
  foreach ($child in @($cell.ChildNodes)) {
    if ($child.LocalName -eq 'v' -or $child.LocalName -eq 'is') {
      $cell.RemoveChild($child) | Out-Null
    }
  }
}

function Set-CellNumber($doc, $row, [string]$column, [int]$rowNumber, [double]$value, $ns) {
  $cell = Ensure-Cell $doc $row $column $rowNumber $ns
  Clear-CellValue $cell $ns
  $v = $doc.CreateElement('v', $ns)
  $v.InnerText = $value.ToString('0.##', [System.Globalization.CultureInfo]::InvariantCulture)
  $cell.AppendChild($v) | Out-Null
}

function Get-RowValues($row, $sharedStrings, $ns) {
  $values = @{}
  foreach ($cell in $row.GetElementsByTagName('c', $ns)) {
    $column = Get-ColumnName ([string]$cell.GetAttribute('r'))
    $values[$column] = Get-CellValue $cell $sharedStrings $ns
  }
  return $values
}

function Set-RowNumber($row, [int]$rowNumber, $ns) {
  $row.SetAttribute('r', [string]$rowNumber)
  foreach ($cell in $row.GetElementsByTagName('c', $ns)) {
    $column = Get-ColumnName ([string]$cell.GetAttribute('r'))
    $cell.SetAttribute('r', "$column$rowNumber")
  }
}

$workRoot = Join-Path $env:USERPROFILE 'yh_filter_work'
if (-not (Test-Path -LiteralPath $workRoot)) {
  New-Item -ItemType Directory -Path $workRoot -Force | Out-Null
}
$tempRoot = Join-Path $workRoot ("yh_filter_" + [guid]::NewGuid().ToString('N'))
$zipPath = Join-Path $workRoot ("yh_filter_" + [guid]::NewGuid().ToString('N') + ".zip")

try {
  if (-not (Test-Path -LiteralPath $SourceXlsx)) {
    throw "Source XLSX not found: $SourceXlsx"
  }

  Copy-Item -LiteralPath $SourceXlsx -Destination $zipPath -Force
  [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $tempRoot)

  $sharedStrings = New-Object System.Collections.Generic.List[string]
  $sharedPath = Join-Path $tempRoot 'xl\sharedStrings.xml'
  if (Test-Path -LiteralPath $sharedPath) {
    [xml]$sharedXml = Get-Content -LiteralPath $sharedPath -Raw -Encoding UTF8
    $sharedNs = $sharedXml.DocumentElement.NamespaceURI
    foreach ($si in $sharedXml.GetElementsByTagName('si', $sharedNs)) {
      $texts = $si.GetElementsByTagName('t', $sharedNs) | ForEach-Object { $_.InnerText }
      $sharedStrings.Add(($texts -join '')) | Out-Null
    }
  }

  # 合并账单中“汇总账单”不一定是第一个工作表，按表头定位实际工作表。
  $sheetPath = Join-Path $tempRoot 'xl\worksheets\sheet1.xml'
  $sheetPath = Join-Path $tempRoot 'xl\worksheets\sheet1.xml'
  $headerIndex = -1
  for ($si = 0; $si -lt $sharedStrings.Count; $si++) {
    if ($sharedStrings[$si] -eq $ColumnName) { $headerIndex = $si; break }
  }
  if ($headerIndex -ge 0) {
    foreach ($candidatePath in (Get-ChildItem -LiteralPath (Join-Path $tempRoot 'xl\worksheets') -Filter 'sheet*.xml' | Select-Object -ExpandProperty FullName)) {
      $candidateText = Get-Content -LiteralPath $candidatePath -Raw -Encoding UTF8
      if ($candidateText -match "<v>$headerIndex</v>") { $sheetPath = $candidatePath; break }
    }
  }
  if (-not (Test-Path -LiteralPath $sheetPath)) { throw 'No worksheet found in source workbook.' }

  $workbookPath = [System.IO.Path]::Combine($tempRoot, 'xl', 'workbook.xml')
  $workbookRelsPath = [System.IO.Path]::Combine($tempRoot, 'xl', '_rels', 'workbook.xml.rels')
  if (Test-Path -LiteralPath $workbookPath) {
    [xml]$workbookXml = Get-Content -LiteralPath $workbookPath -Raw -Encoding UTF8
    $workbookNs = $workbookXml.DocumentElement.NamespaceURI
    $sheetFileName = [System.IO.Path]::GetFileName($sheetPath)
    $targetRelId = $null
    if (Test-Path -LiteralPath $workbookRelsPath) {
      [xml]$relsXml = Get-Content -LiteralPath $workbookRelsPath -Raw -Encoding UTF8
      $relsNs = $relsXml.DocumentElement.NamespaceURI
      foreach ($rel in @($relsXml.DocumentElement.GetElementsByTagName('Relationship', $relsNs))) {
        $target = [string]$rel.GetAttribute('Target')
        if ([System.IO.Path]::GetFileName($target) -eq $sheetFileName) {
          $targetRelId = [string]$rel.GetAttribute('Id')
          break
        }
      }
    }
    $sheets = $workbookXml.GetElementsByTagName('sheets', $workbookNs) | Select-Object -First 1
    if ($sheets -and $targetRelId) {
      foreach ($sheet in @($sheets.GetElementsByTagName('sheet', $workbookNs))) {
        if ([string]$sheet.GetAttribute('r:id') -ne $targetRelId) {
          $sheets.RemoveChild($sheet) | Out-Null
        }
      }
      $workbookWriterSettings = New-Object System.Xml.XmlWriterSettings
      $workbookWriterSettings.Encoding = New-Object System.Text.UTF8Encoding($false)
      $workbookWriterSettings.Indent = $false
      $writer = [System.Xml.XmlWriter]::Create($workbookPath, $workbookWriterSettings)
      $workbookXml.Save($writer)
      $writer.Close()
    }
  }

  [xml]$sheetXml = Get-Content -LiteralPath $sheetPath -Raw -Encoding UTF8
  $ns = $sheetXml.DocumentElement.NamespaceURI
  $sheetData = $sheetXml.GetElementsByTagName('sheetData', $ns) | Select-Object -First 1
  if (-not $sheetData) {
    throw 'sheetData not found in source worksheet.'
  }

  $rows = @($sheetData.GetElementsByTagName('row', $ns))
  if ($rows.Count -lt 1) {
    throw 'No rows found in source worksheet.'
  }

  $headerRow = $rows[0]
  $headerValues = Get-RowValues $headerRow $sharedStrings $ns
  $filterColumn = $null
  $amountColumn = $null
  $totalColumn = $null
  foreach ($key in $headerValues.Keys) {
    if ($headerValues[$key].Trim() -eq $ColumnName) { $filterColumn = $key }
    if ($headerValues[$key].Trim() -eq "$([char]0x91D1)$([char]0x989D)") { $amountColumn = $key }
    if ($headerValues[$key].Trim() -eq "$([char]0x5408)$([char]0x8BA1)") { $totalColumn = $key }
  }
  if (-not $filterColumn) {
    throw "Filter column not found: $ColumnName"
  }

  $sourceTotalGroupByRow = @{}
  if ($totalColumn) {
    foreach ($mergeCell in @($sheetXml.GetElementsByTagName('mergeCell', $ns))) {
      $ref = [string]$mergeCell.GetAttribute('ref')
      if ($ref -match '^([A-Z]+)(\d+):([A-Z]+)(\d+)$' -and $Matches[1] -eq $totalColumn -and $Matches[3] -eq $totalColumn) {
        $startRow = [int]$Matches[2]
        $endRow = [int]$Matches[4]
        for ($r = $startRow; $r -le $endRow; $r++) {
          $sourceTotalGroupByRow[$r] = $startRow
        }
      }
    }
  }

  $newRows = New-Object System.Collections.Generic.List[object]
  $selectedRows = New-Object System.Collections.Generic.List[object]
  $newRows.Add($headerRow.CloneNode($true)) | Out-Null
  $amountTotal = 0

  for ($i = 1; $i -lt $rows.Count; $i++) {
    $values = Get-RowValues $rows[$i] $sharedStrings $ns
    $candidate = if ($values.ContainsKey($filterColumn)) { $values[$filterColumn].Trim() } else { '' }
    if ($candidate -eq $MatchValue) {
      $sourceRowNumberText = [string]$rows[$i].GetAttribute('r')
      $sourceRowNumber = if ($sourceRowNumberText -match '^\d+$') { [int]$sourceRowNumberText } else { $i + 1 }
      $groupKey = if ($sourceTotalGroupByRow.ContainsKey($sourceRowNumber)) { $sourceTotalGroupByRow[$sourceRowNumber] } else { $sourceRowNumber }
      $amount = 0
      if ($amountColumn -and $values.ContainsKey($amountColumn)) {
        $amount = Read-Number $values[$amountColumn]
        $amountTotal += $amount
      }
      $selectedRows.Add([pscustomobject]@{
        Node = $rows[$i].CloneNode($true)
        Group = $groupKey
        Amount = $amount
      }) | Out-Null
    }
  }

  foreach ($selected in $selectedRows) {
    $newRows.Add($selected.Node) | Out-Null
  }

  while ($sheetData.HasChildNodes) {
    $sheetData.RemoveChild($sheetData.FirstChild) | Out-Null
  }

  for ($i = 0; $i -lt $newRows.Count; $i++) {
    $rowNumber = $i + 1
    $row = $newRows[$i]
    Set-RowNumber $row $rowNumber $ns
    $sheetData.AppendChild($row) | Out-Null
  }

  foreach ($mergeNode in @($sheetXml.GetElementsByTagName('mergeCells', $ns))) {
    $mergeNode.ParentNode.RemoveChild($mergeNode) | Out-Null
  }

  $lastColumn = Column-Number-To-Name $headerValues.Count
  $lastRow = [Math]::Max(1, $newRows.Count)
  $rangeRef = "A1:$lastColumn$lastRow"

  $dimension = $sheetXml.GetElementsByTagName('dimension', $ns) | Select-Object -First 1
  if ($dimension) {
    $dimension.SetAttribute('ref', $rangeRef)
  }
  $autoFilter = $sheetXml.GetElementsByTagName('autoFilter', $ns) | Select-Object -First 1
  if ($autoFilter) {
    $autoFilter.SetAttribute('ref', $rangeRef)
  } else {
    $autoFilter = $sheetXml.CreateElement('autoFilter', $ns)
    $autoFilter.SetAttribute('ref', $rangeRef)
    $sheetXml.DocumentElement.AppendChild($autoFilter) | Out-Null
  }

  $mergedGroups = 0
  if ($totalColumn -and $selectedRows.Count -gt 0) {
    $mergeCells = $sheetXml.CreateElement('mergeCells', $ns)
    $selectedIndex = 0
    while ($selectedIndex -lt $selectedRows.Count) {
      $group = $selectedRows[$selectedIndex].Group
      $groupStartIndex = $selectedIndex
      $groupAmount = 0
      while ($selectedIndex -lt $selectedRows.Count -and $selectedRows[$selectedIndex].Group -eq $group) {
        $groupAmount += [double]$selectedRows[$selectedIndex].Amount
        $selectedIndex++
      }
      $groupEndIndex = $selectedIndex - 1

      $startRowNumber = $groupStartIndex + 2
      $endRowNumber = $groupEndIndex + 2
      for ($r = $startRowNumber; $r -le $endRowNumber; $r++) {
        Clear-CellValue (Get-Cell $newRows[$r - 1] $totalColumn $ns) $ns
      }
      Set-CellNumber $sheetXml $newRows[$startRowNumber - 1] $totalColumn $startRowNumber ([Math]::Round($groupAmount, 2)) $ns

      if ($endRowNumber -gt $startRowNumber) {
        $mergeCell = $sheetXml.CreateElement('mergeCell', $ns)
        $mergeCell.SetAttribute('ref', "$totalColumn$startRowNumber`:$totalColumn$endRowNumber")
        $mergeCells.AppendChild($mergeCell) | Out-Null
        $mergedGroups++
      }
    }

    if ($mergedGroups -gt 0) {
      $mergeCells.SetAttribute('count', [string]$mergedGroups)
      $sheetXml.DocumentElement.InsertAfter($mergeCells, $autoFilter) | Out-Null
    }
  }

  $settings = New-Object System.Xml.XmlWriterSettings
  $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
  $settings.Indent = $false
  $writer = [System.Xml.XmlWriter]::Create($sheetPath, $settings)
  $sheetXml.Save($writer)
  $writer.Close()

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

  $tempOut = Join-Path $workRoot ("yh_filter_out_" + [guid]::NewGuid().ToString('N') + ".zip")
  [System.IO.Compression.ZipFile]::CreateFromDirectory($tempRoot, $tempOut)
  Move-Item -LiteralPath $tempOut -Destination $OutputXlsx -Force

  ConvertTo-JsonLine @{
    ok = $true
    sourceFile = $SourceXlsx
    outputFile = $OutputXlsx
    columnName = $ColumnName
    matchValue = $MatchValue
    matchedRows = [Math]::Max(0, $newRows.Count - 1)
    amountTotal = [Math]::Round($amountTotal, 2)
    mergedGroups = $mergedGroups
  }
}
catch {
  ConvertTo-JsonLine @{
    ok = $false
    reason = $_.Exception.Message
    line = $_.InvocationInfo.ScriptLineNumber
    command = ([string]$_.InvocationInfo.Line).Trim()
    sourceFile = $SourceXlsx
    outputFile = $OutputXlsx
    columnName = $ColumnName
    matchValue = $MatchValue
  }
  exit 1
}
finally {
  if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
  }
  if (Test-Path -LiteralPath $tempRoot) {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
  }
}
