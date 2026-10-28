param(
  [Parameter(Mandatory = $true)][string]$SourceXlsx,
  [Parameter(Mandatory = $true)][string]$OutputDir,
  [Parameter(Mandatory = $true)][string]$GroupsBase64
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.Drawing

function ConvertTo-JsonLine($object) {
  $object | ConvertTo-Json -Depth 8 -Compress
}

function U([int[]]$codes) {
  -join ($codes | ForEach-Object { [char]$_ })
}

function Release-ComObject($object) {
  if ($null -ne $object) {
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($object) | Out-Null } catch {}
  }
}

function Safe-FileName([string]$name) {
  $clean = $name -replace '[<>:"/\\|?*\x00-\x1F]', '_'
  if ($clean.Length -gt 120) { return $clean.Substring(0, 120) }
  return $clean
}

function Trim-WhiteMargins([string]$imagePath) {
  $bitmap = [System.Drawing.Bitmap]::FromFile($imagePath)
  try {
    $left = $bitmap.Width
    $top = $bitmap.Height
    $right = -1
    $bottom = -1

    for ($y = 0; $y -lt $bitmap.Height; $y++) {
      for ($x = 0; $x -lt $bitmap.Width; $x++) {
        $pixel = $bitmap.GetPixel($x, $y)
        if ($pixel.R -lt 252 -or $pixel.G -lt 252 -or $pixel.B -lt 252) {
          if ($x -lt $left) { $left = $x }
          if ($y -lt $top) { $top = $y }
          if ($x -gt $right) { $right = $x }
          if ($y -gt $bottom) { $bottom = $y }
        }
      }
    }

    if ($right -lt 0 -or $bottom -lt 0) {
      return
    }

    for ($y = $bitmap.Height - 1; $y -ge 0; $y--) {
      $count = 0
      for ($x = 0; $x -lt $bitmap.Width; $x++) {
        $pixel = $bitmap.GetPixel($x, $y)
        if ($pixel.R -lt 252 -or $pixel.G -lt 252 -or $pixel.B -lt 252) { $count++ }
      }
      if ($count -gt 5 -and $count -lt ($bitmap.Width * 0.85)) {
        $bottom = $y
        break
      }
    }

    for ($x = $bitmap.Width - 1; $x -ge 0; $x--) {
      $count = 0
      for ($y = 0; $y -lt $bitmap.Height; $y++) {
        $pixel = $bitmap.GetPixel($x, $y)
        if ($pixel.R -lt 252 -or $pixel.G -lt 252 -or $pixel.B -lt 252) { $count++ }
      }
      if ($count -gt 5 -and $count -lt ($bitmap.Height * 0.85)) {
        $right = $x
        break
      }
    }

    $padding = 6
    $left = [Math]::Max(0, $left - $padding)
    $top = [Math]::Max(0, $top - $padding)
    $right = [Math]::Min($bitmap.Width - 1, $right + $padding)
    $bottom = [Math]::Min($bitmap.Height - 1, $bottom + $padding)
    $cropRect = New-Object System.Drawing.Rectangle([int]$left, [int]$top, [int]($right - $left + 1), [int]($bottom - $top + 1))
    $cropped = $bitmap.Clone($cropRect, $bitmap.PixelFormat)
    $tempPath = "$imagePath.tmp.png"
    try {
      $bitmap.Dispose()
      $cropped.Save($tempPath, [System.Drawing.Imaging.ImageFormat]::Png)
    } finally {
      $cropped.Dispose()
    }
    Remove-Item -LiteralPath $imagePath -Force
    Move-Item -LiteralPath $tempPath -Destination $imagePath
  } finally {
    $bitmap.Dispose()
  }
}

$excel = $null
$workbook = $null
$sourceSheet = $null
$temporaryWorkbook = $null

try {
  if (-not (Test-Path -LiteralPath $SourceXlsx)) {
    throw "Source workbook not found: $SourceXlsx"
  }
  New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

  $groupsJson = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($GroupsBase64))
  $groups = ConvertFrom-Json -InputObject $groupsJson
  if ($groups.Count -eq 0) {
    throw "No internal reconciliation groups were provided."
  }

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $workbook = $excel.Workbooks.Open($SourceXlsx, 3, $true)
  $summarySheetNames = @(
    (U @(0x6C47,0x603B,0x5BF9,0x8D26,0x5355)),
    (U @(0x6C47,0x603B,0x8D26,0x5355))
  )
  $monthHeader = U @(0x6708,0x4EFD)
  $accountHeader = U @(0x8D26,0x53F7,0x7F16,0x7801)
  $contactHeader = U @(0x5BF9,0x63A5,0x4EBA)
  $taxHeader = U @(0x7EB3,0x7A0E,0x4EBA,0x6807,0x8BC6,0x53F7)
  foreach ($worksheet in $workbook.Worksheets) {
    if ($summarySheetNames -contains ([string]$worksheet.Name).Trim()) {
      $sourceSheet = $worksheet
      break
    }
    Release-ComObject $worksheet
  }
  if ($null -eq $sourceSheet) {
    throw "Worksheet not found. Expected one of: $($summarySheetNames -join ', ')"
  }

  $used = $sourceSheet.UsedRange
  $rowCount = $used.Rows.Count
  $columnCount = $used.Columns.Count
  $headers = @{}
  for ($column = 1; $column -le $columnCount; $column++) {
    $header = ([string]$sourceSheet.Cells.Item(1, $column).Text).Trim()
    if ($header) { $headers[$header] = $column }
  }

  foreach ($requiredHeader in @($monthHeader, $accountHeader, $taxHeader)) {
    if (-not $headers.ContainsKey($requiredHeader)) {
      throw "Required column not found: $requiredHeader"
    }
  }
  $requiresContactColumn = @($groups | Where-Object { [string]$_.rule -eq "contact" }).Count -gt 0
  if ($requiresContactColumn -and -not $headers.ContainsKey($contactHeader)) {
    throw "Required column not found: $contactHeader"
  }

  $firstColumn = [int]$headers[$monthHeader]
  $accountColumn = [int]$headers[$accountHeader]
  $contactColumn = if ($headers.ContainsKey($contactHeader)) { [int]$headers[$contactHeader] } else { 0 }
  $lastColumn = [int]$headers[$taxHeader]
  $outputGroups = @()

  foreach ($group in $groups) {
    $isContactRule = [string]$group.rule -eq "contact"
    if ($isContactRule) {
      $contactAccounts = New-Object System.Collections.Generic.List[string]
      $contactAccountSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
      $matchValue = [string]$group.matchValue
      for ($row = 2; $row -le $rowCount; $row++) {
        $contact = ([string]$sourceSheet.Cells.Item($row, $contactColumn).Text).Trim()
        if ($contact -ne $matchValue) { continue }
        $account = ([string]$sourceSheet.Cells.Item($row, $accountColumn).Text).Trim()
        if ($account -and $contactAccountSet.Add($account)) {
          $contactAccounts.Add($account)
        }
      }
      $accounts = @($contactAccounts)
    } else {
      $accounts = @($group.accounts | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
    }
    $accountSet = @{}
    foreach ($account in $accounts) { $accountSet[$account] = $true }

    $matchingRows = New-Object System.Collections.Generic.List[object]
    $matchedAccounts = New-Object System.Collections.Generic.List[string]
    for ($row = 2; $row -le $rowCount; $row++) {
      $accountValue = ([string]$sourceSheet.Cells.Item($row, $accountColumn).Text).Trim()
      $contactValue = if ($isContactRule) { ([string]$sourceSheet.Cells.Item($row, $contactColumn).Text).Trim() } else { "" }
      if (($isContactRule -and $contactValue -eq [string]$group.matchValue) -or (-not $isContactRule -and $accountSet.ContainsKey($accountValue))) {
        $matchingRows.Add($row)
        if (-not $matchedAccounts.Contains($accountValue)) {
          $matchedAccounts.Add($accountValue)
        }
      }
    }

    $missingAccounts = @($accounts | Where-Object { -not $matchedAccounts.Contains($_) })
    $imageFile = $null

    if ($matchingRows.Count -gt 0) {
      $temporaryWorkbook = $excel.Workbooks.Add()
      $temporarySheet = $temporaryWorkbook.Worksheets.Item(1)
      $temporarySheet.Name = (U @(0x5185,0x90E8,0x5BF9,0x8D26))

      $outputColumnCount = $lastColumn - $firstColumn + 1
      for ($outputColumn = 1; $outputColumn -le $outputColumnCount; $outputColumn++) {
        $sourceColumn = $firstColumn + $outputColumn - 1
        $temporarySheet.Cells.Item(1, $outputColumn).Value2 = $sourceSheet.Cells.Item(1, $sourceColumn).Value2
        $outputRow = 2
        foreach ($sourceRow in $matchingRows) {
          $temporarySheet.Cells.Item($outputRow, $outputColumn).Value2 = [string]$sourceSheet.Cells.Item($sourceRow, $sourceColumn).Text
          $outputRow++
        }
      }

      $range = $temporarySheet.Range(
        $temporarySheet.Cells.Item(1, 1),
        $temporarySheet.Cells.Item($matchingRows.Count + 1, $outputColumnCount)
      )
      $range.Font.Name = "Microsoft YaHei"
      $range.Font.Size = 10
      $range.Borders.LineStyle = 1
      $range.Rows.Item(1).Font.Bold = $true
      $range.Rows.Item(1).Interior.ColorIndex = 15
      $range.Columns.AutoFit() | Out-Null
      for ($column = 1; $column -le $outputColumnCount; $column++) {
        if ($temporarySheet.Columns.Item($column).ColumnWidth -gt 28) {
          $temporarySheet.Columns.Item($column).ColumnWidth = 28
        }
      }
      $temporarySheet.Range("A1").AutoFilter() | Out-Null
      $temporarySheet.Activate() | Out-Null
      $range.CopyPicture(1, 1) | Out-Null
      Start-Sleep -Milliseconds 400

      $width = [Math]::Max(900, $outputColumnCount * 120)
      $height = [Math]::Max(120, ($matchingRows.Count + 1) * 24)
      $chartObject = $temporarySheet.ChartObjects().Add(0, 0, $width, $height)
      $chartObject.Activate() | Out-Null
      Start-Sleep -Milliseconds 200
      $chartObject.Chart.Paste() | Out-Null
      Start-Sleep -Milliseconds 400

      $safeName = Safe-FileName ("{0}_{1}_{2}" -f (U @(0x5185,0x90E8,0x5BF9,0x8D26)), $group.id, ($accounts -join "_"))
      $imageFile = Join-Path $OutputDir "$safeName.png"
      $chartObject.Chart.Export($imageFile, "PNG") | Out-Null
      Trim-WhiteMargins $imageFile

      $chartObject.Delete()
      Release-ComObject $chartObject
      Release-ComObject $range
      $temporaryWorkbook.Close($false) | Out-Null
      Release-ComObject $temporarySheet
      Release-ComObject $temporaryWorkbook
      $temporaryWorkbook = $null
    }

    $outputGroups += @{
      id = [int]$group.id
      name = [string]$group.name
      isSystem = [bool]$group.isSystem
      imageFile = $imageFile
      matchedRows = $matchingRows.Count
      matchedAccounts = @($matchedAccounts)
      missingAccounts = @($missingAccounts)
    }
  }

  ConvertTo-JsonLine @{ ok = $true; groups = $outputGroups }
}
catch {
  ConvertTo-JsonLine @{
    ok = $false
    reason = $_.Exception.Message
    line = $_.InvocationInfo.ScriptLineNumber
    position = $_.InvocationInfo.PositionMessage
    groupId = if ($null -ne $group) { "$($group.id)" } else { $null }
    groupIdType = if ($null -ne $group -and $null -ne $group.id) { $group.id.GetType().FullName } else { $null }
  }
  exit 1
}
finally {
  if ($temporaryWorkbook) {
    try { $temporaryWorkbook.Close($false) | Out-Null } catch {}
    Release-ComObject $temporaryWorkbook
  }
  if ($sourceSheet) { Release-ComObject $sourceSheet }
  if ($workbook) {
    try { $workbook.Close($false) | Out-Null } catch {}
    Release-ComObject $workbook
  }
  if ($excel) {
    try { $excel.Quit() | Out-Null } catch {}
    Release-ComObject $excel
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
