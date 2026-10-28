param(
  [Parameter(Mandatory = $true)][string]$SourceXlsx,
  [Parameter(Mandatory = $true)][string]$GroupsBase64
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

function ConvertTo-JsonLine($object) {
  $object | ConvertTo-Json -Depth 10 -Compress
}

function U([int[]]$codes) {
  -join ($codes | ForEach-Object { [char]$_ })
}

function Release-ComObject($object) {
  if ($null -ne $object) {
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($object) | Out-Null } catch {}
  }
}

$excel = $null
$workbook = $null
$sourceSheet = $null

try {
  if (-not (Test-Path -LiteralPath $SourceXlsx)) {
    throw "Source workbook not found: $SourceXlsx"
  }

  $groupsJson = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($GroupsBase64))
  $groups = ConvertFrom-Json -InputObject $groupsJson
  if ($null -eq $groups -or @($groups).Count -eq 0) {
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
  $accountHeader = U @(0x8D26,0x53F7,0x7F16,0x7801)
  $contactHeader = U @(0x5BF9,0x63A5,0x4EBA)
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
  if (-not $headers.ContainsKey($accountHeader)) {
    throw "Required column not found: $accountHeader"
  }
  $requiresContactColumn = @($groups | Where-Object { [string]$_.rule -eq "contact" }).Count -gt 0
  if ($requiresContactColumn -and -not $headers.ContainsKey($contactHeader)) {
    throw "Required column not found: $contactHeader"
  }

  $accountColumn = [int]$headers[$accountHeader]
  $contactColumn = if ($headers.ContainsKey($contactHeader)) { [int]$headers[$contactHeader] } else { 0 }
  $billAccounts = New-Object System.Collections.Generic.List[string]
  $billAccountSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
  for ($row = 2; $row -le $rowCount; $row++) {
    $account = ([string]$sourceSheet.Cells.Item($row, $accountColumn).Text).Trim()
    if ($account -and $billAccountSet.Add($account)) {
      $billAccounts.Add($account)
    }
  }

  $configuredAccounts = New-Object System.Collections.Generic.List[string]
  $configuredAccountGroups = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::Ordinal)
  $groupResults = New-Object System.Collections.Generic.List[hashtable]

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
    $presentAccounts = New-Object System.Collections.Generic.List[string]
    $missingAccounts = New-Object System.Collections.Generic.List[string]

    foreach ($account in $accounts) {
      if (-not $configuredAccountGroups.ContainsKey($account)) {
        $configuredAccountGroups[$account] = New-Object System.Collections.Generic.List[int]
        $configuredAccounts.Add($account)
      }
      $configuredAccountGroups[$account].Add([int]$group.id)

      if ($billAccountSet.Contains($account)) {
        $presentAccounts.Add($account)
      } else {
        $missingAccounts.Add($account)
      }
    }

    $groupResults.Add(@{
      id = [int]$group.id
      name = [string]$group.name
      isSystem = [bool]$group.isSystem
      accounts = @($accounts)
      presentAccounts = @($presentAccounts)
      missingAccounts = @($missingAccounts)
    })
  }

  $coveredAccounts = New-Object System.Collections.Generic.List[string]
  $uncoveredAccounts = New-Object System.Collections.Generic.List[string]
  $coverage = New-Object System.Collections.Generic.List[hashtable]
  foreach ($account in $billAccounts) {
    if ($configuredAccountGroups.ContainsKey($account)) {
      $coveredAccounts.Add($account)
      $coverage.Add(@{
        account = $account
        groupIds = @($configuredAccountGroups[$account])
      })
    } else {
      $uncoveredAccounts.Add($account)
    }
  }

  $unusedAccounts = New-Object System.Collections.Generic.List[hashtable]
  foreach ($account in $configuredAccounts) {
    if (-not $billAccountSet.Contains($account)) {
      $unusedAccounts.Add(@{
        account = $account
        groupIds = @($configuredAccountGroups[$account])
      })
    }
  }

  $duplicateAccounts = New-Object System.Collections.Generic.List[hashtable]
  foreach ($entry in $configuredAccountGroups.GetEnumerator()) {
    $groupIds = @($entry.Value)
    if ($groupIds.Count -gt 1) {
      $duplicateAccounts.Add(@{
        account = $entry.Key
        groupIds = $groupIds
      })
    }
  }

  ConvertTo-JsonLine @{
    ok = $true
    billAccountCount = $billAccounts.Count
    configuredAccountCount = $configuredAccounts.Count
    coveredAccountCount = $coveredAccounts.Count
    uncoveredAccountCount = $uncoveredAccounts.Count
    unusedAccountCount = $unusedAccounts.Count
    complete = ($uncoveredAccounts.Count -eq 0 -and $duplicateAccounts.Count -eq 0)
    billAccounts = @($billAccounts)
    coveredAccounts = @($coveredAccounts)
    uncoveredAccounts = @($uncoveredAccounts)
    coverage = @($coverage)
    unusedAccounts = @($unusedAccounts)
    duplicateAccounts = @($duplicateAccounts)
    groups = @($groupResults)
  }
}
catch {
  ConvertTo-JsonLine @{
    ok = $false
    reason = $_.Exception.Message
    line = $_.InvocationInfo.ScriptLineNumber
    position = $_.InvocationInfo.PositionMessage
  }
  exit 1
}
finally {
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
