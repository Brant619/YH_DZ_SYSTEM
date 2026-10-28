param(
  [Parameter(Mandatory = $true)][string]$OutputDocx,
  [Parameter(Mandatory = $true)][string]$GroupsJson
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

function U([int[]]$codes) {
  -join ($codes | ForEach-Object { [char]$_ })
}

function Release-ComObject($object) {
  if ($null -ne $object) {
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($object) | Out-Null } catch {}
  }
}

$word = $null
$doc = $null

try {
  $groups = ConvertFrom-Json -InputObject $GroupsJson
  if (-not $groups -or $groups.Count -eq 0) {
    throw "No groups were provided."
  }

  $outDir = Split-Path -Parent $OutputDocx
  if (-not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
  }
  if (Test-Path -LiteralPath $OutputDocx) {
    Remove-Item -LiteralPath $OutputDocx -Force
  }

  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Add()
  $doc.PageSetup.TopMargin = 36
  $doc.PageSetup.BottomMargin = 36
  $doc.PageSetup.LeftMargin = 36
  $doc.PageSetup.RightMargin = 36

  $accountSep = U @(0x3001)
  $colon = U @(0xFF1A)
  $newline = "`r`n"

  foreach ($group in $groups) {
    $accounts = @($group.accounts | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
    $message = [string]$group.message
    $imagePath = [string]$group.assetFile

    $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
    $range.InsertAfter(($accounts -join $accountSep) + $colon + $newline)

    if ($imagePath -and (Test-Path -LiteralPath $imagePath)) {
      $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
      $shape = $doc.InlineShapes.AddPicture($imagePath, $false, $true, $range)
      $shape.LockAspectRatio = $true
      if ($shape.Width -gt 468) {
        $shape.Width = 468
      }
      $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
      $range.InsertAfter($newline)
    }

    $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
    $range.InsertAfter($message + $newline + $newline)
  }

  $doc.SaveAs2($OutputDocx, 16)
  [pscustomobject]@{ ok = $true; outputFile = $OutputDocx } | ConvertTo-Json -Compress
}
catch {
  [pscustomobject]@{ ok = $false; reason = $_.Exception.Message } | ConvertTo-Json -Compress
  exit 1
}
finally {
  if ($doc) {
    try { $doc.Close($false) | Out-Null } catch {}
    Release-ComObject $doc
  }
  if ($word) {
    try { $word.Quit() | Out-Null } catch {}
    Release-ComObject $word
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
