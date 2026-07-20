<#
  Combine-Into-Master.ps1

  Takes the Step 6 output ("MTC_Import_Rows.xlsx", up to 4 tabs: MTC Sales Details,
  Exchange Rate, Net Selling Price, PO comm) and merges it into a NEW copy of the
  master workbook using real Excel automation (COM) — never the original file, and
  never a JS re-serialization — so pivot tables, XLOOKUP formulas, and formatting
  elsewhere in the workbook are untouched.

  For each append-style sheet: FillDown extends every column's formula from the
  template (last existing) row, then each known "raw data" column is individually
  overwritten with the real per-row values from the import file — column by column,
  not via PasteSpecial's SkipBlanks (which doesn't marshal reliably over COM from
  PowerShell and silently pastes over everything instead of skipping blanks).

  Usage:
    powershell -ExecutionPolicy Bypass -File Combine-Into-Master.ps1 -MasterPath "...\Sales-Data-MTC-Mthly Rate.xlsx" -ImportPath "...\MTC_Import_Rows.xlsx" -OutputPath "...\Sales-Data-MTC-Mthly Rate - UPDATED.xlsx"
#>

param(
  [Parameter(Mandatory=$true)][string]$MasterPath,
  [Parameter(Mandatory=$true)][string]$ImportPath,
  [Parameter(Mandatory=$true)][string]$OutputPath
)

$ErrorActionPreference = "Stop"

$xlUp = -4162
$xlPasteValues = -4163

# 1-indexed Excel column numbers that get raw/derived values (everything else is a formula extended via FillDown)
$SALES_SEED_COLS = @(2,3,5,6,11,12,13,22,24,29,34,36,52,53,54,55,74,75,77,78,80,81,99,102,103)
$NSP_SEED_COLS = @(2..35) + @(37)
$POCOMM_SEED_COLS = @(1..35)

if (-not (Test-Path $MasterPath)) { throw "Master file not found: $MasterPath" }
if (-not (Test-Path $ImportPath)) { throw "Import file not found: $ImportPath" }

Write-Output "Copying master file to output path (original is never modified)..."
Copy-Item $MasterPath $OutputPath -Force

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
$excel.ScreenUpdating = $false
$excel.EnableEvents = $false

function Open-WB($path, $readOnly) {
  $wb = $null
  for ($attempt = 1; $attempt -le 3 -and $wb -eq $null; $attempt++) {
    try { $wb = $excel.Workbooks.Open($path, [Type]::Missing, $readOnly) }
    catch { Start-Sleep -Milliseconds 1000 }
  }
  if ($wb -eq $null) { throw "Could not open $path after 3 attempts" }
  return $wb
}

# Extends formulas from the template row down through the new rows, then overwrites
# each seed column with the real per-row values copied from the import sheet.
function Merge-AppendSheet($mWs, $impWs, $lastRow, $n, $numCols, $seedCols, $impFirstDataRow) {
  $fillRange = $mWs.Range($mWs.Cells($lastRow, 1), $mWs.Cells($lastRow + $n, $numCols))
  $fillRange.FillDown()

  foreach ($col in $seedCols) {
    $srcRange = $impWs.Range($impWs.Cells($impFirstDataRow, $col), $impWs.Cells($impFirstDataRow + $n - 1, $col))
    $srcRange.Copy()
    $dstRange = $mWs.Range($mWs.Cells($lastRow + 1, $col), $mWs.Cells($lastRow + $n, $col))
    $dstRange.PasteSpecial($xlPasteValues)
  }
}

try {
  $masterWb = Open-WB $OutputPath $false
  $importWb = Open-WB $ImportPath $true

  $importSheetNames = @()
  foreach ($ws in $importWb.Worksheets) { $importSheetNames += $ws.Name }

  $log = @()

  # ---- MTC Sales Details: append rows A:CY (header row1, data from row2) ----
  if ($importSheetNames -contains "MTC Sales Details") {
    $impWs = $importWb.Worksheets.Item("MTC Sales Details")
    $impLastRow = $impWs.Cells($impWs.Rows.Count, 2).End($xlUp).Row
    $n = $impLastRow - 1
    if ($n -gt 0) {
      $mWs = $masterWb.Worksheets.Item("MTC Sales Details")
      $lastRow = $mWs.Cells($mWs.Rows.Count, 2).End($xlUp).Row
      Merge-AppendSheet $mWs $impWs $lastRow $n 103 $SALES_SEED_COLS 2
      $log += "MTC Sales Details: inserted $n rows ($($lastRow+1)-$($lastRow+$n))."
    }
  } else {
    $log += "MTC Sales Details: not in import file, skipped."
  }

  # ---- Exchange Rate: find the existing pre-built row for the month, fill C:J only ----
  if ($importSheetNames -contains "Exchange Rate") {
    $impWs = $importWb.Worksheets.Item("Exchange Rate")
    $yyyymm = "$($impWs.Cells(2,1).Value2)"
    if ($yyyymm -ne "") {
      $mWs = $masterWb.Worksheets.Item("Exchange Rate")
      $usedRange = $mWs.UsedRange
      $vals = $usedRange.Value2
      $foundRow = 0
      for ($r = 1; $r -le $usedRange.Rows.Count; $r++) {
        if ("$($vals[$r,2])" -eq $yyyymm) { $foundRow = $r; break }
      }
      if ($foundRow -gt 0) {
        $rates = $impWs.Range("B2:I2").Value2
        $mWs.Range($mWs.Cells($foundRow, 3), $mWs.Cells($foundRow, 10)).Value2 = $rates
        $log += "Exchange Rate: filled row $foundRow ($yyyymm), columns C:J."
      } else {
        $log += "Exchange Rate: no pre-built row found for $yyyymm - fill it in manually."
      }
    }
  } else {
    $log += "Exchange Rate: not in import file, skipped."
  }

  # ---- Net Selling Price: append rows (header row2, data from row3); col A Key stays a formula ----
  if ($importSheetNames -contains "Net Selling Price") {
    $impWs = $importWb.Worksheets.Item("Net Selling Price")
    $impLastRow = $impWs.Cells($impWs.Rows.Count, 2).End($xlUp).Row
    $n = $impLastRow - 1
    if ($n -gt 0) {
      $mWs = $masterWb.Worksheets.Item("Net Selling Price")
      $lastRow = $mWs.Cells($mWs.Rows.Count, 2).End($xlUp).Row
      Merge-AppendSheet $mWs $impWs $lastRow $n 44 $NSP_SEED_COLS 2
      $log += "Net Selling Price: inserted $n rows ($($lastRow+1)-$($lastRow+$n))."
    }
  } else {
    $log += "Net Selling Price: not in import file, skipped."
  }

  # ---- PO comm: append rows A:AI (header row1, data from row2); AK Rate formula extended, AJ needs manual entry ----
  if ($importSheetNames -contains "PO comm") {
    $impWs = $importWb.Worksheets.Item("PO comm")
    $impLastRow = $impWs.Cells($impWs.Rows.Count, 1).End($xlUp).Row
    $n = $impLastRow - 1
    if ($n -gt 0) {
      $mWs = $masterWb.Worksheets.Item("PO comm")
      $lastRow = $mWs.Cells($mWs.Rows.Count, 6).End($xlUp).Row
      Merge-AppendSheet $mWs $impWs $lastRow $n 37 $POCOMM_SEED_COLS 2 # 37 cols so AK's Rate formula extends too
      $log += "PO comm: inserted $n rows ($($lastRow+1)-$($lastRow+$n)). AJ (TTL COST USD) still needs manual entry."
    }
  } else {
    $log += "PO comm: not in import file, skipped."
  }

  $masterWb.Save()

  Write-Output "----"
  $log | ForEach-Object { Write-Output $_ }
  Write-Output "----"
  Write-Output "Saved: $OutputPath"
  Write-Output "Original master file was not modified: $MasterPath"
}
finally {
  if ($importWb) { $importWb.Close($false) }
  if ($masterWb) { $masterWb.Close($true) }
  $excel.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
  [System.GC]::Collect()
}
