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

  Before each FillDown, every non-seed column's template-row cell is checked to
  confirm it's actually a formula (starts with "="). If a column that should be
  formula-driven has silently lost its formula (e.g. from a manual paste-as-values
  elsewhere), FillDown would otherwise just copy that frozen static value into
  every new row forever without any warning - this is exactly what happened to
  "MTC Sales Details" column N (Cust_Countries), which spent several months of
  Combine runs stuck on one hardcoded country before being caught and fixed
  manually. A short allowlist of columns already confirmed as legitimately
  static/manual (2026-08 audit) is excluded from the warning to avoid noise.

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
$xlToLeft = -4159
$xlPasteValues = -4163

$SALES_INPUT_HEADERS = @(
  "Invoice", "Invm_date", "Ship_Code", "Ship_to", "Cust_id", "Cust_name", "Region",
  "Part_id", "Part_desc2", "SKU_qty", "Unit_Price_Inv", "curr", "VEND_ID", "VEND_NAME",
  "PO_CURR", "LST_RCP_PR (ORI CUR)", "SO_NO", "PO_NO", "Qtr"
)
$SALES_IGNORED_HEADERS = @(
  "Part_desc1", "no of pans", "Ttl  Pans", "Ttl_Commission1_Inv", "Ttl_Commission2_Inv",
  "Ttl_Direct Commission_Inv", "Ttl_Defective_Allowance_Inv", "Ttl_Return_Rebate_Inv",
  "Ttl_Disc_Rebate_Inv", "Ttl_Royalty_Inv", "Labor_USD", "Matl_brdn_USD", "Labor_brdn_USD",
  "Subcon_cost_USD", "Prime_USD", "Dupont Rebate (USD/unit)", "Ttl Dupont Rebate (USD)",
  "Alum Rebate (USD/unit)", "Ttl Alum Rebate (USD)", "Prime Cost After Rebates (USD)",
  "Net Cont After Rebates (USD)", "Net Cont After Rebates (%)", "Database", "yyyy-mm2",
  "Ttl PO USD", "TTL PO USD3", "TTL PO GROSS CONT USD", "TTL PO NET CONT USD",
  "TTL PO NET CONT %", "unit prime usd", "unit po usd", "Check QTY", "Invoice_date"
)

$NSP_SEED_COLS = @(2..35) + @(37)
$POCOMM_SEED_COLS = @(1..35)

$NSP_EXPECTED_STATIC = @(36)                       # CHECK - confirmed always-static (0) since earliest rows, not a broken formula
$POCOMM_EXPECTED_STATIC = @(36)                    # AJ (TTL COST USD) - already documented as needing manual entry

if (-not (Test-Path $MasterPath)) { throw "Master file not found: $MasterPath" }
if (-not (Test-Path $ImportPath)) { throw "Import file not found: $ImportPath" }

# Excel COM's Workbooks.Open() does not reliably resolve a relative path
# against the calling PowerShell process's working directory - it falls back
# to Excel's own internal "current folder" state instead (often wherever
# Excel last had a file open, which has nothing to do with this script's
# -OutputPath argument as typed). A bare filename like "Foo - UPDATED.xlsx"
# can then fail to open with a generic "file not found" even though the file
# is sitting right there, since Excel looked in the wrong folder for it.
# Resolving every path to absolute here - before Copy-Item, before any COM
# call - removes that ambiguity regardless of whether the caller passed a
# relative or absolute path.
$MasterPath = (Resolve-Path $MasterPath).Path
$ImportPath = (Resolve-Path $ImportPath).Path
if (-not [System.IO.Path]::IsPathRooted($OutputPath)) {
  $OutputPath = Join-Path (Get-Location).Path $OutputPath
}

Write-Output "Copying master file to output path (original is never modified)..."
Copy-Item $MasterPath $OutputPath -Force

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
$excel.ScreenUpdating = $false
$excel.EnableEvents = $false
$saveCompleted = $false

function Open-WB($path, $readOnly) {
  # A freshly Copy-Item'd multi-ten-MB file on a OneDrive-synced folder can
  # stay briefly locked while OneDrive notices the change and scans/uploads
  # it - 3 attempts at 1 second apart (3 seconds total) turned out not to be
  # a long enough window and failed intermittently even though the exact
  # same Open() call against the exact same file succeeded moments later
  # when retried by hand. 10 attempts at 2 seconds apart (up to 20 seconds)
  # comfortably covers that. The real COM exception is also surfaced now
  # instead of silently swallowed - the generic "could not open" message
  # alone gave no way to tell a lock from any other failure.
  $wb = $null
  $lastError = $null
  for ($attempt = 1; $attempt -le 10 -and $wb -eq $null; $attempt++) {
    try { $wb = $excel.Workbooks.Open($path, [Type]::Missing, $readOnly) }
    catch { $lastError = $_; Start-Sleep -Milliseconds 2000 }
  }
  if ($wb -eq $null) { throw "Could not open $path after 10 attempts. Last error: $lastError" }
  return $wb
}

# Checks that every non-seed, non-allowlisted column in the template row is still a
# real formula before FillDown runs. A column that quietly lost its formula (e.g. a
# one-off manual paste-as-values) would otherwise have that frozen static value
# copied into every new row from then on with no indication anything is wrong.
function Test-TemplateFormulas($mWs, $sheetName, $headerRow, $lastRow, $numCols, $seedCols, $expectedStaticCols) {
  $warnings = @()
  for ($col = 1; $col -le $numCols; $col++) {
    if ($seedCols -contains $col) { continue }
    if ($expectedStaticCols -contains $col) { continue }
    $cell = $mWs.Cells($lastRow, $col)
    $formula = "$($cell.Formula)"
    if (-not $formula.StartsWith("=")) {
      $header = "$($mWs.Cells($headerRow, $col).Value2)"
      $warnings += "[$sheetName] Col $col ($header): template row $lastRow has no formula (value=[$($cell.Value2)]) - FillDown will copy this frozen value into every new row unchanged."
    }
  }
  return $warnings
}

# Extends formulas from the template row down through the new rows, then overwrites
# each seed column with the real per-row values copied from the import sheet.
function Merge-AppendSheet($mWs, $impWs, $lastRow, $n, $numCols, $seedCols, $impFirstDataRow) {
  $fillRange = $mWs.Range($mWs.Cells($lastRow, 1), $mWs.Cells($lastRow + $n, $numCols))
  $fillRange.FillDown() | Out-Null

  foreach ($col in $seedCols) {
    $srcRange = $impWs.Range($impWs.Cells($impFirstDataRow, $col), $impWs.Cells($impFirstDataRow + $n - 1, $col))
    $srcRange.Copy() | Out-Null
    $dstRange = $mWs.Range($mWs.Cells($lastRow + 1, $col), $mWs.Cells($lastRow + $n, $col))
    $dstRange.PasteSpecial($xlPasteValues) | Out-Null
  }
}

# Appends MTC Sales Details by header. Only the 19 approved input fields are written;
# all other table columns are preserved as calculated columns.
function Merge-SalesDetailsByHeader($mWs, $impWs) {
  $table = $mWs.ListObjects.Item("MTC_Sales_Details")
  if ($null -eq $table) { throw "Required table 'MTC_Sales_Details' was not found on MTC Sales Details." }

  $masterHeaders = @{}
  for ($c = 1; $c -le $table.ListColumns.Count; $c++) {
    $header = ([string]$table.ListColumns.Item($c).Name).Trim()
    if ([string]::IsNullOrWhiteSpace($header)) { throw "Blank header found in MTC_Sales_Details column $c." }
    if ($masterHeaders.ContainsKey($header)) { throw "Duplicate table header '$header' in MTC_Sales_Details." }
    $masterHeaders[$header] = $c
  }

  $missingMasterInputs = @($SALES_INPUT_HEADERS | Where-Object { -not $masterHeaders.ContainsKey($_) })
  if ($missingMasterInputs.Count) { throw "Required table headers missing from MTC_Sales_Details: $($missingMasterInputs -join ', ')" }

  $importLastCol = $impWs.Cells(1, $impWs.Columns.Count).End($script:xlToLeft).Column
  $importHeaders = @{}
  $unmatchedHeaders = @()
  $duplicateImportHeaders = @()
  for ($c = 1; $c -le $importLastCol; $c++) {
    $header = ([string]$impWs.Cells(1, $c).Value2).Trim()
    if ([string]::IsNullOrWhiteSpace($header)) { continue }
    if ($importHeaders.ContainsKey($header)) { $duplicateImportHeaders += $header; continue }
    $importHeaders[$header] = $c
    if (-not $masterHeaders.ContainsKey($header) -and $header -notin $SALES_IGNORED_HEADERS) { $unmatchedHeaders += $header }
  }

  $missingRequired = @($SALES_INPUT_HEADERS | Where-Object { -not $importHeaders.ContainsKey($_) })
  $problems = @()
  if ($duplicateImportHeaders.Count) { $problems += "Duplicate import headers: $($duplicateImportHeaders -join ', ')" }
  if ($unmatchedHeaders.Count) { $problems += "Import headers neither matched nor approved as ignored: $($unmatchedHeaders -join ', ')" }
  if ($missingRequired.Count) { $problems += "Required input headers missing from import: $($missingRequired -join ', ')" }
  if ($problems.Count) { throw "MTC Sales Details schema validation failed:`n - $($problems -join "`n - ")" }

  $matchedHeaderCount = @($importHeaders.Keys | Where-Object { $masterHeaders.ContainsKey($_) }).Count
  $ignoredPresent = @($SALES_IGNORED_HEADERS | Where-Object { $importHeaders.ContainsKey($_) })
  $importInvoiceCol = $importHeaders["Invoice"]
  $importLastRow = $impWs.Cells($impWs.Rows.Count, $importInvoiceCol).End($script:xlUp).Row
  $n = [Math]::Max(0, $importLastRow - 1)
  if ($n -eq 0) {
    return @("MTC Sales Details: no import data rows; skipped.", "Matched headers: $matchedHeaderCount; ignored headers present: $($ignoredPresent.Count) of $($SALES_IGNORED_HEADERS.Count); input columns written: 0 of $($SALES_INPUT_HEADERS.Count).", "Ignored-value warning count: 0.")
  }

  # The 33 legacy columns are deliberately ignored. Invoice_date is a confirmed
  # duplicate of Invm_date; warn only if a future import disagrees with its kept field.
  $warningLines = @()
  $warningCount = 0
  foreach ($header in $ignoredPresent) {
    $sourceCol = $importHeaders[$header]
    $warningRows = @()
    for ($r = 2; $r -le $importLastRow; $r++) {
      $value = $impWs.Cells($r, $sourceCol).Value2
      $meaningful = $false
      if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
        $number = 0.0
        $isNumeric = [double]::TryParse([string]$value, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)
        $meaningful = -not ($isNumeric -and $number -eq 0)
      }

      if ($header -eq "Invoice_date") {
        $keptValue = $impWs.Cells($r, $importHeaders["Invm_date"]).Value2
        if ([string]$value -eq [string]$keptValue) { $meaningful = $false }
      }
      if ($meaningful) {
        $warningCount++
        $warningRows += "row $r = [$value]"
      }
    }
    if ($warningRows.Count) { $warningLines += "WARNING ignored column '$header' has non-blank/non-zero values: $($warningRows -join '; ')" }
  }

  $lastRow = $table.Range.Row + $table.Range.Rows.Count - 1
  $firstNewRow = $lastRow + 1
  $lastCol = $table.ListColumns.Count

  # Copy each formula column's last available formula into the new table rows.
  # This preserves calculated columns while avoiding any import writes to them.
  $formulaColumnCount = 0
  for ($c = 1; $c -le $lastCol; $c++) {
    $header = [string]$table.ListColumns.Item($c).Name
    if ($header -in $SALES_INPUT_HEADERS) { continue }
    $formulaTemplate = ""
    $templateRow = 0
    for ($r = $lastRow; $r -ge $table.DataBodyRange.Row; $r--) {
      $candidate = [string]$mWs.Cells($r, $c).Formula
      if ($candidate.StartsWith("=")) { $templateRow = $r; $formulaTemplate = [string]$mWs.Cells($r, $c).FormulaR1C1; break }
    }
    if (-not $formulaTemplate) { throw "Expected formula column '$header' has no formula in the existing table." }
    $fillRange = $mWs.Range($mWs.Cells($templateRow, $c), $mWs.Cells($firstNewRow + $n - 1, $c))
    $fillRange.FillDown() | Out-Null
    $formulaColumnCount++
  }

  foreach ($header in $SALES_INPUT_HEADERS) {
    $sourceCol = $importHeaders[$header]
    $targetCol = $masterHeaders[$header]
    $sourceRange = $impWs.Range($impWs.Cells(2, $sourceCol), $impWs.Cells($importLastRow, $sourceCol))
    $targetRange = $mWs.Range($mWs.Cells($firstNewRow, $targetCol), $mWs.Cells($firstNewRow + $n - 1, $targetCol))
    $targetRange.Value2 = $sourceRange.Value2
  }

  $table.Resize($mWs.Range($mWs.Cells($table.Range.Row, $table.Range.Column), $mWs.Cells($lastRow + $n, $table.Range.Column + $lastCol - 1))) | Out-Null
  $logLines = @(
    "MTC Sales Details: appended $n rows ($firstNewRow-$($firstNewRow + $n - 1)) by header; $formulaColumnCount formula columns preserved.",
    "Matched headers: $matchedHeaderCount; ignored headers present: $($ignoredPresent.Count) of $($SALES_IGNORED_HEADERS.Count); input columns written: $($SALES_INPUT_HEADERS.Count) of $($SALES_INPUT_HEADERS.Count).",
    "Ignored-value warning count: $warningCount."
  )
  return $logLines + $warningLines
}

try {
  $masterWb = Open-WB $OutputPath $false
  $importWb = Open-WB $ImportPath $true

  $importSheetNames = @()
  foreach ($ws in $importWb.Worksheets) { $importSheetNames += $ws.Name }

  $log = @()
  $warnings = @()

  # ---- MTC Sales Details: append via table headers; formula columns are preserved ----
  if ($importSheetNames -contains "MTC Sales Details") {
    $impWs = $importWb.Worksheets.Item("MTC Sales Details")
    $mWs = $masterWb.Worksheets.Item("MTC Sales Details")
    $log += Merge-SalesDetailsByHeader $mWs $impWs
  } else {
    throw "Required worksheet 'MTC Sales Details' is missing from the import file."
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
      $warnings += Test-TemplateFormulas $mWs "Net Selling Price" 2 $lastRow 44 $NSP_SEED_COLS $NSP_EXPECTED_STATIC
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
      $warnings += Test-TemplateFormulas $mWs "PO comm" 1 $lastRow 37 $POCOMM_SEED_COLS $POCOMM_EXPECTED_STATIC
      Merge-AppendSheet $mWs $impWs $lastRow $n 37 $POCOMM_SEED_COLS 2 # 37 cols so AK's Rate formula extends too
      $log += "PO comm: inserted $n rows ($($lastRow+1)-$($lastRow+$n)). AJ (TTL COST USD) still needs manual entry."
    }
  } else {
    $log += "PO comm: not in import file, skipped."
  }

  $excel.CalculateFullRebuild()
  $pivotRefreshCount = 0
  foreach ($ws in $masterWb.Worksheets) {
    foreach ($pivot in $ws.PivotTables()) {
      $null = $pivot.RefreshTable()
      $pivotRefreshCount++
    }
  }
  $log += "PivotTables refreshed: $pivotRefreshCount."

  $masterWb.Save()
  $saveCompleted = $true

  Write-Output "----"
  $log | ForEach-Object { Write-Output $_ }
  Write-Output "----"
  Write-Output "Saved: $OutputPath"
  Write-Output "Original master file was not modified: $MasterPath"

  if ($warnings.Count -gt 0) {
    Write-Output "----"
    Write-Output "*** WARNING: $($warnings.Count) column(s) lost their formula and were just FillDown-copied as static values - check these before trusting this month's data: ***"
    $warnings | ForEach-Object { Write-Output "  $_" }
  }
}
catch {
  Write-Output "FAILED at script line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
  Write-Output "Stack: $($_.ScriptStackTrace)"
  Write-Output "Command: $($_.InvocationInfo.Line.Trim())"
  throw
}
finally {
  if ($importWb) { $importWb.Close($false) }
  if ($masterWb) { $masterWb.Close($saveCompleted) }
  $excel.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
  [System.GC]::Collect()
  if (-not $saveCompleted -and (Test-Path -LiteralPath $OutputPath)) {
    Remove-Item -LiteralPath $OutputPath -Force
  }
}
