<#
.SYNOPSIS
  Builds the interactive HTML version of the MTC Sales Report from the monthly
  "FIN_MTC Sales & Margin (...)....xlsx" deliverable.

.DESCRIPTION
  1. Opens the source workbook read-only via Excel COM.
  2. Extracts the full DATA sheet to JSON (drives the Dashboard tab too - computed
     client-side from these rows, not from Excel's static pivot cache, same
     philosophy as the Global Purchase Report reference this was built from).
  3. Extracts only the "Platium Buy Rate : Foreign Currency => USD" table from
     the Exchange Rate sheet (the HKD-rate table next to it is dropped) - same
     pattern as the reference project's own Exchange Rate sheet.
  4. Extracts each of the 5 "2026 v 2025 by ..." year-over-year comparison
     PivotTables (3-row header: metric name row, then 2025/2026 sub-row) by
     locating the header rows directly rather than hardcoding row numbers.
  5. Extracts the "Net Margin by Customer-Quarter" flat summary table.
  6. Runs build_report.js (Node) to assemble the final single-file HTML report.

.PARAMETER SourceExcelPath
  Path to the source "FIN_MTC Sales & Margin (x-y)YYYY_YYYYMMDD_R#.xlsx" file.
  If omitted, auto-detects the most recently modified file matching that
  naming pattern in the project root.

.PARAMETER OutputPath
  Path for the generated HTML file. If omitted, defaults to the fixed working
  file "FIN_MTC_Sales_Report_html_TEMP.html" in the project root - safe to
  regenerate any number of times while iterating (see -Publish below).

.PARAMETER Publish
  After building, also copies the result to the dated, period-stamped
  snapshot "FIN_MTC_Sales_Report_html_<YYYYMM>.html" (YYYYMM from the source
  file's own reporting period). Only do this at the actual business
  checkpoint - not on every test run - per the incremental-report-file-
  chaining pattern: one continuously-overwritten TEMP working file, dated
  snapshots only when finalizing.
#>
param(
  [string]$SourceExcelPath,
  [string]$OutputPath,
  [switch]$Publish
)

$ErrorActionPreference = "Stop"
$projectDir = Split-Path -Parent $PSScriptRoot
$scriptDir = Join-Path $projectDir "HTML_Report_Generator"

if (-not $SourceExcelPath) {
  $namePattern = '^FIN_MTC Sales & Margin \((\d+)-(\d+)\)(\d{4})_(\d{8})_R(\d+)\.xlsx$'
  $candidates = Get-ChildItem -Path $projectDir -Filter "FIN_MTC Sales & Margin (*.xlsx" -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match $namePattern }
  if (-not $candidates) {
    throw "No source report file found matching the expected naming pattern in: $projectDir"
  }
  # Pick by the reporting PERIOD encoded in the filename (end month/year, then
  # generation date, then revision number) - NOT file modification time. A
  # file merely being opened (e.g. in Excel) can bump its LastWriteTime
  # without it actually being a newer reporting period, which would silently
  # pick the wrong source file (see the incremental-report-file-chaining
  # skill for this exact failure mode).
  $scored = $candidates | ForEach-Object {
    $null = $_.Name -match $namePattern
    [PSCustomObject]@{
      File     = $_
      Year     = [int]$Matches[3]
      EndMonth = [int]$Matches[2]
      GenDate  = $Matches[4]
      Revision = [int]$Matches[5]
    }
  }
  $SourceExcelPath = ($scored | Sort-Object Year, EndMonth, GenDate, Revision -Descending | Select-Object -First 1).File.FullName
}
if (-not (Test-Path $SourceExcelPath)) {
  throw "Source file not found: $SourceExcelPath"
}
Write-Host "Source: $SourceExcelPath"

$sourceFileName = Split-Path $SourceExcelPath -Leaf
# Filename encodes the reporting period as "(startMonth-endMonth)YYYY", e.g.
# "(1-7)2026" = Jan-Jul 2026. Use the period's END month/year for the
# published snapshot's name, not the "_YYYYMMDD_" stamp (that's just when
# the source file itself was generated).
if ($sourceFileName -notmatch '\(\d+-(\d+)\)(\d{4})') {
  throw "Could not find a (start-end)YYYY reporting-period pattern in source filename: $sourceFileName"
}
$endMonth = [int]$Matches[1]
$year = $Matches[2]
$yyyymm = "{0}{1:D2}" -f $year, $endMonth

if (-not $OutputPath) {
  $OutputPath = Join-Path $projectDir "FIN_MTC_Sales_Report_html_TEMP.html"
}
$publishedPath = Join-Path $projectDir "FIN_MTC_Sales_Report_html_$yyyymm.html"
Write-Host "Output: $OutputPath"
if ($Publish) { Write-Host "Will also publish to: $publishedPath" }

$workDir = Join-Path $scriptDir "_work"
if (Test-Path $workDir) { Remove-Item $workDir -Recurse -Force }
New-Item -ItemType Directory -Path $workDir | Out-Null

$xlUp = -4162
$xlToLeft = -4159
$noBomUtf8 = New-Object System.Text.UTF8Encoding $false

function Write-JsonFile($obj, $fileName) {
  [System.IO.File]::WriteAllText((Join-Path $workDir $fileName), $obj, $noBomUtf8)
}

function Escape-JsonStr($s) {
  $s = ([string]$s).Replace('\', '\\').Replace('"', '\"')
  return [System.Text.RegularExpressions.Regex]::Replace($s, '[\x00-\x1F]', ' ')
}

# Generic "full sheet, header row 1, data from row 2" extractor - used for DATA.
function Export-FullSheet($ws, $fileName) {
  $lastRow = $ws.Cells($ws.Rows.Count, 1).End($xlUp).Row
  $lastCol = $ws.Cells(1, $ws.Columns.Count).End($xlToLeft).Column
  Write-Host "  $lastRow rows x $lastCol cols"
  $range = $ws.Range($ws.Cells(1, 1), $ws.Cells($lastRow, $lastCol))
  $vals = $range.Value2

  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('{"header":[')
  for ($c = 1; $c -le $lastCol; $c++) {
    if ($c -gt 1) { [void]$sb.Append(',') }
    [void]$sb.Append('"' + (Escape-JsonStr $vals.GetValue(1, $c)) + '"')
  }
  [void]$sb.Append('],"rows":[')
  for ($r = 2; $r -le $lastRow; $r++) {
    if ($r -gt 2) { [void]$sb.Append(',') }
    [void]$sb.Append('[')
    for ($c = 1; $c -le $lastCol; $c++) {
      if ($c -gt 1) { [void]$sb.Append(',') }
      $cellVal = $vals.GetValue($r, $c)
      if ($null -eq $cellVal) { [void]$sb.Append('null') }
      elseif ($cellVal -is [double]) { [void]$sb.Append([string]$cellVal) }
      else { [void]$sb.Append('"' + (Escape-JsonStr $cellVal) + '"') }
    }
    [void]$sb.Append(']')
    if (($r % 5000) -eq 0) { Write-Host "  ...$r rows processed" }
  }
  [void]$sb.Append(']}')
  Write-JsonFile $sb.ToString() $fileName
  Write-Host "  Wrote $fileName"
}

# Extracts one "Platium Buy Rate : Foreign Currency => USD"-style sub-table by
# its own label, same technique proven in the Global Purchase Report reference.
function Export-LabeledSubTable($ws, $labelText, $numCols, $fileName) {
  $labelCell = $ws.Cells.Find($labelText)
  if (-not $labelCell) { throw "Could not find '$labelText' label" }
  $headerRow = $labelCell.Row + 1
  $startCol = $labelCell.Column
  $endCol = $startCol + $numCols - 1

  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('{"header":[')
  for ($c = $startCol; $c -le $endCol; $c++) {
    if ($c -gt $startCol) { [void]$sb.Append(',') }
    [void]$sb.Append('"' + (Escape-JsonStr ($ws.Cells($headerRow, $c).Text.Trim())) + '"')
  }
  [void]$sb.Append('],"rows":[')
  $r = $headerRow + 1
  $first = $true
  while ($ws.Cells($r, $startCol).Text.Trim() -ne "") {
    if (-not $first) { [void]$sb.Append(',') }
    $first = $false
    [void]$sb.Append('[')
    for ($c = $startCol; $c -le $endCol; $c++) {
      if ($c -gt $startCol) { [void]$sb.Append(',') }
      if ($c -eq $startCol) {
        [void]$sb.Append('"' + (Escape-JsonStr ($ws.Cells($r, $c).Text.Trim())) + '"')
      } else {
        $cellVal = $ws.Cells($r, $c).Value2
        if ($null -eq $cellVal) { [void]$sb.Append('null') }
        elseif ($cellVal -is [double]) { [void]$sb.Append([string]$cellVal) }
        else { [void]$sb.Append('"' + (Escape-JsonStr $cellVal) + '"') }
      }
    }
    [void]$sb.Append(']')
    $r++
  }
  [void]$sb.Append(']}')
  Write-JsonFile $sb.ToString() $fileName
  Write-Host "  $($r - $headerRow - 1) rows x $numCols cols -> $fileName"
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($labelCell) | Out-Null
}

# Extracts a "2026 v 2025 by ..." YoY comparison PivotTable: $labelColCount plain
# label columns (e.g. Cust_name, Country of origin, Part_id), followed by metric
# pairs whose name lives in the row above the "2025"/"2026" sub-header (forward-
# filled across the merged pair of columns it spans). Reads via .Text throughout
# so the extracted values match Excel's own displayed formatting ($, %, commas)
# exactly - these are frozen snapshot pivot outputs in a dated deliverable file,
# not something that needs live recomputation.
function Export-YoYPivotSheet($ws, $labelColCount, $metricNameRow, $yearRow, $dataStartRow, $fileName, $forwardFillCols) {
  if (-not $forwardFillCols) { $forwardFillCols = @() }
  $lastCol = $ws.Cells($yearRow, $ws.Columns.Count).End($xlToLeft).Column
  $lastRow = $ws.Cells($ws.Rows.Count, 1).End($xlUp).Row
  # column 1 may have blank continuation cells (pivot outline "repeat labels off"),
  # so re-derive the true last row from the LAST label column that's actually
  # populated on its own final data row instead - use the sheet's own UsedRange
  # bottom edge, which is reliable regardless of which label column has blanks.
  $lastRow = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1
  Write-Host "  header cols 1-$lastCol, data rows $dataStartRow-$lastRow"

  # Forward-fill the metric-name row across each pair of columns it spans.
  $metricNames = @()
  $lastName = ""
  for ($c = 1; $c -le $lastCol; $c++) {
    $t = $ws.Cells($metricNameRow, $c).Text.Trim()
    if ($t -ne "") { $lastName = $t }
    $metricNames += $lastName
  }

  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('{"header":[')
  for ($c = 1; $c -le $lastCol; $c++) {
    if ($c -gt 1) { [void]$sb.Append(',') }
    if ($c -le $labelColCount) {
      $h = $ws.Cells($yearRow, $c).Text.Trim()
    } else {
      $h = ($metricNames[$c - 1] + ' ' + $ws.Cells($yearRow, $c).Text.Trim()).Trim()
    }
    [void]$sb.Append('"' + (Escape-JsonStr $h) + '"')
  }
  [void]$sb.Append('],"rows":[')
  $firstRow = $true
  # Excel's pivot outline display hides a label cell when it repeats the row
  # above (e.g. Cust_name blank on a customer's 2nd+ country-of-origin row) -
  # correct for reading the report in Excel, but leaves a flattened row that's
  # unsortable/unfilterable by that column on its own. Forward-fill just the
  # columns confirmed (by inspecting real extracted rows) to behave this way -
  # NOT every label column, since e.g. Part_id/Country can be genuinely blank
  # for a given row rather than "same as above, hidden".
  $lastLabelValue = @{}
  $rowsEmitted = 0
  for ($r = $dataStartRow; $r -le $lastRow; $r++) {
    # UsedRange can overshoot the real data by many rows (formatting-only
    # padding below the PivotTable's actual Grand Total row) - if we blindly
    # forward-filled through those, every trailing blank row would inherit
    # "Grand Total" (or whatever the real last value was) forever. Detect the
    # true end of data first: read every column's RAW text before any
    # forward-fill: a row where ALL columns are blank is padding, not data -
    # stop here rather than emitting it (or anything past it).
    $rawTexts = @()
    $anyNonBlank = $false
    for ($c = 1; $c -le $lastCol; $c++) {
      $t = $ws.Cells($r, $c).Text.Trim()
      $rawTexts += $t
      if ($t -ne "") { $anyNonBlank = $true }
    }
    if (-not $anyNonBlank) { break }
    # Excel's own "Grand Total" summary row is a real row in the sheet, but
    # mixed into a sortable/filterable table it lands wherever "Grand Total"
    # falls alphabetically (confusing, and not a real record) - drop it from
    # the row data; the totals it shows are still visible in Excel itself.
    if ($rawTexts[0] -eq "Grand Total") { continue }

    if (-not $firstRow) { [void]$sb.Append(',') }
    $firstRow = $false
    [void]$sb.Append('[')
    for ($c = 1; $c -le $lastCol; $c++) {
      if ($c -gt 1) { [void]$sb.Append(',') }
      $text = $rawTexts[$c - 1]
      if ($c -le $labelColCount -and $forwardFillCols -contains $c) {
        if ($text -ne "") { $lastLabelValue[$c] = $text }
        elseif ($lastLabelValue.ContainsKey($c)) { $text = $lastLabelValue[$c] }
      }
      [void]$sb.Append('"' + (Escape-JsonStr $text) + '"')
    }
    [void]$sb.Append(']')
    $rowsEmitted++
  }
  [void]$sb.Append(']}')
  Write-JsonFile $sb.ToString() $fileName
  Write-Host "  Wrote $fileName ($rowsEmitted data rows)"
}

# "Net Margin by Customer-Quarter": simple flat header + data table. This
# sheet's rows are all implicitly scoped to one Year (a page filter, e.g.
# "Year: 2026" a few rows above the header) rather than having per-column
# year-paired metrics like the YoY sheets - captured here (via $yearCell) so
# the drilldown link on the web version knows which Year to filter to.
function Export-FlatTable($ws, $headerRow, $dataStartRow, $lastCol, $fileName, $yearCell) {
  $lastRow = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1
  $year = if ($yearCell) { $ws.Range($yearCell).Text.Trim() } else { "" }
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('{"year":"' + (Escape-JsonStr $year) + '","header":[')
  for ($c = 1; $c -le $lastCol; $c++) {
    if ($c -gt 1) { [void]$sb.Append(',') }
    [void]$sb.Append('"' + (Escape-JsonStr ($ws.Cells($headerRow, $c).Text.Trim())) + '"')
  }
  [void]$sb.Append('],"rows":[')
  $firstRow = $true
  # Excel's own total row is a real row in the sheet, but mixed into a
  # sortable table it lands wherever its (largest) value falls - which put it
  # straight at the top once this page defaulted to a high-to-low sort. Its
  # label here is the Chinese "總計"; built from codepoints rather than typed
  # as a literal because this .ps1 has no BOM, and Windows PowerShell 5.1 can
  # silently misdecode a literal non-ASCII character in an unBOM'd script
  # using the system codepage, leaving the comparison never matching at all.
  $totalLabel = [string]([char]0x7E3D + [char]0x8A08)
  for ($r = $dataStartRow; $r -le $lastRow; $r++) {
    $rowLabel = $ws.Cells($r, 1).Text.Trim()
    if ($rowLabel -eq "") { continue }
    if ($rowLabel -eq $totalLabel -or $rowLabel -match '(?i)^grand total\b') { continue }
    if (-not $firstRow) { [void]$sb.Append(',') }
    $firstRow = $false
    [void]$sb.Append('[')
    for ($c = 1; $c -le $lastCol; $c++) {
      if ($c -gt 1) { [void]$sb.Append(',') }
      [void]$sb.Append('"' + (Escape-JsonStr ($ws.Cells($r, $c).Text.Trim())) + '"')
    }
    [void]$sb.Append(']')
  }
  [void]$sb.Append(']}')
  Write-JsonFile $sb.ToString() $fileName
  Write-Host "  Wrote $fileName (year=$year)"
}

# Family1 -> Family Category (brand) lookup, built from a separate reference
# workbook (not the monthly report file) at Peggy's request - "YoY: Prod
# Family" is grouped by the more granular Family1, but she wants each row's
# broader brand (Family Category, e.g. "PRESTIGE") shown and sortable as its
# own main column, so the table needs both. Many Family1 values map to the
# same Family Category, never the reverse, so a flat last-write-wins
# dictionary keyed by Family1 is enough - no aggregation needed here.
$FAMILY_CATEGORY_SOURCE = [System.IO.Path]::GetFullPath((Join-Path $projectDir "Sales-Data-MTC-Mthly Rate - UPDATED_20261103_CLEANED_95COL.xlsx"))
function Export-FamilyCategoryMap($excel, $fileName) {
  if (-not (Test-Path $FAMILY_CATEGORY_SOURCE)) {
    Write-Host "  WARNING: Family Category source not found at $FAMILY_CATEGORY_SOURCE - skipping (Prod Family tab will show blank brand column)"
    Write-JsonFile '{}' $fileName
    return
  }
  $wb2 = $excel.Workbooks.Open($FAMILY_CATEGORY_SOURCE, [Type]::Missing, $true)
  try {
    $ws2 = $wb2.Worksheets.Item("MTC Sales Details")
    try { $sourceTable = $ws2.ListObjects.Item("MTC_Sales_Details") }
    catch { throw "Required table 'MTC_Sales_Details' not found in $FAMILY_CATEGORY_SOURCE." }
    try {
      $familyColumn = $sourceTable.ListColumns.Item("Family1")
      $categoryColumn = $sourceTable.ListColumns.Item("Family Category")
    } catch {
      throw "Required headers 'Family1' and/or 'Family Category' not found in table 'MTC_Sales_Details' in $FAMILY_CATEGORY_SOURCE."
    }
    $rowCount = $sourceTable.ListRows.Count
    $map = [ordered]@{}
    for ($r = 1; $r -le $rowCount; $r++) {
      $fam1 = $familyColumn.DataBodyRange.Cells($r, 1).Text.Trim()
      if ($fam1 -eq "") { continue }
      $cat = $categoryColumn.DataBodyRange.Cells($r, 1).Text.Trim()
      if ($cat -ne "") { $map[$fam1] = $cat }
    }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('{')
    $first = $true
    foreach ($k in $map.Keys) {
      if (-not $first) { [void]$sb.Append(',') }
      $first = $false
      [void]$sb.Append('"' + (Escape-JsonStr $k) + '":"' + (Escape-JsonStr $map[$k]) + '"')
    }
    [void]$sb.Append('}')
    Write-JsonFile $sb.ToString() $fileName
    Write-Host "  $($map.Count) Family1 -> Family Category mappings -> $fileName"
  } finally {
    $wb2.Close($false)
    if ($categoryColumn) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($categoryColumn) | Out-Null }
    if ($familyColumn) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($familyColumn) | Out-Null }
    if ($sourceTable) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sourceTable) | Out-Null }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws2) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb2) | Out-Null
  }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  Write-Host "Extracting Family1 -> Family Category map..."
  Export-FamilyCategoryMap $excel "family_category_map.json"

  $wb = $excel.Workbooks.Open($SourceExcelPath, [Type]::Missing, $true)

  Write-Host "Extracting DATA sheet..."
  Export-FullSheet ($wb.Worksheets.Item("DATA")) "data_sheet.json"

  Write-Host "Extracting Exchange Rate (Platium Buy Rate USD) table..."
  Export-LabeledSubTable ($wb.Worksheets.Item("Exchange Rate")) "Platium Buy Rate : Foreign Currency => USD" 11 "exrate_usd.json"

  Write-Host "Extracting 2026 v 2025 by Customers..."
  Export-YoYPivotSheet ($wb.Worksheets.Item("2026 v 2025 by Customers")) 3 7 8 9 "yoy_customers.json" @(1)

  Write-Host "Extracting 2026 v 2025 by Fabrication..."
  Export-YoYPivotSheet ($wb.Worksheets.Item("2026 v 2025 by Fabrication")) 4 7 8 9 "yoy_fabrication.json" @(1,2)

  Write-Host "Extracting 2026 v 2025 by Prod Family..."
  Export-YoYPivotSheet ($wb.Worksheets.Item("2026 v 2025 by Prod Family")) 1 7 8 9 "yoy_prodfamily.json"

  Write-Host "Extracting 2026 v 2025 by Part ID..."
  Export-YoYPivotSheet ($wb.Worksheets.Item("2026 v 2025 by Part ID")) 2 7 8 9 "yoy_partid.json" @(1)

  Write-Host "Extracting 2026 v 2025 by NET NET MARGIN..."
  Export-YoYPivotSheet ($wb.Worksheets.Item("2026 v 2025 by NET NET MARGIN")) 3 7 8 9 "yoy_netnetmargin.json" @(1)

  Write-Host "Extracting Net Margin by Customer-Quarter..."
  Export-FlatTable ($wb.Worksheets.Item("Net Margin by Customer-Quarter")) 6 7 6 "net_margin_by_customer.json" "B4"

  $wb.Close($false)
} finally {
  if ($wb) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  Remove-Variable wb, excel -ErrorAction SilentlyContinue
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

Write-Host "Building combined HTML report..."
& node (Join-Path $scriptDir "build_report.js") $workDir $OutputPath
if ($LASTEXITCODE -ne 0) {
  throw "build_report.js failed with exit code $LASTEXITCODE"
}

Remove-Item $workDir -Recurse -Force
Write-Host "Done: $OutputPath"

if ($Publish) {
  Copy-Item -Path $OutputPath -Destination $publishedPath -Force
  Write-Host "Published: $publishedPath"
}
