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
  regenerate any number of times while iterating.
#>
param(
  [string]$SourceExcelPath,
  [string]$OutputPath
)

$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot
$projectDir = "C:\Users\User\Desktop\MTC Sales Report Claude"

if (-not $SourceExcelPath) {
  $namePattern = '^FIN_MTC Sales & Margin \(\d+-\d+\)\d{4}_\d{8}_R\d+\.xlsx$'
  $candidates = Get-ChildItem -Path $projectDir -Filter "FIN_MTC Sales & Margin (*.xlsx" -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match $namePattern }
  if (-not $candidates) {
    throw "No source report file found matching the expected naming pattern in: $projectDir"
  }
  $SourceExcelPath = ($candidates | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
}
if (-not (Test-Path $SourceExcelPath)) {
  throw "Source file not found: $SourceExcelPath"
}
Write-Host "Source: $SourceExcelPath"

if (-not $OutputPath) {
  $OutputPath = Join-Path $projectDir "FIN_MTC_Sales_Report_html_TEMP.html"
}
Write-Host "Output: $OutputPath"

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
  for ($r = $dataStartRow; $r -le $lastRow; $r++) {
    if ($ws.Cells($r, 1).Text.Trim() -eq "") { continue }
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

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
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
