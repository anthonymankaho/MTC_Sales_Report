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
  6. Extracts the "TABLE A"/"TABLE B" blocks from the "Domestic Orders ...
     (for Accounts).xlsx" Summary sheet (resolved by WILDCARD - the real
     filename contains EN-/EM-DASH characters that a no-BOM PS 5.1 script
     cannot hold as a literal) -> domestic.json.
  7. Extracts the HKD-rate table from "STANDARD EXCHANGE RATE-YYYY-MM.xls"
     (sheet "Exchange rate table") and appends a SYNTHETIC INR row: that
     table has no INR, so INR->HKD = (USD row HKD rate) / (monthly-avg
     USD/INR for the reporting month from the newest weekly
     FIN_Ex Rate_PSBUILD_<date>.xlsx in the Weekly Exchange Report output
     folder, selected by filename DATE regex, never LastWriteTime). On any
     miss it falls back to USD/INR = 95.50 with a loud WARNING and the
     island's Source string says FALLBACK -> fx_domestic.json.
     Both extra workbooks reuse the SAME Excel COM session, opened
     read-only and closed with Close($false).
  8. Runs build_report.js (Node) to assemble the final single-file HTML report.

.PARAMETER SourceExcelPath
  Path to the source "FIN_MTC Sales & Margin (x-y)YYYY_YYYYMMDD_R#.xlsx" file.
  If omitted, auto-detects the most recently modified file matching that
  naming pattern in the project root.

.PARAMETER OutputPath
  Path for the generated HTML file. If omitted, defaults to the fixed working
  file "FIN_MTC_Performance_Report_html_TEMP.html" in the project root - safe to
  regenerate any number of times while iterating (see -Publish below).

.PARAMETER Publish
  After building, also copies the result to the dated, period-stamped
  snapshot "FIN_MTC_Performance_Report_html_<YYYYMM>.html" (YYYYMM from the source
  file's own reporting period). Only do this at the actual business
  checkpoint - not on every test run - per the incremental-report-file-
  chaining pattern: one continuously-overwritten TEMP working file, dated
  snapshots only when finalizing.

.PARAMETER KeepWork
  Keep the _work JSON-islands folder after the build instead of deleting it
  (default OFF preserves the previous delete behavior). Lets a smoke run
  inspect the 11 islands without re-running the whole extraction.
#>
param(
  [string]$SourceExcelPath,
  [string]$OutputPath,
  [switch]$Publish,
  [switch]$KeepWork
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
  $OutputPath = Join-Path $projectDir "FIN_MTC_Performance_Report_html_TEMP.html"
}
$publishedPath = Join-Path $projectDir "FIN_MTC_Performance_Report_html_$yyyymm.html"
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

# Returns PSCustomObject{Value,Text,Name} holding the monthly-average USD/INR
# for $monthAbbr read from the "FX Dashboard" sheet of the newest weekly
# FIN_Ex Rate_PSBUILD_<YYYYMMDD>[_RV<n>].xlsx, or $null (after a Write-Host
# WARNING) when the folder/file/sheet/header/period-row is unresolvable - the
# caller then uses the documented 95.50 FALLBACK. The file is picked by
# FILENAME DATE regex, never LastWriteTime (incremental-report-file-chaining
# lesson; TEMP/prev files are excluded by the regex). The G2 "Period"-labelled
# block is located via Cells.Find("USD/INR") plus a "Period" header cell to
# its left (the all-currencies "Month" table also has a USD/INR column and
# must not be picked), then the row whose period label equals $monthAbbr is
# read. .Text (the displayed 2-dp value the user confirmed as 95.50) is
# parsed, not Value2's unrounded 95.5022011.
function Get-WeeklyInrAvg($excel, $monthAbbr) {
  if (-not (Test-Path $WEEKLY_FX_OUTPUT_DIR)) {
    Write-Host "  WARNING: weekly FX output folder not found: $WEEKLY_FX_OUTPUT_DIR"
    return $null
  }
  $psbuildPattern = '^FIN_Ex Rate_PSBUILD_(\d{8})((_RV\d+)?)\.xlsx$'
  $scored = Get-ChildItem -Path $WEEKLY_FX_OUTPUT_DIR -Filter "FIN_Ex Rate_PSBUILD_*.xlsx" -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match $psbuildPattern } | ForEach-Object {
      $null = $_.Name -match $psbuildPattern
      $dateKey = $Matches[1]
      $revKey = 0
      if ($Matches[2]) { $revKey = [int]($Matches[2].Replace("_RV", "")) }
      [PSCustomObject]@{ FullName = $_.FullName; Name = $_.Name; DateKey = $dateKey; RevKey = $revKey }
    }
  if (-not $scored) {
    Write-Host "  WARNING: no date-stamped FIN_Ex Rate_PSBUILD_YYYYMMDD.xlsx files in $WEEKLY_FX_OUTPUT_DIR"
    return $null
  }
  $pick = $scored | Sort-Object DateKey, RevKey -Descending | Select-Object -First 1
  $wbW = $null
  $cell = $null
  try {
    Write-Host "  INR source (newest by filename date): $($pick.Name)"
    $wbW = $excel.Workbooks.Open($pick.FullName, [Type]::Missing, $true)
    $wsD = $wbW.Worksheets.Item("FX Dashboard")
    $hdrD = $null
    $periodCol = 0
    $cell = $wsD.Cells.Find("USD/INR")
    $tries = 0
    while ($cell -and $tries -lt 6) {
      $tries = $tries + 1
      for ($c = 1; $c -lt $cell.Column; $c++) {
        $labH = $wsD.Cells($cell.Row, $c).Text.Trim()
        if ($labH -eq "Period") { $hdrD = $cell; $periodCol = $c; break }
      }
      if ($hdrD) { break }
      $cell = $wsD.Cells.Find("USD/INR", $cell)
    }
    if (-not $hdrD) {
      Write-Host "  WARNING: 'USD/INR' header with a 'Period' label column not found on FX Dashboard of $($pick.Name)"
      return $null
    }
    $r = $hdrD.Row + 1
    $valText = ""
    $val = $null
    while ($wsD.Cells($r, $periodCol).Text.Trim() -ne "") {
      $lab = $wsD.Cells($r, $periodCol).Text.Trim()
      if ($lab -eq $monthAbbr) {
        $valText = $wsD.Cells($r, $hdrD.Column).Text.Trim()
        if ($valText -match '^[\d,]+(\.\d+)?$') {
          $val = [double]::Parse($valText.Replace(",", ""), [System.Globalization.CultureInfo]::InvariantCulture)
        }
        break
      }
      $r = $r + 1
    }
    if (-not $val) {
      Write-Host "  WARNING: 'USD/INR' value for '$monthAbbr' not resolvable on FX Dashboard of $($pick.Name) (got '$valText')"
      return $null
    }
    $result = [PSCustomObject]@{ Value = $val; Text = $valText; Name = $pick.Name }
    return $result
  } catch {
    Write-Host "  WARNING: could not read FX Dashboard from $($pick.Name): $($_.Exception.Message)"
    return $null
  } finally {
    if ($cell) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($cell) | Out-Null }
    if ($wbW) {
      $wbW.Close($false)
      [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wbW) | Out-Null
    }
  }
}

# Extracts the HKD-rate table from the standard FX workbook's "Exchange rate
# table" sheet: header located via Cells.Find("CURRENCY") (never a fixed row),
# data rows scanned while col A is non-blank (this stops at the blank row
# before the MHK/MMC/MTC/MXL signature junk). Provenance captured via Find
# too: the "DATE :" label's right-hand cell (.Text -> "26-Aug-2026") and the
# "EXCHANGE RATE FOR THE MONTH" title cell. The sheet has NO INR row, but the
# Domestic page needs INR->HKD, so a synthetic INR row is appended from
# Get-WeeklyInrAvg: HkdRate = (USD row HkdRate, read from this table, not a
# literal) / (monthly-avg USD/INR). Every row carries a Source tag
# ("standard" / "weekly-avg ..." / "FALLBACK ...") so the page can show it.
function Export-FxHkdTable($excel, $ws, $fileName, $monthAbbr, $year) {
  $hdrCell = $ws.Cells.Find("CURRENCY")
  if (-not $hdrCell) { throw "Could not find the 'CURRENCY' header in sheet '$($ws.Name)' of $((Split-Path $FX_STANDARD_SOURCE -Leaf))" }
  $headerRow = $hdrCell.Row + 1

  $asOf = ""
  $dateCell = $ws.Cells.Find("DATE :")
  if ($dateCell) {
    $dateValCol = $dateCell.Column + 1
    $asOf = $ws.Cells($dateCell.Row, $dateValCol).Text.Trim()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($dateCell) | Out-Null
  } else {
    Write-Host "  WARNING: no 'DATE :' provenance cell found in FX sheet"
  }

  $title = ""
  $titleCell = $ws.Cells.Find("EXCHANGE RATE FOR THE MONTH")
  if ($titleCell) {
    $title = $titleCell.Text.Trim()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($titleCell) | Out-Null
  } else {
    Write-Host "  WARNING: no 'EXCHANGE RATE FOR THE MONTH' title cell found in FX sheet"
  }

  $sb = New-Object System.Text.StringBuilder
  $leaf = $ws.Parent.Name
  [void]$sb.Append('{"file":"' + (Escape-JsonStr $leaf) + '","asOf":"' + (Escape-JsonStr $asOf) + '","title":"' + (Escape-JsonStr $title) + '","header":["Currency","Code","Unit","HKD Rate","Source"],"rows":[')
  $usdHkd = $null
  $r = $headerRow
  $first = $true
  $rowCount = 0
  while ($ws.Cells($r, 1).Text.Trim() -ne "") {
    if (-not $first) { [void]$sb.Append(',') }
    $first = $false
    $name = $ws.Cells($r, 1).Text.Trim()
    $code = $ws.Cells($r, 3).Text.Trim()
    $unit = $ws.Cells($r, 4).Value2
    $rate = $ws.Cells($r, 5).Value2
    if ($code -eq "USD" -and $rate -is [double]) { $usdHkd = [double]$rate }
    [void]$sb.Append('["' + (Escape-JsonStr $name) + '","' + (Escape-JsonStr $code) + '",')
    if ($null -eq $unit) { [void]$sb.Append('null') }
    elseif ($unit -is [double]) { [void]$sb.Append([string]$unit) }
    else { [void]$sb.Append('"' + (Escape-JsonStr $unit) + '"') }
    [void]$sb.Append(',')
    if ($null -eq $rate) { [void]$sb.Append('null') }
    elseif ($rate -is [double]) { [void]$sb.Append([string]$rate) }
    else { [void]$sb.Append('"' + (Escape-JsonStr $rate) + '"') }
    [void]$sb.Append(',"standard"]')
    $rowCount = $rowCount + 1
    $r = $r + 1
  }
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($hdrCell) | Out-Null
  if (-not $usdHkd) { throw "USD row / HKD rate missing from FX table - INR derivation and USD conversions would both be meaningless." }

  $avg = Get-WeeklyInrAvg $excel $monthAbbr
  if ($avg) {
    $inrHkd = $usdHkd / $avg.Value
    $inrProv = "weekly-avg " + $monthAbbr + "-" + $year + " USD/INR " + $avg.Text + " (" + $avg.Name + ")"
  } else {
    Write-Host "  WARNING: FX FALLBACK - synthetic INR row using USD/INR = 95.50 (weekly PSBUILD source unresolvable)"
    $inrHkd = $usdHkd / 95.5
    $inrProv = "FALLBACK weekly-avg USD/INR 95.50 (weekly PSBUILD source unresolvable)"
  }
  [void]$sb.Append(',["INDIAN RUPEE","INR",1,' + [string]$inrHkd + ',"' + (Escape-JsonStr $inrProv) + '"]')
  [void]$sb.Append(']}')
  Write-JsonFile $sb.ToString() $fileName
  Write-Host "  Wrote $fileName ($rowCount standard rows + synthetic INR row, INR->HKD = $inrHkd)"
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

# Extracts the two "TABLE A" / "TABLE B" blocks from the Domestic Orders
# workbook's "Summary" sheet into one island (domestic.json). Anchors located
# via Cells.Find("TABLE A") / Cells.Find("TABLE B"), then each block's header
# row is found by scanning down to the row whose col-A text is exactly
# "Channel" - never fixed rows 5/13. Data rows are collected while col A is
# non-blank, which ends each block at its caution row; those caution rows sit
# in col B on an otherwise-blank-A row ("No combined DOMESTIC SUBTOTAL - do
# not sum ...") and are captured into the island's notes[] instead of being
# hardcoded. Label columns (1-4) come via .Text; metric columns (5-8, the
# 2026 Shipped/Outstanding/Total + 2027 Outstanding figures) come via Value2
# raw numbers with nulls preserved (MCN has only a 2027 figure). Provenance
# ("file", "asOf", "meta", "sourceUrl") is likewise read via Find, and the
# workbook's own name (which contains EN-/EM-DASH characters) reaches the
# island only through $ws.Parent.Name, never as a literal in this no-BOM file.
function Export-DomesticTables($ws, $fileName) {
  $cellA = $ws.Cells.Find("TABLE A")
  if (-not $cellA) { throw "Could not find the 'TABLE A' anchor on the Domestic Orders Summary sheet" }
  $cellB = $ws.Cells.Find("TABLE B")
  if (-not $cellB) { throw "Could not find the 'TABLE B' anchor on the Domestic Orders Summary sheet" }

  $asOf = ""
  $meta = ""
  $metaCell = $ws.Cells.Find("As of")
  if ($metaCell) {
    $meta = $metaCell.Text.Trim()
    if ($meta -match 'As of (\d{1,2} [A-Za-z]{3,9} \d{4})') { $asOf = $Matches[1] }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($metaCell) | Out-Null
  } else {
    Write-Host "  WARNING: no 'As of' provenance line found on the Domestic Summary sheet"
  }

  $sourceUrl = ""
  $srcCell = $ws.Cells.Find("Source")
  if ($srcCell) {
    $srcUrlCol = $srcCell.Column + 1
    $sourceUrl = $ws.Cells($srcCell.Row, $srcUrlCol).Text.Trim()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($srcCell) | Out-Null
  }

  $anchors = @($cellA, $cellB)
  $names = @("tableA", "tableB")
  $blockJson = @("", "")
  $notes = @()
  for ($k = 0; $k -lt 2; $k++) {
    $anchor = $anchors[$k]
    $h = $anchor.Row + 1
    $guard = 0
    while ($ws.Cells($h, 1).Text.Trim() -ne "Channel") {
      $h = $h + 1
      $guard = $guard + 1
      if ($guard -gt 10) { throw "No 'Channel' header row found below the '$($names[$k])' anchor on the Domestic Summary sheet" }
    }
    $hb = New-Object System.Text.StringBuilder
    [void]$hb.Append('{"header":[')
    for ($c = 1; $c -le 8; $c++) {
      if ($c -gt 1) { [void]$hb.Append(',') }
      $ht = $ws.Cells($h, $c).Text.Trim()
      [void]$hb.Append('"' + (Escape-JsonStr $ht) + '"')
    }
    [void]$hb.Append('],"rows":[')
    $r = $h + 1
    $firstRow = $true
    $rowCount = 0
    while ($ws.Cells($r, 1).Text.Trim() -ne "") {
      if (-not $firstRow) { [void]$hb.Append(',') }
      $firstRow = $false
      [void]$hb.Append('[')
      for ($c = 1; $c -le 8; $c++) {
        if ($c -gt 1) { [void]$hb.Append(',') }
        if ($c -le 4) {
          $t = $ws.Cells($r, $c).Text.Trim()
          [void]$hb.Append('"' + (Escape-JsonStr $t) + '"')
        } else {
          $v = $ws.Cells($r, $c).Value2
          if ($null -eq $v) { [void]$hb.Append('null') }
          elseif ($v -is [double]) { [void]$hb.Append([string]$v) }
          else { [void]$hb.Append('"' + (Escape-JsonStr $v) + '"') }
        }
      }
      [void]$hb.Append(']')
      $rowCount = $rowCount + 1
      $r = $r + 1
    }
    [void]$hb.Append('],"entityCol":2,"currencyCol":4,"metricCols":[5,6,7,8]}')
    $blockJson[$k] = $hb.ToString()
    # The block-ending row has a blank col A by definition; if it also has
    # text in col B, that is the workbook's own "do not sum" caution note.
    $noteText = $ws.Cells($r, 2).Text.Trim()
    if ($noteText -ne "") { $notes += $noteText }
    Write-Host "  $($names[$k]): $rowCount rows (header row $h)"
  }
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($cellA) | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($cellB) | Out-Null

  $notesJson = ""
  for ($i = 0; $i -lt $notes.Count; $i++) {
    if ($i -gt 0) { $notesJson += "," }
    $notesJson += '"' + (Escape-JsonStr $notes[$i]) + '"'
  }
  $leaf = $ws.Parent.Name
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('{"file":"' + (Escape-JsonStr $leaf) + '","asOf":"' + (Escape-JsonStr $asOf) + '","meta":"' + (Escape-JsonStr $meta) + '","sourceUrl":"' + (Escape-JsonStr $sourceUrl) + '","notes":[' + $notesJson + '],"tableA":' + $blockJson[0] + ',"tableB":' + $blockJson[1] + '}')
  Write-JsonFile $sb.ToString() $fileName
  Write-Host "  Wrote $fileName ($($notes.Count) caution notes captured)"
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

# --- Sources for the two extra islands (domestic.json / fx_domestic.json) ---
# The Domestic workbook's real filename contains EN-DASH (U+2013) and EM-DASH
# (U+2014) characters: hardcoded as a literal in this no-BOM .ps1, Windows
# PowerShell 5.1 misdecodes it under the system codepage and would silently
# miss the file (the exact bug class already guarded for "總計" above), so it
# is resolved by wildcard instead.
$DOMESTIC_SOURCE = (Get-ChildItem -Path $projectDir -Filter "Domestic Orders*.xlsx" -ErrorAction SilentlyContinue | Select-Object -First 1).FullName
if (-not $DOMESTIC_SOURCE) { throw "Domestic Orders source workbook ('Domestic Orders*.xlsx') not found in: $projectDir" }

# FX standard-rate workbook picked by the PERIOD encoded in its filename
# (STANDARD EXCHANGE RATE-YYYY-MM.xls), preferring the one matching the FIN
# report's own $yyyymm and falling back to the newest period - never by
# LastWriteTime (same discipline as the FIN auto-detection above). A missing
# FX file is a hard throw: without it every HKD conversion would silently
# zero out, so no lenient "{}"-island degrade like the Family map.
$FX_NAME_PATTERN = '^STANDARD EXCHANGE RATE-(\d{4})-(\d{2})\.xls$'
$fxScored = Get-ChildItem -Path $projectDir -Filter "STANDARD EXCHANGE RATE-*.xls*" -ErrorAction SilentlyContinue |
  Where-Object { $_.Name -match $FX_NAME_PATTERN } | ForEach-Object {
    $null = $_.Name -match $FX_NAME_PATTERN
    $fxYear = [int]$Matches[1]
    $fxMonth = [int]$Matches[2]
    $periodKey = "{0}{1:D2}" -f $fxYear, $fxMonth
    [PSCustomObject]@{ FullName = $_.FullName; Name = $_.Name; Year = $fxYear; Month = $fxMonth; PeriodKey = $periodKey }
  }
if (-not $fxScored) { throw "No 'STANDARD EXCHANGE RATE-YYYY-MM.xls' file found in: $projectDir" }
$fxMatch = $fxScored | Where-Object { $_.PeriodKey -eq $yyyymm } | Sort-Object Year, Month -Descending | Select-Object -First 1
if ($fxMatch) {
  $FX_STANDARD_SOURCE = $fxMatch.FullName
} else {
  $FX_STANDARD_SOURCE = ($fxScored | Sort-Object Year, Month -Descending | Select-Object -First 1).FullName
}
if (-not (Test-Path $FX_STANDARD_SOURCE)) { throw "Source file not found: $FX_STANDARD_SOURCE" }

# Weekly FX project output folder - the only place a USD/INR average exists
# (the standard .xls has no INR row at all).
$WEEKLY_FX_OUTPUT_DIR = "C:\Users\User\Desktop\Weekly Exchange Report - CLAUDE\Output"
# Month abbreviation for the reporting period's END month (e.g. (1-9)2026 ->
# Sep), matching the FX Dashboard period-label rows Jan..Dec.
$MONTH_NAMES = @("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
$monthIdx = $endMonth - 1
$fxMonthAbbr = $MONTH_NAMES[$monthIdx]
Write-Host "Domestic source: $DOMESTIC_SOURCE"
Write-Host "FX source: $FX_STANDARD_SOURCE"

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

  # Domestic (7 KB) + standard FX (108 KB) opens are cheap - done here, after
  # the big FIN workbook is closed, in the SAME COM session ($excel); the INR
  # weekly lookup inside Export-FxHkdTable reuses it too. No second
  # Excel.Application anywhere. Read-only opens + Close($false) + the
  # ReleaseComObject chain mirror the Export-FamilyCategoryMap pattern.
  Write-Host "Extracting Domestic Orders TABLE A / TABLE B..."
  $wbDom = $excel.Workbooks.Open($DOMESTIC_SOURCE, [Type]::Missing, $true)
  try {
    $wsDom = $wbDom.Worksheets.Item("Summary")
    Export-DomesticTables $wsDom "domestic.json"
  } finally {
    $wbDom.Close($false)
    if ($wsDom) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wsDom) | Out-Null }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wbDom) | Out-Null
  }

  Write-Host "Extracting FX HKD-rate table (+ synthetic INR row)..."
  $wbFx = $excel.Workbooks.Open($FX_STANDARD_SOURCE, [Type]::Missing, $true)
  try {
    $wsFx = $wbFx.Worksheets.Item("Exchange rate table")
    Export-FxHkdTable $excel $wsFx "fx_domestic.json" $fxMonthAbbr $year
  } finally {
    $wbFx.Close($false)
    if ($wsFx) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wsFx) | Out-Null }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wbFx) | Out-Null
  }
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

if ($KeepWork) {
  $islandFiles = @(Get-ChildItem -Path $workDir -Filter "*.json" -ErrorAction SilentlyContinue)
  $islandCount = $islandFiles.Count
  $islandNames = ($islandFiles | ForEach-Object { $_.Name }) -join ", "
  Write-Host "KeepWork: leaving $workDir with $islandCount islands: $islandNames"
} else {
  Remove-Item $workDir -Recurse -Force
}
Write-Host "Done: $OutputPath"

if ($Publish) {
  Copy-Item -Path $OutputPath -Destination $publishedPath -Force
  Write-Host "Published: $publishedPath"
}
