<#
  Convert-ACT001ToPurDetails.ps1

  Converts a monthly ACT001 purchase export (already imported as its own sheet
  inside the master workbook, e.g. "ACT001_0926") into the "Pur Details" sheet's
  layout and appends it - without touching any existing data, formulas,
  formatting, or workbook structure. This is a new step in the MTC Sales
  Report Import Tool's monthly cycle (see README.md / "MTC Sales Report Import
  SOP.docx"), inserted after the month's ACT001 + Pur_Combined data has already
  been loaded into the master workbook, and is deliberately a separate script
  from Combine-Into-Master.ps1 rather than folded into it: Combine-Into-Master
  merges a *staged external file* (MTC_Import_Rows.xlsx) into a *copy* of the
  master; this script instead reads two sheets that already live *inside* the
  one master workbook and appends into a third sheet of that same workbook -
  a different shape of operation with no staged import file involved.

  Reused conventions from Combine-Into-Master.ps1: absolute-path resolution
  before any Excel COM call (relative paths aren't reliably resolved by
  Workbooks.Open against the calling process's working directory - see that
  script's own comment on this), a retrying Open-WB helper, the same COM
  settings/cleanup discipline, and FillDown-from-the-first-new-row for
  formula columns rather than per-cell formula entry.

  USAGE
    .\PowerShell\Convert-ACT001ToPurDetails.ps1 `
      -WorkbookPath "C:\Users\User\Desktop\MTC Performance Report Claude\Sales-Data-MTC-Mthly Rate - UPDATED.xlsx" `
      -ACT001Sheet "ACT001_0926" `
      -ReportingMonth "2026-09"

  Add -OutputPath to write to a different file instead of modifying
  -WorkbookPath in place (Copy-Item first, same "never touch an input you
  didn't mean to" philosophy as Combine-Into-Master.ps1); omit it to modify
  -WorkbookPath directly - safe either way, since nothing is saved unless
  every validation in step 12 of the spec passes, and a failed validation
  rolls the newly-inserted rows back out and closes without saving.
#>

param(
  [Parameter(Mandatory = $true)][string]$WorkbookPath,
  [Parameter(Mandatory = $true)][string]$ACT001Sheet,
  [Parameter(Mandatory = $true)][string]$ReportingMonth,   # "yyyy-MM", e.g. "2026-09"
  [string]$PurCombinedSheet,                               # default: ACT001Sheet with "ACT001" -> "Pur_Combined"
  [string]$OutputPath                                      # default: modify -WorkbookPath in place
)

$ErrorActionPreference = "Stop"
$sw = [System.Diagnostics.Stopwatch]::StartNew()

$xlUp = -4162
$xlToLeft = -4159
$xlPasteFormats = -4122
$xlCalculationManual = -4135
$xlCalculationAutomatic = -4105

# ---------------------------------------------------------------------------
# Fixed layout facts, confirmed against the real workbook before writing this
# script (not assumed from the spec alone) - see the investigation notes in
# the accompanying summary. Column numbers are 1-indexed.
# ---------------------------------------------------------------------------
$ACT001_EXPECTED_HEADERS = @(
  "COMPANY","TRAN_DATE","VENDOR_ID","VENDOR_NAME","BUYER",
  "ORDER_NO","ORDER_LINE","PO_DATE","VINV_NO","VINV_DATE",
  "PACKSLIP","REFERENCE","STOCK_ID","TERMS","PART_NO",
  "PART_DESC","COMM_CODE","REC_QTY","PO_UOM","PO_CUR",
  "PO_COST","TOT_FOR_PO","TOT_HKD_PO","AVG_COST",
  "TOT_AVG_COST","PPV","EXCH_RATE"
)
# ACT001 source column indexes (1-indexed), by name, for readability below.
$SRC = @{
  COMPANY=1; TRAN_DATE=2; VENDOR_ID=3; VENDOR_NAME=4; BUYER=5
  ORDER_NO=6; ORDER_LINE=7; PO_DATE=8; VINV_NO=9; VINV_DATE=10
  PACKSLIP=11; REFERENCE=12; STOCK_ID=13; TERMS=14; PART_NO=15
  PART_DESC=16; COMM_CODE=17; REC_QTY=18; PO_UOM=19; PO_CUR=20
  PO_COST=21; TOT_FOR_PO=22; TOT_HKD_PO=23; AVG_COST=24
  TOT_AVG_COST=25; PPV=26; EXCH_RATE=27
}
# Pur_Combined_MMYY carries the same first 27 columns as ACT001, plus:
$PC_DEDUCT_AP_UNIT = 28       # DEDUCT FROM AP INVOICE (unit)
$PC_COMM_3RD_UNIT = 30        # Unit PURCHASE COMMISSION PAID TO 3RD PARTY
$PC_KEY_COL = 35              # PO+SKU+QTY (composite key, matches Pur Details col A)
# Pur Details target column indexes (1-indexed; header row 2, data from row 3).
$TGT = @{
  A=1; B=2; C=3; D=4; E=5; F=6; G=7; H=8; I=9; J=10; K=11; L=12; M=13; N=14
  O=15; P=16; Q=17; R=18; S=19; T=20; U=21; V=22; W=23; X=24; Y=25; Z=26
  AA=27; AB=28; AC=29; AD=30; AE=31; AF=32; AG=33; AH=34; AI=35; AJ=36
  AK=37; AL=38; AM=39; AN=40; AO=41; AP=42
}
# Exchange Rate sheet ranges (confirmed): header row 3 holds currency codes in
# O:X, month labels run N4:N39, the rate grid is O4:X39.
$FX_CURR_HEADER_ROW = 3
$FX_CURR_FIRST_COL = 15   # O
$FX_CURR_LAST_COL = 24    # X
$FX_MONTH_FIRST_ROW = 4
$FX_MONTH_LAST_ROW = 39
$FX_MONTH_COL = 14        # N

function Write-Log($msg) { Write-Output "[$([DateTime]::Now.ToString('HH:mm:ss'))] $msg" }

function Resolve-AbsolutePath($path) {
  if (-not (Test-Path $path)) { throw "File not found: $path" }
  return (Resolve-Path $path).Path
}

# Same Workbooks.Open retry pattern as Combine-Into-Master.ps1's Open-WB -
# a freshly touched large file on a OneDrive-synced folder can stay briefly
# locked, and the real COM exception is surfaced rather than swallowed.
function Open-WB($excel, $path, $readOnly) {
  $wb = $null
  $lastError = $null
  for ($attempt = 1; $attempt -le 10 -and $null -eq $wb; $attempt++) {
    try { $wb = $excel.Workbooks.Open($path, [Type]::Missing, $readOnly) }
    catch { $lastError = $_; Start-Sleep -Milliseconds 2000 }
  }
  if ($null -eq $wb) { throw "Could not open $path after 10 attempts. Last error: $lastError" }
  return $wb
}

function Get-ColLetter([int]$col) {
  $letters = ""
  while ($col -gt 0) {
    $rem = ($col - 1) % 26
    $letters = [char](65 + $rem) + $letters
    $col = [int](($col - $rem - 1) / 26)
  }
  return $letters
}

# Reads a bulk 2D Value2 range as a PowerShell array-of-arrays (1-row ranges
# come back as a scalar or 1D array from COM, not a 2D array, so this
# normalizes every case to rows[][] indexed from 0).
function Get-BulkValues($ws, $r1, $c1, $r2, $c2) {
  $raw = $ws.Range($ws.Cells($r1, $c1), $ws.Cells($r2, $c2)).Value2
  $nRows = $r2 - $r1 + 1
  $nCols = $c2 - $c1 + 1
  $out = New-Object 'object[,]' $nRows, $nCols
  if ($raw -is [Array] -and $raw.Rank -eq 2) {
    $rowBase = $raw.GetLowerBound(0)
    $colBase = $raw.GetLowerBound(1)
    for ($r = 0; $r -lt $nRows; $r++) {
      for ($c = 0; $c -lt $nCols; $c++) {
        $out.SetValue($raw.GetValue($rowBase + $r, $colBase + $c), $r, $c)
      }
    }
  } elseif ($raw -is [Array] -and $raw.Rank -eq 1) {
    $base = $raw.GetLowerBound(0)
    if ($nRows -eq 1) {
      for ($c = 0; $c -lt $nCols; $c++) { $out.SetValue($raw.GetValue($base + $c), 0, $c) }
    } else {
      for ($r = 0; $r -lt $nRows; $r++) { $out.SetValue($raw.GetValue($base + $r), $r, 0) }
    }
  } else {
    $out.SetValue($raw, 0, 0)
  }
  return ,$out
}

function To-NumOrZero($v) {
  if ($null -eq $v) { return 0 }
  if ($v -is [double] -or $v -is [int]) { return [double]$v }
  $s = "$v".Trim()
  if ($s -eq "") { return 0 }
  $n = 0.0
  if ([double]::TryParse($s, [ref]$n)) { return $n }
  return 0
}

# Mirrors Excel's own CONCAT(number-cell) text conversion for a whole number
# (e.g. 1008, not 1008.00/1,008.00) - REC_QTY is always a plain quantity in
# this data, never fractional, matching every composite key sampled.
function To-KeyText($v) {
  if ($null -eq $v) { return "" }
  if ($v -is [double]) {
    if ($v -eq [Math]::Floor($v)) { return "{0:0}" -f $v }
    return "$v"
  }
  return "$v".Trim()
}

Write-Log "Starting Convert-ACT001ToPurDetails: sheet=$ACT001Sheet month=$ReportingMonth"

$WorkbookPath = Resolve-AbsolutePath $WorkbookPath
if ($OutputPath) {
  if (-not [System.IO.Path]::IsPathRooted($OutputPath)) { $OutputPath = Join-Path (Get-Location).Path $OutputPath }
  Write-Log "Copying $WorkbookPath -> $OutputPath (input left untouched)"
  Copy-Item $WorkbookPath $OutputPath -Force
  $targetPath = $OutputPath
} else {
  $targetPath = $WorkbookPath
}

if (-not $PurCombinedSheet) {
  if ($ACT001Sheet -match '^ACT001_(.+)$') { $PurCombinedSheet = "Pur_Combined_$($Matches[1])" }
  else { throw "Cannot derive -PurCombinedSheet from '$ACT001Sheet' (expected a name like 'ACT001_0926') - pass -PurCombinedSheet explicitly." }
}
Write-Log "Derived Pur_Combined sheet: $PurCombinedSheet"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.EnableEvents = $false
$priorCalc = $null

$insertedRange = $null   # set once rows are appended, so a failed validation can roll them back
$wb = $null

try {
  $wb = Open-WB $excel $targetPath $false
  $priorCalc = $excel.Calculation
  $excel.Calculation = $xlCalculationManual

  $sheetNames = @()
  foreach ($ws in $wb.Worksheets) { $sheetNames += $ws.Name }
  foreach ($need in @($ACT001Sheet, $PurCombinedSheet, "Pur Details", "Exchange Rate")) {
    if ($sheetNames -notcontains $need) { throw "Required sheet '$need' not found in workbook. Sheets present: $($sheetNames -join ', ')" }
  }
  $hasVendorRef = $sheetNames -contains "COO-VENDOR LIST"
  if (-not $hasVendorRef) { Write-Log "WARNING: 'COO-VENDOR LIST' sheet not found - Grouped Vendor Name will fall back to history-only lookup for every vendor." }

  # ---- Step 3: validate ACT001 source headers -----------------------------
  $wsAct = $wb.Worksheets.Item($ACT001Sheet)
  $actHeaderVals = Get-BulkValues $wsAct 1 1 1 27
  $mismatches = @()
  for ($c = 0; $c -lt 27; $c++) {
    $actual = "$($actHeaderVals[0,$c])".Trim()
    $expected = $ACT001_EXPECTED_HEADERS[$c]
    if ($actual -ne $expected) { $mismatches += "col $(Get-ColLetter ($c+1)): expected '$expected', found '$actual'" }
  }
  if ($mismatches.Count -gt 0) { throw "ACT001 sheet '$ACT001Sheet' header mismatch:`n  $($mismatches -join "`n  ")" }

  $actLastRow = $wsAct.Cells($wsAct.Rows.Count, 1).End($xlUp).Row
  $n = $actLastRow - 1
  if ($n -le 0) { throw "ACT001 sheet '$ACT001Sheet' has no data rows." }
  Write-Log "ACT001 source rows: $n"

  # ---- Step 4: locate Pur Details, find the TRUE last row dynamically -----
  $wsPD = $wb.Worksheets.Item("Pur Details")
  $pdHeaderRow = 2
  $priorLastRow = $wsPD.Cells($wsPD.Rows.Count, $TGT.B).End($xlUp).Row
  if ($priorLastRow -lt $pdHeaderRow) { $priorLastRow = $pdHeaderRow }
  Write-Log "Pur Details current last row: $priorLastRow"

  # Duplicate protection #1: reporting month already present in col Z?
  if ($priorLastRow -gt $pdHeaderRow) {
    $existingZ = Get-BulkValues $wsPD ($pdHeaderRow+1) $TGT.Z $priorLastRow $TGT.Z
    for ($r = 0; $r -lt $existingZ.GetLength(0); $r++) {
      if ("$($existingZ[$r,0])".Trim() -eq $ReportingMonth) {
        Write-Log "STOP: reporting month '$ReportingMonth' already present in Pur Details (row $($pdHeaderRow+1+$r)) - nothing appended, nothing saved."
        exit 0
      }
    }
  }

  # ---- Read ACT001 data (bulk) and build this batch's composite keys ------
  $actVals = Get-BulkValues $wsAct 2 1 $actLastRow 27
  $newKeys = New-Object System.Collections.Generic.List[string]
  for ($r = 0; $r -lt $n; $r++) {
    $k = (To-KeyText $actVals[$r, ($SRC.ORDER_NO - 1)]) + (To-KeyText $actVals[$r, ($SRC.PART_NO - 1)]) + (To-KeyText $actVals[$r, ($SRC.REC_QTY - 1)])
    $newKeys.Add($k)
  }

  # Duplicate protection #2: any of this batch's composite keys already in
  # Pur Details' own key column A?
  if ($priorLastRow -gt $pdHeaderRow) {
    $existingKeys = New-Object 'System.Collections.Generic.HashSet[string]'
    $existingA = Get-BulkValues $wsPD ($pdHeaderRow+1) $TGT.A $priorLastRow $TGT.A
    for ($r = 0; $r -lt $existingA.GetLength(0); $r++) { [void]$existingKeys.Add("$($existingA[$r,0])".Trim()) }
    $dupHit = $null
    foreach ($k in $newKeys) { if ($existingKeys.Contains($k)) { $dupHit = $k; break } }
    if ($dupHit) {
      Write-Log "STOP: composite key '$dupHit' already exists in Pur Details - this batch (or part of it) looks already imported. Nothing appended, nothing saved."
      exit 0
    }
  }

  # ---- Pur_Combined enrichment lookup (bulk read -> in-memory hashtable) --
  $wsPC = $wb.Worksheets.Item($PurCombinedSheet)
  $pcLastRow = $wsPC.Cells($wsPC.Rows.Count, 1).End($xlUp).Row
  $pcLastCol = $wsPC.Cells(1, $wsPC.Columns.Count).End($xlToLeft).Column
  $pcVals = Get-BulkValues $wsPC 2 1 $pcLastRow $pcLastCol
  $pcByKey = @{}
  for ($r = 0; $r -lt $pcVals.GetLength(0); $r++) {
    $k = "$($pcVals[$r, ($PC_KEY_COL - 1)])".Trim()
    if ($k -ne "") { $pcByKey[$k] = $r }
  }
  Write-Log "Pur_Combined '$PurCombinedSheet': $($pcVals.GetLength(0)) rows indexed by composite key."

  # ---- Exchange Rate lookup (bulk read -> currency/month dictionaries) ----
  $wsFx = $wb.Worksheets.Item("Exchange Rate")
  $fxCurrHeader = Get-BulkValues $wsFx $FX_CURR_HEADER_ROW $FX_CURR_FIRST_COL $FX_CURR_HEADER_ROW $FX_CURR_LAST_COL
  $fxCurrCols = @{}
  for ($c = 0; $c -lt $fxCurrHeader.GetLength(1); $c++) { $fxCurrCols["$($fxCurrHeader[0,$c])".Trim().ToUpper()] = $c }
  # Month labels read cell-by-cell via .Text (only 36 cells) - a date-typed
  # cell's Value2 would come back as an OLE serial number, not "2026-09", and
  # ReportingMonth is compared against the latter.
  $fxMonthRow = -1
  for ($r = $FX_MONTH_FIRST_ROW; $r -le $FX_MONTH_LAST_ROW; $r++) {
    if ($wsFx.Cells($r, $FX_MONTH_COL).Text.Trim() -eq $ReportingMonth) { $fxMonthRow = $r; break }
  }
  if ($fxMonthRow -lt 0) { throw "Reporting month '$ReportingMonth' not found in Exchange Rate!N$FX_MONTH_FIRST_ROW:N$FX_MONTH_LAST_ROW." }
  $fxRatesForMonth = Get-BulkValues $wsFx $fxMonthRow $FX_CURR_FIRST_COL $fxMonthRow $FX_CURR_LAST_COL
  $fxRateByCurrency = @{}
  foreach ($cur in $fxCurrCols.Keys) { $fxRateByCurrency[$cur] = To-NumOrZero $fxRatesForMonth[0, $fxCurrCols[$cur]] }
  Write-Log "Exchange Rate row for ${ReportingMonth}: $($fxRateByCurrency.Count) currencies loaded."

  # ---- Vendor grouping reference (bulk read -> hashtable, for unmapped-count only; the live XLOOKUP formula does the real per-row lookup) ----
  $vendorRefSet = New-Object 'System.Collections.Generic.HashSet[string]'
  if ($hasVendorRef) {
    $wsVendor = $wb.Worksheets.Item("COO-VENDOR LIST")
    $vLastRow = $wsVendor.Cells($wsVendor.Rows.Count, 1).End($xlUp).Row
    if ($vLastRow -ge 2) {
      $vVals = Get-BulkValues $wsVendor 2 1 $vLastRow 1
      for ($r = 0; $r -lt $vVals.GetLength(0); $r++) { [void]$vendorRefSet.Add("$($vVals[$r,0])".Trim().ToUpper()) }
    }
  }

  # ---- Build the raw+computed value block (target cols B:AE, 30 cols) -----
  $firstNewRow = $priorLastRow + 1
  $lastNewRow = $priorLastRow + $n
  $block = New-Object 'object[,]' $n, 30   # columns 2(B)..31(AE) => index 0..29
  $missingFxCount = 0
  $unmappedVendors = New-Object 'System.Collections.Generic.HashSet[string]'
  for ($r = 0; $r -lt $n; $r++) {
    $block[$r, ($TGT.B - 2)] = $actVals[$r, ($SRC.TRAN_DATE - 1)]
    $vendorId = "$($actVals[$r, ($SRC.VENDOR_ID - 1)])".Trim()
    $block[$r, ($TGT.C - 2)] = $vendorId
    $block[$r, ($TGT.D - 2)] = $actVals[$r, ($SRC.VENDOR_NAME - 1)]
    $block[$r, ($TGT.E - 2)] = $actVals[$r, ($SRC.BUYER - 1)]
    $block[$r, ($TGT.F - 2)] = $actVals[$r, ($SRC.ORDER_NO - 1)]
    $block[$r, ($TGT.G - 2)] = $actVals[$r, ($SRC.PO_DATE - 1)]
    $block[$r, ($TGT.H - 2)] = $actVals[$r, ($SRC.REFERENCE - 1)]
    $block[$r, ($TGT.I - 2)] = $actVals[$r, ($SRC.STOCK_ID - 1)]
    $block[$r, ($TGT.J - 2)] = $actVals[$r, ($SRC.TERMS - 1)]
    $block[$r, ($TGT.K - 2)] = $actVals[$r, ($SRC.PART_NO - 1)]
    $block[$r, ($TGT.L - 2)] = $null                               # Part Desc1 - left blank, no ACT001 equivalent
    $block[$r, ($TGT.M - 2)] = $actVals[$r, ($SRC.PART_DESC - 1)]
    $block[$r, ($TGT.N - 2)] = $actVals[$r, ($SRC.COMM_CODE - 1)]
    $block[$r, ($TGT.O - 2)] = $actVals[$r, ($SRC.REC_QTY - 1)]
    $block[$r, ($TGT.P - 2)] = $actVals[$r, ($SRC.PO_UOM - 1)]
    $poCur = "$($actVals[$r, ($SRC.PO_CUR - 1)])".Trim()
    $block[$r, ($TGT.Q - 2)] = $poCur
    $block[$r, ($TGT.R - 2)] = $actVals[$r, ($SRC.PO_COST - 1)]
    $block[$r, ($TGT.S - 2)] = $actVals[$r, ($SRC.TOT_FOR_PO - 1)]
    $block[$r, ($TGT.T - 2)] = $actVals[$r, ($SRC.TOT_HKD_PO - 1)]
    $block[$r, ($TGT.U - 2)] = $actVals[$r, ($SRC.AVG_COST - 1)]
    $block[$r, ($TGT.V - 2)] = $actVals[$r, ($SRC.TOT_AVG_COST - 1)]
    $block[$r, ($TGT.W - 2)] = $actVals[$r, ($SRC.PPV - 1)]
    $block[$r, ($TGT.X - 2)] = $actVals[$r, ($SRC.EXCH_RATE - 1)]
    $block[$r, ($TGT.Y - 2)] = $null                                # Order Type - left blank, no ACT001 equivalent

    # AB: Monthly Exchange Rate - computed value (matches historical
    # precedent: every prior row holds a plain number here, not a live
    # formula - see the investigation notes).
    $curKey = $poCur.ToUpper()
    if ($fxRateByCurrency.ContainsKey($curKey)) {
      $block[$r, ($TGT.AB - 2)] = $fxRateByCurrency[$curKey]
    } else {
      $block[$r, ($TGT.AB - 2)] = $null
      $missingFxCount++
    }

    # AE: Unit Commission/deduction at PO currency - computed value, matched
    # by composite key against Pur_Combined (2-for-2 confirmed against this
    # month's own already-populated rows: AE = -(unit deduct-from-AP-invoice
    # + unit 3rd-party commission), both treated as 0 when blank so this
    # still matches exactly in the (so-far only observed) case where the
    # 3rd-party commission column is blank.
    $key = $newKeys[$r]
    if ($pcByKey.ContainsKey($key)) {
      $pcRow = $pcByKey[$key]
      $deductAp = To-NumOrZero $pcVals[$pcRow, ($PC_DEDUCT_AP_UNIT - 1)]
      $comm3rd = To-NumOrZero $pcVals[$pcRow, ($PC_COMM_3RD_UNIT - 1)]
      $block[$r, ($TGT.AE - 2)] = -($deductAp + $comm3rd)
    } else {
      $block[$r, ($TGT.AE - 2)] = 0
      Write-Log "WARNING: no Pur_Combined match for composite key '$key' (row $($r+2) of $ACT001Sheet) - AE set to 0 for this row."
    }

    if ($hasVendorRef -and -not $vendorRefSet.Contains($vendorId.ToUpper())) { [void]$unmappedVendors.Add($vendorId) }
  }

  if ($missingFxCount -gt 0) {
    throw "Validation failed before any write: $missingFxCount row(s) have a PO currency not found in Exchange Rate for $ReportingMonth. Fix the Exchange Rate sheet or the source data, then retry."
  }

  Write-Log "Appending rows $firstNewRow..$lastNewRow into 'Pur Details'."
  $wsPD.Range($wsPD.Cells($firstNewRow, $TGT.B), $wsPD.Cells($lastNewRow, $TGT.AE)).Value2 = $block
  $insertedRange = $wsPD.Range($wsPD.Cells($firstNewRow, $TGT.A), $wsPD.Cells($lastNewRow, $TGT.AP))

  # ---- Formula columns: enter once in the first new row, then FillDown ----
  # (same technique as Combine-Into-Master.ps1's Merge-AppendSheet, just one
  # column range at a time instead of the whole row - needed here because
  # this sheet interleaves formula columns with the raw/computed-value
  # columns already written above, rather than one uniform block.)
  function Fill-FormulaColumn($col, $formula1) {
    $c1 = $wsPD.Cells($firstNewRow, $col)
    $c1.Formula = $formula1
    if ($lastNewRow -gt $firstNewRow) {
      $wsPD.Range($c1, $wsPD.Cells($lastNewRow, $col)).FillDown() | Out-Null
    }
  }
  # Absolute Excel reference, e.g. (Abs 'C' 5) -> '$C$5' - built via plain
  # string concatenation (not interpolation) so there's no backtick-escaping
  # of literal '$' characters to get wrong.
  function Abs($colLetter, $row) { return '$' + $colLetter + '$' + $row }

  $Bc = Get-ColLetter $TGT.B; $Fc = Get-ColLetter $TGT.F; $Kc = Get-ColLetter $TGT.K
  $Oc = Get-ColLetter $TGT.O; $Rc = Get-ColLetter $TGT.R; $Sc = Get-ColLetter $TGT.S
  $Cc = Get-ColLetter $TGT.C; $Dc = Get-ColLetter $TGT.D; $Uc = Get-ColLetter $TGT.U
  $Xc = Get-ColLetter $TGT.X; $Vc = Get-ColLetter $TGT.V; $Wc = Get-ColLetter $TGT.W
  $ABc = Get-ColLetter $TGT.AB; $ACc = Get-ColLetter $TGT.AC; $AEc = Get-ColLetter $TGT.AE
  $AFc = Get-ColLetter $TGT.AF; $AGc = Get-ColLetter $TGT.AG; $AHc = Get-ColLetter $TGT.AH
  $AIc = Get-ColLetter $TGT.AI; $AJc = Get-ColLetter $TGT.AJ; $AMc = Get-ColLetter $TGT.AM
  $ANc = Get-ColLetter $TGT.AN

  Fill-FormulaColumn $TGT.A  ('=CONCAT(' + $Fc + $firstNewRow + ',' + $Kc + $firstNewRow + ',' + $Oc + $firstNewRow + ')')
  Fill-FormulaColumn $TGT.Z  ('=TEXT(' + $Bc + $firstNewRow + ',"yyyy-mm")')
  Fill-FormulaColumn $TGT.AA ('=TEXT(' + $Bc + $firstNewRow + ',"yyyy")')
  Fill-FormulaColumn $TGT.AC ('=' + $Rc + $firstNewRow + '*' + $ABc + $firstNewRow)
  Fill-FormulaColumn $TGT.AD ('=' + $Sc + $firstNewRow + '*' + $ABc + $firstNewRow)
  Fill-FormulaColumn $TGT.AF ('=' + $Rc + $firstNewRow + '+' + $AEc + $firstNewRow)
  Fill-FormulaColumn $TGT.AG ('=' + $AFc + $firstNewRow + '*' + $ABc + $firstNewRow)
  Fill-FormulaColumn $TGT.AH ('=' + $AGc + $firstNewRow + '*' + $Oc + $firstNewRow)
  Fill-FormulaColumn $TGT.AI ('=IFERROR(' + $Uc + $firstNewRow + '/' + $Xc + $firstNewRow + ',0)')
  Fill-FormulaColumn $TGT.AJ ('=' + $AIc + $firstNewRow + '*' + $Oc + $firstNewRow)
  Fill-FormulaColumn $TGT.AK ('=' + $AJc + $firstNewRow + '-' + $AHc + $firstNewRow)
  Fill-FormulaColumn $TGT.AL ('=IFERROR(' + $Wc + $firstNewRow + '/' + $Vc + $firstNewRow + ',0)')

  # AM: dedicated vendor reference first, then the historical self-lookup
  # against every row before this batch, then the raw Vendor Name - per
  # requirement 9's explicit priority order (don't rely on history alone
  # when a dedicated reference sheet exists).
  $histLookup = '=IFERROR(XLOOKUP(' + $Cc + $firstNewRow + ',' + (Abs $Cc ($pdHeaderRow+1)) + ':' + (Abs $Cc $priorLastRow) + ',' + (Abs $AMc ($pdHeaderRow+1)) + ':' + (Abs $AMc $priorLastRow) + '),' + $Dc + $firstNewRow + ')'
  if ($hasVendorRef) {
    $refLookup = 'XLOOKUP(' + $Cc + $firstNewRow + ",'COO-VENDOR LIST'!" + (Abs 'A' 2) + ':' + (Abs 'A' $vLastRow) + ",'COO-VENDOR LIST'!" + (Abs 'C' 2) + ':' + (Abs 'C' $vLastRow) + ')'
    Fill-FormulaColumn $TGT.AM ('=IFERROR(' + $refLookup + ',' + $histLookup.Substring(1) + ')')
  } else {
    Fill-FormulaColumn $TGT.AM $histLookup
  }

  # AN: average AC within THIS batch only (same receipt date + part number) -
  # absolute range fixed to firstNewRow:lastNewRow so FillDown doesn't shift it.
  $anFormula = '=IFERROR(AVERAGEIFS(' + (Abs $ACc $firstNewRow) + ':' + (Abs $ACc $lastNewRow) + ',' + (Abs $Bc $firstNewRow) + ':' + (Abs $Bc $lastNewRow) + ',' + $Bc + $firstNewRow + ',' + (Abs $Kc $firstNewRow) + ':' + (Abs $Kc $lastNewRow) + ',' + $Kc + $firstNewRow + '),0)'
  Fill-FormulaColumn $TGT.AN $anFormula
  Fill-FormulaColumn $TGT.AO ('=IFERROR((' + $ACc + $firstNewRow + '-' + $ANc + $firstNewRow + ')/' + $ANc + $firstNewRow + ',0)')
  Fill-FormulaColumn $TGT.AP ('=(' + $ACc + $firstNewRow + '-' + $ANc + $firstNewRow + ')*' + $Oc + $firstNewRow)

  # ---- Step 11: preserve target formatting (formats only, not values) ----
  $fmtSrc = $wsPD.Range($wsPD.Cells($priorLastRow, 1), $wsPD.Cells($priorLastRow, $TGT.AP))
  $fmtDst = $wsPD.Range($wsPD.Cells($firstNewRow, 1), $wsPD.Cells($lastNewRow, $TGT.AP))
  $fmtSrc.Copy() | Out-Null
  $fmtDst.PasteSpecial($xlPasteFormats) | Out-Null
  # Do not set CutCopyMode here: Excel COM exposes an enum that rejects both
  # Boolean $false and numeric 0 even though VBA accepts False to clear it.

  # Recalculate just this workbook so every formula above has a real value
  # before validation reads them back.
  $wb.Application.Calculate()

  # ---- Step 12: validation --------------------------------------------
  $failures = @()

  $gotN = $lastNewRow - $firstNewRow + 1
  if ($gotN -ne $n) { $failures += "Appended row count ($gotN) does not equal ACT001 source row count ($n)." }

  $zVals = Get-BulkValues $wsPD $firstNewRow $TGT.Z $lastNewRow $TGT.Z
  for ($r = 0; $r -lt $zVals.GetLength(0); $r++) {
    if ("$($zVals[$r,0])".Trim() -ne $ReportingMonth) { $failures += "Row $($firstNewRow+$r): Z = '$($zVals[$r,0])', expected '$ReportingMonth'."; break }
  }

  $aVals = Get-BulkValues $wsPD $firstNewRow $TGT.A $lastNewRow $TGT.A
  $blankAKeys = 0
  $aKeySet = New-Object 'System.Collections.Generic.HashSet[string]'
  for ($r = 0; $r -lt $aVals.GetLength(0); $r++) {
    $k = "$($aVals[$r,0])".Trim()
    if ($k -eq "") { $blankAKeys++ } else { [void]$aKeySet.Add($k) }
  }
  if ($blankAKeys -gt 0) { $failures += "$blankAKeys appended row(s) have a blank column A key." }
  if ($aKeySet.Count -ne $n) { $failures += "Appended column A has duplicate keys within the batch itself ($($aKeySet.Count) distinct out of $n rows)." }

  $errorTexts = @("#REF!","#DIV/0!","#NAME?","#VALUE!","#N/A","#SPILL!")
  $wholeBlockVals = Get-BulkValues $wsPD $firstNewRow 1 $lastNewRow $TGT.AP
  $errorCellsFound = @()
  for ($r = 0; $r -lt $wholeBlockVals.GetLength(0); $r++) {
    for ($c = 0; $c -lt $wholeBlockVals.GetLength(1); $c++) {
      $v = "$($wholeBlockVals[$r,$c])"
      foreach ($errTxt in $errorTexts) { if ($v -eq $errTxt) { $errorCellsFound += "$(Get-ColLetter ($c+1))$($firstNewRow+$r)=$errTxt" } }
    }
  }
  if ($errorCellsFound.Count -gt 0) { $failures += "$($errorCellsFound.Count) error cell(s) found, e.g. $($errorCellsFound[0..([Math]::Min(4,$errorCellsFound.Count-1))] -join ', ')" }

  # Reconciliation: source sums must equal appended target sums exactly.
  function Sum-Col($arr, $idx0) { $s = 0.0; for ($r=0; $r -lt $arr.GetLength(0); $r++) { $s += To-NumOrZero $arr[$r,$idx0] }; return $s }
  $reconPairs = @(
    @{ Name = "REC_QTY";      SrcIdx = $SRC.REC_QTY-1;      TgtCol = $TGT.O },
    @{ Name = "TOT_FOR_PO";   SrcIdx = $SRC.TOT_FOR_PO-1;    TgtCol = $TGT.S },
    @{ Name = "TOT_HKD_PO";   SrcIdx = $SRC.TOT_HKD_PO-1;    TgtCol = $TGT.T },
    @{ Name = "TOT_AVG_COST"; SrcIdx = $SRC.TOT_AVG_COST-1;  TgtCol = $TGT.V },
    @{ Name = "PPV";          SrcIdx = $SRC.PPV-1;           TgtCol = $TGT.W }
  )
  foreach ($p in $reconPairs) {
    $srcSum = Sum-Col $actVals $p.SrcIdx
    $tgtArr = Get-BulkValues $wsPD $firstNewRow $p.TgtCol $lastNewRow $p.TgtCol
    $tgtSum = Sum-Col $tgtArr 0
    if ([Math]::Abs($srcSum - $tgtSum) -gt 0.01) { $failures += "Reconciliation mismatch on $($p.Name): source sum=$srcSum, target sum=$tgtSum." }
  }

  $adArr = Get-BulkValues $wsPD $firstNewRow $TGT.AD $lastNewRow $TGT.AD
  $akArr = Get-BulkValues $wsPD $firstNewRow $TGT.AK $lastNewRow $TGT.AK
  $totalPoPriceUsd = Sum-Col $adArr 0
  $totalPriceVarUsd = Sum-Col $akArr 0

  if ($failures.Count -gt 0) {
    Write-Log "VALIDATION FAILED - rolling back the $gotN appended row(s), nothing will be saved:"
    $failures | ForEach-Object { Write-Log "  - $_" }
    $wsPD.Range($wsPD.Rows($firstNewRow), $wsPD.Rows($lastNewRow)).Delete() | Out-Null
    $wb.Close($false)
    $excel.Calculation = $priorCalc
    throw "Validation failed ($($failures.Count) issue(s)) - rows rolled back, workbook closed without saving. See log above for details."
  }

  $excel.Calculation = $xlCalculationAutomatic
  $wb.Application.Calculate()
  $wb.Save()
  $sw.Stop()

  Write-Log "----"
  Write-Log "Workbook: $targetPath"
  Write-Log "Source sheet: $ACT001Sheet  Target sheet: Pur Details  Reporting month: $ReportingMonth"
  Write-Log "Source row count: $n"
  Write-Log "Appended rows: $firstNewRow - $lastNewRow"
  Write-Log "Total PO Price USD (col AD): $totalPoPriceUsd"
  Write-Log "Total Price Variance USD (col AK): $totalPriceVarUsd"
  Write-Log "Unmapped vendor count (not in COO-VENDOR LIST): $($unmappedVendors.Count)$(if ($unmappedVendors.Count -gt 0) { ' - ' + (($unmappedVendors | Select-Object -First 10) -join ', ') })"
  Write-Log "Missing exchange-rate count: $missingFxCount"
  Write-Log "Duplicate count: 0 (would have stopped earlier otherwise)"
  Write-Log "Validation result: PASSED"
  Write-Log "Elapsed: $($sw.Elapsed)"
  Write-Log "Saved: $targetPath"
}
finally {
  if ($null -ne $priorCalc -and $excel) { try { $excel.Calculation = $priorCalc } catch {} }
  if ($wb) { try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null } catch {} }
  if ($excel) {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  }
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}
