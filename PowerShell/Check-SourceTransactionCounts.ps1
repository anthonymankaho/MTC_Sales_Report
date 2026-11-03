<#
  Check-SourceTransactionCounts.ps1

  Sanity check for the three monthly source files (Step 1 of the SOP, before
  staging with app/index.html): counts real transaction rows in each -
  excluding the title/header row and any blank rows - and confirms all three
  agree on the same row count. A mismatch usually means one export is
  incomplete or was run for the wrong date range before the others.

  Each file keeps its transactions on a different sheet, with a different
  "anchor" column reliably populated on every real row (and blank otherwise):
    - IVR032.xlsx                          : sheet "Page1",     col A (COMPANY)
    - Sales & Commission Report-MTC.xlsx   : sheet "Combined",  col A (Company)
    - EPI_Purchase Commission Combined.xlsx: sheet "Combined",  col B (COMPANY -
      column A itself is unused/blank in this file's layout)
  Header is always row 1; data starts row 2.

  Usage:
    powershell -ExecutionPolicy Bypass -File PowerShell\Check-SourceTransactionCounts.ps1
    powershell -ExecutionPolicy Bypass -File PowerShell\Check-SourceTransactionCounts.ps1 -IvrPath "..." -SalesCommPath "..." -PurchaseCommPath "..."

  Exits with code 0 if all three counts match, 1 otherwise (so it can gate a
  later step in a chained script if wanted).
#>

param(
  [string]$IvrPath = "C:\Users\User\Desktop\MTC Sales Report Claude\IVR032.xlsx",
  [string]$SalesCommPath = "C:\Users\User\Desktop\MTC Sales Report Claude\Sales & Commission Report-MTC.xlsx",
  [string]$PurchaseCommPath = "C:\Users\User\Desktop\MTC Sales Report Claude\EPI_Purchase Commission Combined.xlsx"
)

$ErrorActionPreference = "Stop"

# Same relative-path gotcha as Combine-Into-Master.ps1's Open-WB: Excel COM's
# Workbooks.Open() doesn't reliably resolve a relative path against this
# process's working directory, so every path is made absolute up front.
function Resolve-AbsolutePath($path) {
  if (-not (Test-Path $path)) { throw "File not found: $path" }
  return (Resolve-Path $path).Path
}
$IvrPath = Resolve-AbsolutePath $IvrPath
$SalesCommPath = Resolve-AbsolutePath $SalesCommPath
$PurchaseCommPath = Resolve-AbsolutePath $PurchaseCommPath

# Counts non-blank cells in $anchorCol from row 2 through the sheet's last
# used row - not just End(xlUp) from the bottom, since that alone would miss
# a stray interior blank row and silently over-count.
function Get-TransactionRowCount($ws, $anchorCol, $headerRow) {
  $xlUp = -4162
  $lastRow = $ws.Cells($ws.Rows.Count, $anchorCol).End($xlUp).Row
  if ($lastRow -le $headerRow) { return 0 }
  $vals = $ws.Range($ws.Cells($headerRow + 1, $anchorCol), $ws.Cells($lastRow, $anchorCol)).Value2
  $count = 0
  if ($lastRow -eq $headerRow + 1) {
    if ($null -ne $vals -and "$vals".Trim() -ne "") { $count = 1 }
  } else {
    for ($r = 1; $r -le $vals.GetLength(0); $r++) {
      $v = $vals[$r, 1]
      if ($null -ne $v -and "$v".Trim() -ne "") { $count++ }
    }
  }
  return $count
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $results = @()

  $wb = $excel.Workbooks.Open($IvrPath, [Type]::Missing, $true)
  $ws = $wb.Worksheets.Item("Page1")
  $results += [PSCustomObject]@{ File = "IVR032.xlsx"; Sheet = "Page1"; Rows = Get-TransactionRowCount $ws 1 1 }
  $wb.Close($false)

  $wb = $excel.Workbooks.Open($SalesCommPath, [Type]::Missing, $true)
  $ws = $wb.Worksheets.Item("Combined")
  $results += [PSCustomObject]@{ File = "Sales & Commission Report-MTC.xlsx"; Sheet = "Combined"; Rows = Get-TransactionRowCount $ws 1 1 }
  $wb.Close($false)

  $wb = $excel.Workbooks.Open($PurchaseCommPath, [Type]::Missing, $true)
  $ws = $wb.Worksheets.Item("Combined")
  $results += [PSCustomObject]@{ File = "EPI_Purchase Commission Combined.xlsx"; Sheet = "Combined"; Rows = Get-TransactionRowCount $ws 2 1 }
  $wb.Close($false)

  Write-Output "Transaction row counts (title + blank rows excluded):"
  $results | ForEach-Object { Write-Output ("  {0,-42} [{1}] : {2}" -f $_.File, $_.Sheet, $_.Rows) }

  $distinctCounts = $results.Rows | Select-Object -Unique
  if ($distinctCounts.Count -eq 1) {
    Write-Output "OK: all three files have the same transaction row count ($($distinctCounts[0]))."
    exit 0
  } else {
    Write-Output "MISMATCH: these files do not all have the same transaction row count - check which export is incomplete or for the wrong period before staging."
    exit 1
  }
} finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  [System.GC]::Collect()
}
