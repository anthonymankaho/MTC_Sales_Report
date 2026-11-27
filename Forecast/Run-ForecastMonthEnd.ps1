#requires -Version 5.1
<#
Month-end Forecast automation (SOP Stage 1). Runs at month-end last business day
16:00 via Scheduled Task (daily trigger + this date guard), or manually with -Force.

Data source (owner decision: combine both):
  1. Prefer the newest downloaded SCF017 / SC17 outstanding report (complete,
     current-time) -> build_forecast_from_scf017.py  => forecast == SC17.
  2. Else, if Epicor credentials are stored (DPAPI), fetch SO/PO + combine
     (coverage may be partial - warns).
Then VERIFY vs the newest SCF017 (qty compare) and log.

  -Force          run regardless of date (manual / test)
  -SetCredentials prompt + store Epicor creds DPAPI (one-time setup)
#>
param([switch]$Force, [switch]$SetCredentials)
$ErrorActionPreference = 'Stop'
$FC = $PSScriptRoot
$root = Split-Path -Parent $FC
$log = Join-Path $FC 'forecast_monthend_log.txt'
function Log($m){ $line = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '  ' + $m; Add-Content -LiteralPath $log -Value $line -Encoding UTF8; Write-Host $line }

# ---- one-time Epicor credential setup (DPAPI, current user+machine) ----
$credPath = Join-Path $FC 'epicor_credentials.dat'
if ($SetCredentials) {
  $email = Read-Host 'Epicor user (e.g. alex@meyer-msc.com)'
  $sec = Read-Host 'Epicor password' -AsSecureString
  @{ Email = $email; EncryptedPassword = (ConvertFrom-SecureString $sec) } | ConvertTo-Json | Set-Content -LiteralPath $credPath -Encoding UTF8
  Log 'Epicor credentials stored (DPAPI).'
  exit 0
}

# ---- last business day of month guard ----
$today = (Get-Date).Date
$lastDay = New-Object DateTime ($today.Year, $today.Month, [DateTime]::DaysInMonth($today.Year, $today.Month))
while ($lastDay.DayOfWeek -eq 'Saturday' -or $lastDay.DayOfWeek -eq 'Sunday') { $lastDay = $lastDay.AddDays(-1) }
if (-not $Force -and $today -ne $lastDay.Date) {
  Log ("SKIP: today {0:yyyy-MM-dd} is not month-end last business day ({1:yyyy-MM-dd})." -f $today, $lastDay)
  exit 0
}
Log ("RUN month-end forecast (last business day {0:yyyy-MM-dd})." -f $lastDay)

# ---- locate newest SCF017 / SC17 ----
$scCands = @()
$scCands += Get-ChildItem -Path (Join-Path $env:USERPROFILE 'Downloads') -Filter 'EPI_SCF017_*.xlsx' -ErrorAction SilentlyContinue
$scCands += Get-ChildItem -Path $FC -Filter 'EPI_SCF017_*.xlsx' -ErrorAction SilentlyContinue
$scCands += Get-ChildItem -Path $FC -Filter 'SC17_*.xlsx' -ErrorAction SilentlyContinue
$sc = $scCands | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$out = Join-Path $FC ('MTC_FORECAST_{0:yyyyMMdd}_UPDATED.xlsx' -f (Get-Date))

if ($sc) {
  Log ("Source: SCF017/SC17 = {0}" -f $sc.FullName)
  python (Join-Path $FC 'build_forecast_from_scf017.py') $sc.FullName $out
  if ($LASTEXITCODE -ne 0) { Log 'ERROR: build_forecast_from_scf017 failed'; exit 1 }
} else {
  # fallback: Epicor fetch + combine (partial coverage warning)
  if (Test-Path $credPath) {
    $c = Get-Content $credPath -Raw | ConvertFrom-Json
    $sec = ConvertTo-SecureString $c.EncryptedPassword
    $plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec))
    $env:EPICOR_USER = $c.Email; $env:EPICOR_PASS = $plain
    Log 'No SCF017 found; using stored Epicor creds to fetch SO/PO (coverage may be partial).'
    python (Join-Path $FC 'fetch_mtc_details.py')
    python (Join-Path $FC 'fetch_missing_pos.py')
    python (Join-Path $FC 'fetch_missing_sos.py')
    python (Join-Path $FC 'combine_mtc_data.py')
    $r7 = Get-ChildItem -Path $FC -Filter 'MTC_Data_*_R7.xlsx' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($r7) { Copy-Item $r7.FullName $out -Force; Log ("WARNING: forecast from Epicor combine (partial coverage): {0}" -f $out) }
    else { Log 'ERROR: combine produced no R7 output'; exit 1 }
  } else {
    Log 'ERROR: no SCF017/SC17 found and no stored Epicor creds. Download SCF017 or run -SetCredentials.'
    exit 1
  }
}

# ---- VERIFY vs SC17 ----
if ($sc) {
  Log 'VERIFY forecast vs SC17 (qty compare)...'
  python (Join-Path $FC 'verify_forecast_vs_sc17.py') $out $sc.FullName
  Log ("verify exit code = {0} (0 = IDENTICAL)" -f $LASTEXITCODE)
} else {
  Log 'VERIFY skipped (no SCF017 to compare).'
}
Log 'Month-end forecast run complete.'
exit 0
