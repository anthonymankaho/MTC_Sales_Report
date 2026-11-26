# MTC PowerShell Script List

## Workflow launchers

| Launcher | Menu actions |
|---|---|
| `Builder Data.bat` | Check monthly source transaction counts; merge staged rows using `PowerShell\Combine-Into-Master.ps1`; convert ACT001 purchase rows to Pur Details using a separate output workbook. |
| `Builder HTML.bat` | Open the browser staging tool; generate the interactive HTML report; merge staged rows using `PowerShell\App\Combine-Into-Master.ps1`. |

## PowerShell scripts

Project scripts are grouped under `PowerShell/`.

| Script | Purpose | Used by |
|---|---|---|
| `PowerShell\Check-SourceTransactionCounts.ps1` | Checks transaction row counts in IVR032, Sales & Commission, and Purchase Commission source files. Stops the user from proceeding with mismatched source exports. | `Builder Data.bat`, option 1 |
| `PowerShell\Combine-Into-Master.ps1` | Merges `MTC_Import_Rows.xlsx` into a new copy of the cleaned 95-column master. Maps the approved input fields by header and protects formula columns. Requires master, import, and output paths. | `Builder Data.bat`, option 2 |
| `PowerShell\Convert-ACT001ToPurDetails.ps1` | Converts a loaded `ACT001_MMYY` sheet, enriches it using the matching `Pur_Combined_MMYY` and Exchange Rate data, then appends rows to Pur Details. Use a separate output path. | `Builder Data.bat`, option 3 |
| `PowerShell\Generate-MTCSalesReportHtml.ps1` | Extracts data from the FIN_MTC Excel report (DATA, Exchange Rate, YoY pivots, Domestic, FX) **and the newest `Forecast/MTC_FORECAST_*_UPDATED.xlsx` (forecast_sheet.json)**, then builds the interactive HTML report. Auto-selects a matching source report unless a path is supplied. Forecast absent → Forecast tab skipped, build still succeeds. | `Builder HTML.bat`, option 2 |
| `PowerShell\App\Combine-Into-Master.ps1` | Separate merge-script copy paired with the browser-based staging tool. | `Builder HTML.bat`, option 3 |
| `PowerShell\Backup\Combine-Into-Master- BACKUP.ps1` | Backup copy. Not part of the normal run sequence. | Do not use for routine processing |

## Operating notes

- Use the cleaned 95-column master as the merge input. Keep the original 145-column master as a backup/reference.
- For merge and ACT001 conversion, choose a new output filename for each run. The launchers reject an output path that already exists.
- Close Excel workbooks before running Excel COM automation scripts.
- The launchers have been checked to open and exit their menus. This check did not run the Excel-writing operations.
- The HTML report's **Shipped Data** tab is the 95-column actuals ledger (`data_sheet.json`); the **Forecast Data** tab is the 98-column outstanding-order forecast (`forecast_sheet.json`).

## Forecast data pipeline (Python, `Forecast/` folder)

The Forecast Data tab is produced by Python scripts (not PowerShell) in `Forecast/`:

| Script | Purpose |
|---|---|
| `Forecast\fetch_mtc_details.py` | Logs into the Epicor portal (env `EPICOR_USER`/`EPICOR_PASS`) and Retrieves each SO/PO extended entity (DetailList lines); caches to `_work_mtc_details_raw.json`; writes EPI_SO/PO Detail xlsx. |
| `Forecast\fetch_missing_pos.py` / `fetch_missing_sos.py` | Backfill the 1:1 SO↔PO join gaps (stale header snapshots) into the raw cache. |
| `Forecast\combine_mtc_data.py` | Builds the 98-col `data` sheet: joins SO.PonumE27==PO.PonumJ16 (FIFO by SKU), outstanding-only (excludes COMPLETED/CANCELLED), values at Epidueqty (unshipped), buckets by Ship By Date, FX from `fx_report_island.json` (USD-per-unit). Lookups from Sku/Part/BRAND sheets. Output `MTC_FORECAST_YYYYMM_UPDATED.xlsx`. |

- Sku/Part lookup sheets are refreshed from MMS via the `mms-login` skill (SKU Excel export + `GetPartList`).
- `MTC_FORECAST_YYYYMM_UPDATED.xlsx` is the forecast source the HTML generator reads.
- The `mtc-mcp` Cloudflare Worker (GPR-Web repo `workers/mtc-mcp`) exposes this data over MCP for Grok users; redeploy with `npx wrangler deploy --config workers/mtc-mcp/wrangler.jsonc`.
