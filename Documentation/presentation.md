# MTC Report Workflow

## 1. What the launchers do

Two menu-driven batch files provide entry points to the MTC monthly workflow:

- `Builder Data.bat` starts source checks, sales-data merge, or ACT001-to-Pur-Details conversion.
- `Builder HTML.bat` opens the browser staging tool, starts HTML generation, or runs the app-folder merge-script copy.

The launchers prompt for the workbook paths and reporting period required by each operation. Merge and conversion actions require a new output file path.

## 2. PowerShell roles

MTC PowerShell scripts are collected under `PowerShell/`.

- `PowerShell\Check-SourceTransactionCounts.ps1` checks that the three monthly source files agree on transaction row count.
- `PowerShell\Combine-Into-Master.ps1` merges the staged import into a new copy of the cleaned 95-column master, maps input fields by header, and preserves formula columns.
- `PowerShell\Convert-ACT001ToPurDetails.ps1` converts ACT001 purchase rows and appends them to Pur Details, subject to validation.
- `PowerShell\Generate-MTCSalesReportHtml.ps1` extracts the FIN_MTC report workbook data **plus the newest `Forecast/MTC_FORECAST_*_UPDATED.xlsx`** and creates the interactive HTML report.
- `PowerShell\App\Combine-Into-Master.ps1` is the merge-script copy paired with the browser staging tool.
- `PowerShell\Backup\Combine-Into-Master- BACKUP.ps1` is a backup and is not part of routine processing.

## 3. Written procedure (SOP)

The legacy `MTC Sales Report Import SOP.docx` was **retired 2026-11-10** (archived under
`Documentation/Archive/`). The written procedure is now:

- `README.md` — monthly Steps 1–7 (source checks, header-driven merge, ACT001 append, promotion) + the HTML report + the Forecast pipeline.
- `DOMESTIC-AND-GLOBE-SOP-appendix.md` — Domestic data sources, FX rules, globe drill-down.
- `PowerShell List.md` — script inventory + the Forecast Python pipeline.

Merge basis: the cleaned 95-column master (original 145-column kept as reference); 19 input
fields written by header name; formula columns protected; separate ACT001-to-Pur-Details step.

## 4. Current report structure (finished version, 2026-11-26)

`FIN_MTC_Performance_Report_html_TEMP.html` is the finished single-file report (~10 MB):

- **12 tabs**: Dashboard, **Shipped Data** (95-col actuals), **Forecast Data** (98-col outstanding orders), Exchange Rate, five YoY tabs, Customer-Quarter, Domestic, Purchase: YTD by Vendor.
- **Dashboard 4-row index**: Total Performance / Shipped / Domestic / Forecast, with a Total-only section filter and per-row filter-context lines. Domestic = MTC commission earning; Total = Shipped + Forecast + Domestic commission.
- **Ask AI** is shell-level (every tab); **🤖 Grok Bot** button connects a personal Grok via the `mtc-mcp` Cloudflare MCP worker.
- Forecast source: `Forecast/MTC_FORECAST_YYYYMM_UPDATED.xlsx` (Python pipeline; outstanding-only, Ship By Date basis, MMS-refreshed Sku/Part lookups).

## 5. Validation status

Both batch menus were opened and exited successfully. This only checked launcher startup; the batch files were not used to run Excel-writing operations during this check. Follow the SOP and review script validation output when processing actual monthly workbooks.
