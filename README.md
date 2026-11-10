# MTC Sales Report Claude — MTC Performance Report tooling

Tooling for the monthly MTC sales/commission report cycle. It replaces the manual
data-entry portion of the process — logging into MMS/EPICOR and downloading the source
reports stays manual — by taking those source files, preparing the rows/values that go
into the master workbook, merging them in safely, appending the purchase ledger, and
optionally building the interactive **MTC Performance Report** HTML for distribution.

## Monthly processing steps

1. **Download source files** from MMS/EPICOR (manual, unchanged):

   - `IVR032.xlsx` — sales export
   - `STANDARD EXCHANGE RATE-YYYY-MM.xls` — monthly FX rates
   - `Sales & Commission Report-MTC.xlsx` — commission data (feeds Net Selling Price)
   - `EPI_Purchase Commission Combined.xlsx` — purchase commission/PO data
   - `ACT001.xlsx` — purchase receipts export (needed for the optional Step 6)
   - `Domestic Orders … (for Accounts).xlsx` — monthly domestic-orders snapshot
     supplied by Peggy/Keith for the report's Domestic tab (see
     [appendix](Documentation/DOMESTIC-AND-GLOBE-SOP-appendix.md))
2. **Pre-check the exports**: run `PowerShell\Check-SourceTransactionCounts.ps1`. It
   counts real transaction rows in the three core exports (title/blank rows excluded) and
   fails if they disagree — a mismatch means one export is incomplete or covers the
   wrong period. Do this before staging.
3. **Stage the data** with the browser tool in [app/index.html](app/index.html) (loads
   `app/xlsx.full.min.js`, SheetJS, entirely client-side — no data leaves the browser).
   It reads the four sources plus the current cleaned master (read-only) and produces
   `MTC_Import_Rows.xlsx` — tabs for MTC Sales Details, Exchange Rate, Net Selling
   Price, and PO comm.
4. **Review** `MTC_Import_Rows.xlsx` by hand.
5. **Merge into the master** with [PowerShell/Combine-Into-Master.ps1](PowerShell/Combine-Into-Master.ps1):

   ```powershell
   cd "C:\Users\User\Desktop\MTC Sales Report Claude"
   powershell -ExecutionPolicy Bypass -File PowerShell\Combine-Into-Master.ps1 `
     -MasterPath "Sales-Data-MTC-Mthly Rate - UPDATED_20261103_CLEANED_95COL.xlsx" `
     -ImportPath "MTC_Import_Rows.xlsx" `
     -OutputPath "Sales-Data-MTC-Mthly Rate - UPDATED_YYYYMM.xlsx"
   ```

   Real Excel automation (COM) against a **copy** — the master is never touched. Since
   the 2026-11-03 revision the MTC Sales Details tab is merged **by header name** through
   the `MTC_Sales_Details` table: only the 19 approved input fields are written, every
   calculated column's formula is extended from the nearest real formula found scanning
   upward, headers are schema-validated first (missing/duplicate/unknown = abort with a
   list), and the workbook is recalculated and all pivot tables refreshed before saving.
   A failed run leaves nothing: the partial output is deleted. Net Selling Price and
   PO comm keep the positional FillDown + seed-column merge; Exchange Rate fills the
   month's pre-built row. The script now **requires** the MTC Sales Details tab in the
   import file.
6. **Append the purchase ledger (when ACT001 exists)** with
   [PowerShell/Convert-ACT001ToPurDetails.ps1](PowerShell/Convert-ACT001ToPurDetails.ps1),
   run on a separate copy via `-OutputPath`: it reads the `ACT001_MMYY` +
   `Pur_Combined_MMYY` sheets already inside the workbook, enriches and appends them to
   `Pur Details`, refuses duplicate months/keys, and validates row counts, month labels,
   keys, error cells and five amount-column reconciliations — anything failing rolls the
   new rows back and saves nothing.
7. **Promote** the verified month-dated output as the next cleaned master after
   spot-checking; keep the previous cleaned master archived and never use the
   145-column reference copy as a merge target.

Use `Builder Data.bat` for a menu around Steps 2–7, or ask Claude to run the cycle.

## Interactive HTML report — MTC Performance Report

Once the monthly master is finalized, an interactive, single-file HTML report is built
for email distribution or presenting — no Excel needed to view it:

```powershell
powershell -ExecutionPolicy Bypass -File PowerShell\Generate-MTCSalesReportHtml.ps1
```

`Builder HTML.bat` wraps this. No arguments are required for the normal monthly case —
the script auto-detects the newest `FIN_MTC Sales & Margin (...)_YYYYMMDD_R#.xlsx` in
the project root, reads it read-only via Excel COM, reads the Domestic Orders workbook
(via the `Domestic Orders*.xlsx` wildcard — its name contains em/en dashes) and the
monthly FX `.xls`, stages **eleven JSON data islands** under `HTML_Report_Generator\_work`
(plus `family_category_map.json`), then runs `HTML_Report_Generator/build_report.js` to
assemble `FIN_MTC_Performance_Report_html_TEMP.html` in the project root. Pass `-KeepWork`
to keep the islands for inspection.

The report has **eleven tabs**: Dashboard (KPIs + charts, NET/NET NET toggle), Data (the
full 145-column ledger), Exchange Rate, five YoY (2026 vs 2025) tabs each with a Total
Index and pie chart, Customer-Quarter, **Domestic** (2026–2027 MJP/MCN/MZQ/MIN orders and
commission, per-currency cards, derived-USD memos only), and Purchase: YTD by Vendor.
The Dashboard headline is a live one-line action title above a full-width **3D trade
globe** (Global Sales by destination ⇄ Global Purchase by origin; clicking a country
drills into the Data tab pre-filtered; the Geographic bar and Origin pie remain as the
automatic WebGL fallback). Every tab supports English/繁中 and light/dark toggles; every
exhibit is numbered with source footnotes. The whole thing is one self-contained file
(keep under the 11 MB email budget).

A **Qwen AI assistant panel** ("Ask AI", bottom-right) is built into the report via the
shared local relay at `127.0.0.1:8765` (the same loopback service as the MPL and GPR
tools): it sends only current filters and aggregate figures, never transaction rows or
credentials, can propose filters you confirm with Apply, and the file itself carries no
key. Recipients who don't run the relay simply see "Relay offline" — everything else
works from `file://` with no network at all.

## Current progress (2026-11-10)

- Docs consolidated (2026-11-10): README and the report's welcome intro (now v3) were
  rewritten to match the 11-tab MTC Performance Report; the legacy SOP docx was retired —
  the written procedure is now this README + `DOMESTIC-AND-GLOBE-SOP-appendix.md` +
  `PowerShell List.md` + `presentation.md`.
- 2026-11-09 cycle (all pushed to GitHub `origin/MTC`): Domestic tab + 3D globe + YoY KPI
  cards landed; report retitled to MTC Performance Report; McKinsey-style presentation
  pass (action titles, exhibit numbering, source footnotes).
- 2026-11-03 cycle: header-driven merge promoted; master cleaned 145 → 95 columns
  (`Sales-Data-MTC-Mthly Rate - UPDATED_20261103_CLEANED_95COL.xlsx`, 3,876 rows);
  full-folder reorg into `PowerShell/`, `Documentation/`, `Reports/`, `Reference/`;
  `Builder Data.bat` / `Builder HTML.bat` launchers; ACT001 converter validated
  (282 rows, reconciliation passed).
- The Ask AI panel is built into `build_report.js` and the current TEMP HTML but was not
  yet committed to GitHub as of 2026-11-09.

## Structure

| Path | Purpose |
| --- | --- |
| `Builder Data.bat`, `Builder HTML.bat` | Menu launchers for the data cycle and HTML builder |
| `app/index.html` | Browser staging tool (tool Steps 1–5, output = SOP Step 3–4) |
| `app/xlsx.full.min.js` | SheetJS library used client-side by the tool |
| `PowerShell/Check-SourceTransactionCounts.ps1` | Pre-staging row-count gate on the three core exports |
| `PowerShell/Combine-Into-Master.ps1` | Header-driven Excel-COM merge into the cleaned master |
| `PowerShell/Convert-ACT001ToPurDetails.ps1` | ACT001 + Pur_Combined → Pur Details append, with rollback |
| `PowerShell/Generate-MTCSalesReportHtml.ps1` | Excel-COM extractor → data islands + report builder |
| `PowerShell/App/`, `PowerShell/Backup/` | Staging-workflow copy and retained backup of the combine script |
| `Documentation/` (this README + appendix + lists) | Written procedure — the former `MTC Sales Report Import SOP.docx` was retired 2026-11-10; archived copy under `Documentation/Archive/` |
| `Documentation/DOMESTIC-AND-GLOBE-SOP-appendix.md` | Domestic data sources, FX rules, globe drill-down, troubleshooting |
| `Documentation/PowerShell List.md`, `Documentation/presentation.md` | Script inventory and workflow briefing |
| `HTML_Report_Generator/build_report.js` | Assembles islands + vendored libs into the single-file report |
| `HTML_Report_Generator/globe/globe.js` | 3D earth engine (ported from the Global Purchase Report build) |
| `HTML_Report_Generator/vendor/` | jQuery, DataTables, Chart.js, three.module.min.js, land-110m.json |
| `HTML_Report_Generator/_work/` | Transient JSON islands (rebuilt each run; `-KeepWork` keeps them) |
| `Sales-Data-MTC-Mthly Rate.xlsx` | Original 145-column reference master — never a merge target |
| `Sales-Data-MTC-Mthly Rate - UPDATED_20261103_CLEANED_95COL.xlsx` | Current cleaned 95-column master |
| `Sales-Data-MTC-Mthly Rate - UPDATED_YYYYMM.xlsx` | Month-dated merge output (promote after checking) |
| `MTC_Import_Rows.xlsx` | Staged output from `app/index.html`, ready to merge |
| `IVR032.xlsx`, `STANDARD EXCHANGE RATE-YYYY-MM.xls`, `Sales & Commission Report-MTC.xlsx`, `EPI_Purchase Commission Combined.xlsx`, `ACT001.xlsx` | Current-month source files in the root |
| `Domestic Orders 2026–2027 — MJP _ MCN _ MZQ _ MIN (for Accounts).xlsx` | Domestic snapshot for the Domestic tab (read via wildcard) |
| `FIN_MTC Sales & Margin (...)_YYYYMMDD_R#.xlsx` | Finance source for the HTML report (auto-detected newest) |
| `FIN_MTC_Performance_Report_html_TEMP.html` | Output of the HTML report generator — shareable single file |
| `202609/`, `Pur Comm/` | Archived monthly outputs and purchase-commission history |
| `Reports/Audit Summaries/`, `Reference/Prompts/` | Dated audit exports; reusable reference prompts |
| `MTC SANDBOX/` | Untracked test area (revised-script trials, `TEST_*.xlsx`) — never merge-target |

## Folder map

- Root keeps launchers, active sources, masters and outputs — everything the scripts
  expect by name or wildcard.
- `PowerShell/` — all executable workflow scripts; `App/` and `Backup/` hold secondary
  copies of the combine script (root copy is canonical).
- `Documentation/` — the Domestic & globe appendix (SOP companion), script list, briefing; `Archive/` holds the retired SOP docx (2026-11-10).
- `HTML_Report_Generator/` — builder script, globe engine, vendored libraries, `_work/` islands.
- `MTC SANDBOX/` — renamed former `SANDBOX/`; test artifacts only, untracked (add to
  `.gitignore` before anyone runs `git add .`).
- `202609/`, `Pur Comm/`, `Reports/`, `Reference/` — archives. Monthly folders before
  202609 were removed from the repo on 2026-11-03; they remain in git history.

## Notes

- **Master-promotion caveat:** `PowerShell/Generate-MTCSalesReportHtml.ps1` currently
  pins the `Family1 → Family Category` brand lookup to
  `…UPDATED_20261103_CLEANED_95COL.xlsx` by filename. When you promote a new cleaned
  master next month, update that line (or point it at a wildcard like the Domestic
  workbook) — otherwise the Brand column degrades to blank with a console warning.
- **Re-extraction caveat:** the Domestic INR rate is synthesised from the **newest by
  filename date** `FIN_Ex Rate_PSBUILD_YYYYMMDD.xlsx` in
  `..\Weekly Exchange Report - CLAUDE\Output`. Re-running the full generator after the
  weekly project rolls to a new month changes that figure. For visual/copy-only changes,
  rebuild from the surviving islands instead:
  `node HTML_Report_Generator\build_report.js HTML_Report_Generator\_work <output.html>`.
- The GitHub default branch is still the old `main`; the live branch is **`MTC`**
  (`origin/MTC`). The `backup` remote (AnthonyManhk) cannot be pushed from this machine.
- This repo is **private** and contains real sales, commission, and margin data — keep it
  that way.
- The large `Sales-Data-MTC-Mthly Rate*.xlsx` files are tracked as-is (no Git LFS); the
  10 MB TEMP html also churns per rebuild — expect the repo to grow by roughly their
  size each time they're updated.
