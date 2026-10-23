# MTC Sales Report Processing

Tooling for the monthly MTC sales/commission report cycle. It replaces the manual
data-entry portion of the process — logging into MMS/EPICOR and downloading the source
reports stays manual — by taking those source files, preparing the rows/values that go
into the master workbook, and merging them in safely.

## Monthly processing steps

1. **Download source files** from MMS/EPICOR (manual, unchanged):

   - `IVR032.xlsx` — sales export
   - `STANDARD EXCHANGE RATE-YYYY-MM.xls` — monthly FX rates
   - `EPI_Purchase Commission Combined.xlsx` — commission/PO data
2. **Stage the data** with the browser tool in [app/index.html](app/index.html) (loads
   `app/xlsx.full.min.js`, the SheetJS library, entirely client-side — no data leaves the
   browser). It reads the three source files plus the current master workbook (read-only,
   to auto-detect the next empty row), and produces `MTC_Import_Rows.xlsx` — a single
   file with tabs for MTC Sales Details, Exchange Rate, Net Selling Price, and PO comm.
3. **Review** `MTC_Import_Rows.xlsx` by hand.
4. **Merge into the master** with [Combine-Into-Master.ps1](Combine-Into-Master.ps1):

   ```powershell
   cd "C:\Users\User\Desktop\MTC Sales Report Claude"
   powershell -ExecutionPolicy Bypass -File Combine-Into-Master.ps1 `
     -MasterPath "Sales-Data-MTC-Mthly Rate.xlsx" `
     -ImportPath "MTC_Import_Rows.xlsx" `
     -OutputPath "Sales-Data-MTC-Mthly Rate - UPDATED.xlsx"
   ```

   This drives real Excel automation (COM) against a **copy** of the master — the
   original is never touched. It fills down each column's existing formulas onto the
   new rows, then overwrites just the raw-data columns with the real values, so pivot
   tables, XLOOKUPs, and formatting elsewhere in the workbook stay intact.
5. Spot-check `Sales-Data-MTC-Mthly Rate - UPDATED.xlsx`, then promote it to be the new
   master for next month.

See [MTC Sales Report Import SOP.docx](<MTC Sales Report Import SOP.docx>) for the full
step-by-step SOP.

## Interactive HTML report (optional)

Once the monthly master workbook is finalized (Step 5 above), an interactive, single-file
HTML version of the FIN_MTC Sales & Margin report can be generated for sharing by email or
presenting to others — no Excel required to view it. This is a separate, optional
capability from the monthly import cycle; run it whenever an updated report needs to be
shared.

```powershell
powershell -ExecutionPolicy Bypass -File HTML_Report_Generator\Generate-MTCSalesReportHtml.ps1
```

No arguments are required for the normal monthly case — the script auto-detects the newest
matching `FIN_MTC Sales & Margin (...)_YYYYMMDD_R#.xlsx` file in the project folder, reads
it read-only via Excel COM automation (never modifies the source), and writes
`FIN_MTC_Sales_Report_html_TEMP.html` to the project root.

The report has nine tabs — Dashboard (KPIs + charts, NET/NET NET toggle), Data (the full
transaction ledger), Exchange Rate, five YoY (2026 vs 2025) comparison tabs each with a
Total Index and pie chart, and Customer-Quarter. Every tab supports an English/Traditional
Chinese toggle and light/dark mode (top right), and clicking any highlighted number or name
drills into the underlying transactions on the Data tab. The whole thing is one
self-contained HTML file — safe to attach directly to an email.

See [MTC Sales Report Import SOP.docx](<MTC Sales Report Import SOP.docx>), section 5, for
the full write-up.

## Structure

| Path                                                                     | Purpose                                                                                    |
| ------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------ |
| `app/index.html`                                                       | Browser-based staging tool (Steps 1-3 of the SOP)                                          |
| `app/xlsx.full.min.js`                                                 | SheetJS library used by the tool to read/write`.xlsx` client-side                        |
| `Combine-Into-Master.ps1`                                              | PowerShell + Excel COM script that merges staged rows into the master (Step 4)             |
| `MTC Sales Report Import SOP.docx`                                     | Full written procedure                                                                     |
| `Sales-Data-MTC-Mthly Rate.xlsx`                                       | Current master workbook                                                                    |
| `Sales-Data-MTC-Mthly Rate - UPDATED.xlsx`                             | Output of the merge script for the current month                                           |
| `FIN_MTC Sales & Margin*.xlsx`, `FIN - MTC Sales & Margin Temp.xlsx` | Finance sales & margin report exports                                                      |
| `Sales & Commission Report-MTC.xlsx`                                   | Commission report                                                                          |
| `EPI_Purchase Commission Combined.xlsx`                                | Source: commission/PO data from EPICOR                                                     |
| `IVR032.xlsx`                                                          | Source: monthly sales export                                                               |
| `STANDARD EXCHANGE RATE-YYYY-MM.xls`                                   | Source: monthly FX rate table                                                              |
| `MTC_Import_Rows.xlsx`                                                 | Staged output from`app/index.html`, ready to merge                                       |
| `HTML_Report_Generator/Generate-MTCSalesReportHtml.ps1`                | Extracts data from the finalized FIN_MTC workbook via Excel COM (read-only)                |
| `HTML_Report_Generator/build_report.js`                                | Node script that assembles the extracted data into the single-file interactive HTML report |
| `HTML_Report_Generator/vendor/`                                        | Vendored jQuery, DataTables, and Chart.js used by the generated report                     |
| `FIN_MTC_Sales_Report_html_TEMP.html`                                  | Output of the HTML report generator — the shareable interactive report                    |

## Notes

- This repo is **private** and contains real sales, commission, and margin data — keep
  it that way.
- The large `Sales-Data-MTC-Mthly Rate*.xlsx` files are the full master workbooks and are
  tracked as-is (no Git LFS); expect the repo to grow by roughly their size each time
  they're updated.
