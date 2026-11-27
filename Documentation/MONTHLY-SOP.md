# MTC Performance Report — Monthly SOP (finalized 2026-11-27)

Example period: **2026/11** (reporting month Oct-2026 shipped + Nov-2026 forecast).
Repo: https://github.com/anthonymankaho/MTC_Sales_Report · Local: `C:\Users\User\Desktop\MTC Performance Report Claude`

Three stages. Stage 1 is time-triggered; Stage 2 is email-triggered; Stage 3 builds + publishes.

---

## Stage 1 — Month-end Forecast (SO/PO combine) + SC17 VERIFY

**When:** last working day of the month, **4:00pm** (e.g. 2026-10-31 16:00 for the Nov cycle).

1. Build the outstanding-order forecast (SO/PO combine) →
   `Forecast/MTC_FORECAST_YYYYMMDD_UPDATED.xlsx` (98-col `data` sheet).
   - Source of truth = the **current-time SCF017 / SC17 outstanding report** (Epicor),
     NOT a partial SO header export. Rebuild from SCF017 so coverage is complete
     (a partial header export misses still-open SOs — verified 2026-11-27: 266 keys).
   - Pipeline: `Forecast/fetch_mtc_details.py` + `fetch_missing_pos/sos.py` (Epicor Retrieve,
     env `EPICOR_USER/PASS`) → `Forecast/combine_mtc_data.py` (outstanding-only, Epidueqty,
     Ship By Date bucketing, FX from `fx_report_island.json`). Or rebuild directly from SCF017.
2. **NEW — VERIFY against SC17 (quantity compare):**
   ```powershell
   python Forecast\verify_forecast_vs_sc17.py Forecast\MTC_FORECAST_YYYYMMDD_UPDATED.xlsx <SC17_or_SCF017.xlsx>
   ```
   Reports matched-qty differences, forecast-only and SC17-only keys, grand totals.
   **Accept only when `RESULT: IDENTICAL`** (same cutoff time). Any QTY difference must be
   explained (shipment between cutoffs / missing SO) before proceeding.
   - Note: SCF017 can only be generated at **current time** (cannot back-date to a prior
     cutoff). Compare forecast vs SC17 at the SAME timestamp; a stale forecast vs a newer
     SC17 will show shipment drift, not a method error.

## Stage 2 — Peggy's Sales & Commission email (trigger)

**When:** Gmail receives **「MTC: Sales & Commission Report YYYY - Mmm」** from Peggy.

- Option A (manual): run the browser Html Tool → PowerShell → Verify → Excel Report.
- Option B (automation): trigger the monthly data cycle on email arrival.
- Produce the shipped-data Excel deliverable
  `FIN_MTC Sales & Margin (x-y)YYYY_YYYYMMDD_R#.xlsx` (from the merged master
  `Sales-Data-MTC-Mthly Rate - UPDATED_YYYYMM.xlsx` via `Combine-Into-Master.ps1` +
  `Convert-ACT001ToPurDetails.ps1`, then finance prep).

## Stage 3 — Build HTML (shipped + Domestic + Forecast) + publish

```powershell
powershell -ExecutionPolicy Bypass -File PowerShell\Generate-MTCSalesReportHtml.ps1
```
(`Builder HTML.bat` wraps it.) The PowerShell **combines three sources**:

| # | Input | What it feeds |
|---|---|---|
| 1 | `FIN_MTC Sales & Margin (x-y)YYYY_YYYYMMDD_R#.xlsx` (auto-newest; the **shipped** Excel report derived from the master `Sales-Data-MTC-Mthly Rate - UPDATED_YYYYMM.xlsx`) | Shipped Data tab (`data_sheet.json`) |
| 2 | `Domestic/Domestic Orders *.xlsx` (Domestic/ subfolder preferred, root fallback) | Domestic tab (`domestic.json` + `fx_domestic.json`) |
| 3 | `Forecast/MTC_FORECAST_*_UPDATED.xlsx` (auto-newest) | Forecast Data tab (`forecast_sheet.json`) |

plus FX (`STANDARD EXCHANGE RATE-YYYY-MM.xls` + weekly USD/INR).

**Outputs:**
- **Offline:** `FIN_MTC_Performance_Report_html_TEMP.html` (project root). `-Publish` also
  writes a dated snapshot `FIN_MTC_Performance_Report_html_YYYYMM.html`.
- **Online:** the generator does **NOT** push. Publish separately:
  `git add FIN_MTC_Performance_Report_html_TEMP.html ; git commit ; git push origin MTC`
  → https://anthonymankaho.github.io/MTC_Sales_Report/FIN_MTC_Performance_Report_html_TEMP.html

**Important clarifications (verified 2026-11-27):**
- Input #1 is the **FIN_MTC Sales & Margin** shipped workbook, not the raw
  `Sales-Data-MTC-Mthly Rate - UPDATED.xlsx` directly (the master is the merge input; the FIN
  workbook is the finance-prepared shipped report the generator reads).
- Domestic is read from `Domestic/` (updated 2026-11-27; previously root-only).
- Online publish is a separate git push, not part of the PowerShell.

## Dashboard basis (for reviewers)
4-row index: Total Performance / Shipped / Domestic / Forecast; section filter gates Total only.
Domestic = MTC **commission earning** (not the domestic companies' revenue);
Total Net Sales = Shipped + Forecast + Domestic commission. Forecast = full outstanding snapshot
(Sales_USD / Prime_USD / Gross_Cont). Ask AI on every tab; 🤖 Grok Bot connects via `mtc-mcp`.
