# MTC Performance Report — Standard Operating Procedure

**Version:** 1.0 (drafted 2026-11-27) · **Owner:** Anthony · **Data provider:** Peggy (Sales & Commission), Keith (Domestic)
**Cycle:** monthly · **Example period:** 2026/11 cycle = Oct-2026 shipped actuals + Nov-2026 outstanding forecast
**Local folder:** `C:\Users\User\Desktop\MTC Performance Report Claude` · **Repo:** https://github.com/anthonymankaho/MTC_Sales_Report (branch `MTC`, private)
**Online report:** https://anthonymankaho.github.io/MTC_Sales_Report/FIN_MTC_Performance_Report_html_TEMP.html

This is the full end-to-end SOP. The one-page quick reference is `MONTHLY-SOP.md` (same three stages).

---

## 1. What the report is

One self-contained interactive HTML file (~10 MB, opens offline from `file://`, no Excel needed) with **twelve tabs**:

Dashboard · **Shipped Data** (actuals ledger) · **Forecast Data** (outstanding orders) · Exchange Rate · five YoY 2026-vs-2025 tabs · Customer-Quarter · **Domestic** · Purchase YTD by Vendor.

**Dashboard basis (agree this with reviewers before quoting numbers):**

- Index reads in **four rows**: Total Performance / Shipped / Domestic / Forecast.
- The three-checkbox section filter (Shipped / Forecast / Domestic) gates **only the Total row**.
- **Domestic = MTC commission earning only.** The domestic companies' own sales are NOT MTC revenue.
- **Total Net Sales = Shipped + Forecast + Domestic commission.**
- Forecast = full outstanding snapshot valued at unshipped qty (Sales_USD / Prime_USD / Gross_Cont; no commission yet).
- Every tab: English/繁中 + light/dark. **Ask AI** panel works on every tab. **🤖 Grok Bot** button connects a personal Grok via the `mtc-mcp` Cloudflare MCP server.

## 2. Cycle calendar

| When | Trigger | Stage | What happens |
|---|---|---|---|
| Last business day of month, 16:00 | Scheduled Task `MTC_Forecast_MonthEnd` (daily trigger + date guard) | **1** | Forecast build + SC17 quantity verify (automated) |
| Month-end / early next month | Gmail: **「MTC: Sales & Commission Report YYYY - Mmm」** from Peggy | **2** | Data cycle → merged master → FIN shipped workbook |
| After Stage 2 deliverable is ready | Manual | **3** | Build HTML (shipped + Domestic + forecast) → offline TEMP + online publish |

## 3. Stage 1 — Month-end Forecast + SC17 VERIFY (automated)

**Output:** `Forecast/MTC_FORECAST_YYYYMMDD_UPDATED.xlsx` (98-column `data` sheet, outstanding orders only).

### 3.1 How it runs

`Forecast/Run-ForecastMonthEnd.ps1` (wrapped by **`Builder Forecast.bat`**, also option 4 in `Builder HTML.bat`):

1. **Date guard** — runs only on the last business day of the month (walks back over Sat/Sun). Other days: logs `SKIP`, exits. Manual run: `-Force` (BAT option 1).
2. **Data source — combine both, in priority order:**
   - **Preferred:** newest downloaded `EPI_SCF017_*.xlsx` (Downloads, then `Forecast/`; also accepts `SC17_*.xlsx`) → `build_forecast_from_scf017.py` rebuilds the 98-col sheet directly from it. Complete coverage; forecast == SC17.
   - **Fallback:** if no SCF017 is found and Epicor credentials are stored (DPAPI file `Forecast/epicor_credentials.dat`, set once via BAT option 2 / `-SetCredentials`), fetches SO/PO detail (`fetch_mtc_details.py` + `fetch_missing_pos/sos.py`) and combines (`combine_mtc_data.py`). **Coverage may be partial — the log warns.**
3. **VERIFY** — `verify_forecast_vs_sc17.py` compares forecast vs the SCF017/SC17 by (SO, PART) outstanding quantity.
4. Everything logged to `Forecast/forecast_monthend_log.txt`.

### 3.2 Valuation rules (built into the scripts)

- **Outstanding only** — COMPLETED/CANCELLED lines excluded; value **Epidueqty** (unshipped qty), never ordered qty.
- **Ship By Date is mandatory** and is the bucketing basis (yyyy-mm / Month / Year columns).
- SO↔PO join: `SO.PonumE27 == PO.PonumJ16`, FIFO pairing for duplicate SKUs.
- FX: USD-per-unit from the report's Exchange Rate island (`Forecast/fx_report_island.json`), latest populated month ≤ reporting month (EUR Sep-2026 carry-forward is expected, not a bug).
- Lookups from the workbook's **Sku / Part / BRAND** sheets: Family Category = Brand Name, Family = Product Family Desc, Meyer Brand via BRAND map, Commodity = `SKU-` when the 6th Part_id char is `-`. **`MISC` / `NO BRAND` fallbacks are correct** for unmapped items — do not "fix" them.

### 3.3 Acceptance gate

**Accept only `RESULT: IDENTICAL`** (zero qty difference, zero one-side-only keys) when forecast and SC17 share the same cutoff time.

- SCF017 can only be generated at **current time** — it cannot be back-dated. A stale forecast vs a newer SC17 shows shipment drift, not a method error; rebuild from the newer SCF017 instead of investigating phantom differences.
- Historical lesson (2026-11-27): a forecast built from a partial SO-header export missed **266 still-outstanding keys**. Always prefer the complete current-time SCF017.

### 3.4 Owner preparation (before month-end)

- Download the current-time **SCF017** from Epicor into Downloads (drag it into `Forecast/` to archive), **or** store Epicor creds once (`Builder Forecast.bat` → 2) so the fallback can auto-fetch.
- Scheduled Task registration (one-time; `StartWhenAvailable=True` is mandatory so a missed 16:00 run catches up at next boot — `schtasks /Create` cannot set this flag):

```powershell
$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NoLogo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Users\User\Desktop\MTC Performance Report Claude\Forecast\Run-ForecastMonthEnd.ps1"' -WorkingDirectory 'C:\Users\User\Desktop\MTC Performance Report Claude\Forecast'
$trigger  = New-ScheduledTaskTrigger -Daily -At 16:00
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName 'MTC_Forecast_MonthEnd' -Action $action -Trigger $trigger -Settings $settings
```

- Sku/Part lookup sheets are refreshed from MMS when the catalogue changes (`mms-login` skill: login + email OTP, SKU Excel export + GetPartList).

## 4. Stage 2 — Peggy's email → data cycle → FIN shipped workbook

**Trigger:** Gmail receives **「MTC: Sales & Commission Report YYYY - Mmm」** from Peggy.

### 4.1 Source files (manual downloads, unchanged)

| File | Source | Feeds |
|---|---|---|
| `IVR032.xlsx` | EPICOR | Sales export |
| `STANDARD EXCHANGE RATE-YYYY-MM.xls` | Finance | Monthly FX |
| `Sales & Commission Report-MTC.xlsx` | Peggy's email | Commission / Net Selling Price |
| `EPI_Purchase Commission Combined.xlsx` | EPICOR | Purchase commission / PO |
| `ACT001.xlsx` | EPICOR | Purchase receipts (Step 6) |
| `Domestic Orders … (for Accounts).xlsx` | Peggy/Keith | Domestic tab → keep in `Domestic/` |

### 4.2 Steps (menu: `Builder Data.bat`)

1. **Pre-check counts** — `PowerShell\Check-SourceTransactionCounts.ps1`: the three core exports must agree on real transaction-row counts; a mismatch means an incomplete/wrong-period export. **Do not stage on a mismatch.**
2. **Stage** — browser tool `app/index.html` (SheetJS, fully client-side; no data leaves the browser). Reads the four sources + current master read-only → `MTC_Import_Rows.xlsx` (MTC Sales Details / Exchange Rate / Net Selling Price / PO comm tabs).
3. **Review** `MTC_Import_Rows.xlsx` by hand.
4. **Merge** — `PowerShell\Combine-Into-Master.ps1` (Excel COM on a **copy**; master never touched):
   - Master = `Sales-Data-MTC-Mthly Rate - UPDATED.xlsx` (the live cleaned 95-column master).
   - Header-driven merge through the `MTC_Sales_Details` table: only the 19 approved input fields written; calculated-column formulas extended from the nearest real formula; schema validated first (missing/duplicate/unknown header = abort); workbook recalculated + pivots refreshed; failed run deletes the partial output.
   - Output = month-dated `Sales-Data-MTC-Mthly Rate - UPDATED_YYYYMM.xlsx`.
5. **Append purchase ledger** (when ACT001 exists) — `PowerShell\Convert-ACT001ToPurDetails.ps1` on a separate copy: reads `ACT001_MMYY` + `Pur_Combined_MMYY`, appends to `Pur Details`, refuses duplicate months/keys, validates counts + five amount reconciliations; any failure rolls back and saves nothing.
6. **Promote** — spot-check the month-dated output, then promote it as the next live master (keep the filename `Sales-Data-MTC-Mthly Rate - UPDATED.xlsx`; archive the previous one). **Never merge into the 145-column reference copy** `Sales-Data-MTC-Mthly Rate.xlsx`.
7. **Finance prep** → the shipped deliverable **`FIN_MTC Sales & Margin (x-y)YYYY_YYYYMMDD_R#.xlsx`** (DATA sheet + Exchange Rate + 5 YoY pivots + Customer-Quarter). This FIN workbook — not the raw master — is what Stage 3 reads.

## 5. Stage 3 — Build the HTML report + publish

### 5.1 Build (menu: `Builder HTML.bat` → 2)

```powershell
powershell -ExecutionPolicy Bypass -File PowerShell\Generate-MTCSalesReportHtml.ps1
```

The generator combines **three inputs** (all auto-detected, read-only Excel COM):

| # | Input | Selection rule | Feeds |
|---|---|---|---|
| 1 | `FIN_MTC Sales & Margin (x-y)YYYY_YYYYMMDD_R#.xlsx` | by **filename period** (end month → gen date → R#), never LastWriteTime | Shipped Data + Dashboard + YoY + Customer-Quarter |
| 2 | `Domestic/Domestic Orders *.xlsx` | `Domestic/` preferred, root fallback; wildcard (em/en dashes in name) | Domestic tab |
| 3 | `Forecast/MTC_FORECAST_*_UPDATED.xlsx` | newest by write time; absent → Forecast tab skipped, build still succeeds | Forecast Data tab |

plus FX: `STANDARD EXCHANGE RATE-YYYY-MM.xls` (HKD table) + synthetic INR row from the newest-by-filename-date weekly `FIN_Ex Rate_PSBUILD_*.xlsx` (fallback 95.50 with a loud WARNING). Brand lookup (`Family1 → Family Category`) reads the live master.

### 5.2 Outputs

- **Offline:** `FIN_MTC_Performance_Report_html_TEMP.html` in the project root — the continuously-overwritten working file. **Keep TEMP in root; never move it into a month archive.**
- **Dated snapshot:** add `-Publish` at the actual business checkpoint only → `FIN_MTC_Performance_Report_html_YYYYMM.html`.
- **Online:** the generator does **NOT** push. Publish separately:

```powershell
git add FIN_MTC_Performance_Report_html_TEMP.html
git commit -m "MTC Performance Report YYYY-MM"
git push origin MTC
```

→ live at the Pages URL above. Check the repo is on branch `MTC` and the gh account is `anthonymankaho` first.

### 5.3 Distribute

Email the TEMP html (≤ 11 MB budget) or share the Pages link. Keep the repo **private** — it contains real sales/commission/margin data; the Pages site is the controlled public surface.

## 6. Acceptance checklist (before distributing)

- [ ] Stage 1 log shows `RESULT: IDENTICAL` (or every difference explained in writing).
- [ ] Stage 2 pre-check counts passed; merge + ACT001 reconciliations passed; FIN workbook R# bumped.
- [ ] Generator console: no WARNING lines (FX fallback / missing Family Category source / missing forecast).
- [ ] Dashboard Total = Shipped + Forecast + Domestic **commission**; section filter moves only the Total row.
- [ ] Forecast tab present; Ship By Date populated on every row; MISC/NO BRAND only where genuinely unmapped.
- [ ] Spot-check 2–3 rows per tab against the FIN workbook; EN/繁中 and light/dark render.
- [ ] TEMP opens from `file://` with no network; Ask AI degrades gracefully offline.
- [ ] `-Publish` snapshot written; git push done; Pages URL loads the new period.

## 7. Troubleshooting / known pitfalls

| Symptom | Cause / fix |
|---|---|
| Forecast vs SC17 qty differences | Different cutoff times → rebuild from the current-time SCF017; never back-date. Partial Epicor fetch → prefer SCF017. |
| Brand column blank in YoY tabs | Live master renamed/missing → generator warns; keep the promoted master at `Sales-Data-MTC-Mthly Rate - UPDATED.xlsx`. |
| Domestic INR figure changed after re-run | Weekly FX project rolled to a new month (INR synthesised from newest weekly file). For copy-only edits rebuild from islands: `node HTML_Report_Generator\build_report.js HTML_Report_Generator\_work <out.html>`. |
| Wrong source period picked | FIN files are selected by filename period, not mtime — name new deliverables correctly `(x-y)YYYY_YYYYMMDD_R#`. |
| Scheduled forecast "didn't run" | Check `Forecast/forecast_monthend_log.txt` for SKIP (not month-end) vs ERROR; confirm task `StartWhenAvailable=True`. |
| Data vanished between runs | Incremental-chaining bug pattern — see the `incremental-report-file-chaining` skill; never pick a base file by timestamp across candidates. |
| `git push` rejected / repo not found | Wrong gh account active (anthonymankaho vs AnthonyManhk) or remote PR commits ahead → `git pull --rebase --autostash` then push. |
| MISC / NO BRAND rows | **Correct behaviour** for unmapped SKUs/parts — not an error. |

## 8. Appendix — tooling map

| Tool | Purpose |
|---|---|
| `Builder Data.bat` | Menu: Stage 2 steps (pre-check → stage → merge → ACT001) |
| `Builder HTML.bat` | Menu: staging tool / HTML build / merge / **4 = Forecast** |
| `Builder Forecast.bat` | 1 run month-end now · 2 store Epicor creds · 3 fetch SO/PO · 4 build from SCF017 · 5 verify |
| `Forecast/Run-ForecastMonthEnd.ps1` | Stage-1 runner (date guard, source priority, verify, log) |
| `Forecast/build_forecast_from_scf017.py` | 98-col rebuild from SCF017 (args: SCF017 path, [output]) |
| `Forecast/verify_forecast_vs_sc17.py` | Qty verify (args: [forecast], SC17/SCF017; empty forecast = auto-newest) |
| `PowerShell/Generate-MTCSalesReportHtml.ps1` | Stage-3 extractor + builder (`-Publish`, `-KeepWork`) |
| `HTML_Report_Generator/build_report.js` | Assembles JSON islands + vendored libs into the single file |
| `GPR-Web/workers/mtc-mcp` | Grok Bot MCP worker (Cloudflare, public/free): `mtc_summary / mtc_monthly / mtc_top / mtc_query / mtc_forecast` |
| Archives | `202609/`, `2026-10/`, `Pur Comm/` — month-dated supporting files; **TEMP stays in root** |

**House rules:** `.ps1` files pure ASCII, no BOM · unattended tasks always `StartWhenAvailable=True` · commit early (VS Code host restarts kill sessions) · check for parallel (Codex) edits before commit/merge · credentials never committed (`epicor_credentials.dat` is gitignored).
