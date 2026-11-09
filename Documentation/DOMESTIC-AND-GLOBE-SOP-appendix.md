# SOP Appendix — Domestic Tab & 3D Globe / SOP 附錄——國內訂單分頁與 3D 地球儀

Companion to *MTC Sales Report Import SOP.docx* (section 5), added with the 2026-11-09
build cycle (eleven-tab report). Covers: where the Domestic data comes from, the FX
rules, the reading cautions, globe drill-down behaviour, and troubleshooting.
*《MTC Sales Report Import SOP.docx》第 5 節的配套附錄,配合 2026-11-09 建置週期
(十一個分頁的報表)新增。內容:國內訂單資料來源、匯率規則、閱讀注意事項、地球儀
下鑽行為,以及疑難排解。*

## 1. Where the Domestic data comes from / 國內訂單資料來源

**EN.** The Domestic tab is fed by `Domestic Orders 2026–2027 — MJP _ MCN _ MZQ _ MIN
(for Accounts).xlsx`, the monthly file Peggy/Keith provide for Accounts. Its single
sheet "Summary" holds TABLE A (sales, rows under a "Channel" header) and TABLE B
(commission/royalty), one row per entity (MJP, MIN, MCN, MZQ) in its own local
currency. The generator reads it via the wildcard `Domestic Orders*.xlsx` — never a
hardcoded literal — because the real filename contains en/em dashes that PowerShell 5.1
cannot reliably match. **As-of convention:** the workbook carries an "As of `<date>`
(Asia/Hong_Kong)" line; the generator captures it into `domestic.json` (`asOf`, e.g.
"6 Nov 2026") and the page prints it in the footnote. The MJP figures originate from an
external Google Sheet linked in the workbook — treat that row as the owner's data, not
MTC-recomputed data.

**中文。** Domestic 分頁的資料來源是 `Domestic Orders 2026–2027 — MJP _ MCN _ MZQ _ MIN
(for Accounts).xlsx`,即 Peggy/Keith 每月提供給 Accounts 的檔案。唯一的 "Summary"
工作表內有 TABLE A(銷售)與 TABLE B(佣金/權利金),每列一個實體(MJP、MIN、MCN、
MZQ),以各自的當地貨幣計價。產生器以萬用字元 `Domestic Orders*.xlsx` 讀取此檔——
絕不寫死完整檔名——因為真實檔名含有 PowerShell 5.1 比對不可靠的長短破折號
(EN/EM dash)。**基準日慣例:** 工作簿內有 "As of <日期> (Asia/Hong_Kong)" 一行,
產生器把它寫入 `domestic.json`(`asOf`,例如 "6 Nov 2026"),頁面備註會原樣標示。
MJP 數字來自工作簿內連結的外部 Google 試算表——應視為資料擁有者的數字,而非 MTC
重算的結果。

## 2. FX rules / 匯率規則

**EN.** All conversions are HKD-based (`fx_domestic.json`):

- JPY and CNY (and every other standard currency): taken from
  `STANDARD EXCHANGE RATE-YYYY-MM.xls`, sheet "Exchange rate table" — JPY 0.05,
  CNY 1.16, as-of the date cell (26-Aug-2026 for the 2026-09 file).
- USD: the same table, 7.8 HKD/USD.
- INR: **not** in the standard table. It is synthesised from the monthly-average
  USD/INR read from the newest `FIN_Ex Rate_PSBUILD_YYYYMMDD.xlsx` in
  `C:\Users\User\Desktop\Weekly Exchange Report - CLAUDE\Output` — newest by
  **filename date, never write time** (Sep-2026 average: USD/INR 95.50, so
  INR→HKD = 7.8 / 95.50 = 0.0816754). The Source cell of that row records the exact
  weekly file used.
- USD-equivalent of a local figure = HkdRate ÷ 7.8 (the USD row's HKD rate).
- Join trap: the Domestic workbook writes **"RMB"** while the FX table uses **"CNY"**.
  The generator aliases RMB→CNY; if this ever breaks, both China rows (MCN, MZQ)
  vanish silently — check the Domestic tab row count first.

**中文。** 所有換算以 HKD 為基準(`fx_domestic.json`):

- JPY 與 CNY(及其他標準貨幣):取自 `STANDARD EXCHANGE RATE-YYYY-MM.xls` 的
  "Exchange rate table"——JPY 0.05、CNY 1.16,基準日為表內日期儲存格
  (2026-09 檔為 26-Aug-2026)。
- USD:同一表,7.8 HKD/USD。
- INR:**標準表內沒有**,改由
  `C:\Users\User\Desktop\Weekly Exchange Report - CLAUDE\Output` 中**依檔名日期**
  (絕非寫入時間)最新的 `FIN_Ex Rate_PSBUILD_YYYYMMDD.xlsx` 讀取月均 USD/INR
  合成(2026 年 9 月均價 USD/INR 95.50,故 INR→HKD = 7.8 / 95.50 = 0.0816754)。
  該列 Source 欄會記錄實際使用的週報檔案。
- 當地金額的 USD 等值 = HkdRate ÷ 7.8(USD 列的 HKD 匯率)。
- 對應陷阱:國內工作簿用 **"RMB"**,匯率表用 **"CNY"**。產生器以 RMB→CNY 別名處理;
  此對應一旦失效,兩個中國列(MCN、MZQ)會**無聲消失**——請先檢查 Domestic
  分頁的列數。

## 3. Reading cautions / 閱讀注意事項

**EN.** These cautions come from the source workbook itself and are printed as page
notes; do not "clean them up":

- **Do not sum across currencies.** TABLE A/B show local-currency figures only;
  JPY + INR + RMB has no meaning and the workbook deliberately carries no combined
  DOMESTIC subtotal.
- **The USD column is a derived memo.** It is computed at build time at the rates
  above and is *not* included in the MA/DI USD totals elsewhere in the report —
  never add it to other USD figures.
- **No domestic YoY** until the 2025 comparison set is confirmed with the data
  owners (the workbook marks "2025 omitted (pending confirmation)"); accordingly
  there is no Total Index KPI row on this tab.
- Cancelled orders are excluded from the source counts.

**中文。** 以下注意事項來自來源工作簿本身,頁面照樣標示,請勿「美化」移除:

- **不可跨貨幣加總。** TABLE A/B 僅以當地貨幣列示;JPY + INR + RMB 沒有意義,
  工作簿刻意不含國內合計。
- **USD 欄為建置時推算的備忘數。** 依上述匯率換算,且**未**納入報表其他處的
  MA/DI USD 合計——切勿與其他 USD 數字相加。
- **國內暫不做年度比較:** 在與資料擁有者確認 2025 年比較基礎之前不做 YoY
  (工作簿註明 "2025 omitted (pending confirmation)"),故本分頁不設 Total Index KPI 列。
- 來源已剔除取消的訂單。

## 4. Globe drill-down / 地球儀下鑽

**EN.** The Dashboard's full-width 3D earth replaces the old Geographic bar and Origin
pie visually; both charts stay in the file as an automatic fallback when WebGL fails.
The toggle switches between **Global Sales** (by destination — `Cust_Countries` ×
`Net_Sales_USD`) and **Global Purchase** (by country of origin — `Country of origin` ×
PO cost). Clicking a country marker posts the existing drill-down message
(`mtcDrilldown`) and jumps to the **Data** tab pre-filtered to that country — same
contract as every other clickable number in the report. Mapping note: the
**MIDDLE-EAST** destination is plotted at the **Riyadh** hub (24.71 N, 46.71 E) as a
single constant; if the hub ever moves, update `MIDDLE_EAST_HUB`, not the data.

**中文。** Dashboard 頂端的全幅 3D 地球儀在視覺上取代原本的 Geographic 長條圖與
Origin 圓餅圖;兩張舊圖仍留在檔內,當 WebGL 無法啟用時自動退回。切換模式為
**Global Sales**(依銷售目的地——`Cust_Countries` × `Net_Sales_USD`)與
**Global Purchase**(依原產國——`Country of origin` × PO 成本)。點擊國家標記會沿用
報表既有的下鑽訊息(`mtcDrilldown`),跳到 **Data** 分頁並自動篩選該國——與報表中
其他可點數字的行為一致。對應備註:**MIDDLE-EAST** 目的地依使用者確認繪於
**利雅德 (Riyadh)** 中樞(北緯 24.71、東經 46.71),以單一常數表示;日後若調整,
請改 `MIDDLE_EAST_HUB` 常數,不要改資料。

## 5. Troubleshooting / 疑難排解

**EN.**

- *Missing weekly PSBUILD file (INR source).* The generator prints a loud
  `WARNING: FX FALLBACK ...` line to the console and continues with USD/INR 95.50;
  the Domestic tab footnote then shows a **`FALLBACK`** provenance string instead of
  the normal "weekly-avg `<month>` USD/INR `<rate>` (`<file>`)" text. Fix: run the Weekly
  Exchange Report rebuild so a `FIN_Ex Rate_PSBUILD_YYYYMMDD.xlsx` exists in its
  Output folder, then re-run the MTC generator.
- *Missing Domestic Orders workbook or standard FX .xls.* The generator **throws** —
  unlike the Family-master lookup, these inputs are not allowed to degrade silently.
  Confirm `Domestic Orders*.xlsx` is in the project root and the FX .xls matches
  `STANDARD EXCHANGE RATE-YYYY-MM.xls`.
- *China rows (MCN/MZQ) missing on the Domestic tab.* Suspect the RMB→CNY alias in
  section 2.
- *Globe not rendering.* WebGL/blob import failed → the Geographic bar and Origin
  pie show automatically; check the browser console and report size (< 11 MB email
  budget).

**中文。**

- *週報 PSBUILD 檔案缺失(INR 來源)。* 產生器會在主控台輸出醒目的
  `WARNING: FX FALLBACK ...` 警告,並改用 USD/INR 95.50 繼續執行;Domestic 分頁
  備註此時會顯示 **`FALLBACK`** 字樣,而非正常的
  "weekly-avg `<月份>` USD/INR `<匯率>` (`<檔名>`)"。處理:先執行週報重建,使 Output
  資料夾出現 `FIN_Ex Rate_PSBUILD_YYYYMMDD.xlsx`,再重跑 MTC 產生器。
- *國內訂單工作簿或標準匯率 .xls 缺失。* 產生器會**直接中斷 (throw)**——與
  Family 對照表不同,這兩項輸入不允許無聲降級。請確認專案根部有
  `Domestic Orders*.xlsx`,且匯率檔符合 `STANDARD EXCHANGE RATE-YYYY-MM.xls` 命名。
- *Domestic 分頁看不到中國列 (MCN/MZQ)。* 先查第 2 節的 RMB→CNY 別名對應。
- *地球儀未渲染。* 代表 WebGL/blob 匯入失敗,系統自動退回 Geographic 長條圖與
  Origin 圓餅圖;請檢查瀏覽器主控台,並留意成品大小(電郵上限 11 MB)。
