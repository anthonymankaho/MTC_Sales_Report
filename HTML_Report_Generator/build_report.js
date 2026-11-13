const fs = require('fs');
const path = require('path');

// build_report.js <workDir> <outputHtmlPath>
const SCRIPT_DIR = __dirname; // vendor/ (jquery, DataTables, Chart.js) lives here
const WORK_DIR = process.argv[2];
const OUT_PATH = process.argv[3];
if (!WORK_DIR || !OUT_PATH) {
  console.error('Usage: node build_report.js <workDir> <outputHtmlPath>');
  process.exit(1);
}
const DIR = WORK_DIR;

function neutralize(jsonStr) { return jsonStr.replace(/<\/script/gi, '<\\/script'); }
function readJson(fileName) {
  let raw = fs.readFileSync(path.join(DIR, fileName), 'utf8');
  if (raw.charCodeAt(0) === 0xFEFF) raw = raw.slice(1);
  return JSON.parse(raw);
}

// ---- load DATA sheet (header + rows), build name -> column index map ----
const dataSheet = readJson('data_sheet.json');
const H = dataSheet.header;
const COL = {};
H.forEach((name, i) => { COL[name] = i; });
const REQUIRED_HEADER_GROUPS = [
  ['Year'], ['yyyy-mm'], ['Cust_name'], ['VEND_NAME'], ['Cust_Countries'],
  ['Country of origin'], ['Net_Sales_USD'], ['Ttl PO USD', 'Matl_USD'],
  ['TTL PO NET CONT USD', 'Net_Cont_USD'], ['Family Category'],
  ['Fabrication Code'], ['Family1'], ['Part_id']
];
const missingHeaders = REQUIRED_HEADER_GROUPS
  .filter(group => !group.some(name => Object.prototype.hasOwnProperty.call(COL, name)))
  .map(group => group.join(' or '));
if (missingHeaders.length) {
  throw new Error('Missing required DATA header(s): ' + missingHeaders.join(', '));
}
function resolveColumnIndex(candidates) {
  const name = candidates.find(candidate => Object.prototype.hasOwnProperty.call(COL, candidate));
  return name === undefined ? -1 : COL[name];
}
console.log('Data loaded:', dataSheet.rows.length, 'rows,', H.length, 'cols');

// ---- vendor assets ----
const jqueryJs = fs.readFileSync(path.join(SCRIPT_DIR, 'vendor', 'jquery.min.js'), 'utf8');
const dataTablesJs = fs.readFileSync(path.join(SCRIPT_DIR, 'vendor', 'dataTables.min.js'), 'utf8');
const dataTablesCss = fs.readFileSync(path.join(SCRIPT_DIR, 'vendor', 'dataTables.min.css'), 'utf8');
const chartJs = fs.readFileSync(path.join(SCRIPT_DIR, 'vendor', 'chart.min.js'), 'utf8');
// ---- GLOBE assets: three.js ESM (vendor/lib, byte-identical copy of the GPR
// 3D Earth Dashboard build), land topology (vendor/data) and the ported engine
// (globe/globe.js), all read at build time and inlined into the Dashboard.
const threeSrc = fs.readFileSync(path.join(SCRIPT_DIR, 'vendor', 'lib', 'three.module.min.js'), 'utf8');
const landSrc = fs.readFileSync(path.join(SCRIPT_DIR, 'vendor', 'data', 'land-110m.json'), 'utf8');
const globeJsSrc = fs.readFileSync(path.join(SCRIPT_DIR, 'globe', 'globe.js'), 'utf8');
// Island escape discipline, ported from GPR build/build.mjs:67-76. Every JSON
// island rides as JSON *text* inside an inert <script type="application/json">
// tag; a single raw '<' would let '</script' truncate that tag once the srcdoc
// iframe re-parses the sheet, and the double-embedding here (island-in-sheet,
// sheet-in-SHEETS JSON) makes this file's own neutralize() (:14) insufficient
// on its own. So: replaceAll every '<' with the backslash-u003c unicode escape
// (JSON.parse restores it), then REFUSE to write output if either dangerous sequence is
// still present, exactly like build.mjs:74-76's throw-before-write guard.
function escapeIsland(jsonText, label) {
  const escaped = jsonText.replaceAll('<', '\\u003c');
  if (escaped.includes('</script') || escaped.includes('<!--')) {
    throw new Error('GLOBE BUILD GUARD: ' + label + ' contains a script-terminating sequence - refusing to write output');
  }
  return escaped;
}
function guardInlineScript(src, label) {
  if (src.includes('</script') || src.includes('<!--')) {
    throw new Error('GLOBE BUILD GUARD: ' + label + ' contains a script-terminating sequence - refusing to write output');
  }
  return src;
}
const threeIsland = escapeIsland(JSON.stringify(threeSrc), 'three island');
const landIsland = escapeIsland(JSON.stringify(JSON.parse(landSrc)), 'land island');
guardInlineScript(globeJsSrc, 'globe/globe.js');

// ---------------------------------------------------------------------------
// DOMESTIC tab chrome strings, ENGLISH side, kept as ONE Node map instead of
// being typed into I18N_JS.en by hand: buildDomesticHtml below also has to
// bake the same English copy into the page as its pre-applyLanguage default
// (the INTRO_EN trick at :313 exists for exactly that reason), and INTRO_EN's
// hand-duplicated strings are already known to drift from I18N's copy (see
// the "5 charts vs 6 charts" divergence). So DOM_EN is the single source and
// DOM_EN_JS splices it into I18N_JS.en as generated object entries.
// Its zh twins are written out literally in I18N_JS.zh - the completeness of
// BOTH dictionaries for every key this page uses is asserted, not assumed, by
// _work/../_domestic_assert.js (t() at :143 falls back to the raw key, so a
// missing zh entry leaks English into 中文 mode with no error at all).
// {token} placeholders are filled from the computed data (DOM_FILLS) on both
// sides, so one string serves both languages.
// ---------------------------------------------------------------------------
const DOM_EN = {
  tab_domestic: 'Domestic',
  title_domestic: 'Domestic Orders 2026–2027 (MJP · MCN · MZQ · MIN)',
  introDomestic: "Domestic orders and MTC commission / royalty, shown in each entity's own currency exactly as the source workbook records them. Every USD figure on this page is derived at build time for reference only — the workbook's own MA/DI USD totals do not include them.",
  domActionSales: '2026 domestic sales cover {entities} entities across {currencies} currencies, equivalent to USD {usd} on a derived basis.',
  domActionCommission: 'MTC commission / royalty on those orders derives to USD {usd}, about {pct}% of the sales figure.',
  domActionPending: 'Only {names} carries a 2027 outstanding balance: {sales} of sales and {comm} of commission.',
  domSecSales: 'Exhibit 1 · Table A — domestic sales',
  domSecCommission: 'Exhibit 2 · Table B — MTC commission / royalty',
  domColEntity: 'Entity / Customer',
  domColMarket: 'Market',
  domColShipped: '2026 Shipped',
  domColOutstanding: '2026 Outstanding',
  domColTotal2026: '2026 Total',
  domColOutstanding2027: '2027 Outstanding',
  domColUsd: 'USD equivalent (derived)',
  domColSalesShort: 'Domestic sales (Table A)',
  domColCommissionShort: 'MTC commission (Table B)',
  domSubtotal: 'Local subtotal',
  domCardUsdLabel: 'Derived USD',
  domUsdBoxTitle: 'Derived USD memo — cross-currency, USD only',
  domUsdBasisCol: 'Measure',
  domUsdTotalLabel: '2026 total (all currencies)',
  domUsdBoxCaption: 'Derived at build time — Not included in MA/DI USD totals (source workbook policy)',
  domBasisFlagTitle: 'No 2026 activity on this row — converted from its 2027 outstanding balance.',
  domCautionTitle: 'Caution carried over from the source workbook',
  domCautionNoSum: 'Local-currency amounts must never be added across currencies.',
  domNoYoy: 'No year-on-year comparison — 2025 is omitted pending confirmation.',
  domCancelled: 'Cancelled orders are excluded.',
  domAsOf: 'Domestic data as of',
  domFootnotesTitle: 'Sources and notes',
  domSourceFile: 'Source workbook',
  domSourceSheetNote: 'sheet "Summary" — Table A (sales) and Table B (commission / royalty), read by anchored search, not fixed rows.',
  domFxSource: 'Exchange rates',
  domInrSource: 'INR rate basis',
  domUsdBasisNote: 'USD basis',
  domUsdBasisFormula: 'USD per unit = that currency’s HKD rate ÷ the USD row’s HKD rate.',
  domUsdRowSkipNote: 'No USD subtotal down the column: a currency card can hold rows on different bases (2026 and 2027), so derived USD figures are listed per basis in the memo line under each table.',
  domMjpSource: 'MJP external source',
  domMjpNote: 'MJP arrives from outside this workbook family; confirm against the source before relying on it.',
  domFxMissingLabel: 'Currency with no rate found',
  domFxMissingNote: 'shown with no USD equivalent rather than dropped — its local figures below stay in the report.',
  domInrFallbackWarn: 'Fallback INR rate in use',
  domIslandsLabel: 'Source islands'
};
const DOM_EN_JS = Object.keys(DOM_EN).map(k => '    ' + k + ': ' + JSON.stringify(DOM_EN[k]) + ',').join('\n');

// =========================================================================
// ACTION_EN - the static bilingual "conclusion line" (McKinsey action title)
// for each of the eight table pages, approved as fixed copy by the user in
// the restyle brief (the Dashboard and Domestic pages compute theirs from
// live numbers instead). Same single-source pattern as DOM_EN above: this map
// is spliced into I18N_JS.en by ACTION_EN_JS, the zh twins are the literal
// lines in I18N_JS.zh, and the factory bakes the English value into the page
// as the pre-applyLanguage default - so no INTRO_EN-style drift is possible.
// =========================================================================
const ACTION_EN = {
  actionData: 'Every number in this report recomputes from this one order-level ledger.',
  actionExrate: 'The whole report’s USD basis in a single table - one Platinum Buy Rate per currency.',
  actionYoyCustomers: 'Year-on-year swing customer by customer - the Total Index above holds the whole-business view.',
  actionYoyFabrication: 'The same 2026-vs-2025 swing, cut by fabrication route.',
  actionYoyProdFamily: 'The same swing at product level - Brand rows roll up their Family1 detail.',
  actionYoyPartId: 'The swing taken down to individual part numbers - where the delta actually originates.',
  actionYoyNetNetMargin: 'The tighter Net Net basis, not invoice NET - this comparison reads lower by construction.',
  actionCustomerQuarter: 'Net margin per customer across 2026’s four quarters - a within-year read, not a two-year one.',
};
const ACTION_EN_JS = Object.keys(ACTION_EN).map(k => '    ' + k + ': ' + JSON.stringify(ACTION_EN[k]) + ',').join('\n');

// ---------------------------------------------------------------------------
// Language toggle (EN / 繁體中文) - translates titles, headers, labels, and
// buttons only. Table column headers and every data value stay untouched -
// those are the report's actual data, not UI chrome, and translating e.g.
// customer names or "Cust_name" itself would be wrong. Shared across every
// page (embedded like SHARED_CSS/MULTISELECT_JS) so each iframe keeps its own
// currentLang in sync with the outer shell's toggle.
// ---------------------------------------------------------------------------
const I18N_JS = `
var I18N = {
  en: {
    tab_dashboard: 'Dashboard', tab_data: 'Data', tab_exrate: 'Exchange Rate',
    tab_yoyCustomers: 'YoY: Customers', tab_yoyFabrication: 'YoY: Fabrication',
    tab_yoyProdFamily: 'YoY: Prod Family', tab_yoyPartId: 'YoY: Part ID',
    tab_yoyNetNetMargin: 'YoY: Net Net Margin', tab_customerQuarter: 'Customer-Quarter', tab_purchaseYtdVendor: 'Purchase: YTD by Vendor',
    title_data: 'Data', title_exrate: 'Exchange Rate',
    title_yoyCustomers: '2026 v 2025 by Customers', title_yoyFabrication: '2026 v 2025 by Fabrication',
    title_yoyProdFamily: '2026 v 2025 by Prod Family', title_yoyPartId: '2026 v 2025 by Part ID',
    title_yoyNetNetMargin: '2026 v 2025 by NET NET MARGIN', title_customerQuarter: 'Net Margin by Customer-Quarter',
    rowsSuffix: 'rows', exportCsv: 'Export CSV (filtered)', clearFilters: 'Clear all filters / search',
    clearFiltersTooltip: 'Clear search box + all dropdown filters',
    resetColWidths: 'Reset column widths',
    allLabel: 'All', typeToSearchLabel: 'Type to search...', selectAllLabel: 'Select all', clearSelLabel: 'Clear', selectedLabel: 'selected',
    expandAllLabel: 'Expand', collapseAllLabel: 'Collapse',
    drillTooltip: "Click to see the underlying transactions in the Data tab",
    dashboardH2: 'MTC Sales & Margin Performance',
    filterYear: 'Year', filterMonth: 'Month', filterCustomer: 'Customer', filterVendor: 'Vendor',
    filterCountry: 'Country', resetFilters: 'Reset filters',
    kpiNetSales: 'Net Sales (USD)', kpiNetCost: 'Net Cost (USD)', kpiNetMargin: 'Net Margin (USD)', kpiNetMarginPct: 'Net Margin %',
    kpiVsLabel: 'vs same months last year',
    kpiDomSales: 'Domestic Net Sales (USD, derived)', kpiDomCommission: 'Domestic MTC Commission (USD, derived)',
    domBasis: '2026 total', domAsOf: 'snapshot as of',
    domStripNote: '† Derived USD memo from the Domestic Orders snapshot — Not included in MA/DI USD totals (source workbook policy). Click a card to open the Domestic tab.',
    trendSalesCostTitle: 'Monthly Trend: Sales & Cost (USD), 2025 vs 2026', trendMarginTitle: 'Monthly Trend: Margin (USD), 2025 vs 2026',
    seriesSales: 'Sales', seriesCost: 'Cost', seriesMargin: 'Margin',
    legendNetSales: 'Net Sales', legendNetCost: 'Net Cost', legendNetMargin: 'Net Margin',
    vendorTitle: 'Top Vendors by Cost', customerTitle: 'Top Customers by Sales',
    geoTitle: 'Geographic Performance by Margin', originTitle: 'Country of Origin by Sales',
    globeKicker: 'Global Trade Flow', globeTitle: 'Where the money goes - and where it comes from',
    globeModeSales: 'Global Sales', globeModePurchase: 'Global Purchase',
    globeBuilding: 'Building the 3D earth...', globeFailed: '3D view unavailable on this machine - the Geographic (Margin) and Origin charts below stay in view as the fallback.',
    globeHelp: 'Drag to rotate - hover a marker for the readout - click one to open the Data tab pre-filtered',
    globeNote: 'Marker hubs are indicative capital/economic-city positions; every USD figure comes straight from the embedded DATA ledger (Net Sales / Net Cost).',
    globeUnitSales: 'Net Sales (USD) by destination', globeUnitPurchase: 'Net Cost (USD) by origin',
    globeAllMarkets: 'All markets', globeLegendTitle: 'Geographic performance - click a row to drill in',
    globeSizeLegend: 'marker size scales with value',
    globeUnlocatedNote: 'No hub coordinate - no marker, but still counted in every figure:',
    aiBtn: 'Ask AI', aiTitle: 'AI assistant (Qwen)', aiHello: 'Hi - ask about the numbers on this dashboard (totals, YoY, countries, domestic memo). I can suggest filters to apply too.', aiPh: 'e.g. Which month had the lowest sales?', aiSend: 'Send', aiApply: 'Apply', aiOffline: 'AI service offline - check the internet connection, then close and reopen the assistant. Everything else works normally.', aiOnline: 'Connected', aiClear: 'Clear', aiThinking: 'Thinking...', aiTrunc: '(answer truncated)', aiErrNet: 'No reply from the AI service - check your connection and send again.', aiQuota: 'quota {U} / {C} tokens this month', aiName: 'Name:', aiDisclose: 'Sent to the AI service (Cloudflare; falls back to the local office relay only when Cloudflare is unreachable): current filters and the aggregate figures on this page (KPIs, monthly, country top lists). Never sent: transaction rows or credentials. No access code is needed - the monthly token quota is the limit.',
    otherCategory: 'Others',
    totalIndexHeading: 'Total Index',
    weightedLabel: 'weighted',
    yoyLabel: 'YoY', vsLabel: 'vs', varianceLabel: 'Variance', variancePctLabel: 'Variance %',
    yoyChangeCol: '2025 vs 2026 Net Sales % Change',
    introDashboard: 'Filter by year, month, customer, vendor or country - the 3D trade globe, 4 KPIs and 6 charts update instantly.',
    introData: 'The full transaction ledger behind every number in this report - click any highlighted value on another tab and it lands here, pre-filtered.',
    introExRate: 'The Platinum Buy Rate table used to convert every foreign-currency figure in this report to USD.',
    introYoyCustomers: '2026 vs 2025 by customer, with a Total Index above the table - click any name, country or number to see the transactions behind it.',
    introYoyFabrication: 'The same year-over-year comparison as the Customers tab, grouped by Fabrication Code.',
    introYoyProdFamily: 'The same year-over-year comparison as the Customers tab, grouped by Product Family - click + next to a Brand to break it down by Family1.',
    introYoyPartId: 'The same year-over-year comparison as the Customers tab, taken down to individual Part ID.',
    introYoyNetNetMargin: 'The same comparison as the Customers tab, built on the Net Net (further-deducted) basis instead of the invoice-level NET figure.',
    introCustomerQuarter: 'Net margin by customer across the four quarters of 2026, rather than a two-year comparison.',
    langLabel: 'EN',
    exhibitWord: 'Exhibit', footSource: 'Sources',
    // The backslash-u escapes below are MANDATORY at these two spots: this
    // whole dictionary lives inside one outer Node template literal, so a
    // literal dollar-brace would be evaluated as a (nonexistent) JS
    // interpolation at build time and crash node build_report.js. Cooked,
    // the page's own single-quoted string receives a plain "$" - the
    // {S}/{M}/{P}/{D}/{T} tokens are substituted client-side by renderAction().
    actionDashboard: 'Net sales US\\u0024{S} for {T} — {D} vs prior year; margin US\\u0024{M} ({P}).',
${DOM_EN_JS}
${ACTION_EN_JS}
  },
  zh: {
    tab_dashboard: '總覽', tab_data: '明細資料', tab_exrate: '匯率',
    tab_yoyCustomers: '年度比較：客戶', tab_yoyFabrication: '年度比較：加工類別',
    tab_yoyProdFamily: '年度比較：產品家族', tab_yoyPartId: '年度比較：料號',
    tab_yoyNetNetMargin: '年度比較：淨淨毛利', tab_customerQuarter: '客戶季度毛利', tab_purchaseYtdVendor: '採購：年度供應商排行',
    title_data: '明細資料', title_exrate: '匯率',
    title_yoyCustomers: '2026 對比 2025：依客戶', title_yoyFabrication: '2026 對比 2025：依加工類別',
    title_yoyProdFamily: '2026 對比 2025：依產品家族', title_yoyPartId: '2026 對比 2025：依料號',
    title_yoyNetNetMargin: '2026 對比 2025：依淨淨毛利', title_customerQuarter: '客戶季度淨毛利',
    rowsSuffix: '筆', exportCsv: '匯出CSV（僅目前篩選結果）', clearFilters: '清除所有篩選／搜尋',
    clearFiltersTooltip: '清除搜尋框及所有下拉篩選',
    resetColWidths: '重設欄寬',
    allLabel: '全部', typeToSearchLabel: '輸入以搜尋...', selectAllLabel: '全選', clearSelLabel: '清除', selectedLabel: '項已選',
    expandAllLabel: '展開', collapseAllLabel: '收合',
    drillTooltip: '點擊查看Data頁籤中的相關交易明細',
    dashboardH2: 'MTC 銷售與毛利績效',
    filterYear: '年度', filterMonth: '月份', filterCustomer: '客戶', filterVendor: '供應商',
    filterCountry: '國家', resetFilters: '重設篩選',
    kpiNetSales: 'Net銷售額 (USD)', kpiNetCost: 'Net成本 (USD)', kpiNetMargin: 'Net毛利 (USD)', kpiNetMarginPct: 'Net毛利率',
    kpiVsLabel: '較去年同期',
    kpiDomSales: '國內 Net銷售額 (USD,推算)', kpiDomCommission: '國內 MTC佣金 (USD,推算)',
    domBasis: '2026合計', domAsOf: '快照截至',
    domStripNote: '† 推算 USD 附註(取自 Domestic Orders 快照) — Not included in MA/DI USD totals(未納入 MA／DI 的 USD 合計)。點卡片可開啟 Domestic 頁。',
    trendSalesCostTitle: '月度趨勢：銷售額與成本 (USD)，2025 對比 2026', trendMarginTitle: '月度趨勢：毛利 (USD)，2025 對比 2026',
    seriesSales: '銷售額', seriesCost: '成本', seriesMargin: '毛利',
    legendNetSales: 'Net銷售額', legendNetCost: 'Net成本', legendNetMargin: 'Net毛利',
    vendorTitle: '主要供應商（依採購成本）', customerTitle: '主要客戶（依銷售額）',
    geoTitle: '地理績效（依毛利）', originTitle: '原產國（依銷售額）',
    globeKicker: '全球貿易流向', globeTitle: '資金的去向 — 與貨物的來源',
    globeModeSales: '全球銷售', globeModePurchase: '全球採購',
    globeBuilding: '正在建立 3D 地球…', globeFailed: '此裝置無法顯示 3D 視圖 — 下方的地理（毛利）與原產國圖表將持續顯示作為備援。',
    globeHelp: '拖曳旋轉 — 游標停留於標記查看數字 — 點擊標記可跳到 Data 頁並自動篩選',
    globeNote: '標記位置為示意之都城／經濟城市；所有 USD 數字直接取自內嵌 DATA 明細（Net銷售額／Net成本）。',
    globeUnitSales: '依目的地之 Net銷售額 (USD)', globeUnitPurchase: '依原產國之 Net成本 (USD)',
    globeAllMarkets: '全部市場', globeLegendTitle: '地區績效排行 — 點列可鑽入篩選',
    globeSizeLegend: '標記大小依金額縮放',
    globeUnlocatedNote: '無標示座標 — 不畫標記，但所有數字仍照常計入：',
    aiBtn: '問 AI', aiTitle: 'AI 助理(Qwen)', aiHello: '你好 — 可以問我這份 Dashboard 上的數字(合計、同比、國家、國內附註),我也會建議可直接套用的篩選。', aiPh: '例如:哪個月銷售最低?', aiSend: '送出', aiApply: '套用', aiOffline: 'AI 服務離線 — 請檢查網絡連線,然後關閉再打開助理。其餘功能不受影響。', aiOnline: '已連線', aiClear: '清除', aiThinking: '思考中…', aiTrunc: '(回覆被截斷)', aiErrNet: 'AI 服務沒有回應 — 請檢查連線後再送出。', aiQuota: '本月用量 {U} / {C} tokens', aiName: '名稱:', aiDisclose: '送給 AI 服務的內容(經 Cloudflare 雲端助理;只有喺 Cloudflare 無法連線時先至改用辦公室本機轉接站):目前篩選與頁面彙總數字(KPI、月度、國家前榜)。絕不送出:交易明細或任何憑證。毋須存取碼——每月 token 配額就係上限。',
    otherCategory: '其他',
    totalIndexHeading: '總計指標',
    weightedLabel: '加權',
    yoyLabel: '較去年同期', vsLabel: '對比', varianceLabel: '差異', variancePctLabel: '差異率',
    yoyChangeCol: '2025對比2026 Net銷售額 % 變動',
    introDashboard: '依年度、月份、客戶、供應商或國家篩選 — 3D 貿易地球、4 項 KPI 與 6 個圖表即時更新。',
    introData: '本報表所有數字背後的完整交易明細 — 點擊其他頁籤中任何標示的數值，都會跳到這裡並自動套用篩選。',
    introExRate: '用於將本報表所有外幣數字換算為 USD 的白金買入匯率表。',
    introYoyCustomers: '依客戶呈現 2026 對比 2025，表格上方附總計指標 — 點擊任何名稱、國家或數字可查看其背後的交易明細。',
    introYoyFabrication: '與客戶頁籤相同的年度比較邏輯，改依加工代碼分組。',
    introYoyProdFamily: '與客戶頁籤相同的年度比較邏輯，改依產品家族分組 - 點擊品牌左方的「+」可展開該品牌各Family1明細。',
    introYoyPartId: '與客戶頁籤相同的年度比較邏輯，細分至個別料號層級。',
    introYoyNetNetMargin: '與客戶頁籤相同的比較邏輯，改以 Net Net（進一步扣除後）基礎計算，而非發票層級 NET 數字。',
    introCustomerQuarter: '呈現 2026 年度內各客戶四個季度的淨毛利變化，而非兩年度比較。',
    exhibitWord: '圖表', footSource: '資料來源',
    actionDashboard: '淨銷售額 US\\u0024{S}（{T}）— 較去年同期 {D}；毛利 US\\u0024{M}（{P}）。',
    actionData: '本報表所有數字均由這一張訂單層級帳務明細重新計算。',
    actionExrate: '整份報表的 USD 基礎全在此表 — 每種貨幣一個白金買入匯率。',
    actionYoyCustomers: '逐客戶看年度漲跌 — 全公司總覽請看上方總計指標。',
    actionYoyFabrication: '同樣的 2026 對比 2025 漲跌，改依加工途徑切分。',
    actionYoyProdFamily: '同樣的漲跌拉到產品層級 — 品牌列已加總其 Family1 明細。',
    actionYoyPartId: '漲跌追進單一料號 — 差異實際發生的位置。',
    actionYoyNetNetMargin: '以較嚴的 Net Net 基礎而非發票 NET 比較 — 數字天然偏低。',
    actionCustomerQuarter: '2026 年四個季度各客戶淨毛利 — 年度內解讀，非兩年度比較。',
    langLabel: '中文',
    tab_domestic: '國內訂單',
    title_domestic: '國內訂單 2026–2027（MJP · MCN · MZQ · MIN）',
    introDomestic: '國內訂單與 MTC 佣金／權利金，一律依來源工作簿原樣以各實體自己的貨幣列示。本頁所有 USD 數字均為建置時推算，僅供參考 — 工作簿本身的 MA/DI USD 合計並未納入。',
    domActionSales: '2026 年國內訂單涵蓋 {entities} 個實體、{currencies} 種貨幣，推算等值 USD {usd}。',
    domActionCommission: '對應的 MTC 佣金／權利金推算為 USD {usd}，約占國內銷售額 {pct}%。',
    domActionPending: '僅 {names} 留有 2027 年未結數字：銷售 {sales}、佣金 {comm}。',
    domSecSales: '圖表 1 · 表 A — 國內銷售',
    domSecCommission: '圖表 2 · 表 B — MTC 佣金／權利金',
    domColEntity: '實體／客戶',
    domColMarket: '市場',
    domColShipped: '2026 已出貨',
    domColOutstanding: '2026 未結',
    domColTotal2026: '2026 合計',
    domColOutstanding2027: '2027 未結',
    domColUsd: 'USD 等值（推算）',
    domColSalesShort: '國內銷售（表 A）',
    domColCommissionShort: 'MTC 佣金（表 B）',
    domSubtotal: '本地貨幣小計',
    domCardUsdLabel: '推算 USD',
    domUsdBoxTitle: '推算 USD 備忘 — 僅以 USD 跨貨幣比較',
    domUsdBasisCol: '項目',
    domUsdTotalLabel: '2026 合計（全部貨幣）',
    domUsdBoxCaption: '建置時推算 — Not included in MA/DI USD totals(未納入 MA／DI 的 USD 合計)',
    domBasisFlagTitle: '本列 2026 年無交易 — 以 2027 年未結數字換算。',
    domCautionTitle: '來源工作簿之警示',
    domCautionNoSum: '不同貨幣的本地金額一律不可相加。',
    domNoYoy: '未做年度比較 — 2025 年數字因待確認而省略。',
    domCancelled: '已取消訂單不予計入。',
    domAsOf: '國內訂單資料截至',
    domFootnotesTitle: '資料來源與註示',
    domSourceFile: '來源工作簿',
    domSourceSheetNote: '「Summary」工作表 — 表 A（銷售）與表 B（佣金／權利金），以關鍵字定位擷取，非固定列號。',
    domFxSource: '匯率來源',
    domInrSource: 'INR 換算基礎',
    domUsdBasisNote: 'USD 換算基礎',
    domUsdBasisFormula: '每單位 USD＝該貨幣之 HKD 匯率 ÷ USD 列之 HKD 匯率。',
    domUsdRowSkipNote: 'USD 欄不设欄內小計：同一貨幣卡可能含 2026 與 2027 不同基礎的列，故推算 USD 金額改在各表下方之備忘列按基礎逐一列示。',
    domMjpSource: 'MJP 外部來源',
    domMjpNote: 'MJP 資料來自本工作簿系列之外，引用前請先與來源核對。',
    domFxMissingLabel: '找不到匯率的貨幣',
    domFxMissingNote: '該貨幣僅顯示為無 USD 等值，並不会被移除 — 其本地金額仍保留於本頁。',
    domInrFallbackWarn: 'INR 已改用備援匯率',
    domIslandsLabel: '來源資料島'
  }
};
var currentLang = 'en';
try { currentLang = localStorage.getItem('reportLang') || 'en'; } catch (e) {}
function t(key) { return (I18N[currentLang] && I18N[currentLang][key]) || I18N.en[key] || key; }
`;

// Shared dark-mode CSS block (token-based, same convention across every tab).
// McKinsey-grade retune (restyle stage): ink #051C2C as text, near-white paper
// surfaces, one navy accent family, muted brick #A63D40 reserved for negatives
// and cautions (--neg/--warn; DOMESTIC_CSS already reads var(--warn) with a
// literal fallback that this token now actually supplies), hairline-friendly
// borders, 13px base, generous page padding. Display serif (Georgia) is NOT
// here - it enters through MCKINSEY_CSS (below) on the headline/action-title
// roles only; everything else stays Calibri, no web fonts (offline file://).
const SHARED_CSS = `
:root {
  --bg: #F7F9FB; --card-bg: #FFFFFF; --text: #051C2C; --text-muted: #5B6B79;
  --accent: #1a5c8a; --border: #C9D4DC; --border-light: #E4EAEF; --shadow: rgba(5,28,44,0.06);
  --input-bg: #FFFFFF; --hover-bg: #EEF3F7; --neg: #A63D40; --warn: #A63D40;
}
body[data-theme="dark"] {
  --bg: #08131C; --card-bg: #12212B; --text: #E8EEF2; --text-muted: #9DB0BD;
  --accent: #5FB2C4; --border: #33495A; --border-light: #24353F; --shadow: rgba(0,0,0,0.45);
  --input-bg: #0E1A23; --hover-bg: #1B2B36; --neg: #D98F8E; --warn: #D98F8E;
}
body { font-family: Calibri, Arial, sans-serif; font-size: 13px; margin: 16px 20px; background: var(--bg); color: var(--text); }
table.dataTable { color: var(--text); }
th, td { white-space: nowrap; }
.summary-bar { margin-bottom: 8px; font-weight: bold; color: var(--accent); }
.toolbar { margin-bottom: 8px; display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
#exportBtn, #clearFiltersBtn, #resetBtn, #resetColWidthsBtn { padding: 6px 14px; cursor: pointer; border-radius: 4px; border: 1px solid var(--border); background: var(--input-bg); color: var(--text); }
.filter-group { display: flex; align-items: center; gap: 4px; font-size: 11px; color: var(--text-muted); }
.filter-group select { font-size: 11px; padding: 3px 4px; border-radius: 3px; border: 1px solid var(--border); background: var(--input-bg); color: var(--text); max-width: 160px; }
/* Multi-select + type-to-search filter widget (replaces a plain single-value
   <select> so a page like YoY Product Family can filter to several values at
   once, e.g. two or three Family1 values together) - a small toggle button
   opens a panel with a search box (narrows the checkbox list by substring,
   not the table itself) plus the checkboxes. Sized to match the surrounding
   .filter-group select it replaces; a .total-index-filters override below
   bumps it up to that section's larger KPI-card visual weight instead. */
.ms-widget { position: relative; display: inline-block; }
.ms-toggle { font-size: 11px; padding: 3px 6px; border-radius: 3px; border: 1px solid var(--border); background: var(--input-bg); color: var(--text); cursor: pointer; max-width: 160px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; text-align: left; }
/* display lives on the :not([hidden]) variant, not the base rule - an author
   stylesheet rule at equal specificity to the UA default [hidden]{display:
   none} wins the cascade (author beats UA), so a base "display: flex" here
   would keep the panel visually shown regardless of the hidden property. */
.ms-panel { position: absolute; top: 100%; left: 0; z-index: 50; margin-top: 2px; background: var(--card-bg); border: 1px solid var(--border); border-radius: 4px; box-shadow: 0 2px 8px var(--shadow); min-width: 200px; }
.ms-panel:not([hidden]) { display: flex; flex-direction: column; }
.ms-search { margin: 6px 6px 4px 6px; padding: 4px 6px; border-radius: 3px; border: 1px solid var(--border); background: var(--input-bg); color: var(--text); font-size: 12px; }
.ms-actions { display: flex; gap: 6px; margin: 0 6px 4px 6px; }
.ms-actions button { font-size: 11px; padding: 2px 8px; cursor: pointer; border-radius: 3px; border: 1px solid var(--border); background: var(--input-bg); color: var(--text); }
.ms-options { overflow-y: auto; max-height: 200px; padding: 2px 6px 6px 6px; }
.ms-options label { display: flex; align-items: center; gap: 6px; padding: 3px 4px; font-size: 12px; font-weight: normal; color: var(--text); cursor: pointer; white-space: nowrap; border-radius: 3px; }
.ms-options label:hover { background: var(--hover-bg); }
/* Brand/group collapse toggle (see buildBrandGroupedRows + the GROUP_BY
   client-side logic) - a group-total row's own toggle cell. */
.row-toggle { display: inline-block; width: 18px; height: 18px; line-height: 16px; text-align: center; border: 1px solid var(--border); border-radius: 3px; cursor: pointer; font-weight: bold; color: var(--accent); }
.row-toggle:hover { background: var(--hover-bg); }
body[data-theme="dark"] table.dataTable tbody tr { background-color: var(--card-bg); color: var(--text); }
body[data-theme="dark"] table.dataTable tbody tr.odd, body[data-theme="dark"] table.dataTable tbody tr.even { background-color: var(--card-bg); }
body[data-theme="dark"] table.dataTable.stripe tbody tr.odd, body[data-theme="dark"] table.dataTable.display tbody tr.odd { background-color: var(--hover-bg); }
body[data-theme="dark"] table.dataTable tbody tr:hover { background-color: var(--hover-bg) !important; }
body[data-theme="dark"] table.dataTable thead th, body[data-theme="dark"] table.dataTable thead td { color: var(--text); border-bottom-color: var(--border); }
body[data-theme="dark"] table.dataTable tbody td { border-top-color: var(--border-light); }
body[data-theme="dark"] .dataTables_wrapper .dataTables_length, body[data-theme="dark"] .dataTables_wrapper .dataTables_filter,
body[data-theme="dark"] .dataTables_wrapper .dataTables_info, body[data-theme="dark"] .dataTables_wrapper .dataTables_processing,
body[data-theme="dark"] .dataTables_wrapper .dataTables_paginate { color: var(--text); }
body[data-theme="dark"] .dataTables_wrapper .dataTables_length select, body[data-theme="dark"] .dataTables_wrapper .dataTables_filter input { background-color: var(--input-bg); color: var(--text); border-color: var(--border); }
body[data-theme="dark"] .dataTables_wrapper .dataTables_paginate .paginate_button { color: var(--text) !important; }
body[data-theme="dark"] .dataTables_wrapper .dataTables_paginate .paginate_button.current { background: var(--accent) !important; color: #fff !important; border-color: var(--accent) !important; }
body[data-theme="dark"] .dataTables_wrapper .dataTables_paginate .paginate_button:hover { background: var(--hover-bg) !important; color: var(--text) !important; }
.kpis { display: flex; gap: 16px; margin-bottom: 16px; flex-wrap: wrap; }
.kpi { background: var(--card-bg); padding: 16px 24px; border-radius: 6px; box-shadow: 0 1px 3px var(--shadow); flex: 1; min-width: 180px; }
.kpi .label { font-size: 12px; color: var(--text-muted); }
.kpi .value { font-size: 24px; font-weight: bold; color: var(--accent); text-align: right; }
.kpi .weighted-pct { font-size: 12px; color: var(--text-muted); text-align: right; margin-top: 2px; }
.page-intro { font-size: 13px; color: var(--text-muted); margin: 0 0 14px 0; max-width: 900px; line-height: 1.5; }
.charts { display: flex; flex-wrap: wrap; gap: 16px; margin-bottom: 16px; }
.chart-box { background: var(--card-bg); padding: 12px; border-radius: 6px; box-shadow: 0 1px 3px var(--shadow); flex: 1; min-width: 420px; }
.chart-box.wide { flex-basis: 100%; }
canvas { max-height: 320px; }
.filters { display: flex; gap: 16px; flex-wrap: wrap; background: var(--card-bg); padding: 12px; border-radius: 6px; margin-bottom: 16px; box-shadow: 0 1px 3px var(--shadow); }
.filters label.fl { display: block; font-size: 11px; color: var(--text-muted); margin-bottom: 4px; }
.ms { position: relative; width: 180px; }
.ms-toggle { width: 100%; text-align: left; padding: 5px 8px; background: var(--input-bg); color: var(--text); border: 1px solid var(--border); border-radius: 4px; cursor: pointer; font-size: 12px; }
.ms-panel { position: absolute; z-index: 20; top: 100%; left: 0; background: var(--card-bg); border: 1px solid var(--border); border-radius: 4px; box-shadow: 0 2px 8px var(--shadow); max-height: 260px; overflow-y: auto; width: 240px; padding: 6px; }
.ms-panel label { display: block; font-size: 12px; padding: 2px 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: var(--text); }
.ms-panel label.ms-all-row { border-bottom: 1px solid var(--border-light); margin-bottom: 4px; padding-bottom: 6px; font-weight: bold; }
.ms-search { display: block; width: 100%; box-sizing: border-box; margin-bottom: 6px; padding: 5px 8px; font-size: 12px; border: 1px solid var(--border); border-radius: 4px; background: var(--input-bg); color: var(--text); }
`;

// =========================================================================
// MCKINSEY_CSS - the consulting finish layer, emitted LAST in every style
// assembly site (table pages, Dashboard after GLOBE_CSS, Domestic after
// DOMESTIC_CSS, and the shell before its own rules), so its hairline/typo
// overrides win over the component blocks without editing them. Owns: the
// Georgia serif role for headlines + action titles only (body stays Calibri,
// no web fonts), uppercase letter-spaced exhibit kickers, source footnotes,
// shadows -> 1px hairlines on every card, and tabular-nums on figures.
// It also carries the --shell-* token family + its dark block, which the
// shell chrome (never including SHARED_CSS) reads - that is what makes the
// tab bar and welcome modal follow dark mode for the first time.
// =========================================================================
const MCKINSEY_CSS = `
:root {
  --serif: Georgia, 'Times New Roman', Times, serif;
  /* shell chrome tokens - light */
  --shell-bg: #F7F9FB; --shell-card: #FFFFFF; --shell-text: #051C2C; --shell-muted: #5B6B79;
  --shell-rule: #E4EAEF; --shell-accent: #1a5c8a;
  --shell-nav: #123F5E; --shell-nav-text: #EAF1F6; --shell-nav-active: #051C2C;
  --shell-nav-hover: #0C2C42; --shell-nav-sep: rgba(255,255,255,0.14);
}
body[data-theme="dark"] {
  /* shell chrome tokens - dark */
  --shell-bg: #051C2C; --shell-card: #12212B; --shell-text: #E8EEF2; --shell-muted: #9DB0BD;
  --shell-rule: #33495A; --shell-accent: #5FB2C4;
  --shell-nav: #0B1B26; --shell-nav-text: #D5E3EC; --shell-nav-active: #123F5E;
  --shell-nav-hover: #1B2B36; --shell-nav-sep: rgba(255,255,255,0.08);
}
/* Typographic roles: serif = headlines/action titles only, never body or data. */
h1, h2, .dom-title, .action-title { font-family: var(--serif); font-weight: 600; }
.action-title { font-size: 17px; line-height: 1.35; color: var(--text); margin: 0 0 6px 0; max-width: 1080px; }
.action-title.neg { color: var(--neg); }
.exhibit-kicker { font-size: 10px; letter-spacing: .12em; text-transform: uppercase; font-weight: 700; color: var(--text-muted); margin: 0 0 6px 0; }
.source-foot { margin: 20px 0 6px 0; padding-top: 8px; border-top: 1px solid var(--border-light); font-size: 10.5px; color: var(--text-muted); line-height: 1.55; max-width: 1120px; font-variant-numeric: tabular-nums; }
/* Hairline rules replace the raised-card look. */
.kpi, .chart-box, .filters, .globe-section, .dom-card, .dom-usdbox, .total-index-filters .filter-group, .ms-panel, .ms-search { box-shadow: none; border-radius: 3px; }
.kpi, .chart-box, .filters, .globe-section, .dom-usdbox, .total-index-filters .filter-group, .ms-panel { border: 1px solid var(--border-light); }
#exportBtn, #clearFiltersBtn, #resetBtn, #resetColWidthsBtn, .globe-mode, .row-toggle, .ms-toggle, .ms-panel, .ms-search { border-radius: 2px; }
.kpi { border-bottom: 2px solid var(--accent); }
.kpi-sub { margin-top: 6px; font-size: 11px; color: var(--text-muted); font-variant-numeric: tabular-nums; display: flex; gap: 8px; justify-content: space-between; flex-wrap: wrap; }
.kpi-delta { font-weight: 700; }
.kpi-delta.up { color: var(--accent); }
.kpi-delta.down { color: #A63D40; }
.kpis-dom { margin-top: -4px; }
.kpis-dom .kpi { border-bottom: 2px dashed var(--accent); flex: 0 1 240px; }
.kpi-dom-card { cursor: pointer; }
.kpi-dom-card:hover { background: var(--hover-bg); }
.kpi-dom-note { flex: 1 1 100%; align-self: center; font-size: 11px; color: var(--text-muted); padding: 2px 4px; }
.dom-memo-sup { font-size: 9px; color: var(--text-muted); }
.kpi .label { text-transform: uppercase; letter-spacing: .07em; font-size: 10.5px; }
.kpi .value { font-variant-numeric: tabular-nums; }
.chart-box { padding: 14px 16px; }
/* Every digit rendered as a figure takes tabular spacing - columns of amounts
   and KPI values align vertically instead of dancing with proportional widths. */
table.dataTable td, table.dataTable th, .dom-memo .v, .summary-bar, .kpi .weighted-pct { font-variant-numeric: tabular-nums; }
table.dataTable thead th { border-bottom: 1px solid var(--border); }
.summary-bar { font-weight: 600; }
.page-intro { margin: 0 0 16px 0; }
`;

// Reusable multi-select widget JS (All / one / multi), identical contract on every tab that uses it.
const MULTISELECT_JS = `
function createMultiSelect(containerId, values, onChange, defaultSelected) {
  var container = document.getElementById(containerId);
  // defaultSelected: optional array of values to start pre-checked instead of
  // "All" (e.g. Year defaulting to [2026]) - selected stays empty ("All") only
  // when every value or none is checked, same convention as always.
  var startAll = !defaultSelected || defaultSelected.length === 0 || defaultSelected.length === values.length;
  var selected = startAll ? new Set() : new Set(defaultSelected);
  var toggle = document.createElement('button');
  toggle.type = 'button'; toggle.className = 'ms-toggle';
  container.appendChild(toggle);
  var panel = document.createElement('div');
  panel.className = 'ms-panel'; panel.style.display = 'none';
  panel.addEventListener('click', function (e) { e.stopPropagation(); });
  container.appendChild(panel);
  // Search box: filters which option rows are VISIBLE as you type (substring,
  // case-insensitive). It never changes what's checked - typing to narrow a
  // long list (e.g. customer names) then checking a few, clearing the search
  // still shows every previous check.
  var searchBox = document.createElement('input');
  searchBox.type = 'text'; searchBox.className = 'ms-search'; searchBox.placeholder = 'Search...';
  searchBox.addEventListener('click', function (e) { e.stopPropagation(); });
  panel.appendChild(searchBox);
  var allRow = document.createElement('label');
  allRow.className = 'ms-all-row';
  var allCb = document.createElement('input'); allCb.type = 'checkbox'; allCb.checked = startAll;
  allRow.appendChild(allCb); allRow.appendChild(document.createTextNode(' (Select All)'));
  panel.appendChild(allRow);
  var checkboxes = [];
  var optionRows = [];
  values.forEach(function (v) {
    var row = document.createElement('label');
    var cb = document.createElement('input'); cb.type = 'checkbox'; cb.value = v;
    cb.checked = startAll || selected.has(v);
    row.appendChild(cb); row.appendChild(document.createTextNode(' ' + v));
    panel.appendChild(row); checkboxes.push(cb);
    optionRows.push({ row: row, searchText: String(v).toLowerCase() });
  });
  searchBox.addEventListener('input', function () {
    var q = searchBox.value.toLowerCase();
    optionRows.forEach(function (o) { o.row.style.display = o.searchText.indexOf(q) === -1 ? 'none' : ''; });
  });
  function updateToggleText() {
    if (selected.size === 0 || selected.size === values.length) toggle.textContent = 'All ▾';
    else if (selected.size === 1) toggle.textContent = Array.from(selected)[0] + ' ▾';
    else toggle.textContent = selected.size + ' selected ▾';
  }
  allCb.addEventListener('change', function () {
    checkboxes.forEach(function (cb) { cb.checked = allCb.checked; });
    selected = new Set(); updateToggleText(); onChange();
  });
  checkboxes.forEach(function (cb) {
    cb.addEventListener('change', function () {
      var checkedOnes = checkboxes.filter(function (c) { return c.checked; });
      if (checkedOnes.length === checkboxes.length || checkedOnes.length === 0) { selected = new Set(); allCb.checked = true; }
      else { selected = new Set(checkedOnes.map(function (c) { return c.value; })); allCb.checked = false; }
      updateToggleText(); onChange();
    });
  });
  toggle.addEventListener('click', function (e) {
    e.stopPropagation();
    document.querySelectorAll('.ms-panel').forEach(function (p) { if (p !== panel) p.style.display = 'none'; });
    var opening = panel.style.display === 'none';
    panel.style.display = opening ? 'block' : 'none';
    if (opening) {
      searchBox.value = '';
      optionRows.forEach(function (o) { o.row.style.display = ''; });
      searchBox.focus();
    }
  });
  updateToggleText();
  return {
    getSelected: function () { return Array.from(selected); },
    // AI chip + drill programmability: replace the whole selection with an
    // array of exact values (empty array = no filter, same as "select all").
    setSelected: function (vals) {
      selected = new Set(vals || []);
      checkboxes.forEach(function (cb) { cb.checked = selected.has(cb.value); });
      allCb.checked = selected.size === 0 || selected.size === checkboxes.length;
      updateToggleText();
    },
    reset: function () { selected = new Set(); allCb.checked = true; checkboxes.forEach(function (cb) { cb.checked = true; }); updateToggleText(); }
  };
}
document.addEventListener('click', function () { document.querySelectorAll('.ms-panel').forEach(function (p) { p.style.display = 'none'; }); });
function matchesFilter(value, selectedArr) { return selectedArr.length === 0 || selectedArr.indexOf(value) !== -1; }

`;

// =========================================================================
// English fallback text for each page's intro (${introKey} below) - mirrors
// I18N_JS.en's own copy of these same strings. Used only so the paragraph
// isn't empty before applyLanguage runs on the iframe's own load event (the
// same reason every other data-i18n element in this file carries literal
// English content too, not just a key) - I18N_JS stays the single source of
// truth for the actual translation logic.
const INTRO_EN = {
  introData: 'The full transaction ledger behind every number in this report - click any highlighted value on another tab and it lands here, pre-filtered.',
  introExRate: 'The Platinum Buy Rate table used to convert every foreign-currency figure in this report to USD.',
  introYoyCustomers: '2026 vs 2025 by customer, with a Total Index and Sales pie above the table - click any name, country or number to see the transactions behind it.',
  introYoyFabrication: 'The same year-over-year comparison as the Customers tab, grouped by Fabrication Code.',
  introYoyProdFamily: 'The same year-over-year comparison as the Customers tab, grouped by Product Family - click + next to a Brand to break it down by Family1.',
  introYoyPartId: 'The same year-over-year comparison as the Customers tab, taken down to individual Part ID.',
  introYoyNetNetMargin: 'The same comparison as the Customers tab, built on the Net Net (further-deducted) basis instead of the invoice-level NET figure.',
  introCustomerQuarter: 'Net margin by customer across the four quarters of 2026, rather than a two-year comparison.',
  // Domestic's English fallback is deliberately NOT retyped here - it points
  // at DOM_EN (above I18N_JS), the very object I18N_JS.en is spliced from, so
  // this map and the dictionary can never drift the way the introDashboard
  // copy has ("5 charts" at the dashboard's own <p> vs "6 charts" here).
  introDomestic: DOM_EN.introDomestic
};

// =========================================================================
// Generic DataTables page for a (possibly filtered) set of rows - used for
// Data, Exchange Rate, and every YoY comparison sheet.
// =========================================================================
// defaultSortCol: a column NAME to open sorted high-to-low. Only for sheets
// that have no year-paired metric columns for yoyRankCol to find (the
// Customer-Quarter sheet is quarterly, not 2025/2026-paired), so the two
// options never both apply to the same page.
// groupBy (optional): { groupColIdx, toggleColIdx, detailKeyColIdx } - turns
// on a Brand/parent -> Family1/child collapsible grouping (see
// buildBrandGroupedRows): toggleColIdx holds '+' on a group-total row and ''
// on every detail row, detailKeyColIdx is blank ONLY on a group-total row -
// that blankness is the one signal used everywhere below to tell the two
// row kinds apart, rather than a separate flag column.
// actionKey (13th, optional): I18N key of this page's static bilingual
// conclusion line (ACTION_EN above supplies the English pre-fill, same trick
// as introKey/INTRO_EN). sourceNote (14th, optional): build-time source
// footnote - island filename + report period + FX as-of text. Language-
// neutral data (filenames/dates), so it is emitted verbatim under a
// translated "Sources" label leaf in both modes; it never carries a {token}.
function buildDataTablesHtml(header, rows, title, scrollX, filterColumns, drilldown, titleKey, kpiPie, introKey, yoyRankCol, defaultSortCol, groupBy, actionKey, sourceNote) {
  if (scrollX === undefined) scrollX = true;
  if (!filterColumns) filterColumns = [];
  const defaultSortColIdx = defaultSortCol ? header.indexOf(defaultSortCol) : -1;
  if (defaultSortCol && defaultSortColIdx < 0) throw new Error('defaultSortCol "' + defaultSortCol + '" not found in ' + title + ' header');
  const headerJsonStr = neutralize(JSON.stringify(header));
  const rowsJsonStr = neutralize(JSON.stringify(rows));
  const filterColIdx = filterColumns.map(name => header.indexOf(name)).filter(i => i >= 0);
  const filterColIdxJsonStr = JSON.stringify(filterColIdx);
  const groupByJsonStr = groupBy ? JSON.stringify(groupBy) : 'null';
  // drilldown (optional): { labelColCount, dataColMap: {thisTableColName: DataTabColName} }
  // Every metric column past labelColCount becomes a clickable link - clicking
  // it switches to the Data tab filtered to exactly the transactions behind
  // that number (label columns' values + the year parsed from the clicked
  // column's own header, e.g. "SKU_qty (Data_AC) 2026" -> Year=2026).
  const drilldownJsonStr = drilldown ? neutralize(JSON.stringify(drilldown)) : 'null';
  const rowCountStr = rows.length.toLocaleString();
  return `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>${title}</title>
<style>${dataTablesCss}${SHARED_CSS}
.drill-link { color: var(--accent); cursor: pointer; text-decoration: underline dotted; }
.drill-link:hover { text-decoration: underline solid; }
/* !important: DataTables' own header CSS (table.dataTable thead th) outranks
   a plain .dt-num/th.dt-num rule on specificity alone, so the header cell's
   alignment silently stays left even though the class is correctly applied -
   only the body <td> cells would end up right-aligned without this. */
td.dt-num, th.dt-num { text-align: right !important; }
${scrollX ? '' : `
/* table-layout switches to fixed via JS, only after each column's natural
   width has been measured (see setupColumnResize) - fixed is what makes an
   explicit <th> width (drag-resize or a saved one) actually control the
   rendered column, instead of table-layout:auto silently re-measuring from
   cell content and ignoring it, but switching to fixed BEFORE every column
   has an explicit width makes the browser split space evenly instead,
   squeezing every un-resized column's header text into an illegible pile-up.
   Scoped to non-scrollX tables only - a scrollX table already manages a
   separate cloned header/body width-sync of its own (see Gotcha 7 in the
   skill), and layering manual resize on top of that on a 103-column table
   was not verified safe, so it's intentionally left out of scope here
   rather than shipped half-working. */
#dataTableHeadRow th { position: relative; }
.col-resize-handle { position: absolute; right: 0; top: 0; bottom: 0; width: 8px; cursor: col-resize; user-select: none; z-index: 5; }
.col-resize-handle:hover, .col-resize-handle.resizing { background: var(--accent); opacity: 0.5; }
/* Wrap text (like Excel's own "Wrap Text" toggle) instead of forcing every
   header/cell onto one line - overrides the global "th, td { white-space:
   nowrap }" default (see SHARED_CSS) for these tables specifically, since
   nowrap is what makes a manually-narrowed column just overflow/clip its
   text instead of growing the row to fit it, the way Excel would. */
#dataTable th, #dataTable td { white-space: normal; overflow-wrap: break-word; word-break: break-word; }
/* border-box: a plain style.width set from JS is a CONTENT-box width by
   default, so the cell's actual on-screen size ends up width + DataTables'
   own header/cell padding on top - the fit-to-screen scaling below computed
   widths that summed exactly to the available space, yet the table still
   visibly overflowed it, because every column rendered ~30px wider than its
   own style.width said. border-box makes style.width mean the TOTAL
   rendered width (padding included), matching what it visually occupies. */
#dataTable th, #dataTable td { box-sizing: border-box; }`}
${kpiPie ? `
/* Filter dropdowns for this page's label columns (Cust_name, Country of
   origin, ...) live here, right under "Total Index", instead of the small
   toolbar chips every other page uses - sized up to match the KPI cards'
   own visual weight, since changing one of these now drives the KPI totals
   too (see renderTotalIndex's rowsForCalc), not just the table below -
   they read as part of the Total Index control surface now, not a small
   ancillary table filter. */
.total-index-filters { display: flex; gap: 14px; flex-wrap: wrap; margin-bottom: 16px; }
.total-index-filters .filter-group {
  display: flex; flex-direction: column; align-items: flex-start; gap: 6px;
  background: var(--card-bg); border-radius: 6px; box-shadow: 0 1px 3px var(--shadow);
  padding: 10px 16px; font-size: 13px; color: var(--text-muted); font-weight: 600;
}
.total-index-filters .filter-group select {
  font-size: 15px; padding: 6px 10px; border-radius: 4px; max-width: 260px; min-width: 200px;
  border: 1px solid var(--border); background: var(--input-bg); color: var(--text); font-weight: 500;
}
.total-index-filters .ms-toggle {
  font-size: 15px; padding: 6px 10px; border-radius: 4px; max-width: 260px; min-width: 200px;
  border: 1px solid var(--border); background: var(--input-bg); color: var(--text); font-weight: 500;
}
.total-index-filters .ms-panel { min-width: 260px; }
.total-index-year-row { margin-bottom: 10px; }` : ''}${MCKINSEY_CSS}</style>
</head>
<body>
${actionKey ? `<p class="action-title" data-i18n="${actionKey}">${ACTION_EN[actionKey] || ''}</p>` : ''}
${introKey ? `<p class="page-intro" data-i18n="${introKey}">${INTRO_EN[introKey] || ''}</p>` : ''}
${kpiPie ? `
<h3 data-i18n="totalIndexHeading" style="margin:0 0 8px 0; color: var(--accent);">Total Index</h3>
${filterColumns.length ? '<div class="total-index-filters" id="totalIndexFilters"></div>' : ''}
<div class="kpis total-index-year-row" id="totalIndexKpis2026"></div>
<div class="kpis total-index-year-row" id="totalIndexKpis2025"></div>` : ''}
<div class="exhibit-kicker"><span data-i18n="exhibitWord">Exhibit</span> 1</div>
<div class="summary-bar" id="summaryBar" data-title-key="${titleKey || ''}" data-count-str="${rowCountStr}">${title} - ${rowCountStr} rows</div>
<div class="toolbar">
  <button id="exportBtn" data-i18n="exportCsv">Export CSV (filtered)</button>
  <button id="clearFiltersBtn" data-i18n="clearFilters" data-i18n-title="clearFiltersTooltip" title="Clear search box + all dropdown filters">Clear all filters / search</button>
  ${scrollX ? '' : '<button id="resetColWidthsBtn" data-i18n="resetColWidths" title="Reset all column widths back to default">Reset column widths</button>'}
  <span id="toolbarFilters" style="display:flex; gap:10px; flex-wrap:wrap;"></span>
</div>
<table id="dataTable" class="display">
  <thead><tr id="dataTableHeadRow"></tr></thead>
  <tbody></tbody>
</table>
${sourceNote ? `<div class="source-foot"><span data-i18n="footSource">Sources</span> — ${sourceNote}</div>` : ''}
<script>${jqueryJs}</script>
<script>${dataTablesJs}</script>
<script>
${I18N_JS}
var HEADER = ${headerJsonStr};
var ROWS = ${rowsJsonStr};
var FILTER_COL_IDX = ${filterColIdxJsonStr};
var DRILLDOWN = ${drilldownJsonStr};
var KPI_PIE = ${kpiPie ? 'true' : 'false'};
var YOY_RANK_COL = ${yoyRankCol ? 'true' : 'false'};
var GROUP_BY = ${groupByJsonStr};
// Which brand/group names are currently expanded (Family1 children shown
// instead of the collapsed group-total row) - empty means every group starts
// collapsed, so the table opens on the rolled-up Brand view.
var expandedGroups = {};
function isGroupTotalRow(rowData) { return GROUP_BY && rowData[GROUP_BY.detailKeyColIdx] === ''; }
// Total Index must sum every matching detail row regardless of which brands
// happen to be expanded right now - "rows({search:'applied'})" alone can't
// tell the difference, since it bundles the real filters (search box, column
// dropdowns) together with the group-visibility predicate registered below.
// This flag lets that predicate stand down for one query so 'applied' means
// only the real filters, then the group-total marker rows are stripped out
// separately - giving the true grand total across every brand's own detail
// rows, collapsed or not, rather than just whichever brand is on screen.
var bypassGroupVisibility = false;
function getRowsForTotalIndex(api) {
  if (!GROUP_BY) return api.rows({ search: 'applied' }).data().toArray();
  // rows({search:'applied'}) reads the filtered-row cache built by the last
  // real .draw() - just flipping the bypass flag here does nothing until a
  // draw actually re-runs the search predicates under it. Both draws stay
  // synchronous with no browser paint in between, so the visible table never
  // flashes the bypassed (ungrouped) state - only its own restored draw ever
  // reaches the screen.
  bypassGroupVisibility = true;
  api.draw(false);
  var rows = api.rows({ search: 'applied' }).data().toArray().filter(function (r) { return !isGroupTotalRow(r); });
  bypassGroupVisibility = false;
  api.draw(false);
  return rows;
}
function applyTheme(theme) { document.body.setAttribute('data-theme', theme); }
// Titles/headers/buttons only - column headers and every cell value are the
// report's actual data and are deliberately left untranslated (drill-link
// tooltips on newly-rendered cells still pick up the new language since they
// read t() at creation time via createdCell).
function applyLanguage(lang) {
  currentLang = lang;
  try { localStorage.setItem('reportLang', lang); } catch (e) {}
  document.querySelectorAll('[data-i18n]').forEach(function (el) { el.textContent = t(el.dataset.i18n); });
  document.querySelectorAll('[data-i18n-title]').forEach(function (el) { el.title = t(el.dataset.i18nTitle); });
  document.querySelectorAll('.weighted-pct[data-yoy-val]').forEach(function (el) { el.textContent = (el.dataset.yoyVal >= 0 ? '+' : '') + el.dataset.yoyVal + '% ' + t('yoyLabel'); });
  var sb = document.getElementById('summaryBar');
  if (sb && sb.dataset.titleKey) sb.textContent = t(sb.dataset.titleKey) + ' - ' + sb.dataset.countStr + ' ' + t('rowsSuffix');
  if (table) table.draw(false);
}
function regexEscape(s) {
  // Backslash escaped FIRST (see the excel-to-interactive-html-report skill's
  // Gotcha 8) - each entry here needs exactly 4 backslash CHARACTERS in this
  // source file to survive two rounds of string-literal reduction (once when
  // Node evaluates this outer template literal into the generated page's own
  // source, once more when the browser parses THAT as a JS string literal)
  // and come out as a single real backslash at runtime. 8 characters here
  // (one round too many) was found to silently double-escape every special
  // character instead - "LTD." became a pattern matching "LTD" + any
  // character, not a literal period, so an exact-match filter on any value
  // containing one of these characters matched zero rows.
  var specials = ['\\\\', '.', '*', '+', '?', '^', '$', '(', ')', '|', '[', ']', '{', '}'];
  var str = String(s);
  specials.forEach(function (c) { str = str.split(c).join('\\\\' + c); });
  return str;
}
// Sets each named column's filter (toolbar dropdown, if one exists for it,
// plus the underlying DataTables search either way) - used both by a normal
// dropdown change and by an incoming cross-tab drilldown request.
var table;
// One entry per filterable column, keyed by its HEADER index - populated as
// each multi-select widget is built (see buildMultiSelectFilter below) so
// both a cross-tab drilldown (applyExternalFilter) and "Clear all filters"
// can sync the widget's own checkboxes/toggle label, not just the
// underlying DataTables search.
var msWidgetsByCol = {};
// Multi-select + type-to-search filter widget: a small toggle button opens a
// panel with a search box (narrows the checkbox list shown, not the table
// itself) and one checkbox per unique column value. Replaces a plain
// single-value <select> so pages like YoY Product Family can filter to
// several values at once (e.g. two Family1 values together) instead of only
// ever one - per Peggy's request, illustrated against the old single-select
// screenshot. onChange receives the full array of currently-checked values
// (empty = no filter, matching the old "All" option).
function buildMultiSelectFilter(colIdx, options, onChange) {
  var selected = [];
  var checkboxes = {};
  var wrap = document.createElement('span');
  wrap.className = 'ms-widget';
  wrap.dataset.colIndex = colIdx;
  var toggle = document.createElement('button');
  toggle.type = 'button';
  toggle.className = 'ms-toggle';
  var panel = document.createElement('div');
  panel.className = 'ms-panel';
  panel.hidden = true;
  var search = document.createElement('input');
  search.type = 'text';
  search.className = 'ms-search';
  search.placeholder = t('typeToSearchLabel');
  var actions = document.createElement('div');
  actions.className = 'ms-actions';
  var selectAllBtn = document.createElement('button'); selectAllBtn.type = 'button'; selectAllBtn.textContent = t('selectAllLabel');
  var clearBtn = document.createElement('button'); clearBtn.type = 'button'; clearBtn.textContent = t('clearSelLabel');
  actions.appendChild(selectAllBtn); actions.appendChild(clearBtn);
  var optionsBox = document.createElement('div');
  optionsBox.className = 'ms-options';
  function updateToggleLabel() {
    if (!selected.length) toggle.textContent = t('allLabel') + ' ▾';
    else if (selected.length === 1) toggle.textContent = selected[0] + ' ▾';
    else toggle.textContent = selected.length + ' ' + t('selectedLabel') + ' ▾';
  }
  options.forEach(function (v) {
    var lbl = document.createElement('label');
    var cb = document.createElement('input'); cb.type = 'checkbox'; cb.value = v;
    cb.addEventListener('change', function () {
      if (cb.checked) { if (selected.indexOf(v) < 0) selected.push(v); }
      else { selected = selected.filter(function (x) { return x !== v; }); }
      updateToggleLabel();
      onChange(selected.slice());
    });
    checkboxes[v] = cb;
    lbl.appendChild(cb);
    lbl.appendChild(document.createTextNode(' ' + v));
    optionsBox.appendChild(lbl);
  });
  updateToggleLabel();
  search.addEventListener('input', function () {
    var q = search.value.toLowerCase();
    options.forEach(function (v) {
      checkboxes[v].parentElement.style.display = v.toLowerCase().indexOf(q) >= 0 ? '' : 'none';
    });
  });
  selectAllBtn.addEventListener('click', function () {
    // Only the currently search-narrowed options are affected - selecting
    // "all" while a search is typed is expected to mean all VISIBLE matches,
    // not silently also selecting every hidden option the user hasn't seen.
    var visible = options.filter(function (v) { return checkboxes[v].parentElement.style.display !== 'none'; });
    selected = visible.slice();
    options.forEach(function (v) { checkboxes[v].checked = visible.indexOf(v) >= 0; });
    updateToggleLabel();
    onChange(selected.slice());
  });
  clearBtn.addEventListener('click', function () {
    selected = [];
    options.forEach(function (v) { checkboxes[v].checked = false; });
    updateToggleLabel();
    onChange(selected.slice());
  });
  toggle.addEventListener('click', function (e) {
    e.stopPropagation();
    var willOpen = panel.hidden;
    document.querySelectorAll('.ms-panel').forEach(function (p) { p.hidden = true; });
    if (willOpen) {
      panel.hidden = false;
      search.value = '';
      options.forEach(function (v) { checkboxes[v].parentElement.style.display = ''; });
      search.focus();
    }
  });
  document.addEventListener('click', function (e) { if (!wrap.contains(e.target)) panel.hidden = true; });
  panel.appendChild(search); panel.appendChild(actions); panel.appendChild(optionsBox);
  wrap.appendChild(toggle); wrap.appendChild(panel);
  return {
    el: wrap,
    setSelected: function (values) {
      selected = values.slice();
      options.forEach(function (v) { checkboxes[v].checked = selected.indexOf(v) >= 0; });
      updateToggleLabel();
    }
  };
}
function applyExternalFilter(filterMap) {
  // The parent shell's iframe "load" event can fire before this page's own
  // jQuery-ready callback has assigned the table variable (a real race on a
  // table this size) - if so, queue the request for initComplete (below) to
  // pick up instead of throwing on table.column() against an undefined table.
  if (!table) { window.__pendingDrilldownFilters = filterMap; return; }
  Object.keys(filterMap).forEach(function (colName) {
    var colIdx = HEADER.indexOf(colName);
    if (colIdx < 0) return;
    var val = filterMap[colName];
    table.column(colIdx).search(val === '' ? '' : '^' + regexEscape(val) + '\$', true, false);
    var widget = msWidgetsByCol[colIdx];
    if (widget) widget.setSelected(val === '' ? [] : [val]);
  });
  table.draw();
  if (KPI_PIE) renderTotalIndex(getRowsForTotalIndex(table));
}
// Right-align number-like columns for readability. The Data tab stores
// measures as real JS numbers, but the YoY/Customer-Quarter sheets keep
// Excel's own formatted display text ("6,528", "22.4%", "(1,234)") - typeof
// alone would miss those, so numeric-ness is judged from a sample of each
// column's actual values instead of the column's declared type.
// Leading $ included: these sheets keep Excel's own formatted text ("$6,349,104"),
// and without it every money column failed this test, so it never got the
// dt-num right-alignment or an explicit num-fmt sort type (DataTables' own
// type detection happened to sort them correctly anyway - the alignment and
// the explicit type are what this actually fixes).
var NUMERIC_RE = /^\\$?\\(?-?[\\d,]+(\\.\\d+)?%?\\)?$/;
function isNumericColumn(colIdx) {
  var seen = 0, numericLike = 0;
  for (var i = 0; i < ROWS.length && seen < 200; i++) {
    var v = ROWS[i][colIdx];
    if (v === null || v === undefined || v === '') continue;
    seen++;
    if (typeof v === 'number' || (typeof v === 'string' && NUMERIC_RE.test(v.trim()))) numericLike++;
  }
  return seen > 0 && numericLike / seen > 0.9;
}
var NUMERIC_COL_IDX = HEADER.map(function (h, i) { return i; }).filter(isNumericColumn);
// Per-column resize widths, persisted per sheet (this builder is shared by
// every table, so each sheet needs its own storage slot rather than one
// shared key clobbering every other table's saved widths).
var COL_WIDTHS_KEY = 'mtcColWidths_${(titleKey || title).replace(/[^A-Za-z0-9_]/g, '_')}';
function loadColWidths(n) {
  try {
    var saved = JSON.parse(localStorage.getItem(COL_WIDTHS_KEY) || 'null');
    if (Array.isArray(saved) && saved.length === n) return saved;
  } catch (e) {}
  return new Array(n).fill(null);
}
function saveColWidths(widths) {
  try { localStorage.setItem(COL_WIDTHS_KEY, JSON.stringify(widths)); } catch (e) {}
}
// Recovers a real number from Excel's own formatted display text ("$1,234",
// "22.4%", "(1,234)" for negative) - the YoY sheets keep that text, not raw
// numbers (see the NUMERIC_RE comment above), so summing a metric column for
// the Total Index/pie below has to parse it back first.
function parseFormattedNumber(v) {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined || v === '') return 0;
  var s = String(v).trim();
  var neg = /^\\(.*\\)$/.test(s);
  s = s.replace(/[()$,%]/g, '');
  var n = parseFloat(s);
  if (isNaN(n)) return 0;
  return neg ? -n : n;
}
// "Total Index" KPI cards + a pie chart, above the table - mirrors the
// Global Purchase Report pivot page's own Total/Weighted-% + pie pattern,
// adapted to this table's own year-paired metric columns instead of a
// hardcoded set: every "<Metric> 2025"/"<Metric> 2026" column pair becomes
// one KPI group (each card's "weighted %" is that year's share of the
// 2025+2026 combined total for that metric - NOT a share of some grand total
// across unrelated metrics, which wouldn't mean anything); % ratio columns
// (Net Margin %) are excluded from summing since summing a percentage across
// rows isn't a meaningful total. Cards render as two separate rows - all
// 2026 figures together, all 2025 figures together below - rather than
// interleaved per metric: the 2026 row shows each metric's own YoY % change
// vs 2025 (not a weighted share of a combined total, which read as an odd
// question once there was no pie chart's slices for "share" to describe);
// the 2025 row is the baseline and shows no percentage at all.
var totalIndexPairs = null; // computed once (structure never changes) - only the sums do
// rowsForCalc: which rows to sum - omitted (initial render) uses the full
// ROWS; a filter dropdown or "Clear all filters" passes the table's own
// currently-APPLIED rows instead, so the Total Index stays in sync with
// whatever the user has filtered the table to, not just the table body.
function renderTotalIndex(rowsForCalc) {
  if (!KPI_PIE) return;
  rowsForCalc = rowsForCalc || ROWS;
  // A group-total row (e.g. a Brand's rolled-up totals) duplicates the sum of
  // its own Family1 rows - summing both whenever a group happens to be
  // collapsed (so both are simultaneously "in ROWS") would double-count
  // every card here. The Total Index should read the same true grand total
  // regardless of which groups the user currently has expanded, so it always
  // sums only the real leaf/detail rows, never a synthetic group-total row.
  if (GROUP_BY) rowsForCalc = rowsForCalc.filter(function (r) { return !isGroupTotalRow(r); });
  if (!totalIndexPairs) {
    var labelColCount = DRILLDOWN ? DRILLDOWN.labelColCount : 0;
    var pairs = [], byBase = {};
    HEADER.forEach(function (h, i) {
      if (i < labelColCount) return;
      var m = /^(.*)\\s(20\\d{2})$/.exec(h);
      if (!m) return;
      var base = m[1], year = m[2];
      if (base.indexOf('%') >= 0) return;
      if (!byBase[base]) { byBase[base] = { base: base, cols: {} }; pairs.push(byBase[base]); }
      byBase[base].cols[year] = i;
    });
    totalIndexPairs = pairs.filter(function (p) { return p.cols['2025'] !== undefined && p.cols['2026'] !== undefined; });
  }
  var pairs = totalIndexPairs;
  if (!pairs.length) return;

  var row2026 = document.getElementById('totalIndexKpis2026');
  var row2025 = document.getElementById('totalIndexKpis2025');
  row2026.innerHTML = ''; row2025.innerHTML = ''; // clear the previous filter state's cards before rebuilding
  pairs.forEach(function (p) {
    var sum2025 = 0, sum2026 = 0;
    rowsForCalc.forEach(function (r) {
      sum2025 += parseFormattedNumber(r[p.cols['2025']]);
      sum2026 += parseFormattedNumber(r[p.cols['2026']]);
    });

    var kpi2026 = document.createElement('div'); kpi2026.className = 'kpi';
    var label2026 = document.createElement('div'); label2026.className = 'label'; label2026.textContent = p.base + ' 2026';
    var value2026 = document.createElement('div'); value2026.className = 'value'; value2026.textContent = Math.round(sum2026).toLocaleString();
    var yoyEl = document.createElement('div'); yoyEl.className = 'weighted-pct';
    // Store the raw YoY % so a language switch can rebuild "+x.x% YoY" in the
    // new language without re-summing everything - see applyLanguage.
    if (sum2025) {
      yoyEl.dataset.yoyVal = ((sum2026 - sum2025) / sum2025 * 100).toFixed(1);
      yoyEl.textContent = (yoyEl.dataset.yoyVal >= 0 ? '+' : '') + yoyEl.dataset.yoyVal + '% ' + t('yoyLabel');
    }
    kpi2026.appendChild(label2026); kpi2026.appendChild(value2026); kpi2026.appendChild(yoyEl);
    row2026.appendChild(kpi2026);

    var kpi2025 = document.createElement('div'); kpi2025.className = 'kpi';
    var label2025 = document.createElement('div'); label2025.className = 'label'; label2025.textContent = p.base + ' 2025';
    var value2025 = document.createElement('div'); value2025.className = 'value'; value2025.textContent = Math.round(sum2025).toLocaleString();
    kpi2025.appendChild(label2025); kpi2025.appendChild(value2025);
    row2025.appendChild(kpi2025);
  });
}
$(function () {
  renderTotalIndex();
  // Appends a "YoY % Change" column (2026 vs 2025, on whichever metric pair
  // looks like a sales figure - same pick as the Total Index pie) computed
  // from totalIndexPairs, which renderTotalIndex() (called just above)
  // already found and cached from the ORIGINAL header - appending a column
  // here afterward doesn't disturb any of those earlier column indices.
  var yoyRankRefIdx = null; // which column the default sort targets (the rank metric's 2026 value, not the new % column itself)
  if (YOY_RANK_COL && totalIndexPairs && totalIndexPairs.length) {
    var rankPair = totalIndexPairs.filter(function (p) { return /sales/i.test(p.base); })[0] || totalIndexPairs[0];
    var rc25 = rankPair.cols['2025'], rc26 = rankPair.cols['2026'];
    HEADER.push(t('yoyChangeCol'));
    var yoyColIdx = HEADER.length - 1;
    ROWS.forEach(function (r) {
      var cur = parseFormattedNumber(r[rc26]), prior = parseFormattedNumber(r[rc25]);
      var pct = prior ? (cur - prior) / prior * 100 : null;
      r.push(pct === null ? '' : (pct >= 0 ? '+' : '') + pct.toFixed(1) + '%');
    });
    NUMERIC_COL_IDX.push(yoyColIdx);
    yoyRankRefIdx = rc26;
  }
  var headRow = document.getElementById('dataTableHeadRow');
  HEADER.forEach(function (h, i) {
    var th = document.createElement('th');
    th.textContent = h;
    if (NUMERIC_COL_IDX.indexOf(i) >= 0) th.className = 'dt-num';
    headRow.appendChild(th);
  });
  var toolbarFilters = document.getElementById('toolbarFilters');
  // "num-fmt" DataTables sort type: these Excel-YoY sheets keep formatted
  // display text ("$13,366", "22.4%"), not raw numbers, so the DEFAULT
  // string sort would order them alphabetically, not by magnitude - "high to
  // low" only actually works once sorting is taught to parse the number back
  // out first. Applied to every numeric-looking column (not just the new
  // rank column), fixing click-to-sort too, not just the default order.
  $.fn.dataTable.ext.type.order['num-fmt-pre'] = function (data) {
    if (data === null || data === undefined || data === '') return -Infinity;
    if (typeof data === 'number') return data;
    var neg = /^\\(.*\\)$/.test(String(data).trim());
    var n = parseFloat(String(data).replace(/[$,%()]/g, ''));
    if (isNaN(n)) return -Infinity;
    return neg ? -n : n;
  };
  var columnDefs = [{ targets: NUMERIC_COL_IDX, className: 'dt-num', type: 'num-fmt' }];
  if (DRILLDOWN) {
    // Most YoY sheets pair each metric with its own "... 2025"/"... 2026"
    // column, so the year is parsed per-column from the header text. A sheet
    // like Customer-Quarter has no such pairing - every row is already
    // scoped to one Year via a page filter above the header - so DRILLDOWN.
    // fixedYear (captured from that cell at extraction time) applies to
    // every metric column uniformly instead.
    var yearByCol = {};
    HEADER.forEach(function (h, i) {
      if (i < DRILLDOWN.labelColCount) return;
      if (DRILLDOWN.fixedYear) { yearByCol[i] = DRILLDOWN.fixedYear; return; }
      var m = /\\s(20\\d{2})$/.exec(h);
      if (m) yearByCol[i] = m[1];
    });
    columnDefs.push({
      targets: Object.keys(yearByCol).map(Number),
      createdCell: function (td, cellData, rowData, rowIdx, colIdx) {
        if (cellData === null || cellData === undefined || cellData === '') return;
        var a = document.createElement('span');
        a.className = 'drill-link';
        a.textContent = cellData;
        a.title = t('drillTooltip');
        a.addEventListener('click', function () {
          var filters = { Year: yearByCol[colIdx] };
          Object.keys(DRILLDOWN.dataColMap).forEach(function (srcCol) {
            var srcIdx = HEADER.indexOf(srcCol);
            var v = srcIdx >= 0 ? rowData[srcIdx] : '';
            if (v !== null && v !== undefined && v !== '') filters[DRILLDOWN.dataColMap[srcCol]] = v;
          });
          window.parent.postMessage({ type: 'mtcDrilldown', targetTab: 'Data', filters: filters }, '*');
        });
        td.textContent = '';
        td.appendChild(a);
      }
    });
    // Clicking a label value itself (e.g. a customer name, a country of
    // origin, a fabrication code, a product family, or a part id) drills into
    // the Data tab filtered to just that one field - unlike a metric-column
    // click, there's no single year to scope to here (the row's metrics span
    // both years), so only that field filters. Every label column that has a
    // matching Data-tab column (dataColMap) gets this by default, across
    // every YoY sheet, unless a call site explicitly narrows it down.
    var labelDrillCols = DRILLDOWN.labelDrillCols || Object.keys(DRILLDOWN.dataColMap);
    if (labelDrillCols.length) {
      var labelDrillIdx = {};
      labelDrillCols.forEach(function (colName) {
        var idx = HEADER.indexOf(colName);
        if (idx >= 0) labelDrillIdx[idx] = colName;
      });
      columnDefs.push({
        targets: Object.keys(labelDrillIdx).map(Number),
        createdCell: function (td, cellData, rowData, rowIdx, colIdx) {
          if (cellData === null || cellData === undefined || cellData === '') return;
          var srcCol = labelDrillIdx[colIdx];
          var dataCol = DRILLDOWN.dataColMap[srcCol] || srcCol;
          var a = document.createElement('span');
          a.className = 'drill-link';
          a.textContent = cellData;
          a.title = t('drillTooltip');
          a.addEventListener('click', function () {
            var filters = {};
            filters[dataCol] = cellData;
            window.parent.postMessage({ type: 'mtcDrilldown', targetTab: 'Data', filters: filters }, '*');
          });
          td.textContent = '';
          td.appendChild(a);
        }
      });
    }
  }
  if (GROUP_BY) {
    // Renders '+' / '−' only on a group-total row (its toggle cell is the
    // only one carrying the '+' marker to begin with - a detail row's own
    // toggle cell is always blank, so this naturally never fires for one).
    columnDefs.push({
      targets: [GROUP_BY.toggleColIdx],
      orderable: false,
      createdCell: function (td, cellData, rowData) {
        if (cellData !== '+') return;
        var group = rowData[GROUP_BY.groupColIdx];
        var btn = document.createElement('span'); btn.className = 'row-toggle';
        btn.textContent = expandedGroups[group] ? '−' : '+';
        btn.title = expandedGroups[group] ? t('collapseAllLabel') : t('expandAllLabel');
        btn.addEventListener('click', function () {
          expandedGroups[group] = !expandedGroups[group];
          table.draw(false);
          renderTotalIndex(getRowsForTotalIndex(table));
        });
        td.textContent = '';
        td.appendChild(btn);
      }
    });
    // Every group contributes either its one collapsed total row or all of
    // its expanded detail rows - never both at once (see renderTotalIndex's
    // own comment on why that mutual exclusion matters) - so this alone
    // decides visibility; it composes with the search box/column filters
    // below it, since DataTables ANDs every registered search plugin.
    $.fn.dataTable.ext.search.push(function (settings, data) {
      if (settings.nTable !== document.getElementById('dataTable')) return true;
      if (bypassGroupVisibility) return true;
      var isTotal = data[GROUP_BY.detailKeyColIdx] === '';
      var expanded = !!expandedGroups[data[GROUP_BY.groupColIdx]];
      return isTotal ? !expanded : expanded;
    });
  }
  table = $('#dataTable').DataTable({
    data: ROWS, deferRender: true, pageLength: 50, lengthMenu: [25, 50, 100, 500, 1000],
    scrollX: ${scrollX ? 'true' : 'false'}, autoWidth: false,
    columnDefs: columnDefs,
    // Forces every row of the same group to sort adjacent to each other
    // regardless of which column the user clicks to sort by (that becomes a
    // secondary key within the group) - without this, an expanded group's
    // detail rows would scatter across the table wherever their own values
    // happened to rank, with nothing left visually tying them to their brand.
    orderFixed: GROUP_BY ? [[GROUP_BY.groupColIdx, 'asc']] : undefined,
    // Opens already ranked highest-to-lowest by the rank metric's 2026 value
    // - the whole point of this table is being scannable top-to-bottom
    // without needing to click a column header to sort first.
    order: yoyRankRefIdx !== null ? [[yoyRankRefIdx, 'desc']] : ${defaultSortColIdx >= 0 ? `[[${defaultSortColIdx}, 'desc']]` : '[]'},
    initComplete: function () {
      var api = this.api();
      // DataTables calls initComplete DURING construction, before the outer
      // "table = $(...).DataTable(...)" assignment above has actually
      // completed - assign it here too so applyExternalFilter (which checks
      // the table variable) works correctly even when called from in here.
      table = api;
      // On a KPI_PIE (YoY) page, these filter dropdowns live in a dedicated
      // box right under "Total Index" instead of the table's own toolbar
      // (see #totalIndexFilters in the HTML), and changing one also
      // recomputes the KPI cards + pie from the table's own currently-
      // applied rows - previously they only filtered the table body below,
      // leaving the totals/pie silently stuck on the full, unfiltered data.
      var filtersTarget = document.getElementById('totalIndexFilters') || toolbarFilters;
      FILTER_COL_IDX.forEach(function (colIdx) {
        var group = document.createElement('span'); group.className = 'filter-group';
        var label = document.createElement('label'); label.textContent = HEADER[colIdx] + ':';
        var values = api.column(colIdx).data().toArray().map(function (v) { return (v === null || v === undefined) ? '' : String(v); });
        var unique = values.filter(function (v, i) { return v !== '' && values.indexOf(v) === i; }).sort();
        var widget = buildMultiSelectFilter(colIdx, unique, function (selectedValues) {
          var column = api.column(colIdx);
          if (!selectedValues.length) column.search('');
          else column.search('^(' + selectedValues.map(regexEscape).join('|') + ')\$', true, false);
          // Picking a specific Family1 value while its Brand is still
          // collapsed would otherwise show zero rows - the detail row stays
          // hidden by the group toggle, and the Brand's own total row never
          // matches an exact Family1 filter - so auto-expand every Brand
          // that owns a just-selected value, instead of leaving the user at
          // a confusing dead end they'd have to know to fix by hand.
          if (GROUP_BY && colIdx === GROUP_BY.detailKeyColIdx && selectedValues.length) {
            ROWS.forEach(function (r) {
              if (selectedValues.indexOf(r[GROUP_BY.detailKeyColIdx]) >= 0) expandedGroups[r[GROUP_BY.groupColIdx]] = true;
            });
          }
          api.draw();
          if (KPI_PIE) renderTotalIndex(getRowsForTotalIndex(api));
        });
        msWidgetsByCol[colIdx] = widget;
        group.appendChild(label); group.appendChild(widget.el); filtersTarget.appendChild(group);
      });
      // A pending drilldown may have arrived before this table finished
      // initializing (the parent shell queues it) - apply it now.
      if (window.__pendingDrilldownFilters) {
        applyExternalFilter(window.__pendingDrilldownFilters);
        window.__pendingDrilldownFilters = null;
      }
      // Draggable per-column resize handles, ported from the Global Purchase
      // Report pivot page. Deliberately scoped to non-scrollX tables only -
      // tried it on the scrollX Data tab first and the live drag silently had
      // no visual effect (DataTables' scrollX clones the header into a
      // separate scrolling wrapper and re-derives its own column widths on
      // any columns.adjust()/redraw, discarding a plain style.width change on
      // the original header cell) - see Gotcha 7 in the skill for the same
      // class of scrollX/width interaction. Shipping a resize handle that
      // visibly does nothing on drag is worse than not having one.
      if (${scrollX ? 'false' : 'true'}) (function setupColumnResize() {
        var headerThs = Array.prototype.slice.call(document.querySelectorAll('#dataTableHeadRow th'));
        if (!headerThs.length) return;
        // Capture each column's natural (table-layout:auto) width BEFORE
        // switching to fixed - see the CSS comment above for why every
        // column needs an explicit width the instant layout goes fixed.
        var naturalWidths = headerThs.map(function (th) { return th.getBoundingClientRect().width; });
        // Scale every column proportionally so the whole table exactly fits
        // the current on-screen width - like Excel's "Fit to page" print
        // option - instead of a many-column YoY sheet spilling past a
        // laptop's viewport into a whole-page horizontal scrollbar (this
        // table has no scrollX container of its own to scroll within
        // instead; see the scrollX-scoping note above). This is also what
        // "Reset column widths" restores a column to, not the raw natural
        // (pre-scale) width.
        var container = document.getElementById('dataTable_wrapper') || document.body;
        var availableWidth = container.clientWidth || window.innerWidth;
        var naturalTotal = naturalWidths.reduce(function (a, b) { return a + b; }, 0);
        var fitScale = naturalTotal > 0 && availableWidth > 0 ? availableWidth / naturalTotal : 1;
        var fitWidths = naturalWidths.map(function (w) { return Math.max(24, Math.floor(w * fitScale)); });
        var colWidths = loadColWidths(headerThs.length);
        document.getElementById('dataTable').style.tableLayout = 'fixed';
        headerThs.forEach(function (th, i) {
          th.style.width = (colWidths[i] || fitWidths[i]) + 'px';
          var handle = document.createElement('span');
          handle.className = 'col-resize-handle';
          th.appendChild(handle);
        });
        var dragColIndex = null, dragStartX = 0, dragStartWidth = 0;
        function onMouseMove(e) {
          if (dragColIndex === null) return;
          var newWidth = Math.max(24, dragStartWidth + (e.clientX - dragStartX));
          headerThs[dragColIndex].style.width = newWidth + 'px';
        }
        function onMouseUp() {
          if (dragColIndex === null) return;
          colWidths[dragColIndex] = parseInt(headerThs[dragColIndex].style.width, 10);
          saveColWidths(colWidths);
          var handle = headerThs[dragColIndex].querySelector('.col-resize-handle');
          if (handle) handle.classList.remove('resizing');
          dragColIndex = null;
          document.removeEventListener('mousemove', onMouseMove);
          document.removeEventListener('mouseup', onMouseUp);
          // Nudges DataTables to re-sync its scrollX header/body width
          // bookkeeping after a manual width change - the reference
          // implementation never needed this because it ran with scrollX
          // off; the Data tab here needs scrollX on (103 columns).
          api.columns.adjust();
        }
        // A drag can end with the mouse over a DIFFERENT header cell than it
        // started on (dragging widens/moves the handle out from under the
        // cursor), so the resulting mouseup/click can land on a neighboring
        // <th> instead of the one being resized - stopping propagation only
        // on the handle's own 'mousedown'/'click' (tried first) still let a
        // drag silently flip that OTHER column's sort direction. A one-shot
        // document-level CAPTURING listener swallows the very next click
        // anywhere on the page, regardless of which element it lands on,
        // before DataTables' own bubble-phase sort listener ever sees it.
        function suppressNextClick(e) { e.stopPropagation(); e.preventDefault(); document.removeEventListener('click', suppressNextClick, true); }
        headerThs.forEach(function (th, i) {
          var handle = th.querySelector('.col-resize-handle');
          handle.addEventListener('mousedown', function (e) {
            e.preventDefault();
            e.stopPropagation();
            dragColIndex = i;
            dragStartX = e.clientX;
            dragStartWidth = th.getBoundingClientRect().width;
            handle.classList.add('resizing');
            document.addEventListener('mousemove', onMouseMove);
            document.addEventListener('mouseup', onMouseUp);
            document.addEventListener('click', suppressNextClick, true);
          });
        });
        document.getElementById('resetColWidthsBtn').onclick = function () {
          colWidths = new Array(headerThs.length).fill(null);
          saveColWidths(colWidths);
          // Reset to the fit-to-screen widths, not '' - table-layout is
          // already fixed at this point, so clearing width entirely would
          // reproduce the exact squeezed-together layout this feature is
          // built to avoid, not restore the original fit-to-one-screen look.
          headerThs.forEach(function (th, i) { th.style.width = fitWidths[i] + 'px'; });
          api.columns.adjust();
        };
      })();
    }
  });
  document.getElementById('clearFiltersBtn').onclick = function () {
    table.search('');
    FILTER_COL_IDX.forEach(function (colIdx) {
      table.column(colIdx).search('');
      var widget = msWidgetsByCol[colIdx];
      if (widget) widget.setSelected([]);
    });
    if (GROUP_BY) expandedGroups = {};
    table.draw();
    if (KPI_PIE) renderTotalIndex(getRowsForTotalIndex(table));
  };
  document.getElementById('exportBtn').onclick = function () {
    var rowsData = table.rows({ search: 'applied' }).data();
    var lines = [HEADER.join(',')];
    for (var i = 0; i < rowsData.length; i++) {
      var r = rowsData[i];
      lines.push(r.map(function (v) {
        if (v === null || v === undefined) return '';
        var s = String(v);
        if (s.indexOf(',') >= 0 || s.indexOf('"') >= 0) s = '"' + s.replace(/"/g, '""') + '"';
        return s;
      }).join(','));
    }
    var csv = '\\ufeff' + lines.join('\\r\\n');
    var blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a'); a.href = url; a.download = '${title.replace(/[^a-z0-9]/gi, '_')}_export.csv';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };
});
</script>
</body>
</html>`;
}

// =========================================================================
// Dashboard: KPI cards + Chart.js charts + live multi-select filters,
// computed client-side from the flat DATA rows (not Excel's static pivot
// cache) - mirrors how the "Performance dashboard" tab's own PivotTables
// are defined: Sales=Net_Sales_USD, Cost=Ttl PO USD, Margin=TTL PO NET CONT
// USD, Margin%=TTL PO NET CONT %.
// =========================================================================
const yearIdx = COL['Year'];
const monthIdx = COL['yyyy-mm'];
const custIdx = COL['Cust_name'];
const vendIdx = COL['VEND_NAME'];
const countryIdx = COL['Cust_Countries'];
const originIdx = COL['Country of origin'];
const salesIdx = COL['Net_Sales_USD'];
const costIdx = resolveColumnIndex(['Ttl PO USD', 'Matl_USD']);
const marginIdx = resolveColumnIndex(['TTL PO NET CONT USD', 'Net_Cont_USD']);

const dashRows = dataSheet.rows.map(r => [
  typeof r[yearIdx] === 'number' ? String(r[yearIdx]) : r[yearIdx],
  r[monthIdx], r[custIdx], r[vendIdx], r[countryIdx],
  typeof r[salesIdx] === 'number' ? r[salesIdx] : 0,
  typeof r[costIdx] === 'number' ? r[costIdx] : 0,
  typeof r[marginIdx] === 'number' ? r[marginIdx] : 0,
  r[originIdx]
]);
// dashRows columns: 0=Year,1=Month,2=Customer,3=Vendor,4=Country(customer),5=Sales,
// 6=Cost,7=Margin,8=Country of origin (sourcing)

const yearValues = Array.from(new Set(dashRows.map(r => r[0]))).filter(Boolean).sort();
const monthValues = Array.from(new Set(dashRows.map(r => r[1]))).filter(Boolean).sort();
const custValues = Array.from(new Set(dashRows.map(r => r[2]))).filter(Boolean).sort();
const vendValues = Array.from(new Set(dashRows.map(r => r[3]))).filter(Boolean).sort();
const countryValues = Array.from(new Set(dashRows.map(r => r[4]))).filter(Boolean).sort();
// Display name of the REPORT (user decision 2026-11-09: the deliverable now
// covers sales, margin, purchase and domestic - "Performance", not "Sales".
// File and repo names intentionally keep the historic MTC Sales Report slug.)
const reportTitle = monthValues.length
  ? 'MTC Performance Report - ' + monthValues[0] + (monthValues.length > 1 ? ' to ' + monthValues[monthValues.length - 1] : '')
  : 'MTC Performance Report';
// Consulting convention: every tab ends with a source line (island filename +
// report period + FX basis). These are the shared pieces; the per-tab island
// names are passed at each call site. Language-neutral by design (filenames
// and yyyy-mm dates read the same in EN and 中文 modes).
const reportPeriodRange = monthValues.length
  ? 'Report period ' + monthValues[0] + (monthValues.length > 1 ? ' to ' + monthValues[monthValues.length - 1] : '')
  : 'Report period FY2025-FY2026';
const FX_SOURCE = 'FX: Platinum Buy Rate per the Exchange Rate tab (exrate_usd.json)';
// The Domestic islands are LOADED here (moved up verbatim from their old spot
// in the Domestic section, which still references these same consts) because
// the Dashboard's footnote cites fx_domestic.json's as-of string, and the
// Dashboard template is evaluated at its own line - JS const declarations
// must precede the first use. readJsonOptional is a hoisted function
// declaration, so calling it at this position is safe.
const domesticJson = readJsonOptional('domestic.json');
const fxDomesticJson = readJsonOptional('fx_domestic.json');
const fxAsOfNote = fxDomesticJson
  ? (fxDomesticJson.file + ', as-of ' + fxDomesticJson.asOf)
  : 'fx_domestic.json (not found at build time)';

// ---- Dashboard DOMESTIC index strip ---------------------------------------
// Summary-level memo straight from the two islands. NOT filter-reactive by
// design: the source workbook is a static snapshot (its as-of date is shown
// on the card), and the figures are a DERIVED USD memo - never added to the
// MA/DI totals above them. Same derivation rule as the Domestic tab:
// USD = local x (HkdRate / Unit) / USD-row HkdRate, with the RMB->CNY alias
// or both China rows silently drop out. DOMESTIC_FX_ALIAS is declared here
// (module scope, ahead of the Dashboard template) and shared with the
// Domestic tab builder further down.
const DOMESTIC_FX_ALIAS = { RMB: 'CNY' };
let domMemo = null;
if (domesticJson && fxDomesticJson && domesticJson.tableA && domesticJson.tableB && Array.isArray(fxDomesticJson.rows)) {
  const dFx = {};
  fxDomesticJson.rows.forEach(function (row) { dFx[String(row[1]).trim().toUpperCase()] = row; });
  const dUsdHkd = dFx.USD ? Number(dFx.USD[3]) : 0;
  function dUsd(code, value) {
    if (value === null || value === undefined || value === '' || !dUsdHkd) return null;
    const k = String(code).trim().toUpperCase();
    const row = dFx[DOMESTIC_FX_ALIAS[k] || k];
    if (!row) return null;
    const unit = Number(row[2]) || 1;
    return (Number(row[3]) / unit) / dUsdHkd * Number(value);
  }
  function dTotals(table) {
    let sum = 0, ok = true;
    const curIdx = (table.currencyCol || 4) - 1;
    const mets = table.metricCols || [5, 6, 7, 8];
    const totIdx = mets[2] - 1; // the "2026 Total" column
    table.rows.forEach(function (r) {
      const v = r[totIdx];
      if (v === null || v === undefined || v === '') return; // MCN has no 2026 row - skipped like the tab does
      const u = dUsd(r[curIdx], v);
      if (u === null) { ok = false; return; }
      sum += u;
    });
    return { sum: sum, ok: ok };
  }
  const dA = dTotals(domesticJson.tableA);
  const dB = dTotals(domesticJson.tableB);
  domMemo = { sales: dA.sum, comm: dB.sum, complete: dA.ok && dB.ok, asOf: String(domesticJson.asOf || '').replace(/&/g, '&amp;').replace(/</g, '&lt;') };
  console.log('Domestic memo strip: 2026 derived USD sales ' + Math.round(dA.sum).toLocaleString('en-US') + ' + commission ' + Math.round(dB.sum).toLocaleString('en-US') + (domMemo.complete ? '' : ' (INCOMPLETE - some currencies unresolved)'));
} else {
  console.warn('WARNING: domestic islands absent/incomplete - Dashboard DOMESTIC index strip skipped');
}
const domStripHtml = domMemo ? (
  '<div class="kpis kpis-dom">' +
    '<div class="kpi kpi-dom-card" role="button" tabindex="0">' +
      '<div class="label"><span data-i18n="kpiDomSales">Domestic Net Sales (USD, derived)</span><sup class="dom-memo-sup">†</sup></div>' +
      '<div class="value">' + Math.round(domMemo.sales).toLocaleString('en-US') + '</div>' +
      '<div class="kpi-sub"><span data-i18n="domBasis">2026 total</span> · <span data-i18n="domAsOf">snapshot as of</span> ' + domMemo.asOf + '</div>' +
    '</div>' +
    '<div class="kpi kpi-dom-card" role="button" tabindex="0">' +
      '<div class="label"><span data-i18n="kpiDomCommission">Domestic MTC Commission (USD, derived)</span><sup class="dom-memo-sup">†</sup></div>' +
      '<div class="value">' + Math.round(domMemo.comm).toLocaleString('en-US') + '</div>' +
      '<div class="kpi-sub"><span data-i18n="domBasis">2026 total</span> · <span data-i18n="domAsOf">snapshot as of</span> ' + domMemo.asOf + '</div>' +
    '</div>' +
    '<div class="kpi-dom-note" data-i18n="domStripNote">† Derived USD memo from the Domestic Orders snapshot — Not included in MA/DI USD totals (source workbook policy). Click a card to open the Domestic tab.</div>' +
  '</div>') : '';

const dashRowsJsonStr = neutralize(JSON.stringify(dashRows));
const yearValuesJsonStr = neutralize(JSON.stringify(yearValues));
const monthValuesJsonStr = neutralize(JSON.stringify(monthValues));
const custValuesJsonStr = neutralize(JSON.stringify(custValues));
const vendValuesJsonStr = neutralize(JSON.stringify(vendValues));
const countryValuesJsonStr = neutralize(JSON.stringify(countryValues));

// ---- GLOBE data island (< 10 KB budget): per-year country aggregates for the
// Dashboard's 3D earth. Sales mode = Cust_Countries (dashRows col 4) x
// Net_Sales_USD (col 5); Purchase mode = Country of origin (col 8) x cost
// (col 6, 'Ttl PO USD' or its 'Matl_USD' alias - same resolved costIdx the KPIs
// use). Bucketed BY YEAR because the dashboard's Year multiselect defaults to
// 2026 and the globe must match that default view exactly (MEXICO/COLOMBIA/
// DENMARK are 2025-only destinations and must be absent from it); the client
// sums the buckets for whichever years are selected, so selecting both years
// reproduces the all-time KPI totals too. Marker size + share% are computed
// client-side from these sums (GPR sizes 12..100, main.js:181-190).
// Ledger names never in COUNTRY_COORDS get h:null: console-warned HERE at
// build time AND surfaced on the page - never silently re-pointed to some
// default city like GPR's '上海' fallback does (main.js:185).
const MIDDLE_EAST_HUB = { city: '利雅得', lon: 46.71, lat: 24.71 }; // user-confirmed: Riyadh
const COUNTRY_COORDS = {
  // 15 sales-side destinations (ledger values, normalized):
  AUSTRALIA: '悉尼', CANADA: '多伦多', COLOMBIA: '波哥大', DENMARK: '哥本哈根',
  GERMANY: '柏林', HK: '香港', INDIA: '孟买', JAPAN: '东京', KOREA: '首尔',
  MEXICO: '墨西哥城', 'MIDDLE-EAST': MIDDLE_EAST_HUB.city, SINGAPORE: '新加坡',
  THAILAND: '曼谷', UK: '伦敦', USA: '纽约',
  // Country-of-origin side (BRIEF addendum: 6 more coords or 87.27% of 2026
  // cost would drop into the unlocated list - CHINA alone is 10.77M USD):
  CHINA: '上海', TAIWAN: '台北', VIETNAM: '河内', ITALY: '米兰',
  INDONESIA: '雅加达', BANGLADESH: '达卡',
};
const globeByYear = {};
function globeBucket(year, mode, name) {
  if (!globeByYear[year]) globeByYear[year] = { sales: {}, purchase: {} };
  const b = globeByYear[year][mode];
  if (!b[name]) b[name] = { usd: 0, n: 0 };
  return b[name];
}
dashRows.forEach(function (r) {
  const y = String(r[0] == null ? '' : r[0]).trim();
  if (!y) return;
  const dest = String(r[4] == null ? '' : r[4]).trim();
  if (dest) { const e = globeBucket(y, 'sales', dest); e.usd += r[5]; e.n += 1; }
  const orig = String(r[8] == null ? '' : r[8]).trim();
  if (orig) { const e = globeBucket(y, 'purchase', orig); e.usd += r[6]; e.n += 1; }
});
const globeUnlocated = { sales: [], purchase: [] };
const globeAggYears = {};
Object.keys(globeByYear).sort().forEach(function (y) {
  globeAggYears[y] = {};
  ['sales', 'purchase'].forEach(function (mode) {
    const map = globeByYear[y][mode];
    const total = Object.keys(map).reduce(function (s, k) { return s + map[k].usd; }, 0);
    const list = Object.keys(map).sort(function (a, b) { return map[b].usd - map[a].usd; }).map(function (name) {
      const hub = COUNTRY_COORDS[name.trim().toUpperCase()];
      if (!hub && globeUnlocated[mode].indexOf(name) < 0) globeUnlocated[mode].push(name);
      return { c: name, h: hub || null, usd: map[name].usd, n: map[name].n, s: total ? map[name].usd / total * 100 : 0 };
    });
    globeAggYears[y][mode] = list;
    globeAggYears[y][mode + 'Total'] = total;
  });
});
if (globeUnlocated.sales.length) console.warn('GLOBE: unlocated sales-destination value(s), no marker drawn: ' + globeUnlocated.sales.join(', '));
if (globeUnlocated.purchase.length) console.warn('GLOBE: unlocated country-of-origin value(s), no marker drawn: ' + globeUnlocated.purchase.join(', '));
const globeAgg = {
  v: 1,
  unit: { sales: 'Net_Sales_USD (dashRows col 5)', purchase: 'Ttl PO USD or Matl_USD (dashRows col 6)' },
  years: globeAggYears,
  unlocated: globeUnlocated,
};
const globeDataIsland = escapeIsland(JSON.stringify(globeAgg), 'globe data island');
if (globeDataIsland.length > 10240) {
  throw new Error('GLOBE BUILD GUARD: mtc-globe-data island is ' + globeDataIsland.length + ' B - over the 10 KB budget');
}
console.log('Globe islands: three ' + (threeIsland.length / 1024).toFixed(0) + ' KB, land ' + (landIsland.length / 1024).toFixed(0) +
  ' KB, data ' + (globeDataIsland.length / 1024).toFixed(2) + ' KB (' + Object.keys(globeAggYears).length +
  ' year bucket(s); unlocated: sales=[' + globeUnlocated.sales.join(',') + '] purchase=[' + globeUnlocated.purchase.join(',') + '])');

// ---- GLOBE CSS for the Dashboard sheet. The 15 --globe-* token names are
// IDENTICAL to the GPR engine's (build/styles-original/app.css :37-51 dark and
// :72-86 light) so globe.js needs zero color-code changes - only the VALUES are
// restated into MTC's navy family (ink #051C2C, accent #1a5c8a, deep navy
// #123f5e, muted gold + brick as the only warm notes). They are defined on
// :root / :root[data-theme="dark"] because the engine's tok() reads
// getComputedStyle(documentElement) - body-scoped vars would be invisible to
// it (exactly the skill's dark-smear common mistake). Numerics keep GPR's
// per-theme values (.28/.9/.88/.1/.28 dark; .86/.2/.75/.15/.2 light).
const GLOBE_CSS = `
  /* globe tokens - light (default). Deliberately deeper than GPR's paper
     values (#dfe7f1/#f4f8fc read near-white against a white card - headless
     screenshot review 2026-11-09): the lit ocean stays a clear cool gray-blue
     so the planet silhouette separates from --card-bg, and the land ink is
     the navy family's darkest. Token NAMES identical to GPR; only values. */
  :root {
    --globe-deep: #b9cbdc; --globe-sea: #d3e0ec; --globe-land: #174d73; --globe-landnight: #7d99b4;
    --globe-city: #b4690e; --globe-coast: #123f5e; --globe-grid: #9db0c4; --globe-air: #7fa0c2;
    --globe-node: #1a5c8a; --globe-node-warn: #A63D40;
    --globe-night-k: .86; --globe-air-a: .2; --globe-coast-a: .75; --globe-grid-a: .15; --globe-orbit-a: .2;
  }
  /* globe tokens - dark */
  :root[data-theme="dark"] {
    --globe-deep: #041220; --globe-sea: #0e3350; --globe-land: #4f8fc0; --globe-landnight: #1c3d5a;
    --globe-city: #C4B891; --globe-coast: #8fb8d8; --globe-grid: #274761; --globe-air: #4d8fb8;
    --globe-node: #5fb2c4; --globe-node-warn: #d9a441;
    --globe-night-k: .28; --globe-air-a: .9; --globe-coast-a: .88; --globe-grid-a: .1; --globe-orbit-a: .28;
  }
  .globe-section { background: var(--card-bg); border: 1px solid var(--border); border-radius: 8px; padding: 12px 14px 10px; margin: 0 0 14px 0; box-shadow: var(--shadow); }
  .globe-head { display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; flex-wrap: wrap; margin-bottom: 6px; }
  .globe-eyebrow { font-size: 10px; letter-spacing: .14em; text-transform: uppercase; color: var(--text-muted); }
  .globe-head h3 { margin: 2px 0 0 0; font-size: 15px; color: var(--accent); }
  .globe-modes { display: flex; gap: 6px; }
  .globe-mode { border: 1px solid var(--border); background: var(--input-bg); color: var(--text); padding: 5px 12px; border-radius: 5px; cursor: pointer; font-size: 12px; }
  .globe-mode.on { background: var(--accent); border-color: var(--accent); color: #ffffff; font-weight: bold; }
  .globe-stage { position: relative; }
  #globeEarth { position: relative; width: 100%; height: min(56vh, 600px); min-height: 360px; }
  /* The report-wide "canvas { max-height: 320px }" default would crush the
     globe canvas into the same band as the Chart.js charts - exempt it here. */
  #globeEarth canvas { max-height: none; }
  .globe-status { position: absolute; inset: 0; display: flex; align-items: center; justify-content: center; color: var(--text-muted); font-size: 13px; pointer-events: none; }
  .globe-ok .globe-status { display: none; }
  .globe-ok .geo-fallback { display: none; }
  .globe-section.globe-fail #globeEarth { display: none; }
  .globe-section.globe-fail .globe-status { position: static; min-height: 40px; }
  .globe-readout { position: absolute; left: 10px; bottom: 10px; background: var(--card-bg); border: 1px solid var(--border); border-radius: 6px; padding: 7px 10px; min-width: 190px; color: var(--text); }
  .globe-readout span { display: block; font-size: 10px; color: var(--text-muted); text-transform: uppercase; letter-spacing: .08em; }
  .globe-readout strong { display: block; font-size: 14px; margin-top: 1px; }
  .globe-readout b { display: block; font-size: 13px; color: var(--accent); font-variant-numeric: tabular-nums; }
  .globe-legend { position: absolute; right: 10px; top: 10px; background: var(--card-bg); border: 1px solid var(--border); border-radius: 3px; padding: 8px 10px; font-size: 11px; color: var(--text); width: 272px; max-height: calc(100% - 20px); overflow-y: auto; }
  .globe-legend .lg-k { font-size: 9px; letter-spacing: .12em; text-transform: uppercase; color: var(--text-muted); margin-bottom: 4px; }
  .globe-legend .lg-row { display: grid; grid-template-columns: 20px minmax(0, 1fr) 52px 38px; align-items: center; gap: 6px; width: 100%; text-align: left; background: none; border: 0; border-bottom: 1px solid var(--border-light); padding: 4px 2px; cursor: pointer; font-variant-numeric: tabular-nums; color: var(--text); font-size: 11px; }
  .globe-legend .lg-row[disabled] { cursor: default; opacity: .55; }
  .globe-legend .lg-row:hover, .globe-legend .lg-row.sel { background: var(--hover-bg); }
  .globe-legend .lg-no { color: var(--text-muted); font-size: 9px; }
  .globe-legend .lg-c { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .globe-legend .lg-v { text-align: right; }
  .globe-legend .lg-p { color: var(--text-muted); text-align: right; }
  .globe-legend .lg-track { grid-column: 2 / -1; height: 2px; background: var(--border-light); }
  .globe-legend .lg-track i { display: block; height: 2px; background: var(--accent); }
  .globe-legend .lg-s { margin-top: 6px; font-size: 9px; color: var(--text-muted); }
  .globe-foot { margin-top: 6px; font-size: 10.5px; color: var(--text-muted); display: flex; flex-wrap: wrap; gap: 4px 14px; }
  .globe-foot .globe-unlocated { color: var(--accent); }
  @media (max-width: 700px) { .globe-legend { position: static; width: auto; max-height: none; margin-top: 8px; } #globeEarth { min-height: 300px; } }
`;

const dashboardHtml = `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Dashboard</title>
<style>${dataTablesCss}${SHARED_CSS}${GLOBE_CSS}
  h2 { color: var(--accent); margin-top: 0; }
  /* Overrides the report-wide "canvas { max-height: 320px }" default - each
     trend chart's per-point labels carry a YoY % ("1.8M (-35.8% YoY)"),
     noticeably longer than a bare value, and 16 bars (2025+2026 x 8 months)
     side by side need more vertical room than the standard height allows. */
  #chartSalesCost, #chartMargin { max-height: 420px; }
${MCKINSEY_CSS}
</style>
</head>
<body>
<h2 data-i18n="dashboardH2">MTC Sales &amp; Margin Performance</h2>
<!-- Action title: the page's one-line conclusion, computed by renderAction()
     from the same aggregates as the KPI cards (never hand-typed, so it cannot
     drift). Filled on the first render() below and on every filter/theme/
     language re-render - no data-i18n attribute here on purpose: applyLanguage's
     textContent sweep would fight renderAction for this element. -->
<p class="action-title" id="dashAction"></p>
<p class="page-intro" data-i18n="introDashboard">Filter by year, month, customer, vendor or country - the 3D trade globe, 4 KPIs and 6 charts update instantly.</p>
<div class="filters">
  <div><label class="fl" data-i18n="filterYear">Year</label><div class="ms" id="ms-year"></div></div>
  <div><label class="fl" data-i18n="filterMonth">Month</label><div class="ms" id="ms-month"></div></div>
  <div><label class="fl" data-i18n="filterCustomer">Customer</label><div class="ms" id="ms-cust"></div></div>
  <div><label class="fl" data-i18n="filterVendor">Vendor</label><div class="ms" id="ms-vend"></div></div>
  <div><label class="fl" data-i18n="filterCountry">Country</label><div class="ms" id="ms-country"></div></div>
  <div style="align-self:flex-end"><button id="resetBtn" data-i18n="resetFilters">Reset filters</button></div>
</div>
<div class="kpis">
  <div class="kpi"><div class="label" data-i18n="kpiNetSales">Net Sales (USD)</div><div class="value" id="kpi-sales">-</div><div class="kpi-sub"><span id="kpi-sales-prior">-</span><span class="kpi-delta" id="kpi-sales-delta"></span></div></div>
  <div class="kpi"><div class="label" data-i18n="kpiNetCost">Net Cost (USD)</div><div class="value" id="kpi-cost">-</div><div class="kpi-sub"><span id="kpi-cost-prior">-</span><span class="kpi-delta" id="kpi-cost-delta"></span></div></div>
  <div class="kpi"><div class="label" data-i18n="kpiNetMargin">Net Margin (USD)</div><div class="value" id="kpi-margin">-</div><div class="kpi-sub"><span id="kpi-margin-prior">-</span><span class="kpi-delta" id="kpi-margin-delta"></span></div></div>
  <div class="kpi"><div class="label" data-i18n="kpiNetMarginPct">Net Margin %</div><div class="value" id="kpi-marginpct">-</div><div class="kpi-sub"><span id="kpi-marginpct-prior">-</span><span class="kpi-delta" id="kpi-marginpct-delta"></span></div></div>
</div>
${domStripHtml}
<div class="globe-section" id="globeSection">
  <div class="globe-head">
    <div>
      <div class="globe-eyebrow"><span data-i18n="exhibitWord">Exhibit</span> 1 — <span data-i18n="globeKicker">Global Trade Flow</span></div>
      <h3 data-i18n="globeTitle">Where the money goes - and where it comes from</h3>
    </div>
    <div class="globe-modes" role="group" aria-label="Globe mode">
      <button type="button" class="globe-mode on" id="globeModeSales" data-i18n="globeModeSales">Global Sales</button><button type="button" class="globe-mode" id="globeModePurchase" data-i18n="globeModePurchase">Global Purchase</button>
    </div>
  </div>
  <div class="globe-stage">
    <div class="globe-status" id="globeStatus" data-i18n="globeBuilding">Building the 3D earth...</div>
    <div id="globeEarth" aria-label="3D globe of MTC trade flows"></div>
    <div class="globe-readout"><span id="globeReadoutLabel">-</span><strong id="globeReadoutName">-</strong><b id="globeReadoutValue">-</b></div>
    <div class="globe-legend" id="globeLegend"></div>
  </div>
  <div class="globe-foot">
    <span class="globe-help" data-i18n="globeHelp">Drag to rotate - hover a marker for the readout - click one to open the Data tab pre-filtered</span>
    <span class="globe-note" data-i18n="globeNote">Marker hubs are indicative capital/economic-city positions; every USD figure comes straight from the embedded DATA ledger (Net Sales / Net Cost).</span>
    <span class="globe-unlocated" id="globeUnlocatedNote"></span>
  </div>
</div>
<div class="charts">
  <div class="chart-box wide">
    <div class="exhibit-kicker"><span data-i18n="exhibitWord">Exhibit</span> 2</div>
    <canvas id="chartSalesCost"></canvas>
  </div>
</div>
<div class="charts">
  <div class="chart-box wide">
    <div class="exhibit-kicker"><span data-i18n="exhibitWord">Exhibit</span> 3</div>
    <canvas id="chartMargin"></canvas>
  </div>
</div>
<div class="charts">
  <div class="chart-box">
    <div class="exhibit-kicker"><span data-i18n="exhibitWord">Exhibit</span> 4</div>
    <canvas id="chartVendor"></canvas>
  </div>
  <div class="chart-box">
    <div class="exhibit-kicker"><span data-i18n="exhibitWord">Exhibit</span> 5</div>
    <canvas id="chartCustomer"></canvas>
  </div>
</div>
<div class="charts">
  <div class="chart-box geo-fallback">
    <div class="exhibit-kicker"><span data-i18n="exhibitWord">Exhibit</span> 6</div>
    <canvas id="chartGeo"></canvas>
  </div>
  <div class="chart-box geo-fallback">
    <div class="exhibit-kicker"><span data-i18n="exhibitWord">Exhibit</span> 7</div>
    <canvas id="chartOrigin"></canvas>
  </div>
</div>
<div class="source-foot"><span data-i18n="footSource">Sources</span> — data_sheet.json (DATA ledger island) · mtc-globe-data island (per-year country aggregates, this page) · ${reportPeriodRange} · ${FX_SOURCE} · derived domestic FX basis: ${fxAsOfNote}</div>
<script>${jqueryJs}</script>
<script>${chartJs}</script>
<script>
${I18N_JS}
${MULTISELECT_JS}
var ROWS = ${dashRowsJsonStr};
var YEAR_VALUES = ${yearValuesJsonStr};
var MONTH_VALUES = ${monthValuesJsonStr};
var CUST_VALUES = ${custValuesJsonStr};
var VEND_VALUES = ${vendValuesJsonStr};
var COUNTRY_VALUES = ${countryValuesJsonStr};
// Fixed categorical colors, restyled to the consulting palette: one navy
// family (deep -> light blue ramp + steel grays) plus a single warm note
// (muted gold). Muted brick #A63D40 is NOT a series color - the token pass
// reserves it for negatives/cautions only. All six-digit hex on purpose: the
// bar-in-bar / prior-year datasets append '55' alpha suffixes (e.g.
// C_SALES + '55'), which breaks on 3-digit or rgb() forms. Legible on both
// light (#FFFFFF cards) and dark (#12212B cards) backgrounds.
var C_SALES = '#1A5C8A', C_COST = '#C4A34A', C_MARGIN = '#0F3D5C';
var CATEGORY_PALETTE = ['#1A5C8A', '#0F3D5C', '#2E7DAF', '#6FA3C7', '#C4A34A', '#9A7B2E', '#3A5566', '#6E8494', '#A3B4C0', '#17456B', '#D8C89A', '#253038'];

// Compact on-chart number format (1.2M / 563K / 850) - full precision stays
// available on hover (tooltip). Shorter labels are the single biggest lever
// against overlap on a long-tail bar chart or a pie with one dominant slice.
function fmtCompact(n) {
  var a = Math.abs(n);
  if (a >= 1e6) return (n / 1e6).toFixed(1).replace(/\.0$/, '') + 'M';
  if (a >= 1e3) return (n / 1e3).toFixed(1).replace(/\.0$/, '') + 'K';
  return Math.round(n).toLocaleString();
}
// Draws "value (pct%)" directly on every bar/slice/point - every chart on this
// dashboard shows exact figures on the chart itself, not just on hover, per
// explicit requirement. A stroke outline keeps text legible over any fill
// color/theme. Opt in per-chart via options.plugins.valuePctLabels = { enabled:
// true, showPct: false to omit the percentage (e.g. Monthly Trend, where "% of
// all 19 months" isn't a meaningful per-point read), dy: vertical offset to
// separate overlapping series' labels on a multi-line chart.
Chart.register({
  id: 'valuePctLabels',
  afterDatasetsDraw: function (chart) {
    var opt = chart.options.plugins && chart.options.plugins.valuePctLabels;
    if (!opt || !opt.enabled) return;
    var ctx = chart.ctx;
    var isDark = document.body.getAttribute('data-theme') === 'dark';
    var isPie = chart.config.type === 'pie' || chart.config.type === 'doughnut';
    var isHBar = chart.config.type === 'bar' && chart.options.indexAxis === 'y';
    var isVBar = chart.config.type === 'bar' && chart.options.indexAxis !== 'y';
    chart.data.datasets.forEach(function (dataset, dsIndex) {
      var meta = chart.getDatasetMeta(dsIndex);
      if (meta.hidden) return;
      // dataset.total (when the caller supplies it) is the TRUE grand total
      // across every category, not just whatever subset is currently drawn -
      // required once a chart truncates to a top-N slice (e.g. Top 15
      // vendors/customers), where summing only the visible bars would
      // overstate each one's real share of the whole. Falls back to summing
      // the visible data for any chart that doesn't truncate/doesn't set it.
      var total = typeof dataset.total === 'number' ? dataset.total : dataset.data.reduce(function (s, v) { return s + (typeof v === 'number' ? v : 0); }, 0);
      var dy = (opt.dyPerDataset ? opt.dyPerDataset * dsIndex : 0) + (opt.dy || 0);
      meta.data.forEach(function (element, index) {
        var value = dataset.data[index];
        if (typeof value !== 'number') return;
        var pct = total ? value / total * 100 : 0;
        var label = fmtCompact(value);
        // dataset.yoyPct (only the Dashboard's Monthly Trend chart sets this)
        // is a per-point % change vs the SAME MONTH last year - a completely
        // different question from "this point's share of the dataset total"
        // (which showPct:false already correctly suppresses for a trend
        // line), so it's checked first and shown instead of, not alongside,
        // the normal share-of-total percentage.
        var yoyPct = dataset.yoyPct && typeof dataset.yoyPct[index] === 'number' ? dataset.yoyPct[index] : null;
        if (yoyPct !== null) {
          // compactYoy drops the " YoY" word on-chart (kept in the tooltip's
          // own fuller text) - useful once a chart packs enough bars that
          // every character of on-bar text narrows the margin before
          // adjacent bars' labels start overlapping horizontally.
          label += ' (' + (yoyPct >= 0 ? '+' : '') + yoyPct.toFixed(1) + '%' + (opt.compactYoy ? '' : ' ' + t('yoyLabel')) + ')';
        } else if (opt.showPct !== false && total) {
          label += ' (' + pct.toFixed(1) + '%)';
        }
        var x, y, align = 'center', baseline = 'middle';
        if (isHBar) {
          // Anchored just past the bar's own end (not centered inside it) -
          // a short bar's centered label would spill past the bar and
          // collide with the category axis to the left.
          x = element.x + 6; y = element.y; align = 'left';
        } else if (isVBar) {
          // A long tail of many adjacent small bars leaves too little
          // horizontal room per label - zigzag alternate bars' labels to a
          // second, higher tier so neighbors' text doesn't run together.
          var vTier = (opt.zigzag && index % 2 === 1) ? 16 : 0;
          x = element.x; y = element.y - 6 - vTier; align = 'center'; baseline = 'bottom';
        } else if (isPie && pct < (opt.smallSlicePct || 8)) {
          // A thin slice's own centroid sits right next to its neighbors'
          // centroids - push the label out past the pie's edge along the
          // slice's own angle, with a thin leader line back to the slice
          // (staggering the leader length by index too, so several tiny
          // adjacent slices don't all land on the same ring and overlap) -
          // the standard fix for label crowding on a skewed pie.
          var midAngle = (element.startAngle + element.endAngle) / 2;
          var tier = index % 2;
          var rInner = element.outerRadius + 8 + tier * 18;
          var rOuter = element.outerRadius + 24 + tier * 18;
          var lx1 = element.x + Math.cos(midAngle) * rInner, ly1 = element.y + Math.sin(midAngle) * rInner;
          var lx2 = element.x + Math.cos(midAngle) * rOuter, ly2 = element.y + Math.sin(midAngle) * rOuter;
          ctx.save();
          ctx.beginPath(); ctx.moveTo(lx1, ly1); ctx.lineTo(lx2, ly2);
          ctx.strokeStyle = isDark ? '#7a8a94' : '#888888'; ctx.lineWidth = 1; ctx.stroke();
          ctx.restore();
          x = lx2 + Math.cos(midAngle) * 3; y = ly2;
          align = Math.cos(midAngle) > 0.15 ? 'left' : (Math.cos(midAngle) < -0.15 ? 'right' : 'center');
        } else {
          var pos = element.getCenterPoint ? element.getCenterPoint() : element.tooltipPosition();
          x = pos.x; y = pos.y + dy;
        }
        ctx.save();
        ctx.font = (opt.fontSize || 10) + 'px Calibri, Arial, sans-serif';
        ctx.textAlign = align;
        ctx.textBaseline = baseline;
        ctx.lineWidth = 3;
        ctx.strokeStyle = isDark ? 'rgba(16,22,28,0.85)' : 'rgba(255,255,255,0.85)';
        ctx.strokeText(label, x, y);
        ctx.fillStyle = opt.color || (isDark ? '#f0f0f0' : '#222');
        ctx.fillText(label, x, y);
        ctx.restore();
      });
    });
  }
});

// Year defaults to 2026 only on first load, per requirement - every chart and
// KPI starts scoped to the current year, not all-time, until the user widens it.
var msYear = createMultiSelect('ms-year', YEAR_VALUES, onFilterChange, YEAR_VALUES.indexOf('2026') >= 0 ? ['2026'] : []);
var msMonth = createMultiSelect('ms-month', MONTH_VALUES, onFilterChange);
var msCust = createMultiSelect('ms-cust', CUST_VALUES, onFilterChange);
var msVend = createMultiSelect('ms-vend', VEND_VALUES, onFilterChange);
var msCountry = createMultiSelect('ms-country', COUNTRY_VALUES, onFilterChange);
function onFilterChange() { render(); }

function getFiltered() {
  var selYear = msYear.getSelected(), selMonth = msMonth.getSelected(), selCust = msCust.getSelected(),
      selVend = msVend.getSelected(), selCountry = msCountry.getSelected();
  return ROWS.filter(function (r) {
    return matchesFilter(r[0], selYear) && matchesFilter(r[1], selMonth) && matchesFilter(r[2], selCust) &&
           matchesFilter(r[3], selVend) && matchesFilter(r[4], selCountry);
  });
}
function groupSum(rows, keyIdx, valIdx) {
  var map = {};
  rows.forEach(function (r) { var k = r[keyIdx]; if (!map[k]) map[k] = 0; map[k] += (typeof r[valIdx] === 'number' ? r[valIdx] : 0); });
  return map;
}
// "2026-03" -> "2025-03" (same month, prior year) - the Monthly Trend chart's
// yyyy-mm values are always this exact format.
function priorYearMonth(m) {
  var parts = String(m).split('-');
  return (parseInt(parts[0], 10) - 1) + '-' + parts[1];
}
// Sums for each shown month's SAME MONTH last year, so the trend chart can
// show a real YoY %, not just this year's own value. Deliberately reads from
// the full, unfiltered ROWS (not getFiltered()'s result) and skips the
// Year/Month multiselects entirely - the whole point is looking outside
// whatever year(s)/month(s) those filters currently have selected - but
// still respects Customer/Vendor/Country, so the comparison stays apples-to-
// apples with whichever subset of the business the user is currently viewing.
function priorYearSums(months) {
  var selCust = msCust.getSelected(), selVend = msVend.getSelected(), selCountry = msCountry.getSelected();
  var sums = { sales: {}, cost: {}, margin: {} };
  var priorMonths = {};
  months.forEach(function (m) { priorMonths[priorYearMonth(m)] = m; });
  ROWS.forEach(function (r) {
    var curMonth = priorMonths[r[1]];
    if (!curMonth) return;
    if (!matchesFilter(r[2], selCust) || !matchesFilter(r[3], selVend) || !matchesFilter(r[4], selCountry)) return;
    sums.sales[curMonth] = (sums.sales[curMonth] || 0) + r[5];
    sums.cost[curMonth] = (sums.cost[curMonth] || 0) + r[6];
    sums.margin[curMonth] = (sums.margin[curMonth] || 0) + r[7];
  });
  return sums;
}
function fmt(n) { return Math.round(n).toLocaleString(); }
function withLabelHeadroom(values) {
  var max = values.reduce(function (m, v) { return Math.max(m, v); }, 0);
  return { x: { suggestedMax: max * 1.35 } };
}
var charts = {};
var lastRows = [];

// Two-year same-period comparison on every card (GPR dashboard convention):
// the big number stays the filtered current total; the sub-line shows the
// SAME CALENDAR MONTHS last year plus the +/- change, using priorYearSums -
// the identical basis the action title and the trend chart's YoY annotations
// use, so a card can never contradict the exhibits below it. Colors encode
// DIRECTION only (up=accent, down=brick), not good/bad: a cost decline reads
// as a fall, viewers judge whether that is good. No prior-year data for the
// shown months (e.g. Year=2025 with no 2024 rows) prints an em dash, never a
// fake 0% or a crash.
function renderKpis() {
  var rows = lastRows;
  var sales = 0, cost = 0, margin = 0;
  rows.forEach(function (r) { sales += r[5]; cost += r[6]; margin += r[7]; });
  document.getElementById('kpi-sales').textContent = fmt(sales);
  document.getElementById('kpi-cost').textContent = fmt(cost);
  document.getElementById('kpi-margin').textContent = fmt(margin);
  document.getElementById('kpi-marginpct').textContent = sales ? (margin / sales * 100).toFixed(1) + '%' : '-';
  var months = Array.from(new Set(rows.map(function (r) { return r[1]; }))).filter(Boolean).sort();
  var prior = priorYearSums(months);
  var pS = 0, pC = 0, pM = 0;
  Object.keys(prior.sales).forEach(function (k) { pS += prior.sales[k] || 0; pC += prior.cost[k] || 0; pM += prior.margin[k] || 0; });
  var hasPrior = Object.keys(prior.sales).length > 0;
  var priorLabel = '—';
  if (months.length) {
    var p0 = priorYearMonth(months[0]), p1 = priorYearMonth(months[months.length - 1]);
    priorLabel = (p0 === p1) ? p0 : p0 + ' → ' + p1;
  }
  function card(id, curText, prvOk) {
    var el = document.getElementById('kpi-' + id + '-prior');
    var dl = document.getElementById('kpi-' + id + '-delta');
    if (!hasPrior || !prvOk) { el.textContent = t('kpiVsLabel') + ' ' + priorLabel + ': —'; dl.textContent = ''; dl.className = 'kpi-delta'; return; }
    el.textContent = t('kpiVsLabel') + ' ' + priorLabel + ': ' + curText.p;
    dl.textContent = curText.d;
    dl.className = 'kpi-delta ' + (curText.dir > 0 ? 'up' : (curText.dir < 0 ? 'down' : ''));
  }
  function pctTxt(cur, prv) {
    if (!prv) return { p: '—', d: '—', dir: 0 };
    var ch = (cur - prv) / prv * 100;
    return { p: fmt(prv), d: (ch >= 0 ? '+' : '') + ch.toFixed(1) + '%', dir: Math.sign(ch) };
  }
  card('sales', pctTxt(sales, pS), pS);
  card('cost', pctTxt(cost, pC), pC);
  card('margin', pctTxt(margin, pM), pM);
  var curPct = sales ? margin / sales * 100 : null;
  var prvPct = pS ? pM / pS * 100 : null;
  var pp = (curPct !== null && prvPct !== null) ? { p: prvPct.toFixed(1) + '%', d: (curPct - prvPct >= 0 ? '+' : '') + (curPct - prvPct).toFixed(1) + ' pp', dir: Math.sign(curPct - prvPct) } : { p: '—', d: '—', dir: 0 };
  card('marginpct', pp, prvPct !== null);
}
// ---- Action title (consulting convention: a chart page opens with its
// conclusion, not its label). One line computed from the SAME aggregates the
// KPI cards show, plus the true YoY delta over exactly the months currently
// displayed - priorYearSums is the same same-months-last-year basis the two
// trend charts' YoY annotations use, so the headline can never contradict the
// exhibits under it. Placeholders {S}/{D}/{M}/{P}/{T} live in the
// actionDashboard i18n string (en + zh), so language toggle only swaps the
// template; the numbers recompute from the current filters either way.
// .neg marks a down-year (muted brick, the reserved negative color).
function renderAction() {
  var rows = lastRows;
  var sales = 0, margin = 0;
  rows.forEach(function (r) { sales += r[5]; margin += r[7]; });
  var months = Array.from(new Set(rows.map(function (r) { return r[1]; }))).filter(Boolean).sort();
  var prior = priorYearSums(months);
  var priorSales = Object.keys(prior.sales).reduce(function (s, k) { return s + (prior.sales[k] || 0); }, 0);
  var delta = priorSales ? (sales - priorSales) / priorSales * 100 : null;
  var deltaTxt = delta === null ? '—' : (delta >= 0 ? '+' : '') + delta.toFixed(1) + '%';
  var period = months.length ? (months.length > 1 ? months[0] + ' to ' + months[months.length - 1] : months[0]) : '—';
  var s = t('actionDashboard')
    .replace('{S}', fmtCompact(sales))
    .replace('{D}', deltaTxt)
    .replace('{M}', fmtCompact(margin))
    .replace('{P}', sales ? (margin / sales * 100).toFixed(1) + '%' : '—')
    .replace('{T}', period);
  var el = document.getElementById('dashAction');
  el.textContent = s;
  el.classList.toggle('neg', delta === null ? margin < 0 : delta < 0);
}

// ---- Monthly Trend: Sales/Cost/Margin, as paired bars (prior-year month
// immediately followed by this year's same month) instead of a line chart -
// per explicit stakeholder request (Peggy, relaying Keith's team): "can we
// show 2025 vs 2026 at the same time" as literal side-by-side bars, not just
// a hover tooltip. Same null-when-no-prior-year-data rule as before - a
// brand-new customer/vendor with no history shows no YoY % rather than a
// meaningless "+Infinity%"/"-100%" against a zero base.
function buildYoyBarSeries(rows) {
  var months = Array.from(new Set(rows.map(function (r) { return r[1]; }))).filter(Boolean).sort();
  var salesByMonth = groupSum(rows, 1, 5), costByMonth = groupSum(rows, 1, 6), marginByMonth = groupSum(rows, 1, 7);
  var prior = priorYearSums(months);
  var labels = [], sales = [], cost = [], margin = [], yoySales = [], yoyCost = [], yoyMargin = [];
  months.forEach(function (m) {
    labels.push(priorYearMonth(m), m);
    sales.push(prior.sales[m] || 0, salesByMonth[m] || 0);
    cost.push(prior.cost[m] || 0, costByMonth[m] || 0);
    margin.push(prior.margin[m] || 0, marginByMonth[m] || 0);
    // null on the prior-year (odd) tick - that bar IS the baseline, so it has
    // no "YoY" of its own to show; only its paired current-year tick does.
    yoySales.push(null, prior.sales[m] ? (salesByMonth[m] - prior.sales[m]) / prior.sales[m] * 100 : null);
    yoyCost.push(null, prior.cost[m] ? (costByMonth[m] - prior.cost[m]) / prior.cost[m] * 100 : null);
    yoyMargin.push(null, prior.margin[m] ? (marginByMonth[m] - prior.margin[m]) / prior.margin[m] * 100 : null);
  });
  return { labels: labels, sales: sales, cost: cost, margin: margin, yoySales: yoySales, yoyCost: yoyCost, yoyMargin: yoyMargin };
}
// Shared tooltip title for both trend charts: "2025-01 vs 2026-01" for
// EITHER bar in the pair, not just the current-year one, so hovering either
// side of the pair orients the viewer the same way.
function yoyPairTooltipTitle(labels, items) {
  var idx = items[0].dataIndex;
  var pairStart = idx - (idx % 2);
  return labels[pairStart] + ' ' + t('vsLabel') + ' ' + labels[pairStart + 1];
}
function yoyPairTooltipLabel(ctx) {
  var yoy = ctx.dataset.yoyPct ? ctx.dataset.yoyPct[ctx.dataIndex] : null;
  var line = ctx.dataset.label + ': ' + Math.round(ctx.parsed.y).toLocaleString();
  if (typeof yoy === 'number') line += '  (' + (yoy >= 0 ? '+' : '') + yoy.toFixed(1) + '% ' + t('yoyLabel') + ')';
  return line;
}
// Sales & Cost chart's own tooltip body: always both years' Sales together
// (2026 with its YoY %, 2025 as the plain baseline), then the same pair for
// Cost - both metrics' 2026 line carries its own YoY % - regardless of which
// of the 3 datasets (Sales bar, Cost bar, Sales dot) is actually under the
// cursor, replacing the default "whichever series you happen to be pointing
// at" per-item line.
function yoySalesPairTooltipBody(s, items) {
  var idx = items[0].dataIndex;
  var pairStart = idx - (idx % 2);
  var i2026 = pairStart + 1, i2025 = pairStart;
  function line2026(label, data, yoyArr) {
    var text = s.labels[i2026] + ' ' + label + ': ' + Math.round(data[i2026]).toLocaleString();
    var yoy = yoyArr[i2026];
    if (typeof yoy === 'number') text += '  (' + (yoy >= 0 ? '+' : '') + yoy.toFixed(1) + '% ' + t('yoyLabel') + ')';
    return text;
  }
  function line2025(label, data) {
    return s.labels[i2025] + ' ' + label + ': ' + Math.round(data[i2025]).toLocaleString();
  }
  return [
    line2026(t('seriesSales'), s.sales, s.yoySales),
    line2025(t('seriesSales'), s.sales),
    line2026(t('seriesCost'), s.cost, s.yoyCost),
    line2025(t('seriesCost'), s.cost)
  ];
}
// "Bar-in-bar + dot markers" combo: Sales and Cost both render as low-
// opacity ("shadow") bars OVERLAID at the same month position (Sales wide/
// outer, Cost narrow/inner, both starting from 0) rather than side by side -
// grouped:false is what makes Chart.js overlay same-category datasets
// instead of offsetting them; each dataset's own barPercentage is what makes
// Cost visibly narrower than Sales. Sales additionally gets a layer of solid
// dot markers (own dataset, showLine:false so there's no connecting line
// between them, just the points) carrying its own on-chart value/YoY % label.
// order controls draw sequence - empirically the LOWER order value draws
// last/on top, so dots (1) paint over Cost (2), which paints over Sales (3).
// The dot dataset shares Sales' own label/color family but is excluded from
// the legend (hideFromLegend) so "Sales" doesn't appear twice.
function renderSalesCostTrend() {
  var s = buildYoyBarSeries(lastRows);
  var datasets = [
    { type: 'bar', label: t('seriesSales'), data: s.sales, backgroundColor: C_SALES + '55', borderColor: C_SALES + '55', barPercentage: 0.85, order: 3 },
    { type: 'bar', label: t('seriesCost'), data: s.cost, backgroundColor: C_COST + '55', borderColor: C_COST + '55', barPercentage: 0.45, yoyPct: s.yoyCost, order: 2 },
    { type: 'line', label: t('seriesSales'), data: s.sales, showLine: false, borderColor: C_SALES, backgroundColor: C_SALES, pointBackgroundColor: C_SALES, pointRadius: 4, yoyPct: s.yoySales, order: 1, hideFromLegend: true }
  ];
  if (charts.salesCost) charts.salesCost.destroy();
  var maxVal = Math.max.apply(null, s.sales.concat(s.cost));
  charts.salesCost = new Chart(document.getElementById('chartSalesCost'), {
    type: 'bar',
    data: { labels: s.labels, datasets: datasets },
    options: {
      responsive: true,
      // nearest+intersect (not the other charts' index+no-intersect): the
      // tooltip content here is built entirely from afterBody, always both
      // years' Sales together, regardless of which single element triggered
      // it - index mode would additionally pass all 3 co-located datasets'
      // own items in, which isn't needed and would add noise if label ever
      // stops being fully suppressed below.
      interaction: { mode: 'nearest', intersect: true },
      // categoryPercentage < the Chart.js default (0.8) opens up a visible
      // gap BETWEEN month-pair groups - with 16 groups on one chart, a wide
      // on-bar label like "1.8M (-35.8%)" is wider than its own bar and was
      // spilling sideways into the neighboring group with default spacing.
      datasets: { bar: { categoryPercentage: 0.65, grouped: false } },
      scales: { y: { suggestedMax: maxVal * 1.3 } },
      plugins: {
        title: { display: true, text: t('trendSalesCostTitle') },
        legend: { position: 'bottom', labels: { filter: function (item) { return !datasets[item.datasetIndex].hideFromLegend; } } },
        valuePctLabels: { enabled: true, showPct: false, fontSize: 11, compactYoy: true },
        // Per explicit request: group both years' Sales into ONE tooltip
        // ("2025-01 vs 2026-01" / 2026 Sales + YoY% / 2025 Sales) instead of
        // showing only whichever single bar/dot is under the cursor - label
        // returns null to suppress that default per-item line entirely, and
        // afterBody supplies the real (always-both-years) content instead.
        tooltip: { callbacks: {
          title: function (items) { return yoyPairTooltipTitle(s.labels, items); },
          label: function () { return null; },
          afterBody: function (items) { return yoySalesPairTooltipBody(s, items); }
        } }
      }
    }
  });
}
// Margin as a true year-over-year comparison: ONE x-tick per month pair
// ("2025-01 vs 2026-01"), with both years overlaid bar-in-bar at that tick -
// 2025 as the wide/outer shadow bar, 2026 as the narrow/inner solid bar -
// instead of the interleaved 16-tick layout. The shared tooltip shows both
// years' Margin plus the variance amount and % (per explicit request), and
// is identical whichever of the two bars is under the cursor. order follows
// the same empirical rule as the Sales & Cost chart: lower draws on top.
function renderMarginTrend() {
  var s = buildYoyBarSeries(lastRows);
  var pairLabels = [], m2025 = [], m2026 = [], yoy2026 = [];
  for (var i = 0; i + 1 < s.labels.length; i += 2) {
    pairLabels.push(s.labels[i] + ' ' + t('vsLabel') + ' ' + s.labels[i + 1]);
    m2025.push(s.margin[i]);
    m2026.push(s.margin[i + 1]);
    yoy2026.push(s.yoyMargin[i + 1]);
  }
  var datasets = [
    { type: 'bar', label: t('seriesMargin') + ' 2025', data: m2025, backgroundColor: C_MARGIN + '55', borderColor: C_MARGIN + '55', barPercentage: 0.85, order: 2 },
    { type: 'bar', label: t('seriesMargin') + ' 2026', data: m2026, backgroundColor: C_MARGIN, borderColor: C_MARGIN, barPercentage: 0.45, yoyPct: yoy2026, order: 1 }
  ];
  if (charts.margin) charts.margin.destroy();
  var maxVal = Math.max.apply(null, m2025.concat(m2026));
  charts.margin = new Chart(document.getElementById('chartMargin'), {
    type: 'bar',
    data: { labels: pairLabels, datasets: datasets },
    options: {
      responsive: true,
      interaction: { mode: 'nearest', intersect: true },
      datasets: { bar: { categoryPercentage: 0.7, grouped: false } },
      scales: { y: { suggestedMax: maxVal * 1.3 } },
      plugins: {
        title: { display: true, text: t('trendMarginTitle') }, legend: { position: 'bottom' },
        valuePctLabels: { enabled: true, showPct: false, fontSize: 11, compactYoy: true },
        tooltip: { callbacks: {
          title: function (items) { return pairLabels[items[0].dataIndex]; },
          label: function () { return null; },
          afterBody: function (items) {
            var idx = items[0].dataIndex;
            var a = m2025[idx], b = m2026[idx], diff = b - a;
            var pct = a ? (diff / a * 100) : null;
            var sign = diff >= 0 ? '+' : '';
            return [
              t('seriesMargin') + ' 2025: ' + Math.round(a).toLocaleString(),
              t('seriesMargin') + ' 2026: ' + Math.round(b).toLocaleString(),
              t('varianceLabel') + ': ' + sign + Math.round(diff).toLocaleString(),
              t('variancePctLabel') + ': ' + (pct === null ? '-' : sign + pct.toFixed(1) + '%')
            ];
          }
        } }
      }
    }
  });
}

// ---- Vendor / Customer / Origin: value basis defaults to Sales (col 5) -
// Vendor overrides to Cost (col 6), since a vendor is who MTC buys FROM, so
// ranking vendors by what's spent with them is the more meaningful reading
// than by Sales (a customer-side figure a vendor has no direct bearing on). ----
function renderCategoryChart(opts) {
  // opts: { chartKey, elId, groupKeyIdx, title, color, sortTop, type, indexAxis, valIdx, legendKey }
  var rows = lastRows;
  var netSum = groupSum(rows, opts.groupKeyIdx, opts.valIdx || 5);
  var basisSum = netSum;
  // Grand totals across EVERY category (before sortTop truncates the label
  // list to the top 15) - the on-chart "%" must read as this category's share
  // of the true overall total, not of just the top-15 slice that happens to
  // be drawn. Using dataset.data.reduce() post-slice (the plugin's fallback)
  // would silently inflate every shown bar's % once a report has more
  // categories than sortTop keeps.
  var netGrandTotal = Object.keys(netSum).reduce(function (s, k) { return s + netSum[k]; }, 0);
  var labels = Object.keys(basisSum).sort(function (a, b) { return basisSum[b] - basisSum[a]; });
  // Collapse everything past the top N into one "Others" bar so the chart
  // still visually accounts for the full total, instead of silently dropping
  // the remaining categories with no on-chart trace of them.
  var othersKey = null;
  if (opts.sortTop && labels.length > opts.sortTop) {
    var shownLabels = labels.slice(0, opts.sortTop);
    othersKey = t('otherCategory');
    netSum[othersKey] = netGrandTotal - shownLabels.reduce(function (s, k) { return s + (netSum[k] || 0); }, 0);
    labels = shownLabels.concat([othersKey]);
  }

  var datasets = [{
    label: t(opts.legendKey || 'legendNetSales'),
    data: labels.map(function (v) { return basisSum[v]; }),
    backgroundColor: labels.map(function (v) { return v === othersKey ? '#9099a3' : opts.color; }),
    total: netGrandTotal
  }];
  if (opts.type === 'pie') {
    datasets = [{ label: datasets[0].label, data: datasets[0].data, backgroundColor: labels.map(function (v, i) { return CATEGORY_PALETTE[i % CATEGORY_PALETTE.length]; }), total: datasets[0].total }];
  }

  if (charts[opts.chartKey]) charts[opts.chartKey].destroy();
  var scales = {};
  if (opts.indexAxis === 'y') {
    var allVals = datasets.reduce(function (acc, ds) { return acc.concat(ds.data); }, []);
    scales = withLabelHeadroom(allVals);
  }
  charts[opts.chartKey] = new Chart(document.getElementById(opts.elId), {
    type: opts.type || 'bar',
    data: { labels: labels, datasets: datasets },
    options: {
      indexAxis: opts.indexAxis, responsive: true, scales: scales,
      // Clicking a bar/slice drills into the Data tab filtered to that one
      // category (Others excluded - it's a rollup of many real values, not
      // one to filter by), the same postMessage contract every table-cell
      // drill-link already uses - so both the summary tables AND the
      // Dashboard's own charts can jump to the underlying transactions.
      onClick: function (evt, elements) {
        if (!opts.dataCol || !elements.length) return;
        var label = labels[elements[0].index];
        if (label === othersKey) return;
        var filters = {};
        filters[opts.dataCol] = label;
        window.parent.postMessage({ type: 'mtcDrilldown', targetTab: 'Data', filters: filters }, '*');
      },
      onHover: function (evt, elements) {
        if (evt.native && evt.native.target) evt.native.target.style.cursor = (opts.dataCol && elements.length) ? 'pointer' : 'default';
      },
      // A pie sized to fill its whole canvas leaves no room for the small-
      // slice leader-line labels (see valuePctLabels above) - they get drawn
      // past the pie's own edge and collide with the title text right above
      // it. Shrinking the pie itself (not just adding canvas padding) frees
      // that ring of space on every side, not just the top.
      radius: opts.type === 'pie' ? '65%' : undefined,
      layout: opts.type === 'pie' ? { padding: { top: 26, bottom: 10 } } : {},
      plugins: {
        title: { display: true, text: opts.title },
        legend: { display: false, position: 'bottom' },
        valuePctLabels: { enabled: true, zigzag: opts.zigzag, fontSize: opts.fontSize || 10 },
        // Same weighted %-of-true-total the on-chart labels use (dataset.total,
        // not a sum of only the bars currently drawn - see valuePctLabels
        // above) - without this the hover tooltip on a truncated top-15 bar
        // chart would show no % at all (bar's default tooltip) or, if summed
        // from just the visible bars, a misleadingly inflated one.
        tooltip: { callbacks: { label: function (ctx) {
          var total = typeof ctx.dataset.total === 'number' ? ctx.dataset.total : ctx.dataset.data.reduce(function (s, v) { return s + v; }, 0);
          var value = typeof ctx.parsed === 'number' ? ctx.parsed : (opts.indexAxis === 'y' ? ctx.parsed.x : ctx.parsed.y);
          var pct = total ? (value / total * 100).toFixed(1) : '0.0';
          var prefix = opts.type === 'pie' ? (ctx.dataset.label + ' - ' + ctx.label + ': ') : (ctx.dataset.label + ': ');
          return prefix + Math.round(value).toLocaleString() + ' (' + pct + '%)';
        } } }
      }
    }
  });
}

function renderVendor() { renderCategoryChart({ chartKey: 'vendor', elId: 'chartVendor', groupKeyIdx: 3, title: t('vendorTitle'), color: C_SALES, sortTop: 15, type: 'bar', indexAxis: 'y', dataCol: 'VEND_NAME', valIdx: 6, legendKey: 'legendNetCost' }); }
function renderCustomer() { renderCategoryChart({ chartKey: 'customer', elId: 'chartCustomer', groupKeyIdx: 2, title: t('customerTitle'), color: C_COST, sortTop: 15, type: 'bar', indexAxis: 'y', dataCol: 'Cust_name' }); }
function renderOrigin() { renderCategoryChart({ chartKey: 'origin', elId: 'chartOrigin', groupKeyIdx: 8, title: t('originTitle'), type: 'pie', dataCol: 'Country of origin' }); }

// ---- Geographic: value basis is Margin (col 7), not Sales - separate from
// renderCategoryChart above, which is Sales-only. ----
function renderGeo() {
  var rows = lastRows;
  var netSum = groupSum(rows, 4, 7);
  var basisSum = netSum;
  // No sortTop here (every country is shown), so this already equals the sum
  // of what's drawn - kept explicit anyway so this chart's % stays correct
  // even if a top-N truncation is ever added here later (see renderCategoryChart).
  var netGrandTotal = Object.keys(netSum).reduce(function (s, k) { return s + netSum[k]; }, 0);
  var labels = Object.keys(basisSum).sort(function (a, b) { return basisSum[b] - basisSum[a]; });
  var datasets = [{ label: t('legendNetMargin'), data: labels.map(function (v) { return basisSum[v]; }), backgroundColor: C_MARGIN, total: netGrandTotal }];
  if (charts.geo) charts.geo.destroy();
  var maxVal = datasets.reduce(function (m, ds) { return Math.max(m, ds.data.reduce(function (mm, v) { return Math.max(mm, v); }, 0)); }, 0);
  charts.geo = new Chart(document.getElementById('chartGeo'), {
    type: 'bar',
    data: { labels: labels, datasets: datasets },
    options: {
      responsive: true,
      scales: { y: { suggestedMax: maxVal * 1.2 } },
      onClick: function (evt, elements) {
        if (!elements.length) return;
        var label = labels[elements[0].index];
        window.parent.postMessage({ type: 'mtcDrilldown', targetTab: 'Data', filters: { 'Cust_Countries': label } }, '*');
      },
      onHover: function (evt, elements) {
        if (evt.native && evt.native.target) evt.native.target.style.cursor = elements.length ? 'pointer' : 'default';
      },
      plugins: {
        title: { display: true, text: t('geoTitle') },
        legend: { display: false, position: 'bottom' },
        valuePctLabels: { enabled: true, zigzag: true, fontSize: 9 },
        tooltip: { callbacks: { label: function (ctx) {
          var total = typeof ctx.dataset.total === 'number' ? ctx.dataset.total : ctx.dataset.data.reduce(function (s, v) { return s + v; }, 0);
          var pct = total ? (ctx.parsed.y / total * 100).toFixed(1) : '0.0';
          return ctx.dataset.label + ': ' + Math.round(ctx.parsed.y).toLocaleString() + ' (' + pct + '%)';
        } } }
      }
    }
  });
}

function render() {
  lastRows = getFiltered();
  renderKpis();
  renderAction();
  renderSalesCostTrend();
  renderMarginTrend();
  renderVendor();
  renderCustomer();
  // Globe module (deferred) registers window.globeUpdateHook and sets
  // window.globeSuccess on a live 3D earth. Until the earth proves itself the
  // two geo charts keep rendering as the visible fallback; once it succeeds
  // they're display:none (.globe-ok .geo-fallback) so re-drawing hidden
  // canvases on every filter change would be pure waste - skipped.
  if (window.globeUpdateHook) window.globeUpdateHook();
  if (!window.globeSuccess) { renderGeo(); renderOrigin(); }
}

function applyTheme(theme) {
  document.body.setAttribute('data-theme', theme);
  // GLOBE contract: globe.js caches CSS tokens off documentElement and
  // invalidates on documentElement[data-theme] mutations + the kit:theme event
  // (globe.js:13-17 + 496). MTC's toggle historically only touched BODY -
  // without these two lines the earth would stay dark on the light theme
  // (skill's named "dark smear" failure). Chart.defaults work stays as before.
  document.documentElement.setAttribute('data-theme', theme);
  window.dispatchEvent(new CustomEvent('kit:theme'));
  Chart.defaults.color = theme === 'dark' ? '#c8ccd4' : '#333333';
  Chart.defaults.borderColor = theme === 'dark' ? 'rgba(255,255,255,0.12)' : 'rgba(0,0,0,0.1)';
}
// Titles/labels/KPI captions/chart titles+legends/toggle buttons only - the
// filtered data itself and every KPI/chart VALUE are unaffected by language.
function applyLanguage(lang) {
  currentLang = lang;
  try { localStorage.setItem('reportLang', lang); } catch (e) {}
  document.querySelectorAll('[data-i18n]').forEach(function (el) { el.textContent = t(el.dataset.i18n); });
  render();
}
document.getElementById('resetBtn').onclick = function () {
  msYear.reset(); msMonth.reset(); msCust.reset(); msVend.reset(); msCountry.reset(); render();
};
render();
</script>
<script id="mtc-three" type="application/json">${threeIsland}</script>
<script id="mtc-land" type="application/json">${landIsland}</script>
<script id="mtc-globe-data" type="application/json">${globeDataIsland}</script>
<script type="module">
${globeJsSrc}
/* ===== MTC globe boot glue (Dashboard sheet) =====
   Ported from GPR main.js:144-190. The spike (2026-11-09) proved dynamic
   import() of a blob:text/javascript ESM module works inside an iframe
   document created via srcdoc from a file:// parent, in BOTH Chromium and
   WebKit - so the srcdoc architecture stays and fallback B (standalone
   blob-document iframe) is not needed. */
var GLOBE_AGG = JSON.parse(document.getElementById('mtc-globe-data').textContent);
var globeMode = 'sales';
var globeApi = null, globeTried = false, globeThreePromise = null;
// loadThree: verbatim port of GPR main.js:149-158 (island id mtc-three,
// memoized promise, revoke after 5s like the shipped GPR page).
function globeLoadThree() {
  if (!globeThreePromise) {
    globeThreePromise = (async function () {
      var src = JSON.parse(document.getElementById('mtc-three').textContent);
      var url = URL.createObjectURL(new Blob([src], { type: 'text/javascript' }));
      try { return await import(url); } finally { setTimeout(function () { URL.revokeObjectURL(url); }, 5000); }
    })();
  }
  return globeThreePromise;
}
// Year multiselect drives which island buckets are summed; [] means "All".
// Other dashboard filters (customer/vendor/month/country) deliberately do
// not move the globe - it is driven by the build-time island (10 KB budget),
// not the full live row set.
function globeYearsSelected() {
  var sel = (window.msYear && msYear.getSelected()) || [];
  return sel.length ? sel : (window.YEAR_VALUES || Object.keys(GLOBE_AGG.years));
}
// Markers are computed LIVE from lastRows - the EXACT filtered rows the KPI
// cards and action title sum (user fix 2026-11-09: the old per-year island
// honoured the Year filter but ignored Month/Country/Customer/Vendor, so the
// earth could contradict the numbers above it). Sales mode sums col5
// (Net_Sales_USD) by col4 Cust_Countries; purchase mode sums col6 cost by col8
// Country of origin. GLOBE_AGG stays in the file purely as the name->hub
// coordinate reference (+ the build-time unlocated warnings).
var globeHubMap = null;
function globeHubs() {
  if (globeHubMap) return globeHubMap;
  globeHubMap = {};
  Object.keys(GLOBE_AGG.years || {}).forEach(function (y) {
    var b = GLOBE_AGG.years[y];
    ['sales', 'purchase'].forEach(function (m) {
      (b[m] || []).forEach(function (e) { if (!(e.c in globeHubMap)) globeHubMap[e.c] = e.h; });
    });
  });
  return globeHubMap;
}
function globeData() {
  var hubs = globeHubs();
  var salesMode = globeMode === 'sales';
  var byName = {}, total = 0;
  lastRows.forEach(function (r) {
    var name = String(salesMode ? r[4] : r[8] || '').trim();
    if (!name) name = String(salesMode ? r[4] : r[8]).trim();
    if (!name || name === 'null' || name === 'undefined') return;
    var v = salesMode ? r[5] : r[6];
    if (typeof v !== 'number' || !isFinite(v)) v = 0;
    if (!byName[name]) byName[name] = { c: name, h: hubs[name] || null, usd: 0, n: 0 };
    byName[name].usd += v; byName[name].n += 1;
    total += v;
  });
  var list = Object.keys(byName).map(function (k) { return byName[k]; });
  list.sort(function (a, b) { return b.usd - a.usd; });
  return { list: list, total: total };
}
// Marker sizing 12..100 exactly like GPR globeMarkers (main.js:184-188), but
// entries whose hub is null are console-warned + skipped - NEVER re-pointed
// to a default city (GPR's main.js:185 '上海' fallback is deliberately not
// ported).
function globeMarkerList() {
  var d = globeData();
  var max = d.list.reduce(function (m, e) { return Math.max(m, e.usd); }, 0) || 1;
  var un = {};
  var out = [];
  d.list.forEach(function (e) {
    if (!e.h) { un[e.c] = 1; return; }
    out.push({ city: e.h, country: e.c, usd: e.usd, n: e.n, v: Math.max(12, e.usd / max * 100) });
  });
  var names = Object.keys(un);
  if (names.length) console.warn('GLOBE: no hub coordinate, no marker drawn (' + globeMode + '): ' + names.join(', '));
  return out;
}
function globeEsc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function globeReadout(mk) {
  var d = globeData();
  document.getElementById('globeReadoutLabel').textContent = t(globeMode === 'sales' ? 'globeUnitSales' : 'globeUnitPurchase');
  if (mk) {
    var share = d.total ? (mk.usd / d.total * 100) : 0;
    document.getElementById('globeReadoutName').textContent = mk.country;
    document.getElementById('globeReadoutValue').textContent = Math.round(mk.usd).toLocaleString() + '  (' + share.toFixed(1) + '%)';
  } else {
    document.getElementById('globeReadoutName').textContent = t('globeAllMarkets') + ' - ' + d.list.length;
    document.getElementById('globeReadoutValue').textContent = Math.round(d.total).toLocaleString();
  }
  // 1:1 row<->marker highlight: whatever drove this readout (marker hover via
  // onHover, or a rank row's mouseover) marks the matching row .sel.
  var box = document.getElementById('globeLegend');
  if (box) {
    var rws = box.querySelectorAll('.lg-row');
    for (var i = 0; i < rws.length; i++) {
      rws[i].classList.toggle('sel', !!mk && rws[i].getAttribute('data-country') === mk.country);
    }
  }
}
// GPR-style performance rank list, 1:1 with the drawn markers: same source
// (globeData()'s mode- and filter-aware buckets), same desc-by-value order,
// one row per marker (hubless values stay in totals but get a dimmed,
// non-clickable row instead of vanishing). Rows are buttons: hover previews
// in the readout + highlights the matching marker's row, click spins the
// earth to the hub and fires the SAME mtcDrilldown contract as a marker
// click (Data tab pre-filtered by Cust_Countries / Country of origin).
function globeLegend(d) {
  var html = '<div class="lg-k">' + t('globeLegendTitle') + '</div>';
  var max = d.list.length ? d.list[0].usd : 1;
  d.list.forEach(function (e, i) {
    var pct = d.total ? (e.usd / d.total * 100) : 0;
    var no = (i + 1 < 10 ? '0' : '') + (i + 1);
    var tw = max ? (e.usd / max * 100).toFixed(1) : '0';
    html = html + '<button type="button" class="lg-row"' + (e.h ? '' : ' disabled') + ' data-country="' + globeEsc(e.c) + '" data-hub="' + globeEsc(e.h || '') + '" data-usd="' + e.usd + '">'
      + '<span class="lg-no">' + no + '</span>'
      + '<span class="lg-c">' + globeEsc(e.c) + '</span>'
      + '<span class="lg-v">' + fmtCompact(e.usd) + '</span>'
      + '<span class="lg-p">' + pct.toFixed(1) + '%</span>'
      + '<span class="lg-track"><i style="width:' + tw + '%"></i></span></button>';
  });
  html = html + '<div class="lg-s">' + t('globeSizeLegend') + '</div>';
  document.getElementById('globeLegend').innerHTML = html;
}
function globePaint() {
  var d = globeData();
  if (globeApi) globeApi.setMarkers(globeMarkerList());
  globeLegend(d);
  var names = (GLOBE_AGG.unlocated && GLOBE_AGG.unlocated[globeMode]) || [];
  document.getElementById('globeUnlocatedNote').innerHTML = names.length
    ? '<span>' + t('globeUnlocatedNote') + '</span> ' + globeEsc(names.join(', ')) : '';
  globeReadout(null);
}
async function globeEnsure() {
  if (globeTried) return;
  globeTried = true;
  try {
    var probe = document.createElement('canvas');
    if (!(probe.getContext('webgl2') || probe.getContext('webgl'))) throw new Error('WebGL unavailable');
    // Mirror any theme the shell already applied onto documentElement BEFORE
    // the first token read, so a dark-mode viewer never sees a light frame.
    var bt = document.body.getAttribute('data-theme');
    if (bt && document.documentElement.getAttribute('data-theme') !== bt) {
      document.documentElement.setAttribute('data-theme', bt);
      window.dispatchEvent(new CustomEvent('kit:theme'));
    }
    var THREE = await globeLoadThree();
    if (globeApi) return;
    globeApi = makeGlobe(THREE, document.getElementById('globeEarth'), {
      routes: [],
      markers: globeMarkerList(),
      onPick: function (marker) {
        // Same proven drilldown contract as the geo chart bar (and the table
        // cells): shell switches to the Data tab and pre-filters it.
        var filters = {};
        filters[globeMode === 'sales' ? 'Cust_Countries' : 'Country of origin'] = marker.country;
        window.parent.postMessage({ type: 'mtcDrilldown', targetTab: 'Data', filters: filters }, '*');
      },
      onHover: globeReadout
    });
    window.globeSuccess = true;
    document.body.classList.add('globe-ok');
    globePaint();
    globeFocusTop(); // start the story facing the biggest market (GPR main.js:173 pattern)
  } catch (err) {
    console.warn('Globe initialization failed', err);
    document.getElementById('globeSection').classList.add('globe-fail');
    var st = document.getElementById('globeStatus');
    st.removeAttribute('data-i18n'); // else a later language sweep rewrites the failure line back to "Building..."
    st.textContent = t('globeFailed');
    // The two fallback charts are still rendered and visible (render() only
    // skips them once window.globeSuccess is true) - nothing else to undo.
  }
}
// Spin the globe to face the mode's biggest market - on init and on an
// explicit mode switch only (never on filter repaints, so a user drag is
// never fought). api.focus is the shipped GPR easing (globe.js:481-493).
function globeFocusTop() {
  if (!globeApi) return;
  var top = globeData().list.filter(function (e) { return !!e.h; })[0];
  if (top) globeApi.focus(top.h);
}
document.getElementById('globeModeSales').addEventListener('click', function () { globeMode = 'sales'; globeSetModeOn(); globePaint(); globeFocusTop(); });
document.getElementById('globeModePurchase').addEventListener('click', function () { globeMode = 'purchase'; globeSetModeOn(); globePaint(); globeFocusTop(); });
function globeSetModeOn() {
  document.getElementById('globeModeSales').className = 'globe-mode' + (globeMode === 'sales' ? ' on' : '');
  document.getElementById('globeModePurchase').className = 'globe-mode' + (globeMode === 'purchase' ? ' on' : '');
}
// DOMESTIC memo cards -> open the Domestic tab through the shell's proven
// drilldown postMessage (empty filters; the shell typeof-guards the missing
// applyExternalFilter on that sheet, so nothing else happens).
Array.prototype.forEach.call(document.querySelectorAll('.kpi-dom-card'), function (c) {
  function go() { window.parent.postMessage({ type: 'mtcDrilldown', targetTab: 'Domestic', filters: {} }, '*'); }
  c.addEventListener('click', go);
  c.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); go(); } });
});

// Rank-list interactions (event delegation survives globePaint re-renders).
(function () {
  var box = document.getElementById('globeLegend');
  box.addEventListener('mouseover', function (e) {
    var row = e.target.closest ? e.target.closest('.lg-row') : null;
    if (row && !row.disabled) {
      globeReadout({ country: row.getAttribute('data-country'), usd: parseFloat(row.getAttribute('data-usd')) || 0 });
    }
  });
  box.addEventListener('mouseleave', function () { globeReadout(null); });
  box.addEventListener('click', function (e) {
    var row = e.target.closest ? e.target.closest('.lg-row') : null;
    if (!row || row.disabled) return;
    var country = row.getAttribute('data-country');
    var hub = row.getAttribute('data-hub');
    if (hub && globeApi) globeApi.focus(hub);
    var filters = {};
    filters[globeMode === 'sales' ? 'Cust_Countries' : 'Country of origin'] = country;
    window.parent.postMessage({ type: 'mtcDrilldown', targetTab: 'Data', filters: filters }, '*');
  });
})();
window.globeUpdateHook = globePaint;
// Lazy first build: only when the host actually scrolls into view (panels
// paint first, three.js compile lands ~1-2 s later - same UX as GPR).
(function () {
  var host = document.getElementById('globeEarth');
  if ('IntersectionObserver' in window) {
    var io = new IntersectionObserver(function (ents) {
      if (ents[0].isIntersecting) { io.disconnect(); globeEnsure(); }
    }, { rootMargin: '200px 0px' });
    io.observe(host);
  } else {
    globeEnsure();
  }
})();

/* ---- AI assistant panel (Qwen; hosted-first, GPR-style) ----
   Default: the hosted meyer-ai-assistant Cloudflare Worker (/mtc/chat,
   /mtc/health) - no access code needed since 2026-11-12 (Anthony's decision,
   auditor/team convenience); the per-app 5M monthly token quota is the guard.
   Emailed copies work for colleagues without any local install. If the hosted
   service is unreachable, the panel automatically falls back to the legacy
   loopback relay at 127.0.0.1:8765 when it happens to run on this PC. This
   file carries NO credential; the worker and relay hold the bearer
   server-side, enforce the monthly cap, and fail closed. If neither endpoint
   answers, the panel shows a friendly offline line and the dashboard is
   untouched. Model text is always rendered with textContent; chat history
   lives in memory only (page session); the panel nodes are tagged
   data-ai-exclude for save-copy hygiene. */
var AI_RELAY = 'http://127.0.0.1:8765';
var AI_HOSTED = 'https://meyer-ai-assistant.anthonym-354.workers.dev/mtc';
var DOM_MEMO = ${domMemo ? JSON.stringify(domMemo) : 'null'};
var aiState = { history: [], busy: false, online: null, mode: 'hosted', lastError: '' };
var aiUser = 'MTC-report';
try { aiUser = localStorage.getItem('mtc-ai-user') || 'MTC-report'; } catch (e) { /* srcdoc storage may be unavailable */ }

function aiUsd(v) { return Math.round(v).toLocaleString('en-US'); }
function aiProbe() {
  aiState.mode = 'hosted';
  aiState.lastError = '';
  function hostedOnce() {
    return fetch(AI_HOSTED + '/health', { signal: (function () { var c = new AbortController(); setTimeout(function () { c.abort(); }, 8000); return c.signal; })() })
      .then(function (r) { return r.json().then(function (b) { return { ok: r.ok, b: b }; }, function () { return { ok: false, b: null }; }); })
      .then(function (res) {
        if (res.ok && res.b && res.b.ok === true) { aiState.online = true; aiState.lastError = ''; return true; }
        aiState.lastError = (res.b && res.b.error) || 'Cloudflare AI service unavailable. Check the internet connection.';
        return false;
      })
      .catch(function () { aiState.lastError = 'Cannot reach Cloudflare right now - your network path timed out. Close and reopen the assistant to try again.'; return false; });
  }
  return hostedOnce().then(function (ok) {
    if (ok) return true;
    // One automatic retry: brief network-path glitches (DNS/QUIC/route hiccups)
    // otherwise surface as a misleading "check your connection" failure.
    return hostedOnce().then(function (ok2) {
      if (ok2) return true;
      // Last resort: the legacy loopback relay, if it runs on this PC.
      aiState.mode = 'relay';
      return fetch(AI_RELAY + '/health', { signal: (function () { var c = new AbortController(); setTimeout(function () { c.abort(); }, 5000); return c.signal; })() })
        .then(function (r) { aiState.online = r.ok; if (r.ok) aiState.lastError = ''; return r.ok; })
        .catch(function () { aiState.online = false; aiState.mode = 'hosted'; return false; });
    });
  });
}
function aiContext() {
  var rows = lastRows;
  var sales = 0, cost = 0, margin = 0;
  rows.forEach(function (r) { sales += r[5]; cost += r[6]; margin += r[7]; });
  var months = Array.from(new Set(rows.map(function (r) { return r[1]; }))).filter(Boolean).sort();
  var prior = priorYearSums(months);
  var pS = 0, pC = 0, pM = 0;
  Object.keys(prior.sales).forEach(function (k) { pS += prior.sales[k] || 0; pC += prior.cost[k] || 0; pM += prior.margin[k] || 0; });
  var byM = groupSum(rows, 1, 5);
  var monthly = months.slice(0, 24).map(function (m) { return m + ':' + aiUsd(byM[m] || 0); }).join(', ');
  var gd = null;
  try { if (window.globeSuccess) gd = globeData(); } catch (e) { gd = null; }
  var L = [];
  L.push('CURRENT APP: MTC Performance Report - single-file interactive report (NOT MPL_System, NOT the Global Purchase dashboard). Answer only about this MTC report.');
  L.push('UI_LANGUAGE: ' + (currentLang === 'zh' ? 'zh-Hant (reply in Traditional Chinese)' : 'en (reply in English)'));
  L.push('DATA PERIOD: ' + (MONTH_VALUES[0] || '?') + ' to ' + (MONTH_VALUES[MONTH_VALUES.length - 1] || '?') + '; ledger rows ' + ROWS.length.toLocaleString('en-US') + '; rows under current filters ' + rows.length.toLocaleString('en-US'));
  L.push('ACTIVE FILTERS: ' + JSON.stringify({ Year: msYear.getSelected(), Month: msMonth.getSelected(), Customer: msCust.getSelected(), Vendor: msVend.getSelected(), Country: msCountry.getSelected() }));
  L.push('KPI UNDER FILTERS: netSalesUsd=' + aiUsd(sales) + ', netCostUsd=' + aiUsd(cost) + ', netMarginUsd=' + aiUsd(margin) + ', marginPct=' + (sales ? (margin / sales * 100).toFixed(1) : '0') + '%; sameMonthsLastYear: sales=' + aiUsd(pS) + ', cost=' + aiUsd(pC) + ', margin=' + aiUsd(pM) + ((pS && sales) ? ', salesYoYpct=' + ((sales - pS) / pS * 100).toFixed(1) : ''));
  if (monthly) L.push('MONTHLY NET SALES (USD): ' + monthly);
  if (gd) L.push('GLOBE MODE: ' + globeMode + ' - TOP BY USD: ' + gd.list.slice(0, 8).map(function (e) { return e.c + ':' + aiUsd(e.usd) + ' (' + (gd.total ? (e.usd / gd.total * 100).toFixed(1) : '0') + '%)'; }).join(', ') + ' - earth, rank list and KPIs all read ONE filtered dataset');
  if (DOM_MEMO) L.push('DOMESTIC MEMO (separate snapshot workbook; derived USD, NOT included in the MA/DI totals above): sales2026Usd=' + aiUsd(DOM_MEMO.sales) + ', commission2026Usd=' + aiUsd(DOM_MEMO.comm) + ', asOf=' + DOM_MEMO.asOf + '; entities MJP/JPY, MIN/INR, MCN+MZQ/RMB; no 2025 domestic data yet (pending confirmation)');
  L.push('ALLOWED FILTER VALUES: Year=[' + YEAR_VALUES.join(', ') + '] Month=[' + MONTH_VALUES.join(', ') + '] Country=[' + COUNTRY_VALUES.join(', ') + '] Customer=[' + CUST_VALUES.slice(0, 40).join(', ') + '] Vendor=[' + VEND_VALUES.slice(0, 40).join(', ') + ']');
  L.push('ACTION PROPOSALS: when a filter or navigation would help, add up to 3 lines AT THE END, each exactly one of: FILTER Year=<value> / FILTER Month=<value> / FILTER Country=<value> / FILTER Customer=<value> / FILTER Vendor=<value> / Tab=<Data|Domestic|Exchange Rate> / Mode=<GlobeSales|GlobePurchase>. Values must be copied EXACTLY from ALLOWED FILTER VALUES - never invent. Lines render as buttons; nothing changes until the user clicks.');
  return L.join('\\n').slice(0, 5800);
}
function aiParseActions(reply) {
  var out = [], keep = [];
  String(reply).split(/\\r?\\n/).forEach(function (line) {
    var l = line.trim();
    var m = l.match(/^FILTER\\s+(Year|Month|Country|Customer|Vendor)=(.+)$/);
    var mt = l.match(/^Tab=(Data|Domestic|Exchange Rate)$/);
    var md = l.match(/^Mode=(GlobeSales|GlobePurchase)$/);
    var okv = false;
    if (m) {
      var k = m[1], v = m[2].trim();
      okv = (k === 'Year') ? YEAR_VALUES.indexOf(v) >= 0
        : (k === 'Month') ? MONTH_VALUES.indexOf(v) >= 0
        : (k === 'Country') ? COUNTRY_VALUES.indexOf(v) >= 0
        : (k === 'Customer') ? CUST_VALUES.indexOf(v) >= 0
        : VEND_VALUES.indexOf(v) >= 0;
      if (okv && out.length < 3) out.push({ kind: 'filter', key: k, value: v }); else keep.push(line);
    } else if (mt && out.length < 3) { out.push({ kind: 'tab', value: mt[1] }); }
    else if (md && out.length < 3) { out.push({ kind: 'mode', value: md[1] }); }
    else keep.push(line);
  });
  return { text: keep.join('\\n').trim(), actions: out };
}
function aiApplyAction(a) {
  if (a.kind === 'filter') {
    if (a.key === 'Year') msYear.setSelected([a.value]);
    else if (a.key === 'Month') msMonth.setSelected([a.value]);
    else if (a.key === 'Country') msCountry.setSelected([a.value]);
    else if (a.key === 'Customer') msCust.setSelected([a.value]);
    else msVend.setSelected([a.value]);
    render();
  } else if (a.kind === 'tab') {
    window.parent.postMessage({ type: 'mtcDrilldown', targetTab: a.value, filters: {} }, '*');
  } else {
    globeMode = (a.value === 'GlobeSales') ? 'sales' : 'purchase';
    globeSetModeOn(); globePaint(); globeFocusTop();
  }
}
function aiBubble(role, text, actions) {
  var wrap = document.createElement('div');
  wrap.className = 'ai-msg ' + role;
  wrap.setAttribute('data-ai-exclude', '1');
  var body = document.createElement('div');
  body.textContent = text;
  wrap.appendChild(body);
  (actions || []).forEach(function (a) {
    var chip = document.createElement('button');
    chip.type = 'button'; chip.className = 'ai-chip';
    var lbl = a.kind === 'filter' ? (a.key + ': ' + a.value) : (a.kind === 'tab' ? ('Go: ' + a.value) : ('Earth: ' + a.value));
    chip.textContent = t('aiApply') + ' ' + lbl;
    chip.addEventListener('click', function () {
      try { aiApplyAction(a); chip.disabled = true; } catch (e) { chip.textContent = t('aiErrNet'); }
    });
    wrap.appendChild(chip);
  });
  return wrap;
}
function aiStatusHtml() {
  var s = document.getElementById('ai-status');
  if (!s) return;
  if (aiState.online === false) {
    s.innerHTML = '<i class="ai-dot off"></i><span></span>';
    s.lastElementChild.textContent = (aiState.mode === 'hosted' && aiState.lastError) ? aiState.lastError : t('aiOffline');
    var i0 = document.getElementById('ai-input'); if (i0) i0.disabled = true;
    var b0 = document.getElementById('ai-send'); if (b0) b0.disabled = true;
  } else {
    s.innerHTML = '<i class="ai-dot on"></i><span></span><span class="ai-quota"></span><button type="button" class="ai-clear"></button>';
    s.children[1].textContent = t('aiOnline') + (aiState.mode === 'hosted' ? ' - Cloudflare' : ' - relay');
    s.querySelector('.ai-clear').textContent = t('aiClear');
    s.querySelector('.ai-clear').addEventListener('click', function () {
      aiState.history = [];
      var logEl = document.getElementById('ai-log');
      logEl.innerHTML = '';
      logEl.appendChild(aiBubble('sys', t('aiHello'), null));
    });
    var i1 = document.getElementById('ai-input'); if (i1) i1.disabled = aiState.busy;
    var b1 = document.getElementById('ai-send'); if (b1) b1.disabled = aiState.busy;
  }
}
function aiSend(message) {
  var logEl = document.getElementById('ai-log');
  logEl.appendChild(aiBubble('user', message, null));
  var thinking = aiBubble('sys', t('aiThinking'), null);
  logEl.appendChild(thinking);
  logEl.scrollTop = logEl.scrollHeight;
  aiState.busy = true;
  var inp = document.getElementById('ai-input'), snd = document.getElementById('ai-send');
  if (inp) inp.disabled = true; if (snd) snd.disabled = true;
  var ctl = new AbortController();
  setTimeout(function () { ctl.abort(); }, 90000);
  var aiBase = aiState.mode === 'hosted' ? AI_HOSTED : AI_RELAY;
  var aiHeaders = { 'content-type': 'application/json' };
  fetch(aiBase + '/chat', {
    method: 'POST', headers: aiHeaders, signal: ctl.signal,
    body: JSON.stringify({ username: aiUser, message: message, history: aiState.history.slice(-12), context: aiContext() })
  }).then(function (r) {
    return r.json().catch(function () { return null; }).then(function (j) { return { ok: r.ok, j: j }; });
  }).then(function (res) {
    thinking.remove();
    if (!res.ok || !res.j || res.j.ok !== true) {
      logEl.appendChild(aiBubble('sys', '⚠ ' + ((res.j && res.j.error) || t('aiErrNet')), null));
    } else {
      var parsed = aiParseActions(res.j.reply);
      var tail = parsed.text + (res.j.truncated ? ('\\n' + t('aiTrunc')) : '');
      logEl.appendChild(aiBubble('bot', tail || '(empty)', parsed.actions));
      aiState.history.push({ role: 'user', content: message.slice(0, 2000) });
      aiState.history.push({ role: 'assistant', content: parsed.text.slice(0, 2000) });
      if (aiState.history.length > 24) aiState.history.splice(0, aiState.history.length - 24);
      var q = document.querySelector('.ai-quota');
      if (q && res.j.quota) q.textContent = t('aiQuota').replace('{U}', Number(res.j.quota.used).toLocaleString('en-US')).replace('{C}', Number(res.j.quota.cap).toLocaleString('en-US'));
      if (res.j.warn) logEl.appendChild(aiBubble('sys', '⚠ ' + res.j.warn, null));
    }
  }).catch(function () {
    thinking.remove();
    logEl.appendChild(aiBubble('sys', '⚠ ' + (aiState.mode === 'hosted' ? 'Cannot reach Cloudflare right now - your message was NOT sent. Check your connection and send again.' : t('aiErrNet')), null));
    aiProbe().then(function (ok) { if (!ok) aiStatusHtml(); });
  }).then(function () {
    aiState.busy = false;
    if (inp) inp.disabled = false; if (snd) snd.disabled = false;
    logEl.scrollTop = logEl.scrollHeight;
    if (inp) inp.focus();
  });
}
function aiClose() {
  var root = document.getElementById('ai-root');
  if (root) root.innerHTML = '';
  if (aiState.winResizeClamp) { window.removeEventListener('resize', aiState.winResizeClamp); aiState.winResizeClamp = null; }
  // NOTE: #ai-css stays in <head> on purpose - it also styles the #aiBtn FAB,
  // which the old drawer code removed on close (button lost its styling after
  // the first open/close cycle). Fixed 2026-11-10 alongside the floating-window change.
  var btn = document.getElementById('aiBtn'); if (btn) btn.focus();
}
/* Floating draggable chat window, per the shared qwen-chat-assistant UI rule:
   drag by the header (mouse + touch pointer events), always clamped inside the
   viewport, re-clamped on window resize, header close button stays clickable. */
function aiMakeDraggable(win) {
  var head = win.querySelector('.ai-head');
  var r0 = win.getBoundingClientRect();
  win.style.right = 'auto'; win.style.bottom = 'auto';
  win.style.left = r0.left + 'px'; win.style.top = r0.top + 'px';
  function clamp() {
    if (!win.isConnected) return;
    var b = win.getBoundingClientRect();
    var x = parseFloat(win.style.left), y = parseFloat(win.style.top);
    if (isNaN(x)) x = b.left; if (isNaN(y)) y = b.top;
    x = Math.min(Math.max(6, x), Math.max(6, window.innerWidth - b.width - 6));
    y = Math.min(Math.max(6, y), Math.max(6, window.innerHeight - b.height - 6));
    win.style.left = x + 'px'; win.style.top = y + 'px';
  }
  clamp();
  head.addEventListener('pointerdown', function (e) {
    if (e.button) return;
    if (e.target.closest && e.target.closest('button')) return; // controls stay clickable
    var b = win.getBoundingClientRect();
    var dx = e.clientX - b.left, dy = e.clientY - b.top;
    try { head.setPointerCapture(e.pointerId); } catch (err) {}
    function mv(ev) {
      win.style.left = (ev.clientX - dx) + 'px';
      win.style.top = (ev.clientY - dy) + 'px';
      clamp();
    }
    function up(ev) {
      try { head.releasePointerCapture(ev.pointerId); } catch (err) {}
      head.removeEventListener('pointermove', mv);
      head.removeEventListener('pointerup', up);
      head.removeEventListener('pointercancel', up);
    }
    head.addEventListener('pointermove', mv);
    head.addEventListener('pointerup', up);
    head.addEventListener('pointercancel', up);
    e.preventDefault();
  });
  window.addEventListener('resize', clamp);
  aiState.winResizeClamp = clamp;
}
function aiOpen() {
  if (aiState.winResizeClamp) { window.removeEventListener('resize', aiState.winResizeClamp); aiState.winResizeClamp = null; }
  var root = document.getElementById('ai-root');
  root.innerHTML = '<aside class="ai-win" role="dialog" aria-labelledby="ai-title">'
    + '<header class="ai-head" title="Drag to move / 拖曳移動"><span class="ai-grip">⠿</span><h3 id="ai-title">' + t('aiTitle') + '</h3><button type="button" class="ai-x" id="ai-close" aria-label="close">✕</button></header>'
    + '<div class="ai-status" id="ai-status"></div>'
    + '<div class="ai-log" id="ai-log"></div>'
    + '<p class="ai-disclose">' + t('aiDisclose') + '</p>'
    + '<form class="ai-compose" id="ai-form"><input id="ai-input" type="text" maxlength="4000"/><button class="ai-send" id="ai-send" type="submit"></button></form>'
    + '<footer class="ai-userbar"><span id="ai-namelbl"></span><input id="ai-user" maxlength="40"/></footer>'
    + '</aside>';
  root.querySelector('#ai-input').placeholder = t('aiPh');
  root.querySelector('#ai-send').textContent = t('aiSend');
  root.querySelector('#ai-namelbl').textContent = t('aiName');
  root.querySelector('#ai-user').value = aiUser;
  aiStatusHtml();
  document.getElementById('ai-log').appendChild(aiBubble('sys', t('aiHello'), null));
  document.getElementById('ai-close').addEventListener('click', aiClose);
  aiMakeDraggable(root.querySelector('.ai-win'));
  document.getElementById('ai-form').addEventListener('submit', function (e) {
    e.preventDefault();
    var v = document.getElementById('ai-input').value.trim();
    if (!v || aiState.busy) return;
    document.getElementById('ai-input').value = '';
    aiSend(v);
  });
  document.getElementById('ai-user').addEventListener('change', function () {
    aiUser = this.value.trim() || 'MTC-report';
    try { localStorage.setItem('mtc-ai-user', aiUser); } catch (e) { /* fine */ }
  });
  aiProbe().then(function (ok) { aiStatusHtml(); });
  var inp = document.getElementById('ai-input');
  if (aiState.online !== false && inp) inp.focus();
}
(function aiInit() {
  var st = document.createElement('style');
  st.id = 'ai-css';
  st.textContent = '#aiBtn{position:fixed;right:18px;bottom:18px;z-index:60;background:var(--accent);color:#fff;border:0;border-radius:20px;padding:9px 16px;font-size:13px;cursor:pointer;box-shadow:0 2px 10px rgba(0,0,0,.25);}#aiBtn:hover{filter:brightness(1.12);}'
    + '.ai-win{position:fixed;right:16px;bottom:64px;width:390px;max-width:92vw;height:min(560px,84vh);background:var(--bg);color:var(--text);display:flex;flex-direction:column;border:1px solid var(--border);border-radius:10px;box-shadow:0 12px 36px rgba(5,28,44,.35);z-index:70;overflow:hidden;}'
    + '.ai-head{display:flex;align-items:center;gap:8px;padding:12px 14px;border-bottom:1px solid var(--border-light);cursor:move;user-select:none;-webkit-user-select:none;touch-action:none;background:var(--card-bg);flex:none;}'
    + '.ai-grip{color:var(--text-muted);font-size:14px;flex:none;}'
    + '.ai-head h3{flex:1;margin:0;font-size:15px;font-family:Georgia,serif;}'
    + '.ai-x{background:none;border:0;font-size:17px;cursor:pointer;color:var(--text-muted);}'
    + '.ai-status{padding:6px 14px;font-size:11px;color:var(--text-muted);display:flex;gap:8px;align-items:center;border-bottom:1px solid var(--border-light);}'
    + '.ai-access{display:flex;gap:8px;padding:8px 14px;border-bottom:1px solid var(--border-light);}'
    + '.ai-access input{flex:1;padding:6px 8px;border:1px solid var(--border);border-radius:4px;background:var(--input-bg);color:var(--text);font-size:11.5px;}'
    + '.ai-access button{background:var(--accent);color:#fff;border:0;border-radius:4px;padding:6px 12px;font-size:11.5px;cursor:pointer;flex:none;}'
    + '.ai-dot{width:8px;height:8px;border-radius:50%;display:inline-block;flex:none;}'
    + '.ai-dot.on{background:#1c7a63;}.ai-dot.off{background:#A63D40;}'
    + '.ai-quota{margin-left:auto;}'
    + '.ai-clear{background:none;border:1px solid var(--border-light);border-radius:3px;color:var(--text-muted);font-size:10.5px;padding:3px 8px;cursor:pointer;}'
    + '.ai-log{flex:1;overflow-y:auto;padding:10px 14px;display:flex;flex-direction:column;gap:8px;}'
    + '.ai-msg{max-width:92%;padding:8px 10px;border-radius:8px;font-size:12.5px;line-height:1.45;white-space:pre-wrap;}'
    + '.ai-msg.user{align-self:flex-end;background:var(--accent);color:#fff;}'
    + '.ai-msg.bot{align-self:flex-start;background:var(--card-bg);border:1px solid var(--border-light);}'
    + '.ai-msg.sys{align-self:center;color:var(--text-muted);font-size:11px;}'
    + '.ai-chip{display:inline-block;margin:6px 6px 0 0;background:var(--card-bg);border:1px solid var(--accent);color:var(--accent);border-radius:14px;padding:4px 10px;font-size:11.5px;cursor:pointer;}'
    + '.ai-chip:disabled{opacity:.5;cursor:default;border-color:var(--border);}'
    + '.ai-disclose{margin:0;padding:8px 14px;font-size:10.5px;color:var(--text-muted);border-top:1px solid var(--border-light);}'
    + '.ai-compose{display:flex;gap:8px;padding:10px 14px;border-top:1px solid var(--border-light);}'
    + '.ai-compose input{flex:1;padding:8px 10px;border:1px solid var(--border);border-radius:4px;background:var(--input-bg);color:var(--text);font-size:12.5px;}'
    + '.ai-send{background:var(--accent);color:#fff;border:0;border-radius:4px;padding:8px 14px;cursor:pointer;font-size:12.5px;}'
    + '.ai-userbar{display:flex;gap:8px;align-items:center;padding:6px 14px;font-size:11px;color:var(--text-muted);border-top:1px solid var(--border-light);}'
    + '.ai-userbar input{flex:1;padding:4px 6px;border:1px solid var(--border-light);border-radius:3px;background:var(--input-bg);color:var(--text);font-size:11px;}';
  document.head.appendChild(st);
  var btn = document.createElement('button');
  btn.id = 'aiBtn'; btn.type = 'button';
  btn.setAttribute('data-ai-exclude', '1');
  var sp = document.createElement('span');
  sp.setAttribute('data-i18n', 'aiBtn');
  sp.textContent = 'Ask AI';
  btn.textContent = '✦ ';
  btn.appendChild(sp);
  document.body.appendChild(btn);
  var root = document.createElement('div');
  root.id = 'ai-root';
  root.setAttribute('data-ai-exclude', '1');
  document.body.appendChild(root);
  btn.addEventListener('click', function () { if (root.firstChild) aiClose(); else aiOpen(); });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && root.firstChild) aiClose();
  });
})();
</script>
</body>
</html>`;
console.log('Dashboard built (' + (dashboardHtml.length / 1024 / 1024).toFixed(2) + ' MB, incl. globe islands)');

// =========================================================================
// Data sheet (full detail table)
// =========================================================================
const dataSheetHtml = buildDataTablesHtml(H, dataSheet.rows, 'Data', true,
  // The last 4 columns here exist so a YoY-tab drilldown click always has a
  // matching toolbar filter/dropdown to land on and visibly show as applied.
  ['Cust_name', 'VEND_NAME', 'Cust_Countries', 'yyyy-mm', 'Year', 'Country of origin', 'Fabrication Code', 'Family1', 'Part_id'],
  null, 'title_data', false, 'introData', null, null, null, 'actionData',
  'data_sheet.json - full DATA ledger island · ' + reportPeriodRange + ' · ' + FX_SOURCE);
console.log('Data sheet built (' + (dataSheetHtml.length / 1024 / 1024).toFixed(2) + ' MB)');

// =========================================================================
// Exchange Rate (USD table only)
// =========================================================================
const exRateData = readJson('exrate_usd.json');
const exRateHtml = buildDataTablesHtml(exRateData.header, exRateData.rows, 'Exchange Rate', false, null, null, 'title_exrate', false, 'introExRate', null, null, null, 'actionExrate',
  'exrate_usd.json - "Exchange Rate" sheet of the report workbook (Platinum Buy Rate -> USD) · ' + reportPeriodRange);

// =========================================================================
// YoY comparison sheets - already Excel-formatted display text, plain
// DataTables pages (narrow enough sheets that scrollX is unneeded except
// Data itself).
// =========================================================================
// Drops named columns from an already-extracted {header, rows} sheet.
function dropColumns(sheet, colsToRemove) {
  const dropIdx = colsToRemove.map(n => sheet.header.indexOf(n)).filter(i => i >= 0);
  const keepIdx = sheet.header.map((h, i) => i).filter(i => dropIdx.indexOf(i) === -1);
  return {
    header: keepIdx.map(i => sheet.header[i]),
    rows: sheet.rows.map(r => keepIdx.map(i => r[i]))
  };
}
// "Country of origin"/"Part_id" dropped per explicit request - Cust_name is
// the only label column this table keeps.
const yoyCustomers = dropColumns(readJson('yoy_customers.json'), ['Country of origin', 'Part_id']);
const yoyFabrication = readJson('yoy_fabrication.json');
// Prepends the broader brand grouping (e.g. "PRESTIGE") ahead of the more
// granular Family1 ("PRESTIGE SMART PLUS PRESSURE COOKER") - per Peggy's
// request, sourced from a separate reference workbook (not this month's
// report file) since Family Category isn't tracked per-Family1 anywhere in
// this sheet's own pivot output. Labeled "Brand" on-page (Peggy's own word
// for it) even though the underlying field name is "Family Category" - the
// drilldown map below points this "Brand" column at the Data tab's real
// "Family Category" column by name. Also satisfies her separate "sort by
// Brand" ask for free - it's just another sortable/filterable column.
function addFamilyCategoryColumn(sheet, categoryMap) {
  const famIdx = sheet.header.indexOf('Family1');
  return {
    header: ['Brand', ...sheet.header],
    rows: sheet.rows.map(r => [categoryMap[r[famIdx]] || '', ...r])
  };
}
// Parses Excel's own formatted display text ("$1,234", "22.4%", "(1,234)")
// back into a real number, purely for summing into a Brand-level total row -
// same technique as the browser-side parseFormattedNumber, just needed here
// too since this aggregation happens at build time (Node), not in the page.
function parseFormattedNumberServer(v) {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined || v === '') return 0;
  let s = String(v).trim();
  const neg = /^\(.*\)$/.test(s);
  s = s.replace(/[()$,%]/g, '');
  const n = parseFloat(s);
  if (isNaN(n)) return 0;
  return neg ? -n : n;
}
function fmtIntServer(n) { return Math.round(n).toLocaleString(); }
function fmtUsdServer(n) { const sign = n < 0 ? '-' : ''; return sign + '$' + Math.abs(Math.round(n)).toLocaleString(); }
function fmtPctServer(n) { return isFinite(n) ? n.toFixed(1) + '%' : ''; }
// Inserts a "+/-" toggle column (index 1, between Brand and Family1) and one
// synthetic Brand-total row ahead of each brand's own Family1 rows - lets the
// page collapse to just Brand-level totals by default and expand a single
// brand to see its Family1 breakdown, per Peggy's request to understand
// "Brand total performance" at a glance. A brand-total row is identified
// downstream purely by its Family1 cell being blank (never true of a real
// Family1 row) - no separate flag/column needed for that. The 4 USD/qty
// metric pairs are summed directly; the 2 "Net Margin %" columns are ratios
// (summing a percentage across rows is meaningless) so they're recomputed
// from the already-summed margin/sales USD instead.
function buildBrandGroupedRows(sheet) {
  const H = sheet.header;
  const brandIdx = 0, fam1Idx = 1;
  const metricStart = 2; // first metric column, before the toggle column is inserted
  const pairCount = (H.length - metricStart) / 2; // 5 pairs: SKU_qty, Net_Sales, TTL PO, PO Net Margin, Net Margin %
  const isPctPair = pairCount - 1; // last pair (Net Margin %) is a ratio, not summable
  const salesPairIdx = H.findIndex(h => /net_sales_usd/i.test(h));
  const marginPairIdx = H.findIndex(h => /net margin usd/i.test(h));

  const brandOrder = [];
  const brandRows = new Map();
  sheet.rows.forEach(r => {
    const brand = r[brandIdx] || '';
    if (!brandRows.has(brand)) { brandRows.set(brand, []); brandOrder.push(brand); }
    brandRows.get(brand).push(r);
  });

  const newHeader = [H[brandIdx], '', H[fam1Idx], ...H.slice(metricStart)];
  const newRows = [];
  brandOrder.forEach(brand => {
    const rows = brandRows.get(brand);
    const sums = new Array(pairCount * 2).fill(0);
    rows.forEach(r => {
      for (let p = 0; p < pairCount; p++) {
        if (p === isPctPair) continue; // recomputed below, not summed
        sums[p * 2] += parseFormattedNumberServer(r[metricStart + p * 2]);
        sums[p * 2 + 1] += parseFormattedNumberServer(r[metricStart + p * 2 + 1]);
      }
    });
    const salesP = (salesPairIdx - metricStart) / 2, marginP = (marginPairIdx - metricStart) / 2;
    const totalRowMetrics = [];
    for (let p = 0; p < pairCount; p++) {
      if (p === isPctPair) {
        const pct2025 = sums[salesP * 2] ? (sums[marginP * 2] / sums[salesP * 2]) * 100 : NaN;
        const pct2026 = sums[salesP * 2 + 1] ? (sums[marginP * 2 + 1] / sums[salesP * 2 + 1]) * 100 : NaN;
        totalRowMetrics.push(fmtPctServer(pct2025), fmtPctServer(pct2026));
      } else if (p === salesP || p === marginP || /usd/i.test(H[metricStart + p * 2])) {
        totalRowMetrics.push(fmtUsdServer(sums[p * 2]), fmtUsdServer(sums[p * 2 + 1]));
      } else {
        totalRowMetrics.push(fmtIntServer(sums[p * 2]), fmtIntServer(sums[p * 2 + 1]));
      }
    }
    newRows.push([brand, '+', '', ...totalRowMetrics]);
    rows.forEach(r => newRows.push([brand, '', r[fam1Idx], ...r.slice(metricStart)]));
  });
  return { header: newHeader, rows: newRows };
}
const familyCategoryMap = readJson('family_category_map.json');
const yoyProdFamily = buildBrandGroupedRows(addFamilyCategoryColumn(readJson('yoy_prodfamily.json'), familyCategoryMap));
const yoyPartId = readJson('yoy_partid.json');
const yoyNetNetMargin = readJson('yoy_netnetmargin.json');
const netMarginByCustomer = readJson('net_margin_by_customer.json');

// drilldown.dataColMap maps THIS sheet's label column name -> the matching
// Data-tab column name to filter by; labelColCount marks where the year-paired
// metric columns start (see Export-YoYPivotSheet in the .ps1 for the same counts).
const yoyCustomersHtml = buildDataTablesHtml(yoyCustomers.header, yoyCustomers.rows, '2026 v 2025 by Customers', false, ['Cust_name'],
  { labelColCount: 1, dataColMap: { 'Cust_name': 'Cust_name' } }, 'title_yoyCustomers', true, 'introYoyCustomers', true, null, null, 'actionYoyCustomers',
  'yoy_customers.json - Excel pivot extract (2026 vs 2025 by customer) · ' + reportPeriodRange + ' · ' + FX_SOURCE);
const yoyFabricationHtml = buildDataTablesHtml(yoyFabrication.header, yoyFabrication.rows, '2026 v 2025 by Fabrication', false, ['Fabrication Code'],
  { labelColCount: 4, dataColMap: { 'Fabrication Code': 'Fabrication Code', 'Part_id': 'Part_id' } }, 'title_yoyFabrication', true, 'introYoyFabrication', true, null, null, 'actionYoyFabrication',
  'yoy_fabrication.json - Excel pivot extract (2026 vs 2025 by Fabrication Code) · ' + reportPeriodRange + ' · ' + FX_SOURCE);
const yoyProdFamilyHtml = buildDataTablesHtml(yoyProdFamily.header, yoyProdFamily.rows, '2026 v 2025 by Prod Family', false, ['Brand', 'Family1'],
  { labelColCount: 3, dataColMap: { 'Brand': 'Family Category', 'Family1': 'Family1' } }, 'title_yoyProdFamily', true, 'introYoyProdFamily', true, null,
  { groupColIdx: 0, toggleColIdx: 1, detailKeyColIdx: 2 }, 'actionYoyProdFamily',
  'yoy_prodfamily.json + family_category_map.json (Brand roll-up) · 2026 vs 2025 by Product Family · ' + reportPeriodRange + ' · ' + FX_SOURCE);
const yoyPartIdHtml = buildDataTablesHtml(yoyPartId.header, yoyPartId.rows, '2026 v 2025 by Part ID', false, ['Part_id'],
  { labelColCount: 2, dataColMap: { 'Part_id': 'Part_id' } }, 'title_yoyPartId', true, 'introYoyPartId', true, null, null, 'actionYoyPartId',
  'yoy_partid.json - Excel pivot extract (2026 vs 2025 by Part ID) · ' + reportPeriodRange + ' · ' + FX_SOURCE);
const yoyNetNetMarginHtml = buildDataTablesHtml(yoyNetNetMargin.header, yoyNetNetMargin.rows, '2026 v 2025 by NET NET MARGIN', false, ['Cust_name', 'Country of origin'],
  { labelColCount: 3, dataColMap: { 'Cust_name': 'Cust_name', 'Country of origin': 'Country of origin', 'Part_id': 'Part_id' } }, 'title_yoyNetNetMargin', true, 'introYoyNetNetMargin', true, null, null, 'actionYoyNetNetMargin',
  'yoy_netnetmargin.json - Excel pivot extract, Net Net basis (2026 vs 2025) · ' + reportPeriodRange + ' · ' + FX_SOURCE);
const netMarginByCustomerHtml = buildDataTablesHtml(netMarginByCustomer.header, netMarginByCustomer.rows, 'Net Margin by Customer-Quarter', false, ['Cust_name'],
  { labelColCount: 1, dataColMap: { 'Cust_name': 'Cust_name' }, fixedYear: netMarginByCustomer.year }, 'title_customerQuarter', false, 'introCustomerQuarter',
  false, 'Net_Sales_USD (Data_AX)', null, 'actionCustomerQuarter',
  'net_margin_by_customer.json - 2026 net margin by customer-quarter · ' + FX_SOURCE);

// =========================================================================
// Domestic tab - the 11th page, hand-authored as its own template literal.
// It is deliberately NOT built by buildDataTablesHtml, for two reasons:
//  (1) that factory's Total Index / kpiPie path pairs metric columns by the
//      header pattern "<Metric> 2025" + "<Metric> 2026" (see renderTotalIndex
//      around :717-726). The Domestic Orders workbook omits 2025 entirely
//      ("pending confirmation"), so the factory would silently render an empty
//      KPI block rather than error, and the tab must not imply a YoY read.
//  (2) its drill-link contract posts {targetTab:'Data'} and the shell resolves
//      that against the DATA ledger (shell :2039), which holds no domestic
//      rows at all - so this page is self-contained and offers no drilldown.
// What it does reuse verbatim: SHARED_CSS tokens for light/dark, I18N_JS for
// language (the English half is spliced from DOM_EN near :49), and the
// window-level applyTheme(theme) + applyLanguage(lang) pair that the shell's
// applyThemeToFrame/applyLanguageToFrame require - without both, the toggle
// silently does nothing on this tab.
// Content contract (confirmed with the user, see BRIEF.md): Table A (sales)
// and Table B (commission/royalty) rendered as one card per currency, so no
// local amount is ever added across JPY/INR/RMB; the workbook's own caution
// strings stay visible; USD is derived here from fx_domestic.json (USD per
// unit = currency HKD rate / the USD row's HKD rate) and captioned as a memo
// figure that is NOT part of the MA/DI USD totals.
// =========================================================================
function readJsonOptional(fileName) {
  try { return readJson(fileName); } catch (e) {
    if (e && e.code === 'ENOENT') return null;
    throw e;
  }
}
// DOMESTIC_FX_ALIAS (RMB -> CNY) is declared once at module scope above the
// Dashboard template, shared with the memo strip - do not redeclare here.
// domesticJson / fxDomesticJson load once, above the Dashboard template (the
// restyle stage's source-footnotes need fx_domestic.json's as-of string there).

const DOMESTIC_CSS = `
  .dom-title { font-size: 20px; font-weight: 600; margin: 0 0 8px 0; color: var(--text); }
  .dom-rule { border: 0; border-top: 1px solid var(--accent); margin: 0 0 14px 0; }
  .dom-action { font-size: 14px; line-height: 1.5; margin: 0 0 6px 0; max-width: 1040px; color: var(--text); }
  .dom-action .dom-v { font-weight: 600; font-variant-numeric: tabular-nums; }
  /* --warn is defined by the report-wide token pass when it runs; the literal
     brick is this page's own fallback so the caution strip is already correct
     today rather than waiting on that other edit. */
  .dom-caution { background: var(--card-bg); border: 1px solid var(--border-light); border-left: 3px solid var(--warn, #a63d40); border-radius: 4px; padding: 10px 14px; margin: 14px 0 18px 0; max-width: 1040px; }
  body[data-theme="dark"] .dom-caution { border-left-color: var(--warn, #c9706f); }
  .dom-caution-k { font-size: 11px; letter-spacing: .06em; text-transform: uppercase; color: var(--text-muted); margin-bottom: 4px; }
  .dom-caution ul { margin: 0; padding-left: 18px; font-size: 12px; line-height: 1.65; color: var(--text); }
  .dom-caution li.verbatim { color: var(--text-muted); font-style: italic; }
  .dom-cards { display: flex; flex-wrap: wrap; gap: 16px; align-items: flex-start; }
  .dom-card { flex: 1 1 460px; min-width: 300px; background: var(--card-bg); border: 1px solid var(--border-light); border-radius: 6px; box-shadow: 0 1px 3px var(--shadow); padding: 12px 14px 14px 14px; }
  .dom-card-head { border-bottom: 1px solid var(--border); padding-bottom: 6px; margin-bottom: 2px; }
  .dom-cur { display: flex; align-items: baseline; gap: 8px; }
  .dom-cur-code { font-size: 15px; font-weight: 700; color: var(--text); }
  .dom-cur-market { font-size: 12px; color: var(--text-muted); }
  .dom-rate { font-size: 11px; color: var(--text-muted); margin-top: 3px; font-variant-numeric: tabular-nums; }
  .dom-kicker { font-size: 10px; letter-spacing: .06em; text-transform: uppercase; color: var(--text-muted); margin: 12px 0 2px 0; }
  /* The card is one currency's whole story (entity rows x 2 tables), so it has
     to survive three-up on a 1920 screen: 11px figures, header text allowed to
     wrap onto two lines, and the entity name allowed to wrap too. Numbers keep
     nowrap - a split amount ("395,01 6.86") is worse than a wrapped caption. */
  table.dom-tbl { width: 100%; border-collapse: collapse; font-size: 11px; margin-top: 4px; }
  .dom-tbl th { font-weight: 600; font-size: 10px; color: var(--text-muted); text-align: right; border-bottom: 1px solid var(--border); padding: 3px 5px; white-space: normal; vertical-align: bottom; line-height: 1.25; }
  .dom-tbl th.lbl { text-align: left; }
  .dom-tbl td { text-align: right; padding: 4px 5px; border-bottom: 1px solid var(--border-light); color: var(--text); font-variant-numeric: tabular-nums; white-space: nowrap; }
  .dom-tbl td.lbl { text-align: left; white-space: normal; line-height: 1.3; }
  .dom-tbl td.nil, .dom-tbl th.nil { color: var(--text-muted); }
  .dom-tbl td.tot, .dom-tbl th.tot { border-left: 1px solid var(--border); }
  .dom-tbl td.usd, .dom-tbl th.usd { border-left: 1px solid var(--border); }
  .dom-tbl tfoot td { border-bottom: 0; border-top: 1px solid var(--border); font-weight: 700; }
  .dom-memo { font-size: 11px; color: var(--text-muted); margin-top: 6px; line-height: 1.6; }
  .dom-memo .v { color: var(--text); font-weight: 600; font-variant-numeric: tabular-nums; }
  .dom-sup { font-size: 9px; color: var(--text-muted); }
  .dom-usdbox { margin: 18px 0 4px 0; background: var(--card-bg); border: 1px solid var(--border); border-radius: 6px; padding: 12px 14px; max-width: 720px; }
  .dom-usdbox h3 { margin: 0 0 6px 0; font-size: 12px; letter-spacing: .04em; text-transform: uppercase; color: var(--text-muted); font-weight: 600; }
  table.dom-usdtbl { border-collapse: collapse; font-size: 12px; width: 100%; }
  .dom-usdtbl th { text-align: right; font-size: 11px; color: var(--text-muted); font-weight: 600; border-bottom: 1px solid var(--border); padding: 4px 8px; white-space: nowrap; }
  .dom-usdtbl th.lbl { text-align: left; }
  .dom-usdtbl td { text-align: right; padding: 4px 8px; border-bottom: 1px solid var(--border-light); font-variant-numeric: tabular-nums; color: var(--text); }
  .dom-usdtbl td.lbl { text-align: left; color: var(--text-muted); }
  .dom-usdtbl tr.grand td { border-top: 1px solid var(--border); border-bottom: 0; font-weight: 700; }
  .dom-usdnote { font-size: 11px; color: var(--text-muted); margin-top: 6px; }
  .dom-notes { margin-top: 20px; border-top: 1px solid var(--border); padding-top: 8px; font-size: 11px; color: var(--text-muted); max-width: 1040px; line-height: 1.6; }
  .dom-notes-k { font-size: 10px; letter-spacing: .06em; text-transform: uppercase; margin-bottom: 6px; }
  .dom-note { margin-bottom: 5px; }
  .dom-note .k { font-weight: 600; color: var(--text); }
  .dom-note.warn { color: #a63d40; }
  body[data-theme="dark"] .dom-note.warn { color: #d98f8e; }
`;

function buildDomesticHtml(domestic, fx) {
  if (!domestic || !domestic.tableA || !domestic.tableB) throw new Error('domestic.json is missing tableA/tableB');
  if (!fx || !Array.isArray(fx.rows)) throw new Error('fx_domestic.json is missing its rows array');

  function esc(value) {
    return String(value === null || value === undefined ? '' : value)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  // Islands ship Value2 numbers, but the documented fallback is the workbook's
  // own .Text - accept either, and treat a blank/"—" as "no figure" (null),
  // never as zero: a zero would make a subtotal row look like a real balance.
  function num(value) {
    if (typeof value === 'number') return isFinite(value) ? value : null;
    if (value === null || value === undefined) return null;
    const s = String(value).trim();
    if (!s || s === '-' || s === '—' || /not applicable|pending/i.test(s)) return null;
    const n = Number(s.replace(/,/g, ''));
    return isFinite(n) ? n : null;
  }
  function fmt2(n) { return Number(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  function fmtRate(n) { return String(Math.round(Number(n) * 1e7) / 1e7); }
  const NIL = '<span class="nil">—</span>';

  // ---- FX: USD per unit of the local currency -----------------------------
  const fxByCode = {};
  fx.rows.forEach(function (row) { fxByCode[String(row[1]).trim().toUpperCase()] = row; });
  if (!fxByCode.USD) throw new Error('fx_domestic.json has no USD row - no USD equivalent can be derived');
  const usdHkd = Number(fxByCode.USD[3]);
  function fxRowFor(code) {
    const key = String(code).trim().toUpperCase();
    return fxByCode[DOMESTIC_FX_ALIAS[key] || key] || null;
  }
  function usdPerUnit(code) {
    const row = fxRowFor(code);
    if (!row) return null;
    const unit = Number(row[2]) || 1; // the sheet's Unit column (100 = rate per 100)
    return (Number(row[3]) / unit) / usdHkd;
  }

  // ---- island table -> model (the island's column indexes are 1-based) -----
  const METRIC_KEYS = {
    '2026 Shipped': 'domColShipped', '2026 Outstanding': 'domColOutstanding',
    '2026 Total': 'domColTotal2026', '2027 Outstanding': 'domColOutstanding2027'
  };
  function tableModel(sheet, label) {
    const metricIdxs = sheet.metricCols.map(function (n) { return n - 1; });
    metricIdxs.forEach(function (i) { if (i < 0 || i >= sheet.header.length) throw new Error('domestic.json ' + label + ': metricCols out of range'); });
    const entityIdx = sheet.entityCol - 1;
    const currencyIdx = sheet.currencyCol - 1;
    const marketIdx = sheet.header.indexOf('Market');
    // c.i = the column's index in the SOURCE row (1-based island number minus
    // one); c.k = its index in the trimmed per-row vals array. Reading a value
    // with c.i instead of c.k silently yields undefined -> "NaN" cells, so the
    // two are kept as separate named fields rather than reused.
    const cols = metricIdxs.map(function (i, k) {
      const h = String(sheet.header[i]).trim();
      return { i: i, k: k, header: h, key: METRIC_KEYS[h] || null, isTotal: /total/i.test(h) };
    });
    const totalCol = cols.filter(function (c) { return c.key === 'domColTotal2026' || c.isTotal; })[0] || null;
    const laterCol = cols.filter(function (c) { return c !== totalCol && /^20(2[7-9])/.test(c.header); })[0] || null;
    const priority = cols.filter(function (c) { return c === totalCol; }).concat(cols.filter(function (c) { return c !== totalCol; }));
    return {
      label: label, header: sheet.header, cols: cols, totalCol: totalCol, laterCol: laterCol, priority: priority,
      rows: sheet.rows.map(function (r) {
        return {
          entity: String(r[entityIdx] === null || r[entityIdx] === undefined ? '' : r[entityIdx]).trim(),
          market: marketIdx >= 0 ? String(r[marketIdx] === null || r[marketIdx] === undefined ? '' : r[marketIdx]).trim() : '',
          currency: String(r[currencyIdx] === null || r[currencyIdx] === undefined ? '' : r[currencyIdx]).trim(),
          vals: cols.map(function (c) { return num(r[c.i]); })
        };
      })
    };
  }
  const modelA = tableModel(domestic.tableA, 'Table A');
  const modelB = tableModel(domestic.tableB, 'Table B');
  const models = [modelA, modelB];

  const currencyOrder = [];
  models.forEach(function (m) { m.rows.forEach(function (r) { if (r.currency && currencyOrder.indexOf(r.currency) < 0) currencyOrder.push(r.currency); }); });
  const unresolvedFx = currencyOrder.filter(function (code) { return usdPerUnit(code) === null; });
  unresolvedFx.forEach(function (code) { console.warn('WARNING: Domestic currency ' + code + ' has no row in fx_domestic.json (alias ' + (DOMESTIC_FX_ALIAS[code] || code) + ') - its USD equivalents are blank, the local figures still show.'); });

  function localSum(rows, col) {
    let s = null;
    rows.forEach(function (r) { const v = r.vals[col.k]; if (v !== null) s = (s === null ? 0 : s) + v; });
    return s;
  }
  function usdSum(rows, col) {
    let s = null;
    rows.forEach(function (r) {
      const v = r.vals[col.k];
      if (v === null) return;
      const rate = usdPerUnit(r.currency);
      if (rate === null) return;
      s = (s === null ? 0 : s) + v * rate;
    });
    return s;
  }

  // ---- data-driven action titles + the 2027-only sentence -----------------
  const fills = {};
  const aTotalCol = modelA.totalCol, bTotalCol = modelB.totalCol;
  const aUsd2026 = aTotalCol ? usdSum(modelA.rows, aTotalCol) : null;
  const bUsd2026 = bTotalCol ? usdSum(modelB.rows, bTotalCol) : null;
  fills.domActionSales = { entities: String(modelA.rows.length), currencies: String(currencyOrder.length), usd: aUsd2026 === null ? '—' : fmt2(aUsd2026) };
  fills.domActionCommission = { usd: bUsd2026 === null ? '—' : fmt2(bUsd2026), pct: (aUsd2026 && bUsd2026 !== null) ? (bUsd2026 / aUsd2026 * 100).toFixed(1) : '—' };
  // Entities with no 2026 figure at all, only a later-year outstanding balance.
  const pendingRows = modelA.rows.filter(function (r) {
    return r.vals.every(function (v, i) { return modelA.cols[i] === modelA.laterCol ? true : v === null; })
      && modelA.laterCol && r.vals[modelA.cols.indexOf(modelA.laterCol)] !== null;
  });
  let pendingHtml = '';
  if (pendingRows.length) {
    const sameCcy = pendingRows.every(function (r) { return r.currency === pendingRows[0].currency; });
    const bByName = {};
    modelB.rows.forEach(function (r) { bByName[r.entity] = r; });
    const pSum = [0, 0];
    let pOk = sameCcy;
    pendingRows.forEach(function (r) {
      const aV = r.vals[modelA.cols.indexOf(modelA.laterCol)];
      const bRow = bByName[r.entity];
      const bV = (bRow && modelB.laterCol) ? bRow.vals[modelB.cols.indexOf(modelB.laterCol)] : null;
      if (aV === null && bV === null) { pOk = false; return; }
      pSum[0] += aV || 0; pSum[1] += bV || 0;
    });
    if (pOk) {
      const ccy = pendingRows[0].currency;
      fills.domActionPending = {
        names: pendingRows.map(function (r) { return r.entity.split('—')[0].trim(); }).join(', '),
        sales: ccy + ' ' + fmt2(pSum[0]), comm: ccy + ' ' + fmt2(pSum[1])
      };
    }
  }
  function fillEn(key) {
    const toks = fills[key] || {};
    return String(DOM_EN[key]).replace(/\{(\w+)\}/g, function (m, name) { return Object.prototype.hasOwnProperty.call(toks, name) ? String(toks[name]) : m; });
  }
  function actionHtml(key, extraClass) {
    if (!fills[key]) return '';
    const withValues = fillEn(key).replace(/(USD [0-9,.]+|[A-Z]{3} [0-9,.]+)/g, '<span class="dom-v">$1</span>');
    return '<div class="dom-action' + (extraClass ? ' ' + extraClass : '') + '" data-dom-fill="' + key + '">' + withValues + '</div>';
  }

  // ---- one card per currency ----------------------------------------------
  function cellValue(v) { return v === null ? NIL : fmt2(v); }
  function usdCellHtml(r, model) {
    let principalCol = null, value = null;
    for (let i = 0; i < model.priority.length; i++) {
      const c = model.priority[i];
      if (r.vals[c.k] !== null) { principalCol = c; value = r.vals[c.k]; break; }
    }
    if (principalCol === null) return '<td class="usd">' + NIL + '</td>';
    const rate = usdPerUnit(r.currency);
    if (rate === null) return '<td class="usd">' + NIL + '</td>';
    const isLater = !!model.laterCol && principalCol === model.laterCol;
    const flag = isLater ? '<sup class="dom-sup">†</sup>' : '';
    const tip = isLater ? ' data-i18n-title="domBasisFlagTitle" title="' + esc(DOM_EN.domBasisFlagTitle) + '"' : '';
    return '<td class="usd"' + tip + '>' + fmt2(value * rate) + flag + '</td>';
  }
  // showMarket: every row of a one-currency card normally sits in the same
  // market (JPY=Japan, INR=India, RMB=China), and that card header already says
  // so, so repeating it in a column costs width the figures need. The column
  // comes back automatically the moment a card ever holds two markets.
  function tableHtml(model, rows, showMarket) {
    let html = '<table class="dom-tbl"><thead><tr>';
    html += '<th class="lbl" scope="col">' + '<span data-i18n="domColEntity">' + esc(DOM_EN.domColEntity) + '</span>' + '</th>';
    if (showMarket) html += '<th class="lbl" scope="col"><span data-i18n="domColMarket">' + esc(DOM_EN.domColMarket) + '</span></th>';
    model.cols.forEach(function (c) {
      const label = c.key ? '<span data-i18n="' + c.key + '">' + esc(model.header[c.i]) + '</span>' : esc(c.header);
      html += '<th class="' + (c.isTotal ? 'tot' : '') + '" scope="col">' + label + '</th>';
    });
    html += '<th class="usd" scope="col"><span data-i18n="domColUsd">' + esc(DOM_EN.domColUsd) + '</span><sup class="dom-sup">†</sup></th>';
    html += '</tr></thead><tbody>';
    rows.forEach(function (r) {
      html += '<tr><td class="lbl">' + esc(r.entity) + '</td>' + (showMarket ? '<td class="lbl">' + esc(r.market) + '</td>' : '');
      model.cols.forEach(function (c) {
        const cls = (c.isTotal ? 'tot' : '') + (r.vals[c.k] === null ? ' nil' : '');
        html += '<td class="' + cls.trim() + '">' + cellValue(r.vals[c.k]) + '</td>';
      });
      html += usdCellHtml(r, model) + '</tr>';
    });
    html += '</tbody><tfoot><tr>';
    html += '<td class="lbl" colspan="' + (showMarket ? 2 : 1) + '"><span data-i18n="domSubtotal">' + esc(DOM_EN.domSubtotal) + '</span></td>';
    model.cols.forEach(function (c) {
      const s = localSum(rows, c);
      const cls = (c.isTotal ? 'tot' : '') + (s === null ? ' nil' : '');
      html += '<td class="' + cls.trim() + '">' + (s === null ? NIL : fmt2(s)) + '</td>';
    });
    // No USD figure in the subtotal row on purpose: the rows above can sit on
    // different bases (2026 Total vs a 2027-only balance), so adding the USD
    // column would mix them. The memo line below states each basis separately.
    html += '<td class="usd nil" title="see the memo line">' + NIL + '</td>';
    html += '</tr></tfoot></table>';
    return html;
  }
  function memoHtml(model, rows) {
    const parts = model.cols.map(function (c) {
      const label = c.key ? '<span data-i18n="' + c.key + '">' + esc(model.header[c.i]) + '</span>' : esc(c.header);
      const s = usdSum(rows, c);
      return label + ' <span class="v">' + (s === null ? '—' : fmt2(s)) + '</span>';
    });
    return '<div class="dom-memo"><span data-i18n="domCardUsdLabel">' + esc(DOM_EN.domCardUsdLabel) + '</span> · ' + parts.join(' · ') + '</div>';
  }
  const cardsHtml = currencyOrder.map(function (code) {
    const rowA = modelA.rows.filter(function (r) { return r.currency === code; });
    const rowB = modelB.rows.filter(function (r) { return r.currency === code; });
    const markets = [];
    rowA.concat(rowB).forEach(function (r) { if (r.market && markets.indexOf(r.market) < 0) markets.push(r.market); });
    const market = markets.join(' / ');
    const showMarket = markets.length > 1;
    const frow = fxRowFor(code);
    const per = usdPerUnit(code);
    const hkd = frow ? fmtRate(Number(frow[3]) / (Number(frow[2]) || 1)) + ' HKD' : '— HKD';
    const perUsdTxt = per === null ? '—' : (Math.round(per * 1e6) / 1e6).toFixed(6);
    // Spelled out on the card itself: the workbook's "RMB" is the FX sheet's
    // "CNY", and a reader auditing the derived column must be able to see that
    // the join crossed a naming gap rather than found an exact match.
    const fxCcy = frow ? String(frow[1]).trim().toUpperCase() : '';
    const aliasTxt = (frow && fxCcy !== code.toUpperCase()) ? ' (the FX sheet names this ' + fxCcy + ')' : '';
    let inner = '';
    if (rowA.length) inner += '<div class="dom-kicker"><span data-i18n="domSecSales">' + esc(DOM_EN.domSecSales) + '</span></div>' + tableHtml(modelA, rowA, showMarket) + memoHtml(modelA, rowA);
    if (rowB.length) inner += '<div class="dom-kicker"><span data-i18n="domSecCommission">' + esc(DOM_EN.domSecCommission) + '</span></div>' + tableHtml(modelB, rowB, showMarket) + memoHtml(modelB, rowB);
    return '<section class="dom-card">'
      + '<div class="dom-card-head"><div class="dom-cur"><span class="dom-cur-code">' + esc(code) + '</span><span class="dom-cur-market">' + esc(market) + '</span></div>'
      + '<div class="dom-rate">' + esc(code) + ' 1 = ' + hkd + aliasTxt + ' ÷ USD ' + fmtRate(usdHkd) + ' = ' + perUsdTxt + ' USD</div></div>'
      + inner + '</section>';
  }).join('\n');

  // ---- derived USD memo block (the only cross-currency view on the page) ---
  function memoRowHtml(labelKey, labelEn, colA, colB, isGrand) {
    const sA = colA ? usdSum(modelA.rows, colA) : null;
    const sB = colB ? usdSum(modelB.rows, colB) : null;
    return '<tr' + (isGrand ? ' class="grand"' : '') + '><td class="lbl"><span data-i18n="' + labelKey + '">' + esc(labelEn) + '</span></td>'
      + '<td' + (sA === null ? ' class="nil"' : '') + '>' + (sA === null ? '—' : fmt2(sA)) + '</td>'
      + '<td' + (sB === null ? ' class="nil"' : '') + '>' + (sB === null ? '—' : fmt2(sB)) + '</td></tr>';
  }
  function colByModel(model, key) { return model.cols.filter(function (c) { return c.key === key; })[0] || null; }
  let usdBoxRows = '';
  // 2026-basis lines first, then the 2026 grand line, then the later-year
  // balance last: the totals above it must never read as if they included it.
  models[0].cols.forEach(function (c) {
    if (c.isTotal || c === models[0].laterCol) return; // both drawn explicitly below
    usdBoxRows += memoRowHtml(c.key || 'domColShipped', c.header, c, colByModel(modelB, c.key), false);
  });
  usdBoxRows += memoRowHtml('domUsdTotalLabel', DOM_EN.domUsdTotalLabel, modelA.totalCol, modelB.totalCol, true);
  const laterA = modelA.laterCol;
  if (laterA) usdBoxRows += memoRowHtml(laterA.key || 'domColOutstanding2027', laterA.header, laterA, colByModel(modelB, laterA.key), false);

  // ---- footnotes: provenance straight out of the two islands ---------------
  const usedFxRows = [];
  currencyOrder.forEach(function (code) { const row = fxRowFor(code); if (row && usedFxRows.indexOf(row) < 0) usedFxRows.push(row); });
  const fxRatesTxt = usedFxRows.map(function (row) { return String(row[1]).trim() + ' ' + fmtRate(Number(row[3]) / (Number(row[2]) || 1)) + ' HKD'; }).join(' · ') + ' · USD ' + fmtRate(usdHkd) + ' HKD';
  const syntheticFx = usedFxRows.filter(function (row) { const s = String(row[4] || '').trim(); return s !== '' && s.toLowerCase() !== 'standard'; });
  let notes = '';
  notes += '<div class="dom-note"><span class="k" data-i18n="domSourceFile">' + esc(DOM_EN.domSourceFile) + '</span> '
    + '<span>' + esc(domestic.file) + ' <span data-i18n="domSourceSheetNote">' + esc(DOM_EN.domSourceSheetNote) + '</span>'
    + (domestic.meta ? ' — ' + esc(domestic.meta) : '') + '</span></div>';
  // Every prose fragment below is its own data-i18n LEAF (a tagged element
  // must never wrap another element - applyLanguage sets textContent, which
  // would wipe the nested span and any marker inside it); everything the
  // islands themselves supply (file names, titles, dates, rate strings) is
  // verbatim data and stays untagged in both languages.
  notes += '<div class="dom-note"><span class="k" data-i18n="domFxSource">' + esc(DOM_EN.domFxSource) + '</span> '
    + '<span>' + esc(fx.file) + ' — &ldquo;' + esc(fx.title) + '&rdquo; · ' + esc(fx.asOf) + ' · ' + esc(fxRatesTxt) + ' — '
    + '<span data-i18n="domUsdBasisFormula">' + esc(DOM_EN.domUsdBasisFormula) + '</span></span></div>';
  syntheticFx.forEach(function (row) {
    const code = String(row[1]).trim().toUpperCase();
    const fb = /fallback/i.test(String(row[4] || ''));
    notes += '<div class="dom-note' + (fb ? ' warn' : '') + '"><span class="k" data-i18n="' + (code === 'INR' ? 'domInrSource' : 'domFxSource') + '">'
      + esc(code === 'INR' ? DOM_EN.domInrSource : DOM_EN.domFxSource) + '</span> '
      + (fb ? '<span data-i18n="domInrFallbackWarn">' + esc(DOM_EN.domInrFallbackWarn) + '</span> · ' : '')
      + '<span>' + esc(row[4]) + '</span></div>';
  });
  notes += '<div class="dom-note"><span class="k" data-i18n="domAsOf">' + esc(DOM_EN.domAsOf) + '</span> <span>' + esc(domestic.asOf) + '</span></div>';
  notes += '<div class="dom-note"><span class="k" data-i18n="domUsdBasisNote">' + esc(DOM_EN.domUsdBasisNote) + '</span> <span><sup class="dom-sup">†</sup> '
    + '<span data-i18n="domUsdRowSkipNote">' + esc(DOM_EN.domUsdRowSkipNote) + '</span> · <span data-i18n="domUsdBoxCaption">' + esc(DOM_EN.domUsdBoxCaption) + '</span></span></div>';
  notes += '<div class="dom-note"><span class="k" data-i18n="domNoYoy">' + esc(DOM_EN.domNoYoy) + '</span> · <span data-i18n="domCancelled">' + esc(DOM_EN.domCancelled) + '</span></div>';
  if (domestic.sourceUrl) {
    notes += '<div class="dom-note"><span class="k" data-i18n="domMjpSource">' + esc(DOM_EN.domMjpSource) + '</span> <span>' + esc(domestic.sourceUrl) + ' — <span data-i18n="domMjpNote">' + esc(DOM_EN.domMjpNote) + '</span></span></div>';
  }
  if (unresolvedFx.length) {
    notes += '<div class="dom-note warn"><span class="k" data-i18n="domFxMissingLabel">' + esc(DOM_EN.domFxMissingLabel) + '</span> <span>' + esc(unresolvedFx.join(', ')) + ' — <span data-i18n="domFxMissingNote">' + esc(DOM_EN.domFxMissingNote) + '</span></span></div>';
  }
  // Source-island provenance line, so an auditor can trace this page back to
  // the exact _work extracts the PowerShell generator wrote for it.
  notes += '<div class="dom-note"><span class="k" data-i18n="domIslandsLabel">' + esc(DOM_EN.domIslandsLabel) + '</span> <span>domestic.json · fx_domestic.json</span></div>';
  const cautionNotes = Array.from(new Set((domestic.notes || []).map(function (s) { return String(s).trim(); })).values())
    .map(function (s) { return '<li class="verbatim">' + esc(s) + '</li>'; }).join('');

  const domFillsJson = neutralize(JSON.stringify(fills));
  return `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Domestic</title>
<style>${SHARED_CSS}${DOMESTIC_CSS}${MCKINSEY_CSS}</style>
</head>
<body>
<h1 class="dom-title" data-i18n="title_domestic">${esc(DOM_EN.title_domestic)}</h1>
<hr class="dom-rule">
<p class="page-intro" data-i18n="introDomestic">${esc(INTRO_EN.introDomestic)}</p>
${actionHtml('domActionSales')}
${actionHtml('domActionCommission')}
${actionHtml('domActionPending')}
<div class="dom-caution">
  <div class="dom-caution-k" data-i18n="domCautionTitle">${esc(DOM_EN.domCautionTitle)}</div>
  <ul>
    <li data-i18n="domCautionNoSum">${esc(DOM_EN.domCautionNoSum)}</li>
    ${cautionNotes}
    <li data-i18n="domNoYoy">${esc(DOM_EN.domNoYoy)}</li>
    <li data-i18n="domCancelled">${esc(DOM_EN.domCancelled)}</li>
  </ul>
</div>
<div class="dom-cards">
${cardsHtml}
</div>
<div class="dom-usdbox">
  <h3 data-i18n="domUsdBoxTitle">${esc(DOM_EN.domUsdBoxTitle)}</h3>
  <table class="dom-usdtbl">
    <thead><tr><th class="lbl" scope="col"><span data-i18n="domUsdBasisCol">${esc(DOM_EN.domUsdBasisCol)}</span></th><th scope="col"><span data-i18n="domColSalesShort">${esc(DOM_EN.domColSalesShort)}</span></th><th scope="col"><span data-i18n="domColCommissionShort">${esc(DOM_EN.domColCommissionShort)}</span></th></tr></thead>
    <tbody>
${usdBoxRows}
    </tbody>
  </table>
  <div class="dom-usdnote"><sup class="dom-sup">†</sup> <span data-i18n="domUsdBoxCaption">${esc(DOM_EN.domUsdBoxCaption)}</span></div>
</div>
<div class="dom-notes">
  <div class="dom-notes-k" data-i18n="domFootnotesTitle">${esc(DOM_EN.domFootnotesTitle)}</div>
  ${notes}
</div>
<script>
${I18N_JS}
var DOM_FILLS = ${domFillsJson};
function fillKey(key) {
  var s = t(key);
  var toks = DOM_FILLS[key] || {};
  return String(s).replace(/\\{(\\w+)\\}/g, function (m, name) { return Object.prototype.hasOwnProperty.call(toks, name) ? String(toks[name]) : m; });
}
function applyTheme(theme) { document.body.setAttribute('data-theme', theme); }
// Same contract as every other page: chrome only, never an amount. This page
// has no DataTables instance, so swapping text is the whole job - and it is
// called once on load as well, so the tab also reads correctly when the file
// is opened outside the shell (applyLanguageToFrame never fires there).
function applyLanguage(lang) {
  currentLang = lang;
  try { localStorage.setItem('reportLang', lang); } catch (e) {}
  document.querySelectorAll('[data-i18n]').forEach(function (el) { el.textContent = t(el.getAttribute('data-i18n')); });
  document.querySelectorAll('[data-i18n-title]').forEach(function (el) { el.title = t(el.getAttribute('data-i18n-title')); });
  document.querySelectorAll('[data-dom-fill]').forEach(function (el) { el.textContent = fillKey(el.getAttribute('data-dom-fill')); });
}
applyLanguage(currentLang);
</script>
</body>
</html>`;
}

let domesticHtml = null;
if (domesticJson && fxDomesticJson) {
  domesticHtml = buildDomesticHtml(domesticJson, fxDomesticJson);
  console.log('Domestic sheet built (' + (domesticHtml.length / 1024).toFixed(1) + ' KB)');
} else {
  console.warn('WARNING: domestic.json and/or fx_domestic.json not in ' + DIR + ' - skipping the "Domestic" tab');
}

// =========================================================================
// Assemble final combined file - tab shell + iframe srcdoc, single self-
// contained file (no separate .files folder), theme toggle propagated into
// each iframe via a shared applyTheme(theme) contract.
// =========================================================================
// "name" is the stable English id used for the postMessage targetTab lookup
// (drilldown always targets literal 'Data') and must never change with the
// display language - "key" is only used to look up the tab's translated label.
const sheetsData = [
  { key: 'dashboard', name: 'Dashboard', html: dashboardHtml },
  { key: 'data', name: 'Data', html: dataSheetHtml },
  { key: 'exrate', name: 'Exchange Rate', html: exRateHtml },
  { key: 'yoyCustomers', name: 'YoY: Customers', html: yoyCustomersHtml },
  { key: 'yoyFabrication', name: 'YoY: Fabrication', html: yoyFabricationHtml },
  { key: 'yoyProdFamily', name: 'YoY: Prod Family', html: yoyProdFamilyHtml },
  { key: 'yoyPartId', name: 'YoY: Part ID', html: yoyPartIdHtml },
  { key: 'yoyNetNetMargin', name: 'YoY: Net Net Margin', html: yoyNetNetMarginHtml },
  { key: 'customerQuarter', name: 'Customer-Quarter', html: netMarginByCustomerHtml },
  // Order IS tab order: last of the Sales-side tabs, before the Purchase tab
  // that gets pushed on the next line. name:'Domestic' stays English for the
  // postMessage targetTab contract (see the note above), tab_domestic is what
  // the viewer sees.
  ...(domesticHtml ? [{ key: 'domestic', name: 'Domestic', html: domesticHtml }] : []),
];

// Purchase Report's "YTD by Vendor" tab, lifted verbatim from that project's
// own already-built HTML (per Peggy's request to have it inside the Sales
// report). Its page is a self-contained iframe document, so it drops in as
// just another SHEETS entry. Its own vendor drill-links target a "Data" tab
// that only exists in the Purchase report, so they no-op here (silently -
// the handler checks for the parent function first). Missing source file
// only warns, so a Sales-only rebuild never fails because of it.
const PURCHASE_HTML_PATH = 'C:\\Users\\User\\Desktop\\MTC Purchase Report Claude\\FIN_MTC_Purchase_Report_html_TEMP_RV1.html';
function extractPurchaseYtdTab() {
  if (!fs.existsSync(PURCHASE_HTML_PATH)) {
    console.warn('WARNING: Purchase report HTML not found, skipping "Purchase: YTD by Vendor" tab:', PURCHASE_HTML_PATH);
    return null;
  }
  const html = fs.readFileSync(PURCHASE_HTML_PATH, 'utf8');
  const anchor = 'var SHEETS = ';
  const start = html.indexOf(anchor);
  const endMarker = html.indexOf("\nvar tabsEl = document.getElementById('tabs');", start);
  if (start < 0 || endMarker < 0) throw new Error('Purchase report HTML has an unexpected shell structure - cannot locate its SHEETS array');
  const jsonEnd = html.lastIndexOf(';', endMarker);
  const sheets = JSON.parse(html.slice(start + anchor.length, jsonEnd).replace(/<\\\/script/gi, '</script'));
  const tab = sheets.find(s => s.key === 'purchaseYtd');
  if (!tab) throw new Error('Purchase report HTML has no "purchaseYtd" tab');
  // Restyle-stage rule: every tab closes with a source line. This page belongs
  // to the Purchase project (we must not restyle its internals), so the only
  // edit is appending a self-styled footer before its own </body> - inline
  // styles because we cannot assume its document carries our CSS tokens.
  const foot = '<div style="margin:18px 8px 6px;padding:8px 0 0;border-top:1px solid rgba(128,128,128,.4);font:10.5px/1.55 Calibri,Arial,sans-serif;color:rgba(128,128,128,.95)">Sources / 資料來源: FIN_MTC Purchase Report html (MTC Purchase Report Claude project, embedded as-built) — see that report for its period &amp; FX basis · 期間與匯率基礎詳該採購報表。</div>';
  const bodyEnd = tab.html.lastIndexOf('</body>');
  const patchedHtml = bodyEnd >= 0 ? tab.html.slice(0, bodyEnd) + foot + tab.html.slice(bodyEnd) : tab.html;
  return { key: 'purchaseYtdVendor', name: 'Purchase: YTD by Vendor', html: patchedHtml };
}
const purchaseYtdTab = extractPurchaseYtdTab();
if (purchaseYtdTab) { sheetsData.push(purchaseYtdTab); console.log('Added Purchase: YTD by Vendor tab from', PURCHASE_HTML_PATH); }
const sheetsJsonStr = neutralize(JSON.stringify(sheetsData));

const combinedHtml = `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>${reportTitle}</title>
<style>${MCKINSEY_CSS}
  /* Restyle stage: the shell chrome now reads the --shell-* custom properties
     from MCKINSEY_CSS (defined for light on :root and for dark under
     body[data-theme="dark"] there) - no literal navy/white rules survive here,
     so the tab bar and the welcome modal finally follow the theme toggle.
     Geometry: a flex column, NOT the old calc(100vh - 44px) magic constant -
     that constant assumed a one-row tab bar; with 11 tabs (#tabs wraps,
     flex-wrap: wrap) plus the serif masthead, any fixed subtraction either
     clipped the bottom of the page or left a dead band. The iframe simply
     takes the remaining height (flex:1; min-height:0) whatever the chrome
     above it grows to. */
  body { font-family: Calibri, Arial, sans-serif; margin: 0; padding: 0; background: var(--shell-bg); color: var(--shell-text); display: flex; flex-direction: column; height: 100vh; }
  #masthead { flex: none; background: var(--shell-card); border-bottom: 1px solid var(--shell-rule); padding: 10px 16px 9px; }
  #masthead .mast-title { font-family: var(--serif); font-size: 17px; font-weight: 600; color: var(--shell-text); }
  #masthead .mast-sub { font-size: 11px; color: var(--shell-muted); margin-top: 3px; }
  #tabbar { display: flex; background: var(--shell-nav); align-items: stretch; flex: none; }
  #tabs { display: flex; padding: 0; margin: 0; list-style: none; flex-wrap: wrap; flex: 1; }
  #tabs button { background: var(--shell-nav); color: var(--shell-nav-text); border: none; border-right: 1px solid var(--shell-nav-sep); padding: 9px 13px; cursor: pointer; font-size: 12.5px; }
  #tabs button.active { background: var(--shell-nav-active); color: var(--shell-nav-text); font-weight: bold; box-shadow: inset 0 -2px 0 var(--shell-accent); }
  #tabs button:hover { background: var(--shell-nav-hover); }
  #themeToggle, #langToggle { background: var(--shell-nav-active); color: var(--shell-nav-text); border: none; padding: 9px 13px; cursor: pointer; font-size: 12.5px; white-space: nowrap; }
  #themeToggle:hover, #langToggle:hover { background: var(--shell-nav-hover); }
  #langToggle { border-left: 1px solid var(--shell-nav-sep); }
  #tabs button:focus-visible, #themeToggle:focus-visible, #langToggle:focus-visible { outline: 2px solid var(--shell-accent); outline-offset: -2px; }
  iframe { flex: 1 1 auto; min-height: 0; width: 100%; border: none; display: block; }
  #introOverlay { position: fixed; inset: 0; background: rgba(5,28,44,0.55); display: flex; align-items: center; justify-content: center; z-index: 1000; }
  #introOverlay.hidden { display: none; }
  #introModal { background: var(--shell-card); color: var(--shell-text); border: 1px solid var(--shell-rule); border-radius: 4px; box-shadow: 0 8px 30px rgba(0,0,0,0.35); max-width: 640px; width: 90%; max-height: 85vh; overflow-y: auto; padding: 24px 28px; }
  #introModal h2 { color: var(--shell-accent); margin: 0 0 4px 0; font-size: 20px; }
  #introModal .intro-sub { color: var(--shell-muted); font-size: 13px; margin: 0 0 14px 0; }
  #introModal ul { margin: 0 0 16px 0; padding-left: 20px; font-size: 13px; line-height: 1.7; }
  #introModal ul b { color: var(--shell-text); }
  #introModal li { margin-bottom: 4px; }
  #introModal ul + p.intro-note { font-size: 12px; color: var(--shell-muted); margin: 0 0 16px 0; border-top: 1px solid var(--shell-rule); padding-top: 10px; }
  #introFooter { display: flex; align-items: center; justify-content: space-between; gap: 12px; flex-wrap: wrap; }
  #introFooter label { font-size: 13px; color: var(--shell-muted); display: flex; align-items: center; gap: 6px; cursor: pointer; }
  #introClose { background: var(--shell-accent); color: #ffffff; border: none; border-radius: 3px; padding: 9px 20px; font-size: 14px; cursor: pointer; }
  #introClose:hover { background: var(--shell-nav-active); }
</style>
</head>
<body>
<div id="introOverlay" class="hidden">
  <div id="introModal">
    <h2>Welcome / 歡迎使用</h2>
    <p class="intro-sub">${reportTitle} — interactive web version of the official monthly pack, one offline file · report period ${reportPeriodRange} / 正式月報互動網頁版，單一離線檔案 · 報表期間 ${reportPeriodRange}</p>
    <!-- Intro copy v3 (doc rewrite, 2026-11-10): full rewrite of the v2 restage
     copy - adds what/when provenance (interpolated from the real build data
     instead of hardcoded counts), the Ask AI panel (added but uncommitted at
     v2 time), and explicit reading rules for the Domestic tab. Dismiss key
     bumped to reportIntroDismissed_v3 alongside this text, per the rule. -->
    <ul>
      <li><b>What this is / 內容</b> — ${sheetsData.length} tabs: Dashboard, Data (the full ${H.length}-column ledger), Exchange Rate, five YoY comparisons (2026 vs 2025, each with Total Index), Customer-Quarter, <b>Domestic</b>, and Purchase: YTD by Vendor.<br>${sheetsData.length} 個頁籤：總覽、明細（完整 ${H.length} 欄交易帳）、匯率、五個「2026 對比 2025」年度比較頁（各附總計指標）、客戶季度毛利、<b>國內訂單</b>、採購：年度供應商排行。</li>
      <li><b>Dashboard / 總覽</b> — a one-line conclusion computed live from your filters, above the 3D trade globe: toggle Global Sales (destinations) ⇄ Global Purchase (origins), click a country marker to open the Data tab pre-filtered. The Geographic bar and Origin pie stay as the automatic fallback.<br>一行結論隨篩選即時計算；下方 3D 貿易地球可切換全球銷售（目的地）／全球採購（原產國），點擊國家標記即跳至明細並自動篩選；地理長條圖與原產國圓餅圖保留為自動備援。</li>
      <li><b>Domestic tab / 國內訂單頁</b> — 2026–2027 domestic orders &amp; MTC commission (MJP · MCN · MZQ · MIN), per-currency cards: local amounts are never summed across currencies, and the derived-USD memos are <i>not</i> included in the MA/DI USD totals. As-of date shown on the tab.<br>按貨幣分卡，絕不跨貨幣相加；推算 USD 備忘<b>未</b>計入 MA/DI USD 總數；基準日見頁內標示。</li>
      <li><b>Ask AI / 問 AI</b> — the "Ask AI" button (bottom-right) answers questions about the figures on screen and can propose filters you confirm before they apply. It sees filters and aggregates only — never transaction rows — and works out of the box via the cloud AI service (no access code needed); if the service is unreachable, everything else works unchanged.<br>右下角「問 AI」可詢問頁面數字並建議可套用的篩選（須經你確認）；AI 僅見篩選與彙總值、絕不讀交易明細；經雲端 AI 服務即時可用（毋須存取碼），服務無法連線時其他功能不受影響。</li>
      <li><b>How to work with it / 操作方式</b> — Year · Month · Customer · Vendor · Country filters update KPIs and charts together; "Export CSV" saves exactly the filtered view; EN/中文 and light/dark toggles sit top-right; exhibits are numbered and every tab closes with its sources, FX basis and period.<br>篩選同步更新 KPI 與圖表；「Export CSV」僅匯出目前篩選結果；右上角切換中英文與深淺色；圖表統一編號，每頁頁尾附資料來源、匯率基礎與報表期間。</li>
    </ul>
    <p class="intro-note">Same data as the official Excel pack — a convenience for browsing and presenting, not a replacement. / 資料與正式 Excel 報表一致，本頁僅為瀏覽與簡報便利之用，非取代 Excel 報表。</p>
    <div id="introFooter">
      <label><input type="checkbox" id="introDontShow"> Don't show this again / 下次不再顯示</label>
      <button id="introClose">Get Started / 開始使用</button>
    </div>
  </div>
</div>
<div id="masthead">
  <div class="mast-title">${reportTitle}</div>
  <div class="mast-sub">${sheetsData.length} tabs — Sales, Domestic &amp; Purchase in one offline file · ${reportPeriodRange} · FX per the Exchange Rate tab; Domestic memo basis: ${fxAsOfNote}</div>
</div>
<div id="tabbar">
  <div id="tabs"></div>
  <button id="langToggle">中文</button>
  <button id="themeToggle">🌙 Dark Mode</button>
</div>
<iframe id="sheetFrame" title="Report sheet"></iframe>
<script>
${I18N_JS}
var SHEETS = ${sheetsJsonStr};
var tabsEl = document.getElementById('tabs');
var frameEl = document.getElementById('sheetFrame');
var themeBtn = document.getElementById('themeToggle');
var langBtn = document.getElementById('langToggle');
var currentTheme = 'light';
try { currentTheme = localStorage.getItem('reportTheme') || 'light'; } catch (e) {}
// Restyle stage: the shell chrome (masthead/tab bar/intro modal) is token-
// driven via body[data-theme="dark"] in MCKINSEY_CSS, so the shell must now
// apply the theme to its OWN body too - previously only the iframes got it
// and the shell stayed navy/white forever (half-done dark mode).
function applyShellTheme() { document.body.setAttribute('data-theme', currentTheme); }
function applyThemeToFrame() {
  var win = frameEl.contentWindow;
  if (!win || typeof win.applyTheme !== 'function') return;
  win.applyTheme(currentTheme);
  if (typeof win.render === 'function') win.render();
}
function updateThemeBtnLabel() { themeBtn.textContent = currentTheme === 'dark' ? '☀️ Light Mode' : '🌙 Dark Mode'; }
themeBtn.onclick = function () {
  currentTheme = currentTheme === 'dark' ? 'light' : 'dark';
  try { localStorage.setItem('reportTheme', currentTheme); } catch (e) {}
  updateThemeBtnLabel(); applyShellTheme(); applyThemeToFrame();
};
applyShellTheme();

// Language toggle: relabels the tab bar itself plus, via each iframe's own
// applyLanguage(lang) contract (parallel to applyTheme above), every page's
// titles/headers/buttons/chart text - never the underlying data or column
// headers, and never SHEETS[i].name (the stable id postMessage drilldowns rely on).
function updateLangBtnLabel() { langBtn.textContent = currentLang === 'zh' ? 'EN' : '中文'; }
function relabelTabs() {
  Array.prototype.forEach.call(tabsEl.children, function (btn, i) { btn.textContent = t('tab_' + SHEETS[i].key); });
}
function applyLanguageToFrame() {
  var win = frameEl.contentWindow;
  if (!win || typeof win.applyLanguage !== 'function') return;
  win.applyLanguage(currentLang);
}
langBtn.onclick = function () {
  currentLang = currentLang === 'zh' ? 'en' : 'zh';
  try { localStorage.setItem('reportLang', currentLang); } catch (e) {}
  updateLangBtnLabel(); relabelTabs(); applyLanguageToFrame();
};

// Cross-tab drilldown: a YoY table's clicked cell posts {type:'mtcDrilldown',
// targetTab, filters} up to this shell; switch to that tab and, once its
// iframe finishes loading, hand the filters to its own applyExternalFilter.
var pendingDrilldownFilters = null;
window.addEventListener('message', function (e) {
  if (!e.data || e.data.type !== 'mtcDrilldown') return;
  var idx = SHEETS.findIndex(function (s) { return s.name === e.data.targetTab; });
  if (idx < 0) return;
  pendingDrilldownFilters = e.data.filters;
  showSheet(idx);
});
frameEl.addEventListener('load', function () {
  applyThemeToFrame();
  applyLanguageToFrame();
  var win = frameEl.contentWindow;
  if (pendingDrilldownFilters && win && typeof win.applyExternalFilter === 'function') {
    win.applyExternalFilter(pendingDrilldownFilters);
    pendingDrilldownFilters = null;
  }
});
function showSheet(i) {
  frameEl.srcdoc = SHEETS[i].html;
  Array.prototype.forEach.call(tabsEl.children, function (btn, idx) { btn.className = idx === i ? 'active' : ''; });
}
SHEETS.forEach(function (s, i) {
  var btn = document.createElement('button');
  btn.textContent = t('tab_' + s.key);
  btn.onclick = function () { showSheet(i); };
  tabsEl.appendChild(btn);
});
updateThemeBtnLabel();
updateLangBtnLabel();
showSheet(0);
var introOverlay = document.getElementById('introOverlay');
var introDontShow = document.getElementById('introDontShow');
var introCloseBtn = document.getElementById('introClose');
// Versioned key - bump the _vN suffix whenever the intro's content changes
// materially, so a returning viewer who dismissed an older version sees the
// update once, instead of staying permanently hidden by a stale dismissal.
var introDismissed = false;
try { introDismissed = localStorage.getItem('reportIntroDismissed_v3') === '1'; } catch (e) {}
if (!introDismissed) introOverlay.classList.remove('hidden');
introCloseBtn.onclick = function () {
  if (introDontShow.checked) { try { localStorage.setItem('reportIntroDismissed_v3', '1'); } catch (e) {} }
  introOverlay.classList.add('hidden');
};
</script>
</body>
</html>`;

fs.writeFileSync(OUT_PATH, combinedHtml, 'utf8');
const stat = fs.statSync(OUT_PATH);
console.log('Wrote combined report:', OUT_PATH);
console.log('Total size:', (stat.size / 1024 / 1024).toFixed(2), 'MB');
