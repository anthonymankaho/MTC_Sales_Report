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
    trendSalesCostTitle: 'Monthly Trend: Sales & Cost (USD), 2025 vs 2026', trendMarginTitle: 'Monthly Trend: Margin (USD), 2025 vs 2026',
    seriesSales: 'Sales', seriesCost: 'Cost', seriesMargin: 'Margin',
    legendNetSales: 'Net Sales', legendNetCost: 'Net Cost', legendNetMargin: 'Net Margin',
    vendorTitle: 'Top Vendors by Cost', customerTitle: 'Top Customers by Sales',
    geoTitle: 'Geographic Performance by Margin', originTitle: 'Country of Origin by Sales',
    otherCategory: 'Others',
    totalIndexHeading: 'Total Index',
    weightedLabel: 'weighted',
    yoyLabel: 'YoY', vsLabel: 'vs', varianceLabel: 'Variance', variancePctLabel: 'Variance %',
    yoyChangeCol: '2025 vs 2026 Net Sales % Change',
    introDashboard: 'Filter by year, month, customer, vendor or country - 4 KPIs and 6 charts update instantly.',
    introData: 'The full transaction ledger behind every number in this report - click any highlighted value on another tab and it lands here, pre-filtered.',
    introExRate: 'The Platinum Buy Rate table used to convert every foreign-currency figure in this report to USD.',
    introYoyCustomers: '2026 vs 2025 by customer, with a Total Index above the table - click any name, country or number to see the transactions behind it.',
    introYoyFabrication: 'The same year-over-year comparison as the Customers tab, grouped by Fabrication Code.',
    introYoyProdFamily: 'The same year-over-year comparison as the Customers tab, grouped by Product Family - click + next to a Brand to break it down by Family1.',
    introYoyPartId: 'The same year-over-year comparison as the Customers tab, taken down to individual Part ID.',
    introYoyNetNetMargin: 'The same comparison as the Customers tab, built on the Net Net (further-deducted) basis instead of the invoice-level NET figure.',
    introCustomerQuarter: 'Net margin by customer across the four quarters of 2026, rather than a two-year comparison.',
    langLabel: 'EN'
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
    trendSalesCostTitle: '月度趨勢：銷售額與成本 (USD)，2025 對比 2026', trendMarginTitle: '月度趨勢：毛利 (USD)，2025 對比 2026',
    seriesSales: '銷售額', seriesCost: '成本', seriesMargin: '毛利',
    legendNetSales: 'Net銷售額', legendNetCost: 'Net成本', legendNetMargin: 'Net毛利',
    vendorTitle: '主要供應商（依採購成本）', customerTitle: '主要客戶（依銷售額）',
    geoTitle: '地理績效（依毛利）', originTitle: '原產國（依銷售額）',
    otherCategory: '其他',
    totalIndexHeading: '總計指標',
    weightedLabel: '加權',
    yoyLabel: '較去年同期', vsLabel: '對比', varianceLabel: '差異', variancePctLabel: '差異率',
    yoyChangeCol: '2025對比2026 Net銷售額 % 變動',
    introDashboard: '依年度、月份、客戶、供應商或國家篩選 — 4 項 KPI 與 6 個圖表即時更新。',
    introData: '本報表所有數字背後的完整交易明細 — 點擊其他頁籤中任何標示的數值，都會跳到這裡並自動套用篩選。',
    introExRate: '用於將本報表所有外幣數字換算為 USD 的白金買入匯率表。',
    introYoyCustomers: '依客戶呈現 2026 對比 2025，表格上方附總計指標 — 點擊任何名稱、國家或數字可查看其背後的交易明細。',
    introYoyFabrication: '與客戶頁籤相同的年度比較邏輯，改依加工代碼分組。',
    introYoyProdFamily: '與客戶頁籤相同的年度比較邏輯，改依產品家族分組 - 點擊品牌左方的「+」可展開該品牌各Family1明細。',
    introYoyPartId: '與客戶頁籤相同的年度比較邏輯，細分至個別料號層級。',
    introYoyNetNetMargin: '與客戶頁籤相同的比較邏輯，改以 Net Net（進一步扣除後）基礎計算，而非發票層級 NET 數字。',
    introCustomerQuarter: '呈現 2026 年度內各客戶四個季度的淨毛利變化，而非兩年度比較。',
    langLabel: '中文'
  }
};
var currentLang = 'en';
try { currentLang = localStorage.getItem('reportLang') || 'en'; } catch (e) {}
function t(key) { return (I18N[currentLang] && I18N[currentLang][key]) || I18N.en[key] || key; }
`;

// Shared dark-mode CSS block (token-based, same convention across every tab).
const SHARED_CSS = `
:root {
  --bg: #f4f6f9; --card-bg: #ffffff; --text: #222222; --text-muted: #666666;
  --accent: #1a5c8a; --border: #cccccc; --border-light: #eeeeee; --shadow: rgba(0,0,0,0.1);
  --input-bg: #ffffff; --hover-bg: #f5f5f5;
}
body[data-theme="dark"] {
  --bg: #10161c; --card-bg: #182029; --text: #e7edf0; --text-muted: #a3b3bc;
  --accent: #5fb2c4; --border: #3a4a56; --border-light: #2a3641; --shadow: rgba(0,0,0,0.4);
  --input-bg: #10161c; --hover-bg: #202a33;
}
body { font-family: Calibri, Arial, sans-serif; font-size: 12px; margin: 8px; background: var(--bg); color: var(--text); }
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
  introCustomerQuarter: 'Net margin by customer across the four quarters of 2026, rather than a two-year comparison.'
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
function buildDataTablesHtml(header, rows, title, scrollX, filterColumns, drilldown, titleKey, kpiPie, introKey, yoyRankCol, defaultSortCol, groupBy) {
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
.total-index-year-row { margin-bottom: 10px; }` : ''}</style>
</head>
<body>
${introKey ? `<p class="page-intro" data-i18n="${introKey}">${INTRO_EN[introKey] || ''}</p>` : ''}
${kpiPie ? `
<h3 data-i18n="totalIndexHeading" style="margin:0 0 8px 0; color: var(--accent);">Total Index</h3>
${filterColumns.length ? '<div class="total-index-filters" id="totalIndexFilters"></div>' : ''}
<div class="kpis total-index-year-row" id="totalIndexKpis2026"></div>
<div class="kpis total-index-year-row" id="totalIndexKpis2025"></div>` : ''}
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
const reportTitle = monthValues.length
  ? 'MTC Sales Report - ' + monthValues[0] + (monthValues.length > 1 ? ' to ' + monthValues[monthValues.length - 1] : '')
  : 'MTC Sales Report';

const dashRowsJsonStr = neutralize(JSON.stringify(dashRows));
const yearValuesJsonStr = neutralize(JSON.stringify(yearValues));
const monthValuesJsonStr = neutralize(JSON.stringify(monthValues));
const custValuesJsonStr = neutralize(JSON.stringify(custValues));
const vendValuesJsonStr = neutralize(JSON.stringify(vendValues));
const countryValuesJsonStr = neutralize(JSON.stringify(countryValues));

const dashboardHtml = `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Dashboard</title>
<style>${dataTablesCss}${SHARED_CSS}
  h2 { color: var(--accent); margin-top: 0; }
  /* Overrides the report-wide "canvas { max-height: 320px }" default - each
     trend chart's per-point labels carry a YoY % ("1.8M (-35.8% YoY)"),
     noticeably longer than a bare value, and 16 bars (2025+2026 x 8 months)
     side by side need more vertical room than the standard height allows. */
  #chartSalesCost, #chartMargin { max-height: 420px; }
</style>
</head>
<body>
<h2 data-i18n="dashboardH2">MTC Sales &amp; Margin Performance</h2>
<p class="page-intro" data-i18n="introDashboard">Filter by year, month, customer, vendor or country - 4 KPIs and 5 charts update instantly.</p>
<div class="filters">
  <div><label class="fl" data-i18n="filterYear">Year</label><div class="ms" id="ms-year"></div></div>
  <div><label class="fl" data-i18n="filterMonth">Month</label><div class="ms" id="ms-month"></div></div>
  <div><label class="fl" data-i18n="filterCustomer">Customer</label><div class="ms" id="ms-cust"></div></div>
  <div><label class="fl" data-i18n="filterVendor">Vendor</label><div class="ms" id="ms-vend"></div></div>
  <div><label class="fl" data-i18n="filterCountry">Country</label><div class="ms" id="ms-country"></div></div>
  <div style="align-self:flex-end"><button id="resetBtn" data-i18n="resetFilters">Reset filters</button></div>
</div>
<div class="kpis">
  <div class="kpi"><div class="label" data-i18n="kpiNetSales">Net Sales (USD)</div><div class="value" id="kpi-sales">-</div></div>
  <div class="kpi"><div class="label" data-i18n="kpiNetCost">Net Cost (USD)</div><div class="value" id="kpi-cost">-</div></div>
  <div class="kpi"><div class="label" data-i18n="kpiNetMargin">Net Margin (USD)</div><div class="value" id="kpi-margin">-</div></div>
  <div class="kpi"><div class="label" data-i18n="kpiNetMarginPct">Net Margin %</div><div class="value" id="kpi-marginpct">-</div></div>
</div>
<div class="charts">
  <div class="chart-box wide">
    <canvas id="chartSalesCost"></canvas>
  </div>
</div>
<div class="charts">
  <div class="chart-box wide">
    <canvas id="chartMargin"></canvas>
  </div>
</div>
<div class="charts">
  <div class="chart-box">
    <canvas id="chartVendor"></canvas>
  </div>
  <div class="chart-box">
    <canvas id="chartCustomer"></canvas>
  </div>
</div>
<div class="charts">
  <div class="chart-box">
    <canvas id="chartGeo"></canvas>
  </div>
  <div class="chart-box">
    <canvas id="chartOrigin"></canvas>
  </div>
</div>
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
// Fixed categorical colors (validated CVD-safe trio, same as the earlier MTC
// executive infographic): Sales=navy, Cost=copper, Margin=teal.
var C_SALES = '#1a5c8a', C_COST = '#b96a2c', C_MARGIN = '#1c7a63';
var CATEGORY_PALETTE = ['#1a5c8a', '#b96a2c', '#1c7a63', '#7a4fa3', '#c94f4f', '#4f7ac9', '#9c8a1a', '#1a9c8a', '#8a1a5c', '#5c8a1a', '#4f9ec9', '#c98a4f'];

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

function renderKpis() {
  var rows = lastRows;
  var sales = 0, cost = 0, margin = 0;
  rows.forEach(function (r) { sales += r[5]; cost += r[6]; margin += r[7]; });
  document.getElementById('kpi-sales').textContent = fmt(sales);
  document.getElementById('kpi-cost').textContent = fmt(cost);
  document.getElementById('kpi-margin').textContent = fmt(margin);
  document.getElementById('kpi-marginpct').textContent = sales ? (margin / sales * 100).toFixed(1) + '%' : '-';
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
  renderSalesCostTrend();
  renderMarginTrend();
  renderVendor();
  renderCustomer();
  renderGeo();
  renderOrigin();
}

function applyTheme(theme) {
  document.body.setAttribute('data-theme', theme);
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
</body>
</html>`;
console.log('Dashboard built (' + (dashboardHtml.length / 1024 / 1024).toFixed(2) + ' MB)');

// =========================================================================
// Data sheet (full detail table)
// =========================================================================
const dataSheetHtml = buildDataTablesHtml(H, dataSheet.rows, 'Data', true,
  // The last 4 columns here exist so a YoY-tab drilldown click always has a
  // matching toolbar filter/dropdown to land on and visibly show as applied.
  ['Cust_name', 'VEND_NAME', 'Cust_Countries', 'yyyy-mm', 'Year', 'Country of origin', 'Fabrication Code', 'Family1', 'Part_id'],
  null, 'title_data', false, 'introData');
console.log('Data sheet built (' + (dataSheetHtml.length / 1024 / 1024).toFixed(2) + ' MB)');

// =========================================================================
// Exchange Rate (USD table only)
// =========================================================================
const exRateData = readJson('exrate_usd.json');
const exRateHtml = buildDataTablesHtml(exRateData.header, exRateData.rows, 'Exchange Rate', false, null, null, 'title_exrate', false, 'introExRate');

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
  { labelColCount: 1, dataColMap: { 'Cust_name': 'Cust_name' } }, 'title_yoyCustomers', true, 'introYoyCustomers', true);
const yoyFabricationHtml = buildDataTablesHtml(yoyFabrication.header, yoyFabrication.rows, '2026 v 2025 by Fabrication', false, ['Fabrication Code'],
  { labelColCount: 4, dataColMap: { 'Fabrication Code': 'Fabrication Code', 'Part_id': 'Part_id' } }, 'title_yoyFabrication', true, 'introYoyFabrication', true);
const yoyProdFamilyHtml = buildDataTablesHtml(yoyProdFamily.header, yoyProdFamily.rows, '2026 v 2025 by Prod Family', false, ['Brand', 'Family1'],
  { labelColCount: 3, dataColMap: { 'Brand': 'Family Category', 'Family1': 'Family1' } }, 'title_yoyProdFamily', true, 'introYoyProdFamily', true, null,
  { groupColIdx: 0, toggleColIdx: 1, detailKeyColIdx: 2 });
const yoyPartIdHtml = buildDataTablesHtml(yoyPartId.header, yoyPartId.rows, '2026 v 2025 by Part ID', false, ['Part_id'],
  { labelColCount: 2, dataColMap: { 'Part_id': 'Part_id' } }, 'title_yoyPartId', true, 'introYoyPartId', true);
const yoyNetNetMarginHtml = buildDataTablesHtml(yoyNetNetMargin.header, yoyNetNetMargin.rows, '2026 v 2025 by NET NET MARGIN', false, ['Cust_name', 'Country of origin'],
  { labelColCount: 3, dataColMap: { 'Cust_name': 'Cust_name', 'Country of origin': 'Country of origin', 'Part_id': 'Part_id' } }, 'title_yoyNetNetMargin', true, 'introYoyNetNetMargin', true);
const netMarginByCustomerHtml = buildDataTablesHtml(netMarginByCustomer.header, netMarginByCustomer.rows, 'Net Margin by Customer-Quarter', false, ['Cust_name'],
  { labelColCount: 1, dataColMap: { 'Cust_name': 'Cust_name' }, fixedYear: netMarginByCustomer.year }, 'title_customerQuarter', false, 'introCustomerQuarter',
  false, 'Net_Sales_USD (Data_AX)');

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
  return { key: 'purchaseYtdVendor', name: 'Purchase: YTD by Vendor', html: tab.html };
}
const purchaseYtdTab = extractPurchaseYtdTab();
if (purchaseYtdTab) { sheetsData.push(purchaseYtdTab); console.log('Added Purchase: YTD by Vendor tab from', PURCHASE_HTML_PATH); }
const sheetsJsonStr = neutralize(JSON.stringify(sheetsData));

const combinedHtml = `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>${reportTitle}</title>
<style>
  body { font-family: Calibri, Arial, sans-serif; margin: 0; padding: 0; }
  #tabbar { display: flex; background: #1a5c8a; align-items: stretch; }
  #tabs { display: flex; padding: 0; margin: 0; list-style: none; flex-wrap: wrap; flex: 1; }
  #tabs button { background: #1a5c8a; color: white; border: none; padding: 10px 16px; cursor: pointer; font-size: 13px; }
  #tabs button.active { background: #123f5e; font-weight: bold; }
  #tabs button:hover { background: #123f5e; }
  #themeToggle, #langToggle { background: #123f5e; color: white; border: none; padding: 10px 16px; cursor: pointer; font-size: 13px; white-space: nowrap; }
  #themeToggle:hover, #langToggle:hover { background: #0c2c42; }
  #langToggle { border-left: 1px solid rgba(255,255,255,0.2); }
  iframe { width: 100%; height: calc(100vh - 44px); border: none; display: block; }
  #introOverlay { position: fixed; inset: 0; background: rgba(0,0,0,0.55); display: flex; align-items: center; justify-content: center; z-index: 1000; }
  #introOverlay.hidden { display: none; }
  #introModal { background: #ffffff; color: #222; border-radius: 10px; box-shadow: 0 8px 30px rgba(0,0,0,0.3); max-width: 560px; width: 90%; max-height: 85vh; overflow-y: auto; padding: 24px 28px; }
  #introModal h2 { color: #1a5c8a; margin: 0 0 4px 0; font-size: 20px; }
  #introModal .intro-sub { color: #666; font-size: 13px; margin: 0 0 16px 0; }
  #introModal ul { margin: 0 0 18px 0; padding-left: 20px; font-size: 14px; line-height: 1.7; }
  #introModal ul + p.intro-note { font-size: 12px; color: #777; margin: 0 0 18px 0; }
  #introFooter { display: flex; align-items: center; justify-content: space-between; gap: 12px; flex-wrap: wrap; }
  #introFooter label { font-size: 13px; color: #555; display: flex; align-items: center; gap: 6px; cursor: pointer; }
  #introClose { background: #1a5c8a; color: #fff; border: none; border-radius: 6px; padding: 9px 20px; font-size: 14px; cursor: pointer; }
  #introClose:hover { background: #123f5e; }
</style>
</head>
<body>
<div id="introOverlay" class="hidden">
  <div id="introModal">
    <h2>👋 Welcome / 歡迎使用</h2>
    <p class="intro-sub">MTC Sales Report — Interactive Web Version / 互動式網頁版</p>
    <ul>
      <li><b>Live filters / 即時篩選</b> — Year, Month, Customer, Vendor, Country on the Dashboard tab — charts and KPIs update together instantly.<br>Dashboard 頁可依年度、月份、客戶、供應商、國家篩選，圖表與 KPI 同步更新。</li>
      <li><b>Sortable tables + CSV export / 明細表排序與匯出</b> — click a column header to sort; "Export CSV" downloads exactly what's currently filtered.<br>點欄位標題可排序；「Export CSV」只會匯出目前篩選畫面的內容。</li>
      <li><b>Data tab</b> — the full transaction-level DATA sheet, with quick filters for Customer, Vendor, Country, Month, Year.<br>Data 頁為完整交易明細，可快速篩選客戶、供應商、國家、月份、年度。</li>
      <li><b>YoY comparison tabs</b> — the five "2026 v 2025 by ..." sheets, exactly as shown in the Excel report.<br>五個「2026 v 2025 by...」年度比較分析頁，與 Excel 報表內容一致。</li>
      <li><b>🌙 Light/Dark mode</b> — toggle top-right.<br>右上角可切換淡色/深色模式。</li>
    </ul>
    <p class="intro-note">This is the same data as the official Excel report — a convenience for browsing, not a replacement. / 資料與正式 Excel 報表一致，本頁僅為方便瀏覽之用，非取代 Excel 報表。</p>
    <div id="introFooter">
      <label><input type="checkbox" id="introDontShow"> Don't show this again / 下次不再顯示</label>
      <button id="introClose">Get Started / 開始使用</button>
    </div>
  </div>
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
  updateThemeBtnLabel(); applyThemeToFrame();
};

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
try { introDismissed = localStorage.getItem('reportIntroDismissed_v1') === '1'; } catch (e) {}
if (!introDismissed) introOverlay.classList.remove('hidden');
introCloseBtn.onclick = function () {
  if (introDontShow.checked) { try { localStorage.setItem('reportIntroDismissed_v1', '1'); } catch (e) {} }
  introOverlay.classList.add('hidden');
};
</script>
</body>
</html>`;

fs.writeFileSync(OUT_PATH, combinedHtml, 'utf8');
const stat = fs.statSync(OUT_PATH);
console.log('Wrote combined report:', OUT_PATH);
console.log('Total size:', (stat.size / 1024 / 1024).toFixed(2), 'MB');
