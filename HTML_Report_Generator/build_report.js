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
    tab_yoyNetNetMargin: 'YoY: Net Net Margin', tab_customerQuarter: 'Customer-Quarter',
    title_data: 'Data', title_exrate: 'Exchange Rate',
    title_yoyCustomers: '2026 v 2025 by Customers', title_yoyFabrication: '2026 v 2025 by Fabrication',
    title_yoyProdFamily: '2026 v 2025 by Prod Family', title_yoyPartId: '2026 v 2025 by Part ID',
    title_yoyNetNetMargin: '2026 v 2025 by NET NET MARGIN', title_customerQuarter: 'Net Margin by Customer-Quarter',
    rowsSuffix: 'rows', exportCsv: 'Export CSV (filtered)', clearFilters: 'Clear all filters / search',
    clearFiltersTooltip: 'Clear search box + all dropdown filters',
    drillTooltip: "Click to see the underlying transactions in the Data tab",
    dashboardH2: 'MTC Sales & Margin Performance',
    filterYear: 'Year', filterMonth: 'Month', filterCustomer: 'Customer', filterVendor: 'Vendor',
    filterCountry: 'Country', resetFilters: 'Reset filters',
    netSectionLabel: 'NET (invoice basis)', netnetSectionLabel: 'NET NET (further-deducted basis)',
    kpiNetSales: 'Net Sales (USD)', kpiNetCost: 'Net Cost (USD)', kpiNetMargin: 'Net Margin (USD)', kpiNetMarginPct: 'Net Margin %',
    kpiNNSales: 'Net Net Sales (USD)', kpiNNCost: 'Net Net Cost (USD)', kpiNNMargin: 'Net Net Margin (USD)', kpiNNMarginPct: 'Net Net Margin %',
    basisNet: 'NET', basisNetNet: 'NET NET', basisAll: 'ALL',
    trendTitle: 'Monthly Trend', trendSeries: 'Sales / Cost / Margin (USD)',
    seriesSales: 'Sales', seriesCost: 'Cost', seriesMargin: 'Margin',
    seriesNNSales: 'Net Net Sales', seriesNNCost: 'Net Net Cost', seriesNNMargin: 'Net Net Margin',
    legendNetSales: 'Net Sales', legendNetNetSales: 'Net Net Sales',
    legendNetMargin: 'Net Margin', legendNetNetMargin: 'Net Net Margin',
    vendorTitle: 'Top Vendors by Sales', customerTitle: 'Top Customers by Sales',
    geoTitle: 'Geographic Performance by Margin', originTitle: 'Country of Origin by Sales',
    otherCategory: 'Others',
    langLabel: 'EN'
  },
  zh: {
    tab_dashboard: '總覽', tab_data: '明細資料', tab_exrate: '匯率',
    tab_yoyCustomers: '年度比較：客戶', tab_yoyFabrication: '年度比較：加工類別',
    tab_yoyProdFamily: '年度比較：產品家族', tab_yoyPartId: '年度比較：料號',
    tab_yoyNetNetMargin: '年度比較：淨淨毛利', tab_customerQuarter: '客戶季度毛利',
    title_data: '明細資料', title_exrate: '匯率',
    title_yoyCustomers: '2026 對比 2025：依客戶', title_yoyFabrication: '2026 對比 2025：依加工類別',
    title_yoyProdFamily: '2026 對比 2025：依產品家族', title_yoyPartId: '2026 對比 2025：依料號',
    title_yoyNetNetMargin: '2026 對比 2025：依淨淨毛利', title_customerQuarter: '客戶季度淨毛利',
    rowsSuffix: '筆', exportCsv: '匯出CSV（僅目前篩選結果）', clearFilters: '清除所有篩選／搜尋',
    clearFiltersTooltip: '清除搜尋框及所有下拉篩選',
    drillTooltip: '點擊查看Data頁籤中的相關交易明細',
    dashboardH2: 'MTC 銷售與毛利績效',
    filterYear: '年度', filterMonth: '月份', filterCustomer: '客戶', filterVendor: '供應商',
    filterCountry: '國家', resetFilters: '重設篩選',
    netSectionLabel: 'NET（發票基礎）', netnetSectionLabel: 'NET NET（進一步扣除後基礎）',
    kpiNetSales: 'Net銷售額 (USD)', kpiNetCost: 'Net成本 (USD)', kpiNetMargin: 'Net毛利 (USD)', kpiNetMarginPct: 'Net毛利率',
    kpiNNSales: 'Net Net銷售額 (USD)', kpiNNCost: 'Net Net成本 (USD)', kpiNNMargin: 'Net Net毛利 (USD)', kpiNNMarginPct: 'Net Net毛利率',
    basisNet: 'NET', basisNetNet: 'NET NET', basisAll: '全部',
    trendTitle: '月度趨勢', trendSeries: '銷售額／成本／毛利 (USD)',
    seriesSales: '銷售額', seriesCost: '成本', seriesMargin: '毛利',
    seriesNNSales: 'Net Net銷售額', seriesNNCost: 'Net Net成本', seriesNNMargin: 'Net Net毛利',
    legendNetSales: 'Net銷售額', legendNetNetSales: 'Net Net銷售額',
    legendNetMargin: 'Net毛利', legendNetNetMargin: 'Net Net毛利',
    vendorTitle: '主要供應商（依銷售額）', customerTitle: '主要客戶（依銷售額）',
    geoTitle: '地理績效（依毛利）', originTitle: '原產國（依銷售額）',
    otherCategory: '其他',
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
#exportBtn, #clearFiltersBtn, #resetBtn { padding: 6px 14px; cursor: pointer; border-radius: 4px; border: 1px solid var(--border); background: var(--input-bg); color: var(--text); }
.filter-group { display: flex; align-items: center; gap: 4px; font-size: 11px; color: var(--text-muted); }
.filter-group select { font-size: 11px; padding: 3px 4px; border-radius: 3px; border: 1px solid var(--border); background: var(--input-bg); color: var(--text); max-width: 160px; }
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
.kpi-section-label { font-size: 12px; font-weight: 700; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.04em; margin: 4px 0 6px; }
.kpis { display: flex; gap: 16px; margin-bottom: 16px; flex-wrap: wrap; }
.kpi { background: var(--card-bg); padding: 16px 24px; border-radius: 6px; box-shadow: 0 1px 3px var(--shadow); flex: 1; min-width: 180px; }
.kpi .label { font-size: 12px; color: var(--text-muted); }
.kpi .value { font-size: 24px; font-weight: bold; color: var(--accent); }
.charts { display: flex; flex-wrap: wrap; gap: 16px; margin-bottom: 16px; }
.chart-box { background: var(--card-bg); padding: 12px; border-radius: 6px; box-shadow: 0 1px 3px var(--shadow); flex: 1; min-width: 420px; }
.chart-box.wide { flex-basis: 100%; }
canvas { max-height: 320px; }
.chart-toggle { display: flex; justify-content: flex-end; gap: 4px; margin-bottom: 4px; }
.chart-toggle button { font-size: 11px; padding: 4px 10px; border-radius: 4px; border: 1px solid var(--border); background: var(--input-bg); color: var(--text-muted); cursor: pointer; }
.chart-toggle button.active { background: var(--accent); color: #fff; border-color: var(--accent); font-weight: 600; }
.filters { display: flex; gap: 16px; flex-wrap: wrap; background: var(--card-bg); padding: 12px; border-radius: 6px; margin-bottom: 16px; box-shadow: 0 1px 3px var(--shadow); }
.filters label.fl { display: block; font-size: 11px; color: var(--text-muted); margin-bottom: 4px; }
.ms { position: relative; width: 180px; }
.ms-toggle { width: 100%; text-align: left; padding: 5px 8px; background: var(--input-bg); color: var(--text); border: 1px solid var(--border); border-radius: 4px; cursor: pointer; font-size: 12px; }
.ms-panel { position: absolute; z-index: 20; top: 100%; left: 0; background: var(--card-bg); border: 1px solid var(--border); border-radius: 4px; box-shadow: 0 2px 8px var(--shadow); max-height: 260px; overflow-y: auto; width: 240px; padding: 6px; }
.ms-panel label { display: block; font-size: 12px; padding: 2px 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: var(--text); }
.ms-panel label.ms-all-row { border-bottom: 1px solid var(--border-light); margin-bottom: 4px; padding-bottom: 6px; font-weight: bold; }
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
  var allRow = document.createElement('label');
  allRow.className = 'ms-all-row';
  var allCb = document.createElement('input'); allCb.type = 'checkbox'; allCb.checked = startAll;
  allRow.appendChild(allCb); allRow.appendChild(document.createTextNode(' (Select All)'));
  panel.appendChild(allRow);
  var checkboxes = [];
  values.forEach(function (v) {
    var row = document.createElement('label');
    var cb = document.createElement('input'); cb.type = 'checkbox'; cb.value = v;
    cb.checked = startAll || selected.has(v);
    row.appendChild(cb); row.appendChild(document.createTextNode(' ' + v));
    panel.appendChild(row); checkboxes.push(cb);
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
    panel.style.display = panel.style.display === 'none' ? 'block' : 'none';
  });
  updateToggleText();
  return {
    getSelected: function () { return Array.from(selected); },
    reset: function () { selected = new Set(); allCb.checked = true; checkboxes.forEach(function (cb) { cb.checked = true; }); updateToggleText(); }
  };
}
document.addEventListener('click', function () { document.querySelectorAll('.ms-panel').forEach(function (p) { p.style.display = 'none'; }); });
function matchesFilter(value, selectedArr) { return selectedArr.length === 0 || selectedArr.indexOf(value) !== -1; }

// NET / NET NET / ALL toggle, inserted into every Dashboard chart so each one
// can independently show the invoice-basis figures, the further-deducted
// "Net Net" figures, or both together.
function createBasisToggle(containerId, onChange, initial) {
  var container = document.getElementById(containerId);
  var current = initial || 'NET';
  var options = ['NET', 'NET NET', 'ALL'];
  // Internal comparison/state always stays on the English value (opt) - only
  // the displayed textContent is translated, so a language switch can relabel
  // buttons without disturbing which one is "active" or what onChange receives.
  var keyFor = { 'NET': 'basisNet', 'NET NET': 'basisNetNet', 'ALL': 'basisAll' };
  var buttons = options.map(function (opt) {
    var btn = document.createElement('button');
    btn.type = 'button'; btn.textContent = t(keyFor[opt]); btn.dataset.value = opt;
    if (opt === current) btn.className = 'active';
    btn.addEventListener('click', function () {
      current = opt;
      buttons.forEach(function (b) { b.className = b.dataset.value === current ? 'active' : ''; });
      onChange(current);
    });
    container.appendChild(btn);
    return btn;
  });
  return {
    get: function () { return current; },
    relabel: function () { buttons.forEach(function (b) { b.textContent = t(keyFor[b.dataset.value]); }); }
  };
}
`;

// =========================================================================
// Generic DataTables page for a (possibly filtered) set of rows - used for
// Data, Exchange Rate, and every YoY comparison sheet.
// =========================================================================
function buildDataTablesHtml(header, rows, title, scrollX, filterColumns, drilldown, titleKey) {
  if (scrollX === undefined) scrollX = true;
  if (!filterColumns) filterColumns = [];
  const headerJsonStr = neutralize(JSON.stringify(header));
  const rowsJsonStr = neutralize(JSON.stringify(rows));
  const filterColIdx = filterColumns.map(name => header.indexOf(name)).filter(i => i >= 0);
  const filterColIdxJsonStr = JSON.stringify(filterColIdx);
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
td.dt-num, th.dt-num { text-align: right !important; }</style>
</head>
<body>
<div class="summary-bar" id="summaryBar" data-title-key="${titleKey || ''}" data-count-str="${rowCountStr}">${title} - ${rowCountStr} rows</div>
<div class="toolbar">
  <button id="exportBtn" data-i18n="exportCsv">Export CSV (filtered)</button>
  <button id="clearFiltersBtn" data-i18n="clearFilters" data-i18n-title="clearFiltersTooltip" title="Clear search box + all dropdown filters">Clear all filters / search</button>
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
  var sb = document.getElementById('summaryBar');
  if (sb && sb.dataset.titleKey) sb.textContent = t(sb.dataset.titleKey) + ' - ' + sb.dataset.countStr + ' ' + t('rowsSuffix');
  if (table) table.draw(false);
}
function regexEscape(s) {
  var specials = ['\\\\\\\\', '.', '*', '+', '?', '^', '$', '(', ')', '|', '[', ']', '{', '}'];
  var str = String(s);
  specials.forEach(function (c) { str = str.split(c).join('\\\\\\\\' + c); });
  return str;
}
// Sets each named column's filter (toolbar dropdown, if one exists for it,
// plus the underlying DataTables search either way) - used both by a normal
// dropdown change and by an incoming cross-tab drilldown request.
var table;
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
    var sel = document.querySelector('#toolbarFilters select[data-col-index="' + colIdx + '"]');
    if (sel) sel.value = val;
  });
  table.draw();
}
// Right-align number-like columns for readability. The Data tab stores
// measures as real JS numbers, but the YoY/Customer-Quarter sheets keep
// Excel's own formatted display text ("6,528", "22.4%", "(1,234)") - typeof
// alone would miss those, so numeric-ness is judged from a sample of each
// column's actual values instead of the column's declared type.
var NUMERIC_RE = /^\\(?-?[\\d,]+(\\.\\d+)?%?\\)?$/;
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
$(function () {
  var headRow = document.getElementById('dataTableHeadRow');
  HEADER.forEach(function (h, i) {
    var th = document.createElement('th');
    th.textContent = h;
    if (NUMERIC_COL_IDX.indexOf(i) >= 0) th.className = 'dt-num';
    headRow.appendChild(th);
  });
  var toolbarFilters = document.getElementById('toolbarFilters');
  var columnDefs = [{ targets: NUMERIC_COL_IDX, className: 'dt-num' }];
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
  table = $('#dataTable').DataTable({
    data: ROWS, deferRender: true, pageLength: 50, lengthMenu: [25, 50, 100, 500, 1000],
    scrollX: ${scrollX ? 'true' : 'false'}, autoWidth: false,
    columnDefs: columnDefs,
    initComplete: function () {
      var api = this.api();
      // DataTables calls initComplete DURING construction, before the outer
      // "table = $(...).DataTable(...)" assignment above has actually
      // completed - assign it here too so applyExternalFilter (which checks
      // the table variable) works correctly even when called from in here.
      table = api;
      FILTER_COL_IDX.forEach(function (colIdx) {
        var group = document.createElement('span'); group.className = 'filter-group';
        var label = document.createElement('label'); label.textContent = HEADER[colIdx] + ':';
        var sel = document.createElement('select'); sel.dataset.colIndex = colIdx;
        var optAll = document.createElement('option'); optAll.value = ''; optAll.textContent = 'All'; sel.appendChild(optAll);
        var values = api.column(colIdx).data().toArray().map(function (v) { return (v === null || v === undefined) ? '' : String(v); });
        var unique = values.filter(function (v, i) { return v !== '' && values.indexOf(v) === i; }).sort();
        unique.forEach(function (v) { var opt = document.createElement('option'); opt.value = v; opt.textContent = v; sel.appendChild(opt); });
        sel.addEventListener('change', function () {
          var val = sel.value; var column = api.column(colIdx);
          if (val === '') column.search(''); else column.search('^' + regexEscape(val) + '\$', true, false);
          api.draw();
        });
        group.appendChild(label); group.appendChild(sel); toolbarFilters.appendChild(group);
      });
      // A pending drilldown may have arrived before this table finished
      // initializing (the parent shell queues it) - apply it now.
      if (window.__pendingDrilldownFilters) {
        applyExternalFilter(window.__pendingDrilldownFilters);
        window.__pendingDrilldownFilters = null;
      }
    }
  });
  document.getElementById('clearFiltersBtn').onclick = function () {
    table.search('');
    FILTER_COL_IDX.forEach(function (colIdx) {
      table.column(colIdx).search('');
      var sel = toolbarFilters.querySelector('select[data-col-index="' + colIdx + '"]');
      if (sel) sel.value = '';
    });
    table.draw();
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
const costIdx = COL['Ttl PO USD'];
const marginIdx = COL['TTL PO NET CONT USD'];
// "Net Net" = a second, more-deducted sales/margin basis (see the "2026 v 2025
// by NET NET MARGIN" sheet) - cost has no separate Net Net figure, the same
// Ttl PO USD basis is used for both.
const nnSalesIdx = COL['Total Net Net Sales (USD)'];
const nnMarginIdx = COL['Total Net Net Margin (USD)'];

const dashRows = dataSheet.rows.map(r => [
  typeof r[yearIdx] === 'number' ? String(r[yearIdx]) : r[yearIdx],
  r[monthIdx], r[custIdx], r[vendIdx], r[countryIdx],
  typeof r[salesIdx] === 'number' ? r[salesIdx] : 0,
  typeof r[costIdx] === 'number' ? r[costIdx] : 0,
  typeof r[marginIdx] === 'number' ? r[marginIdx] : 0,
  r[originIdx],
  typeof r[nnSalesIdx] === 'number' ? r[nnSalesIdx] : 0,
  typeof r[nnMarginIdx] === 'number' ? r[nnMarginIdx] : 0
]);
// dashRows columns: 0=Year,1=Month,2=Customer,3=Vendor,4=Country(customer),5=Sales,
// 6=Cost,7=Margin,8=Country of origin (sourcing),9=NetNetSales,10=NetNetMargin

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
</style>
</head>
<body>
<h2 data-i18n="dashboardH2">MTC Sales &amp; Margin Performance</h2>
<div class="filters">
  <div><label class="fl" data-i18n="filterYear">Year</label><div class="ms" id="ms-year"></div></div>
  <div><label class="fl" data-i18n="filterMonth">Month</label><div class="ms" id="ms-month"></div></div>
  <div><label class="fl" data-i18n="filterCustomer">Customer</label><div class="ms" id="ms-cust"></div></div>
  <div><label class="fl" data-i18n="filterVendor">Vendor</label><div class="ms" id="ms-vend"></div></div>
  <div><label class="fl" data-i18n="filterCountry">Country</label><div class="ms" id="ms-country"></div></div>
  <div style="align-self:flex-end"><button id="resetBtn" data-i18n="resetFilters">Reset filters</button></div>
</div>
<div class="kpi-section-label" data-i18n="netSectionLabel">NET (invoice basis)</div>
<div class="kpis">
  <div class="kpi"><div class="label" data-i18n="kpiNetSales">Net Sales (USD)</div><div class="value" id="kpi-sales">-</div></div>
  <div class="kpi"><div class="label" data-i18n="kpiNetCost">Net Cost (USD)</div><div class="value" id="kpi-cost">-</div></div>
  <div class="kpi"><div class="label" data-i18n="kpiNetMargin">Net Margin (USD)</div><div class="value" id="kpi-margin">-</div></div>
  <div class="kpi"><div class="label" data-i18n="kpiNetMarginPct">Net Margin %</div><div class="value" id="kpi-marginpct">-</div></div>
</div>
<div class="kpi-section-label" data-i18n="netnetSectionLabel">NET NET (further-deducted basis)</div>
<div class="kpis">
  <div class="kpi"><div class="label" data-i18n="kpiNNSales">Net Net Sales (USD)</div><div class="value" id="kpi-nnsales">-</div></div>
  <div class="kpi"><div class="label" data-i18n="kpiNNCost">Net Net Cost (USD)</div><div class="value" id="kpi-nncost">-</div></div>
  <div class="kpi"><div class="label" data-i18n="kpiNNMargin">Net Net Margin (USD)</div><div class="value" id="kpi-nnmargin">-</div></div>
  <div class="kpi"><div class="label" data-i18n="kpiNNMarginPct">Net Net Margin %</div><div class="value" id="kpi-nnmarginpct">-</div></div>
</div>
<div class="charts">
  <div class="chart-box wide">
    <div class="chart-toggle" id="tog-trend"></div>
    <canvas id="chartTrend"></canvas>
  </div>
</div>
<div class="charts">
  <div class="chart-box">
    <div class="chart-toggle" id="tog-vendor"></div>
    <canvas id="chartVendor"></canvas>
  </div>
  <div class="chart-box">
    <div class="chart-toggle" id="tog-customer"></div>
    <canvas id="chartCustomer"></canvas>
  </div>
</div>
<div class="charts">
  <div class="chart-box">
    <div class="chart-toggle" id="tog-geo"></div>
    <canvas id="chartGeo"></canvas>
  </div>
  <div class="chart-box">
    <div class="chart-toggle" id="tog-origin"></div>
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
// Translates a chart's NET/NET NET/ALL basis word for display in a title -
// the underlying comparison value (opts.basis / togXxx.get()) always stays English.
function basisLabel(basis) {
  return basis === 'NET NET' ? t('basisNetNet') : (basis === 'ALL' ? t('basisAll') : t('basisNet'));
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
        if (opt.showPct !== false && total) label += ' (' + pct.toFixed(1) + '%)';
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
function fmt(n) { return Math.round(n).toLocaleString(); }
function withLabelHeadroom(values) {
  var max = values.reduce(function (m, v) { return Math.max(m, v); }, 0);
  return { x: { suggestedMax: max * 1.35 } };
}
var C_NET = C_SALES, C_NETNET = '#7a4fa3'; // distinct color for the Net Net series/dataset in ALL mode

var charts = {};
var lastRows = [];

// ---- KPIs: both bases always shown together (the 8-card requirement),
// independent of any chart's own NET/NET NET/ALL toggle. ----
function renderKpis() {
  var rows = lastRows;
  var sales = 0, cost = 0, margin = 0, nnSales = 0, nnMargin = 0;
  rows.forEach(function (r) { sales += r[5]; cost += r[6]; margin += r[7]; nnSales += r[9]; nnMargin += r[10]; });
  document.getElementById('kpi-sales').textContent = fmt(sales);
  document.getElementById('kpi-cost').textContent = fmt(cost);
  document.getElementById('kpi-margin').textContent = fmt(margin);
  document.getElementById('kpi-marginpct').textContent = sales ? (margin / sales * 100).toFixed(1) + '%' : '-';
  document.getElementById('kpi-nnsales').textContent = fmt(nnSales);
  document.getElementById('kpi-nncost').textContent = fmt(cost);
  document.getElementById('kpi-nnmargin').textContent = fmt(nnMargin);
  document.getElementById('kpi-nnmarginpct').textContent = nnSales ? (nnMargin / nnSales * 100).toFixed(1) + '%' : '-';
}

// ---- Monthly Trend: NET = Sales/Cost/Margin, NET NET = Net Net Sales/Cost/
// Net Net Margin (dashed, to stay visually distinct when ALL overlays both). ----
function renderTrend() {
  var rows = lastRows, basis = togTrend.get();
  var months = Array.from(new Set(rows.map(function (r) { return r[1]; }))).filter(Boolean).sort();
  var salesByMonth = groupSum(rows, 1, 5), costByMonth = groupSum(rows, 1, 6), marginByMonth = groupSum(rows, 1, 7);
  var nnSalesByMonth = groupSum(rows, 1, 9), nnMarginByMonth = groupSum(rows, 1, 10);
  var datasets = [];
  if (basis === 'NET' || basis === 'ALL') {
    datasets.push({ label: t('seriesSales'), data: months.map(function (m) { return salesByMonth[m] || 0; }), borderColor: C_SALES, backgroundColor: C_SALES, tension: 0.2, pointRadius: 3 });
    datasets.push({ label: t('seriesCost'), data: months.map(function (m) { return costByMonth[m] || 0; }), borderColor: C_COST, backgroundColor: C_COST, tension: 0.2, pointRadius: 3 });
    datasets.push({ label: t('seriesMargin'), data: months.map(function (m) { return marginByMonth[m] || 0; }), borderColor: C_MARGIN, backgroundColor: C_MARGIN, tension: 0.2, pointRadius: 3 });
  }
  if (basis === 'NET NET' || basis === 'ALL') {
    datasets.push({ label: t('seriesNNSales'), data: months.map(function (m) { return nnSalesByMonth[m] || 0; }), borderColor: C_SALES, backgroundColor: C_SALES, borderDash: [5, 3], tension: 0.2, pointRadius: 3 });
    datasets.push({ label: t('seriesNNCost'), data: months.map(function (m) { return costByMonth[m] || 0; }), borderColor: C_COST, backgroundColor: C_COST, borderDash: [5, 3], tension: 0.2, pointRadius: 3 });
    datasets.push({ label: t('seriesNNMargin'), data: months.map(function (m) { return nnMarginByMonth[m] || 0; }), borderColor: C_MARGIN, backgroundColor: C_MARGIN, borderDash: [5, 3], tension: 0.2, pointRadius: 3 });
  }
  if (charts.trend) charts.trend.destroy();
  charts.trend = new Chart(document.getElementById('chartTrend'), {
    type: 'line',
    data: { labels: months, datasets: datasets },
    options: {
      responsive: true, interaction: { mode: 'index', intersect: false },
      plugins: {
        title: { display: true, text: t('trendTitle') + ' (' + basisLabel(basis) + '): ' + t('trendSeries') }, legend: { position: 'bottom' },
        // Value only here (no %): "this month's sales as a % of N months of
        // sales" isn't a meaningful per-point read the way it is for a bar/pie
        // category share. dyPerDataset spaces each series' labels vertically
        // so they don't collide on the same month - tighter when ALL doubles
        // the series count to 6.
        valuePctLabels: { enabled: true, showPct: false, fontSize: 9, dy: -12, dyPerDataset: basis === 'ALL' ? -9 : -11 }
      }
    }
  });
}

// ---- Vendor / Customer / Origin: single value basis (Sales), NET vs NET NET
// vs a grouped ALL (two bars/rings per category). ----
function renderCategoryChart(opts) {
  // opts: { chartKey, elId, groupKeyIdx, basis, title, color, sortTop, type, indexAxis }
  var rows = lastRows;
  var netSum = groupSum(rows, opts.groupKeyIdx, 5);
  var nnSum = groupSum(rows, opts.groupKeyIdx, 9);
  var basisSum = opts.basis === 'NET NET' ? nnSum : netSum;
  // Grand totals across EVERY category (before sortTop truncates the label
  // list to the top 15) - the on-chart "%" must read as this category's share
  // of the true overall total, not of just the top-15 slice that happens to
  // be drawn. Using dataset.data.reduce() post-slice (the plugin's fallback)
  // would silently inflate every shown bar's % once a report has more
  // categories than sortTop keeps.
  var netGrandTotal = Object.keys(netSum).reduce(function (s, k) { return s + netSum[k]; }, 0);
  var nnGrandTotal = Object.keys(nnSum).reduce(function (s, k) { return s + nnSum[k]; }, 0);
  var labels = Object.keys(basisSum).sort(function (a, b) { return basisSum[b] - basisSum[a]; });
  if (opts.basis === 'ALL') {
    labels = Array.from(new Set(Object.keys(netSum).concat(Object.keys(nnSum))))
      .sort(function (a, b) { return (netSum[b] || 0) - (netSum[a] || 0); });
  }
  // Collapse everything past the top N into one "Others" bar so the chart
  // still visually accounts for the full total, instead of silently dropping
  // the remaining categories with no on-chart trace of them.
  var othersKey = null;
  if (opts.sortTop && labels.length > opts.sortTop) {
    var shownLabels = labels.slice(0, opts.sortTop);
    othersKey = t('otherCategory');
    netSum[othersKey] = netGrandTotal - shownLabels.reduce(function (s, k) { return s + (netSum[k] || 0); }, 0);
    nnSum[othersKey] = nnGrandTotal - shownLabels.reduce(function (s, k) { return s + (nnSum[k] || 0); }, 0);
    labels = shownLabels.concat([othersKey]);
  }

  var datasets;
  if (opts.basis === 'ALL') {
    datasets = [
      { label: t('legendNetSales'), data: labels.map(function (v) { return netSum[v] || 0; }), backgroundColor: C_NET, total: netGrandTotal },
      { label: t('legendNetNetSales'), data: labels.map(function (v) { return nnSum[v] || 0; }), backgroundColor: C_NETNET, total: nnGrandTotal }
    ];
  } else {
    datasets = [{
      label: opts.basis === 'NET NET' ? t('legendNetNetSales') : t('legendNetSales'),
      data: labels.map(function (v) { return basisSum[v]; }),
      backgroundColor: labels.map(function (v) { return v === othersKey ? '#9099a3' : opts.color; }),
      total: opts.basis === 'NET NET' ? nnGrandTotal : netGrandTotal
    }];
  }
  if (opts.type === 'pie') {
    // A pie can't group two bars side by side - "ALL" renders as two
    // concentric rings instead (Chart.js natively supports multiple pie
    // datasets as rings): inner = Net, outer = Net Net.
    if (opts.basis === 'ALL') {
      datasets = [
        { label: t('legendNetSales'), data: labels.map(function (v) { return netSum[v] || 0; }), backgroundColor: labels.map(function (v, i) { return CATEGORY_PALETTE[i % CATEGORY_PALETTE.length]; }), total: netGrandTotal },
        { label: t('legendNetNetSales'), data: labels.map(function (v) { return nnSum[v] || 0; }), backgroundColor: labels.map(function (v, i) { return CATEGORY_PALETTE[i % CATEGORY_PALETTE.length] + 'aa'; }), total: nnGrandTotal }
      ];
    } else {
      datasets = [{ label: datasets[0].label, data: datasets[0].data, backgroundColor: labels.map(function (v, i) { return CATEGORY_PALETTE[i % CATEGORY_PALETTE.length]; }), total: datasets[0].total }];
    }
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
      // A pie sized to fill its whole canvas leaves no room for the small-
      // slice leader-line labels (see valuePctLabels above) - they get drawn
      // past the pie's own edge and collide with the title text right above
      // it. Shrinking the pie itself (not just adding canvas padding) frees
      // that ring of space on every side, not just the top.
      radius: opts.type === 'pie' ? '65%' : undefined,
      layout: opts.type === 'pie' ? { padding: { top: 26, bottom: 10 } } : {},
      plugins: {
        title: { display: true, text: opts.title + ' (' + basisLabel(opts.basis) + ')' },
        legend: { display: opts.basis === 'ALL', position: 'bottom' },
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

function renderVendor() { renderCategoryChart({ chartKey: 'vendor', elId: 'chartVendor', groupKeyIdx: 3, basis: togVendor.get(), title: t('vendorTitle'), color: C_SALES, sortTop: 15, type: 'bar', indexAxis: 'y' }); }
function renderCustomer() { renderCategoryChart({ chartKey: 'customer', elId: 'chartCustomer', groupKeyIdx: 2, basis: togCustomer.get(), title: t('customerTitle'), color: C_COST, sortTop: 15, type: 'bar', indexAxis: 'y' }); }
function renderOrigin() { renderCategoryChart({ chartKey: 'origin', elId: 'chartOrigin', groupKeyIdx: 8, basis: togOrigin.get(), title: t('originTitle'), type: 'pie' }); }

// ---- Geographic: value basis is Margin (7 / 10), not Sales - separate from
// renderCategoryChart above, which is Sales-only. ----
function renderGeo() {
  var rows = lastRows, basis = togGeo.get();
  var netSum = groupSum(rows, 4, 7);
  var nnSum = groupSum(rows, 4, 10);
  var basisSum = basis === 'NET NET' ? nnSum : netSum;
  // No sortTop here (every country is shown), so these already equal the sum
  // of what's drawn - kept explicit anyway so this chart's % stays correct
  // even if a top-N truncation is ever added here later (see renderCategoryChart).
  var netGrandTotal = Object.keys(netSum).reduce(function (s, k) { return s + netSum[k]; }, 0);
  var nnGrandTotal = Object.keys(nnSum).reduce(function (s, k) { return s + nnSum[k]; }, 0);
  var labels = Object.keys(basisSum).sort(function (a, b) { return basisSum[b] - basisSum[a]; });
  if (basis === 'ALL') labels = Array.from(new Set(Object.keys(netSum).concat(Object.keys(nnSum)))).sort(function (a, b) { return (netSum[b] || 0) - (netSum[a] || 0); });
  var datasets = basis === 'ALL'
    ? [
        { label: t('legendNetMargin'), data: labels.map(function (v) { return netSum[v] || 0; }), backgroundColor: C_NET, total: netGrandTotal },
        { label: t('legendNetNetMargin'), data: labels.map(function (v) { return nnSum[v] || 0; }), backgroundColor: C_NETNET, total: nnGrandTotal }
      ]
    : [{ label: basis === 'NET NET' ? t('legendNetNetMargin') : t('legendNetMargin'), data: labels.map(function (v) { return basisSum[v]; }), backgroundColor: C_MARGIN, total: basis === 'NET NET' ? nnGrandTotal : netGrandTotal }];
  if (charts.geo) charts.geo.destroy();
  var maxVal = datasets.reduce(function (m, ds) { return Math.max(m, ds.data.reduce(function (mm, v) { return Math.max(mm, v); }, 0)); }, 0);
  charts.geo = new Chart(document.getElementById('chartGeo'), {
    type: 'bar',
    data: { labels: labels, datasets: datasets },
    options: {
      responsive: true,
      scales: { y: { suggestedMax: maxVal * 1.2 } },
      plugins: {
        title: { display: true, text: t('geoTitle') + ' (' + basisLabel(basis) + ')' },
        legend: { display: basis === 'ALL', position: 'bottom' },
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

var togTrend = createBasisToggle('tog-trend', renderTrend);
var togVendor = createBasisToggle('tog-vendor', renderVendor);
var togCustomer = createBasisToggle('tog-customer', renderCustomer);
var togGeo = createBasisToggle('tog-geo', renderGeo);
var togOrigin = createBasisToggle('tog-origin', renderOrigin);

function render() {
  lastRows = getFiltered();
  renderKpis();
  renderTrend();
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
  [togTrend, togVendor, togCustomer, togGeo, togOrigin].forEach(function (tg) { tg.relabel(); });
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
  null, 'title_data');
console.log('Data sheet built (' + (dataSheetHtml.length / 1024 / 1024).toFixed(2) + ' MB)');

// =========================================================================
// Exchange Rate (USD table only)
// =========================================================================
const exRateData = readJson('exrate_usd.json');
const exRateHtml = buildDataTablesHtml(exRateData.header, exRateData.rows, 'Exchange Rate', false, null, null, 'title_exrate');

// =========================================================================
// YoY comparison sheets - already Excel-formatted display text, plain
// DataTables pages (narrow enough sheets that scrollX is unneeded except
// Data itself).
// =========================================================================
const yoyCustomers = readJson('yoy_customers.json');
const yoyFabrication = readJson('yoy_fabrication.json');
const yoyProdFamily = readJson('yoy_prodfamily.json');
const yoyPartId = readJson('yoy_partid.json');
const yoyNetNetMargin = readJson('yoy_netnetmargin.json');
const netMarginByCustomer = readJson('net_margin_by_customer.json');

// drilldown.dataColMap maps THIS sheet's label column name -> the matching
// Data-tab column name to filter by; labelColCount marks where the year-paired
// metric columns start (see Export-YoYPivotSheet in the .ps1 for the same counts).
const yoyCustomersHtml = buildDataTablesHtml(yoyCustomers.header, yoyCustomers.rows, '2026 v 2025 by Customers', false, ['Cust_name', 'Country of origin'],
  { labelColCount: 3, dataColMap: { 'Cust_name': 'Cust_name', 'Country of origin': 'Country of origin', 'Part_id': 'Part_id' } }, 'title_yoyCustomers');
const yoyFabricationHtml = buildDataTablesHtml(yoyFabrication.header, yoyFabrication.rows, '2026 v 2025 by Fabrication', false, ['Fabrication Code'],
  { labelColCount: 4, dataColMap: { 'Fabrication Code': 'Fabrication Code', 'Part_id': 'Part_id' } }, 'title_yoyFabrication');
const yoyProdFamilyHtml = buildDataTablesHtml(yoyProdFamily.header, yoyProdFamily.rows, '2026 v 2025 by Prod Family', false, [],
  { labelColCount: 1, dataColMap: { 'Family1': 'Family1' } }, 'title_yoyProdFamily');
const yoyPartIdHtml = buildDataTablesHtml(yoyPartId.header, yoyPartId.rows, '2026 v 2025 by Part ID', false, [],
  { labelColCount: 2, dataColMap: { 'Part_id': 'Part_id' } }, 'title_yoyPartId');
const yoyNetNetMarginHtml = buildDataTablesHtml(yoyNetNetMargin.header, yoyNetNetMargin.rows, '2026 v 2025 by NET NET MARGIN', false, ['Cust_name', 'Country of origin'],
  { labelColCount: 3, dataColMap: { 'Cust_name': 'Cust_name', 'Country of origin': 'Country of origin', 'Part_id': 'Part_id' } }, 'title_yoyNetNetMargin');
const netMarginByCustomerHtml = buildDataTablesHtml(netMarginByCustomer.header, netMarginByCustomer.rows, 'Net Margin by Customer-Quarter', false, ['Cust_name'],
  { labelColCount: 1, dataColMap: { 'Cust_name': 'Cust_name' }, fixedYear: netMarginByCustomer.year }, 'title_customerQuarter');

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
var introDismissed = false;
try { introDismissed = localStorage.getItem('reportIntroDismissed') === '1'; } catch (e) {}
if (!introDismissed) introOverlay.classList.remove('hidden');
introCloseBtn.onclick = function () {
  if (introDontShow.checked) { try { localStorage.setItem('reportIntroDismissed', '1'); } catch (e) {} }
  introOverlay.classList.add('hidden');
};
</script>
</body>
</html>`;

fs.writeFileSync(OUT_PATH, combinedHtml, 'utf8');
const stat = fs.statSync(OUT_PATH);
console.log('Wrote combined report:', OUT_PATH);
console.log('Total size:', (stat.size / 1024 / 1024).toFixed(2), 'MB');
