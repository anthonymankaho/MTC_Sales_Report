'use strict';

// Forecast cost and total cost for the embedded "Purchase: YTD by Vendor"
// page. The Purchase Report project has no forecast pipeline, so this join
// runs in extractPurchaseYtdTab() when that page is copied in. A later
// verbatim re-embed must keep the call — without it the columns disappear.
//
// Forecast cost basis (already applied in forecast_sheet.json / the Forecast
// Data tab): outstanding orders only (COMPLETED/CANCELLED excluded), valued
// at unshipped qty (Epidueqty), bucketed by Ship By Date. Amount = Prime_USD.
// Forecast has no commission, so it is never added into the 2025/2026 PO
// Price w/comm columns and pctChange is not recomputed.
// Total cost = 2026 PO Price w/comm USD + forecast cost. A row with no 2026
// actual shows the forecast amount alone.

const COPY = {
  en: {
    colForecastUsd: 'Forecast cost USD',
    colTotalCost: 'Total cost USD',
    colTotalCostNote: '2026 total PO Price w/comm USD plus forecast cost USD. Forecast has no commission yet and is not inside the 2026 column or the % change.',
    colForecastNote: 'Forecast purchase cost USD. Outstanding orders only (excludes COMPLETED/CANCELLED), valued at unshipped qty (Epidueqty), bucketed by Ship By Date. Purchase cost = Prime_USD. No commission yet - this column is not part of the 2025/2026 PO Price w/comm totals and is not in the % change.',
    totalLabel: 'Total',
    summaryForecast: 'Forecast cost',
    summaryTotalCost: 'Total cost',
    forecastUnmatchedTag: 'unmatched forecast',
    forecastUnmatchedFoot: '{n} forecast vendor name(s) did not match a YTD vendor ({amt}). Listed on their own rows and included in the total.',
    forecastAllMatched: 'Every forecast vendor name matched a YTD vendor.'
  },
  zh: {
    colForecastUsd: '預測成本 USD',
    colTotalCost: '總成本 USD',
    colTotalCostNote: '2026年採購總額（含佣金）USD 加上預測成本 USD。預測尚未含佣金，不併入 2026 欄，亦不計入變動百分比。',
    colForecastNote: '預測採購成本 USD。僅未完成訂單（不含 COMPLETED／CANCELLED），以未出貨數量（Epidueqty）計價，按 Ship By Date 分桶。採購成本 = Prime_USD，尚未含佣金，故本欄不併入 2025／2026 含佣金採購總額，亦不計入變動百分比。',
    totalLabel: '合計',
    summaryForecast: '預測成本',
    summaryTotalCost: '總成本',
    forecastUnmatchedTag: '未對上預測',
    forecastUnmatchedFoot: '{n} 個預測供應商名稱未能對上年度供應商（{amt}）。已另列並計入合計。',
    forecastAllMatched: '所有預測供應商名稱均已對上年度供應商。'
  }
};

function normVendorName(value) {
  return String(value == null ? '' : value)
    .replace(/\u00a0/g, ' ')
    .trim()
    .toUpperCase()
    .replace(/[.,&'"()\u2019]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function aggregateForecastPrime(header, rows) {
  const columns = ['Prime_USD', 'VEND_NAME', 'Combined Vendor', 'Vendor Country'];
  const idx = {};
  columns.forEach(name => {
    idx[name] = header.indexOf(name);
    if (idx[name] < 0) throw new Error('forecast sheet is missing column "' + name + '"');
  });
  const buckets = new Map();
  let total = 0;
  let counted = 0;
  rows.forEach(row => {
    const raw = row[idx.Prime_USD];
    const amount = typeof raw === 'number' ? raw : (raw == null || raw === '' ? 0 : Number(raw));
    if (!Number.isFinite(amount)) return;
    total += amount;
    counted += 1;
    const combined = String(row[idx['Combined Vendor']] == null ? '' : row[idx['Combined Vendor']]).trim();
    const vendor = String(row[idx.VEND_NAME] == null ? '' : row[idx.VEND_NAME]).trim();
    const country = String(row[idx['Vendor Country']] == null ? '' : row[idx['Vendor Country']]).trim();
    const key = combined + '\t' + country + '\t' + vendor;
    buckets.set(key, (buckets.get(key) || 0) + amount);
  });
  return { buckets, total, counted };
}

function blankMoney(extra) {
  return Object.assign({ usd2025: null, usd2026: null, pctChange: null, forecastUsd: 0 }, extra);
}

// Attach forecast Prime_USD onto a purchase GROUPS array (mutates it).
// Match order, most specific first:
//   1. Individual vendor name + country of origin
//   2. Individual vendor name unique across the page (Vendor Country often
//      differs from purchase COO: domicile vs factory)
//   3. Same individual name on several COOs, forecast country matches none
//      -> a labeled COO row under that grouped vendor, not a guess
//   4. Grouped vendor name, then the same country / single-leaf rules
//   5. Anything still unmatched becomes its own forecast-only group
// 2025/2026 amounts and pctChange are left as they were.
function joinForecastOntoGroups(groups, buckets) {
  const leaves = [];
  groups.forEach(group => {
    group.cooRows.forEach(coo => {
      coo.vendorRows.forEach(vendor => {
        vendor.forecastUsd = 0;
        leaves.push({
          g: group,
          c: coo,
          v: vendor,
          gN: normVendorName(group.vendor),
          vN: normVendorName(vendor.vendorName),
          cN: normVendorName(coo.coo)
        });
      });
    });
  });

  function addLeaf(leaf, amount) {
    leaf.v.forecastUsd += amount;
  }

  function addSyntheticIndividual(group, coo, vendorName, amount) {
    const name = vendorName || group.vendor || '(blank vendor)';
    let row = coo.vendorRows.find(v => normVendorName(v.vendorName) === normVendorName(name) && v.forecastSynthetic);
    if (!row) {
      row = blankMoney({ vendorName: name, forecastSynthetic: true });
      coo.vendorRows.push(row);
    }
    row.forecastUsd += amount;
    return true;
  }

  function addSyntheticCoo(group, country, vendorName, amount) {
    const label = (country || 'unspecified') + ' (forecast vendor country / 預測供應商國家)';
    let coo = group.cooRows.find(c => c.coo === label);
    if (!coo) {
      coo = blankMoney({ coo: label, vendorRows: [], forecastSynthetic: true });
      group.cooRows.push(coo);
    }
    return addSyntheticIndividual(group, coo, vendorName, amount);
  }

  const unmatched = [];
  buckets.forEach((amount, key) => {
    const parts = key.split('\t');
    const combined = parts[0];
    const country = parts[1];
    const vendor = parts[2];
    const nComb = normVendorName(combined);
    const nVend = normVendorName(vendor);
    const nCtry = normVendorName(country);
    const indiv = leaves.filter(leaf => (nVend && leaf.vN === nVend) || (nComb && leaf.vN === nComb));
    const indivCountry = nCtry ? indiv.filter(leaf => leaf.cN === nCtry) : [];
    let placed = false;

    if (indivCountry.length === 1) {
      addLeaf(indivCountry[0], amount);
      placed = true;
    } else if (indivCountry.length > 1) {
      const groupsHit = new Set(indivCountry.map(leaf => leaf.g));
      // Same individual name under one grouped vendor: keep the amount on a
      // labeled sibling row instead of splitting it across lookalike leaves.
      // Two grouped vendors sharing that name stay unmatched rather than guessed.
      if (groupsHit.size === 1) placed = addSyntheticIndividual(indivCountry[0].g, indivCountry[0].c, vendor || combined, amount);
    } else if (indiv.length === 1) {
      addLeaf(indiv[0], amount);
      placed = true;
    } else if (indiv.length > 1) {
      const groupsHit = new Set(indiv.map(leaf => leaf.g));
      if (groupsHit.size === 1) placed = addSyntheticCoo(indiv[0].g, country, vendor || combined, amount);
    }

    if (!placed) {
      const grouped = leaves.filter(leaf => (nComb && leaf.gN === nComb) || (nVend && leaf.gN === nVend));
      const groupedHit = new Set(grouped.map(leaf => leaf.g));
      if (groupedHit.size === 1) {
        const group = grouped[0].g;
        const groupedCountry = nCtry ? grouped.filter(leaf => leaf.cN === nCtry) : [];
        if (groupedCountry.length === 1) {
          addLeaf(groupedCountry[0], amount);
          placed = true;
        } else if (groupedCountry.length > 1) {
          placed = addSyntheticIndividual(group, groupedCountry[0].c, vendor || combined, amount);
        } else if (group.cooRows.length === 1 && group.cooRows[0].vendorRows.length === 1) {
          addLeaf(grouped[0], amount);
          placed = true;
        } else if (group.cooRows.length === 1) {
          placed = addSyntheticIndividual(group, group.cooRows[0], vendor || combined, amount);
        } else {
          placed = addSyntheticCoo(group, country, vendor || combined, amount);
        }
      }
    }

    if (!placed) unmatched.push({ combined: combined, country: country, vendor: vendor, amount: amount });
  });

  function rollup(list) {
    list.forEach(group => {
      group.cooRows.forEach(coo => {
        coo.forecastUsd = coo.vendorRows.reduce((sum, vendor) => sum + (Number(vendor.forecastUsd) || 0), 0);
      });
      group.forecastUsd = group.cooRows.reduce((sum, coo) => sum + (Number(coo.forecastUsd) || 0), 0);
    });
  }
  rollup(groups);

  unmatched.forEach(item => {
    const name = item.combined || item.vendor || '(blank vendor)';
    const leafName = item.vendor || name;
    groups.push({
      vendor: name,
      forecastOnly: true,
      usd2025: null,
      usd2026: null,
      pctChange: null,
      forecastUsd: item.amount,
      cooRows: [{
        coo: item.country || '',
        usd2025: null,
        usd2026: null,
        pctChange: null,
        forecastUsd: item.amount,
        vendorRows: [blankMoney({ vendorName: leafName, forecastUsd: item.amount })]
      }]
    });
  });

  const unmatchedAmount = unmatched.reduce((sum, item) => sum + item.amount, 0);
  return {
    unmatchedCount: unmatched.length,
    unmatchedAmount: unmatchedAmount,
    unmatchedNames: unmatched.map(item => item.combined || item.vendor || '(blank vendor)')
  };
}

function mustReplace(html, from, to, label) {
  const count = html.split(from).length - 1;
  if (count !== 1) {
    throw new Error('Purchase YTD forecast injection: expected 1 occurrence of "' + label + '", found ' + count);
  }
  return html.replace(from, to);
}

function scanJsonEnd(source, start) {
  if (source[start] !== '[' && source[start] !== '{') throw new Error('Purchase YTD GROUPS is not a JSON value');
  let depth = 0;
  let inString = false;
  let escaped = false;
  for (let i = start; i < source.length; i++) {
    const ch = source[i];
    if (inString) {
      if (escaped) escaped = false;
      else if (ch === '\\') escaped = true;
      else if (ch === '"') inString = false;
      continue;
    }
    if (ch === '"') inString = true;
    else if (ch === '[' || ch === '{') depth += 1;
    else if (ch === ']' || ch === '}') {
      depth -= 1;
      if (depth === 0) return i + 1;
    }
  }
  throw new Error('Purchase YTD GROUPS JSON is unterminated');
}

function jsString(value) {
  return "'" + String(value).replace(/\\/g, '\\\\').replace(/'/g, '\\\'') + "'";
}

function i18nPairs(lang) {
  const copy = COPY[lang];
  return [
    'colForecastUsd: ' + jsString(copy.colForecastUsd),
    'colTotalCost: ' + jsString(copy.colTotalCost),
    'colTotalCostNote: ' + jsString(copy.colTotalCostNote),
    'colForecastNote: ' + jsString(copy.colForecastNote),
    'totalLabel: ' + jsString(copy.totalLabel),
    'summaryForecast: ' + jsString(copy.summaryForecast),
    'summaryTotalCost: ' + jsString(copy.summaryTotalCost),
    'forecastUnmatchedTag: ' + jsString(copy.forecastUnmatchedTag),
    'forecastUnmatchedFoot: ' + jsString(copy.forecastUnmatchedFoot),
    'forecastAllMatched: ' + jsString(copy.forecastAllMatched)
  ].join(', ');
}

// Appended by extractPurchaseYtdTab after the column injection. Kept here so
// a rebuild of an already-embedded page uses the same source line.
const PURCHASE_YTD_SOURCE_FOOT = '<div style="margin:18px 8px 6px;padding:8px 0 0;border-top:1px solid rgba(128,128,128,.4);font:10.5px/1.55 Calibri,Arial,sans-serif;color:rgba(128,128,128,.95)">Sources / 資料來源: FIN_MTC Purchase Report html (MTC Purchase Report Claude project, embedded as-built) — see that report for its period &amp; FX basis · 期間與匯率基礎詳該採購報表。 Forecast cost USD is injected at embed from forecast_sheet.json (Prime_USD, outstanding / Ship By Date, no commission) and is not in that Purchase report. Total cost USD is 2026 PO Price w/comm plus that forecast cost. · 預測成本於嵌入時由 forecast_sheet.json 注入（Prime_USD、未完成訂單、Ship By Date、尚未含佣金），並非該採購報表欄位。總成本為 2026 含佣金採購總額加上該預測成本。</div>';

function insertForecastNote(html) {
  if (html.indexOf('id="forecastBasisNote"') >= 0) return html;
  const note = '<p class="page-intro forecast-basis" id="forecastBasisNote">' + COPY.en.colForecastNote + '</p>\n';
  const sourceMarker = '<div style="margin:18px 8px 6px;padding:8px 0 0;border-top:1px solid rgba(128,128,128,.4);';
  const atSource = html.lastIndexOf(sourceMarker);
  if (atSource >= 0) return html.slice(0, atSource) + note + html.slice(atSource);
  const body = html.lastIndexOf('</body>');
  if (body < 0) throw new Error('Purchase YTD page has no </body> for the forecast footnote');
  return html.slice(0, body) + note + html.slice(body);
}

function injectForecastPurchaseColumn(tabHtml, forecastSheet) {
  if (!forecastSheet || !Array.isArray(forecastSheet.header) || !Array.isArray(forecastSheet.rows)) {
    return { html: tabHtml, injected: false, reason: 'forecast_sheet.json absent' };
  }
  if (tabHtml.indexOf('data-col="forecastUsd"') >= 0) {
    throw new Error('Purchase YTD page already has the forecast column; refusing to inject twice');
  }

  const marker = 'var GROUPS = ';
  const at = tabHtml.indexOf(marker);
  if (at < 0) throw new Error('Purchase YTD page has no GROUPS array');
  const jsonStart = at + marker.length;
  const jsonEnd = scanJsonEnd(tabHtml, jsonStart);
  if (tabHtml.slice(jsonEnd, jsonEnd + 1) !== ';') throw new Error('Purchase YTD GROUPS is not terminated with a semicolon');
  const groups = JSON.parse(tabHtml.slice(jsonStart, jsonEnd));
  const before = JSON.stringify(groups);

  const agg = aggregateForecastPrime(forecastSheet.header, forecastSheet.rows);
  const join = joinForecastOntoGroups(groups, agg.buckets);
  const joinedTotal = groups.reduce((sum, group) => sum + (Number(group.forecastUsd) || 0), 0);
  if (Math.abs(joinedTotal - agg.total) > 0.02) {
    throw new Error('Forecast join dropped Prime_USD: sheet ' + agg.total + ' vs groups ' + joinedTotal);
  }

  // Actuals and the stored % change must survive the round-trip.
  const original = JSON.parse(before);
  original.forEach((group, gi) => {
    const next = groups[gi];
    if (next.vendor !== group.vendor || next.usd2025 !== group.usd2025 || next.usd2026 !== group.usd2026 || next.pctChange !== group.pctChange) {
      throw new Error('Forecast join changed YTD actuals for ' + group.vendor);
    }
    group.cooRows.forEach((coo, ci) => {
      const nextCoo = next.cooRows[ci];
      if (!nextCoo || nextCoo.coo !== coo.coo || nextCoo.usd2025 !== coo.usd2025 || nextCoo.usd2026 !== coo.usd2026 || nextCoo.pctChange !== coo.pctChange) {
        throw new Error('Forecast join changed a COO actual for ' + group.vendor + ' / ' + coo.coo);
      }
      coo.vendorRows.forEach((vendor, vi) => {
        const nextVendor = nextCoo.vendorRows[vi];
        if (!nextVendor || nextVendor.vendorName !== vendor.vendorName || nextVendor.usd2025 !== vendor.usd2025 || nextVendor.usd2026 !== vendor.usd2026 || nextVendor.pctChange !== vendor.pctChange) {
          throw new Error('Forecast join changed an individual-vendor actual for ' + vendor.vendorName);
        }
      });
    });
  });

  const groupsJson = JSON.stringify(groups).replace(/<\/script/gi, '<\\/script');
  let html = tabHtml.slice(0, jsonStart) + groupsJson + tabHtml.slice(jsonEnd);

  html = mustReplace(
    html,
    "colPctChange: '2025 vs 2026 PO Cost % Change', colIndividualVendorName: 'Individual Vendor Name',",
    "colPctChange: '2025 vs 2026 PO Cost % Change', " + i18nPairs('en') + ", colIndividualVendorName: 'Individual Vendor Name',",
    'English column labels'
  );
  html = mustReplace(
    html,
    "colPctChange: '2025對比2026採購成本%變動', colIndividualVendorName: '個別供應商名稱',",
    "colPctChange: '2025對比2026採購成本%變動', " + i18nPairs('zh') + ", colIndividualVendorName: '個別供應商名稱',",
    'Chinese column labels'
  );
  html = mustReplace(
    html,
    '<th data-col="pctChange" data-i18n="colPctChange">2025 vs 2026 PO Cost % Change</th>',
    '<th data-col="pctChange" data-i18n="colPctChange">2025 vs 2026 PO Cost % Change</th>\n    <th data-col="forecastUsd" data-i18n="colForecastUsd" data-i18n-title="colForecastNote" title="' + COPY.en.colForecastNote.replace(/"/g, '&quot;') + '">' + COPY.en.colForecastUsd + '</th>\n    <th data-col="totalCost" data-i18n="colTotalCost" data-i18n-title="colTotalCostNote" title="' + COPY.en.colTotalCostNote.replace(/"/g, '&quot;') + '">' + COPY.en.colTotalCost + '</th>',
    'forecast column header'
  );
  html = mustReplace(
    html,
    'tr.vname-row td:first-child { padding-left: 46px; }',
    'tr.vname-row td:first-child { padding-left: 46px; }\ntr.total-row td { font-weight: bold; border-top: 2px solid var(--border); background: var(--hover-bg); }\ntr.forecast-only td:first-child { font-style: italic; }\np.forecast-basis { font-size: 11px; max-width: 980px; margin-top: 14px; }',
    'forecast column CSS'
  );
  html = mustReplace(
    html,
    "document.querySelectorAll('[data-i18n]').forEach(function (el) { el.textContent = t(el.dataset.i18n); });",
    "document.querySelectorAll('[data-i18n]').forEach(function (el) { el.textContent = t(el.dataset.i18n); });\n  document.querySelectorAll('[data-i18n-title]').forEach(function (el) { el.title = t(el.getAttribute('data-i18n-title')); });",
    'language title hook'
  );
  html = mustReplace(
    html,
    "[['usd2025', fmtUsd], ['usd2026', fmtUsd], ['pctChange', fmtPctChange]].forEach(function (pair) {\n    var td = document.createElement('td');\n    td.textContent = pair[1](rowData[pair[0]]);",
    "var cells = { usd2025: rowData.usd2025, usd2026: rowData.usd2026, pctChange: rowData.pctChange, forecastUsd: rowData.forecastUsd, totalCost: totalCostOf(rowData) };\n  [['usd2025', fmtUsd], ['usd2026', fmtUsd], ['pctChange', fmtPctChange], ['forecastUsd', fmtUsd], ['totalCost', fmtUsd]].forEach(function (pair) {\n    var td = document.createElement('td');\n    td.textContent = pair[1](cells[pair[0]]);",
    'numeric cells'
  );
  html = mustReplace(
    html,
    "    var link = document.createElement('span');\n    link.className = 'drill-link'; link.textContent = g.vendor; link.title = t('drillTooltip');\n    link.addEventListener('click', function () { drillToVendor(g.vendor); });\n    tdVendor.appendChild(link); tr.appendChild(tdVendor);",
    "    if (g.forecastOnly) tr.classList.add('forecast-only');\n    var link = document.createElement('span');\n    if (g.forecastOnly) {\n      link.textContent = vendorRowLabel(g);\n      link.title = t('forecastUnmatchedTag');\n    } else {\n      link.className = 'drill-link'; link.textContent = g.vendor; link.title = t('drillTooltip');\n      link.addEventListener('click', function () { drillToVendor(g.vendor); });\n    }\n    tdVendor.appendChild(link); tr.appendChild(tdVendor);",
    'vendor drill link'
  );
  html = mustReplace(
    html,
    "    if (sortCol === 'vendor') { av = a.vendor; bv = b.vendor; return av.localeCompare(bv) * sortDir; }\n    av = a[sortCol]; bv = b[sortCol];",
    "    if (sortCol === 'vendor') { av = a.vendor; bv = b.vendor; return av.localeCompare(bv) * sortDir; }\n    av = sortCol === 'totalCost' ? totalCostOf(a) : a[sortCol]; bv = sortCol === 'totalCost' ? totalCostOf(b) : b[sortCol];",
    'total cost sort'
  );
  html = mustReplace(
    html,
    "function render() {",
    "function totalCostOf(row) {\n" +
    "  var hasActual = typeof row.usd2026 === 'number' && !isNaN(row.usd2026);\n" +
    "  var forecast = (typeof row.forecastUsd === 'number' && !isNaN(row.forecastUsd)) ? row.forecastUsd : 0;\n" +
    "  if (!hasActual && !forecast) return null;\n" +
    "  return (hasActual ? row.usd2026 : 0) + forecast;\n" +
    "}\n" +
    "function forecastColumnTotal(groups) {\n" +
    "  var tot25 = 0, tot26 = 0, totFc = 0, any25 = false, any26 = false;\n" +
    "  groups.forEach(function (g) {\n" +
    "    if (typeof g.usd2025 === 'number' && !isNaN(g.usd2025)) { tot25 += g.usd2025; any25 = true; }\n" +
    "    if (typeof g.usd2026 === 'number' && !isNaN(g.usd2026)) { tot26 += g.usd2026; any26 = true; }\n" +
    "    if (typeof g.forecastUsd === 'number' && !isNaN(g.forecastUsd)) totFc += g.forecastUsd;\n" +
    "  });\n" +
    "  // Same rule as the Purchase page hierarchy: a zero 2025 base is N/A when 2026 is non-zero, else 0.\n" +
    "  var pct = !any25 && !any26 ? null : (!tot25 ? (tot26 ? null : 0) : (tot26 - tot25) / tot25 * 100);\n" +
    "  return { usd2025: any25 ? tot25 : null, usd2026: any26 ? tot26 : null, pctChange: pct, forecastUsd: totFc, totalCost: totalCostOf({ usd2026: any26 ? tot26 : null, forecastUsd: totFc }) };\n" +
    "}\n" +
    "function vendorRowLabel(g) {\n" +
    "  return g.forecastOnly ? (g.vendor + ' (' + t('forecastUnmatchedTag') + ')') : g.vendor;\n" +
    "}\n" +
    "function render() {",
    'render helpers'
  );
  html = mustReplace(
    html,
    "  if (sb) sb.textContent = t(sb.dataset.titleKey || '') + ' - ' + groups.length.toLocaleString() + ' ' + t('rowsSuffix');",
    "  var tot = forecastColumnTotal(groups);\n" +
    "  var trT = document.createElement('tr'); trT.className = 'total-row';\n" +
    "  var tdL = document.createElement('td'); tdL.textContent = t('totalLabel'); trT.appendChild(tdL);\n" +
    "  trT.appendChild(document.createElement('td'));\n" +
    "  trT.appendChild(document.createElement('td'));\n" +
    "  trT.appendChild(document.createElement('td'));\n" +
    "  addNumCells(trT, tot);\n" +
    "  tbody.appendChild(trT);\n" +
    "  var note = document.getElementById('forecastBasisNote');\n" +
    "  if (note) {\n" +
    "    var unmatched = GROUPS.filter(function (g) { return g.forecastOnly; });\n" +
    "    var uAmt = unmatched.reduce(function (s, g) { return s + (g.forecastUsd || 0); }, 0);\n" +
    "    note.textContent = t('colForecastNote') + ' ' + (unmatched.length\n" +
    "      ? t('forecastUnmatchedFoot').replace('{n}', unmatched.length.toLocaleString()).replace('{amt}', fmtUsd(uAmt))\n" +
    "      : t('forecastAllMatched'));\n" +
    "  }\n" +
    "  if (sb) sb.textContent = t(sb.dataset.titleKey || '') + ' - ' + groups.length.toLocaleString() + ' ' + t('rowsSuffix') + ' · ' + t('summaryForecast') + ' ' + fmtUsd(tot.forecastUsd) + ' · ' + t('summaryTotalCost') + ' ' + fmtUsd(tot.totalCost);",
    'summary and total row'
  );
  html = mustReplace(
    html,
    "  var header = [t('colIndividualVendorName'), t('colVendorName'), t('colCoo'), t('colUsd2025'), t('colUsd2026'), t('colPctChange')];",
    "  var header = [t('colIndividualVendorName'), t('colVendorName'), t('colCoo'), t('colUsd2025'), t('colUsd2026'), t('colPctChange'), t('colForecastUsd'), t('colTotalCost')];",
    'CSV header'
  );
  html = mustReplace(
    html,
    "    lines.push(['', g.vendor, hasCoo ? '' : (g.cooRows[0] ? g.cooRows[0].coo : ''), fmtCsvNum(g.usd2025), fmtCsvNum(g.usd2026), fmtCsvNum(g.pctChange)].map(csvEscape).join(','));",
    "    lines.push(['', vendorRowLabel(g), hasCoo ? '' : (g.cooRows[0] ? g.cooRows[0].coo : ''), fmtCsvNum(g.usd2025), fmtCsvNum(g.usd2026), fmtCsvNum(g.pctChange), fmtCsvNum(g.forecastUsd), fmtCsvNum(totalCostOf(g))].map(csvEscape).join(','));",
    'CSV group row'
  );
  html = mustReplace(
    html,
    "      if (hasCoo) lines.push(['', g.vendor, c.coo, fmtCsvNum(c.usd2025), fmtCsvNum(c.usd2026), fmtCsvNum(c.pctChange)].map(csvEscape).join(','));",
    "      if (hasCoo) lines.push(['', vendorRowLabel(g), c.coo, fmtCsvNum(c.usd2025), fmtCsvNum(c.usd2026), fmtCsvNum(c.pctChange), fmtCsvNum(c.forecastUsd), fmtCsvNum(totalCostOf(c))].map(csvEscape).join(','));",
    'CSV country row'
  );
  html = mustReplace(
    html,
    "        lines.push([v.vendorName, g.vendor, c.coo, fmtCsvNum(v.usd2025), fmtCsvNum(v.usd2026), fmtCsvNum(v.pctChange)].map(csvEscape).join(','));",
    "        lines.push([v.vendorName, vendorRowLabel(g), c.coo, fmtCsvNum(v.usd2025), fmtCsvNum(v.usd2026), fmtCsvNum(v.pctChange), fmtCsvNum(v.forecastUsd), fmtCsvNum(totalCostOf(v))].map(csvEscape).join(','));",
    'CSV vendor row'
  );
  html = mustReplace(
    html,
    "  var csv = '\\ufeff' + lines.join('\\r\\n');",
    "  var csvTot = forecastColumnTotal(groups);\n" +
    "  lines.push([t('totalLabel'), '', '', fmtCsvNum(csvTot.usd2025), fmtCsvNum(csvTot.usd2026), fmtCsvNum(csvTot.pctChange), fmtCsvNum(csvTot.forecastUsd), fmtCsvNum(csvTot.totalCost)].map(csvEscape).join(','));\n" +
    "  var csv = '\\ufeff' + lines.join('\\r\\n');",
    'CSV total'
  );
  html = insertForecastNote(html);

  return {
    html: html,
    injected: true,
    total: agg.total,
    counted: agg.counted,
    buckets: agg.buckets.size,
    unmatchedCount: join.unmatchedCount,
    unmatchedAmount: join.unmatchedAmount,
    unmatchedNames: join.unmatchedNames,
    groupCount: groups.length
  };
}

module.exports = {
  COPY: COPY,
  PURCHASE_YTD_SOURCE_FOOT: PURCHASE_YTD_SOURCE_FOOT,
  normVendorName: normVendorName,
  aggregateForecastPrime: aggregateForecastPrime,
  joinForecastOntoGroups: joinForecastOntoGroups,
  injectForecastPurchaseColumn: injectForecastPurchaseColumn
};
