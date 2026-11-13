const { chromium } = require('playwright');
const path = require('path');

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage();
  const filePath = 'file:///' + path.resolve('C:/Users/User/Desktop/MTC Performance Report Claude/FIN_MTC_Performance_Report_html_TEMP.html').replace(/\\/g, '/');
  await page.goto(filePath);
  await page.waitForTimeout(500);
  const overlay = await page.$('#introOverlay');
  if (overlay) { await page.evaluate(() => { const el = document.getElementById('introOverlay'); if (el) el.remove(); }); await page.waitForTimeout(200); }

  await page.evaluate(() => { window.__idx = SHEETS.findIndex(s => s.key === 'yoyProdFamily'); });
  const idx = await page.evaluate(() => window.__idx);
  const tabEls = await page.$$('#tabs > *');
  await tabEls[idx].click();
  await page.waitForTimeout(600);
  const frame = page.frameLocator('#sheetFrame');

  const headerTexts = await frame.locator('#dataTableHeadRow th').allInnerTexts();
  console.log('Header (first 4 cols):', JSON.stringify(headerTexts.slice(0, 4)));

  const totalRows = await frame.locator('#dataTable tbody tr').count();
  console.log('Visible rows on default (all-collapsed) load:', totalRows, '(should equal # distinct brands, not 138 detail rows)');

  const firstRowCells = await frame.locator('#dataTable tbody tr').first().locator('td').allInnerTexts();
  console.log('First row (collapsed) cells:', JSON.stringify(firstRowCells.slice(0, 4)));

  // KPI totals at load (should reflect leaf-only sum, i.e. true grand total)
  const kpiValueInitial = await frame.locator('#totalIndexKpis2026 .kpi').first().locator('.value').innerText();
  console.log('Total Index first KPI (2026) at load:', kpiValueInitial);

  // Find PRESTIGE row and expand it
  await frame.locator('#dataTable tbody tr').locator('td:has-text("PRESTIGE")').first().waitFor();
  const rows = frame.locator('#dataTable tbody tr');
  const rowCount = await rows.count();
  let prestigeRowIdx = -1;
  for (let i = 0; i < rowCount; i++) {
    const t = await rows.nth(i).locator('td').first().innerText();
    if (t.trim() === 'PRESTIGE') { prestigeRowIdx = i; break; }
  }
  console.log('PRESTIGE row index:', prestigeRowIdx);
  const toggle = rows.nth(prestigeRowIdx).locator('.row-toggle');
  await toggle.click();
  await page.waitForTimeout(400);

  const rowsAfterExpand = await frame.locator('#dataTable tbody tr').count();
  console.log('Row count after expanding PRESTIGE:', rowsAfterExpand, '(should be initial - 1 + # of PRESTIGE Family1 rows)');

  const kpiValueAfterExpand = await frame.locator('#totalIndexKpis2026 .kpi').first().locator('.value').innerText();
  console.log('Total Index first KPI (2026) after expanding PRESTIGE (should be UNCHANGED):', kpiValueAfterExpand, '| matches initial:', kpiValueAfterExpand === kpiValueInitial);

  // check grouping: are all visible PRESTIGE-brand rows adjacent to each other?
  const allBrandCol = await frame.locator('#dataTable tbody tr td:first-child').allInnerTexts();
  let prestigeIndices = [];
  allBrandCol.forEach((v, i) => { if (v.trim() === 'PRESTIGE') prestigeIndices.push(i); });
  const isContiguous = prestigeIndices.every((v, i) => i === 0 || v === prestigeIndices[i - 1] + 1);
  console.log('PRESTIGE rows contiguous (grouping/orderFixed works):', isContiguous, 'indices:', JSON.stringify(prestigeIndices));

  // collapse again
  await toggle.click();
  await page.waitForTimeout(300);
  const rowsAfterCollapse = await frame.locator('#dataTable tbody tr').count();
  console.log('Row count after re-collapsing (should match initial):', rowsAfterCollapse, '| matches initial:', rowsAfterCollapse === totalRows);

  // Family1 filter auto-expand test
  const family1WidgetToggle = frame.locator('.ms-widget').nth(1).locator('.ms-toggle');
  await family1WidgetToggle.click();
  await page.waitForTimeout(200);
  await frame.locator('.ms-widget').nth(1).locator('.ms-search').fill('PRESTIGE AEROLIFT');
  await page.waitForTimeout(200);
  await frame.locator('.ms-widget').nth(1).locator('.ms-options input[type=checkbox]:visible').first().check();
  await page.waitForTimeout(400);
  const rowsAfterFamily1Filter = await frame.locator('#dataTable tbody tr').count();
  const visibleTexts = await frame.locator('#dataTable tbody tr td:nth-child(2)').allInnerTexts();
  console.log('After filtering Family1="PRESTIGE AEROLIFT..." -> row count:', rowsAfterFamily1Filter, '| visible Family1 values:', JSON.stringify(visibleTexts));

  await browser.close();
})().catch(e => { console.error('ERROR:', e); process.exit(1); });
