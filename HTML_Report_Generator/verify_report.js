// Real-browser verification of the generated MTC Sales Report HTML.
// Usage: node verify_report.js <path-to-html-or-url> [--webkit]
const { chromium, webkit } = require('playwright');
const path = require('path');

const filePath = process.argv[2];
const useWebkit = process.argv.includes('--webkit');
if (!filePath) { console.error('Usage: node verify_report.js <path-to-html-or-url> [--webkit]'); process.exit(1); }

const fileUrl = /^https?:\/\//i.test(filePath) ? filePath : 'file:///' + path.resolve(filePath).replace(/\\/g, '/');

(async () => {
  const engine = useWebkit ? webkit : chromium;
  console.log('Engine:', useWebkit ? 'WebKit' : 'Chromium');
  const browser = await engine.launch();
  const page = await browser.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push('PAGE ERROR: ' + e.message));
  page.on('console', (msg) => { if (msg.type() === 'error') errors.push('CONSOLE ERROR: ' + msg.text()); });

  await page.goto(fileUrl, { waitUntil: 'load', timeout: 60000 });
  await page.waitForTimeout(500);

  // Access gate: reject a wrong code, then accept the configured MTC code.
  const accessGate = page.locator('#accessGate');
  if (await accessGate.count()) {
    const initiallyVisible = await accessGate.isVisible();
    await page.locator('#accessCode').fill('WRONG');
    await page.locator('#accessSubmit').click();
    const errorText = await page.locator('#accessError').textContent();
    const stillVisibleAfterWrong = await accessGate.isVisible();
    await page.locator('#accessCode').fill('MTC');
    await page.locator('#accessSubmit').click();
    await page.waitForTimeout(150);
    const hiddenAfterCorrect = !(await accessGate.isVisible());
    console.log('Access gate test: initial=' + initiallyVisible + ' wrongRejected=' +
      (stillVisibleAfterWrong && /Incorrect/.test(errorText || '')) + ' correctAccepted=' + hiddenAfterCorrect);
  }

  // Dismiss intro modal
  const introBtn = page.locator('#introClose');
  if (await introBtn.count()) { await introBtn.click(); }

  const tabNames = await page.locator('#tabs button').allTextContents();
  console.log('Tabs found:', tabNames.join(', '));

  const results = [];
  for (let i = 0; i < tabNames.length; i++) {
    await page.evaluate((idx) => { showSheet(idx); }, i);
    await page.waitForTimeout(400);
    const frame = page.frameLocator('#sheetFrame');
    let rowCount = null, kpiSales = null, chartCanvasCount = null, filterCount = null;
    try { rowCount = await frame.locator('#dataTable tbody tr').count(); } catch (e) {}
    try { kpiSales = await frame.locator('#kpi-sales').textContent({ timeout: 1000 }); } catch (e) {}
    try { chartCanvasCount = await frame.locator('canvas').count(); } catch (e) {}
    try { filterCount = await frame.locator('.ms-toggle, .filter-group select').count(); } catch (e) {}
    results.push({ tab: tabNames[i], rowCount, kpiSales, chartCanvasCount, filterCount });
  }

  console.log('--- Per-tab results ---');
  results.forEach((r) => console.log(JSON.stringify(r)));

  // Specifically exercise the Dashboard's filters to confirm they actually change output.
  await page.evaluate(() => { showSheet(0); });
  await page.waitForTimeout(400);
  const dashFrame = page.frameLocator('#sheetFrame');
  const salesBefore = await dashFrame.locator('#kpi-sales').textContent();
  // Open Year multi-select, uncheck "select all", check just the first real year option.
  await dashFrame.locator('#ms-year .ms-toggle').click();
  await page.waitForTimeout(150);
  const yearCheckboxes = dashFrame.locator('#ms-year .ms-panel input[type=checkbox]');
  const yearCbCount = await yearCheckboxes.count();
  if (yearCbCount > 1) {
    await yearCheckboxes.nth(1).click(); // first real year (index 0 is "select all")
    await page.waitForTimeout(400);
  }
  const salesAfter = await dashFrame.locator('#kpi-sales').textContent();
  console.log('Dashboard filter test: Sales before=' + salesBefore + ' after Year-filter=' + salesAfter + ' (should differ)');

  // Test theme toggle
  await page.locator('#themeToggle').click();
  await page.waitForTimeout(300);
  const dashBodyTheme = await dashFrame.locator('body').getAttribute('data-theme');
  console.log('Theme toggle test: dashboard iframe data-theme=' + dashBodyTheme + ' (should be dark)');

  // Test Data tab toolbar filter dropdown actually filters
  const dataTabIdx = tabNames.indexOf('Data');
  if (dataTabIdx >= 0) {
    await page.evaluate((idx) => { showSheet(idx); }, dataTabIdx);
    await page.waitForTimeout(500);
    const dataFrame = page.frameLocator('#sheetFrame');
    const rowsBefore = await dataFrame.locator('#dataTable tbody tr').count();
    const firstSelect = dataFrame.locator('.filter-group select').first();
    const optionValues = await firstSelect.locator('option').allTextContents();
    if (optionValues.length > 1) {
      await firstSelect.selectOption({ index: 1 });
      await page.waitForTimeout(400);
      const rowsAfter = await dataFrame.locator('#dataTable tbody tr').count();
      console.log('Data tab filter test: rows before=' + rowsBefore + ' after=' + rowsAfter + ' (after should be smaller)');
    }
  }

  console.log('--- Console/page errors captured ---');
  if (errors.length === 0) console.log('(none)');
  else errors.forEach((e) => console.log(e));

  // Logout clears the session-only gate and returns the viewer to login.
  const logoutBtn = page.locator('#accessLogout');
  if (await logoutBtn.count()) {
    await logoutBtn.click();
    await page.waitForLoadState('load');
    await page.waitForTimeout(150);
    console.log('Logout test: gateVisible=' + await page.locator('#accessGate').isVisible());
  }

  await browser.close();
  console.log('Verification run complete.');
})().catch((e) => { console.error('FATAL:', e); process.exit(1); });
