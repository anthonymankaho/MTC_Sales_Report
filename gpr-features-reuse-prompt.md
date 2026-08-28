# Reusable prompt: bring Global Purchase Report's HTML features to MTC Sales/Purchase Report

Copy this whole block into a new Claude Code conversation. It assumes the
target report was (or will be) built with the iframe-per-sheet single-file
architecture documented in the `excel-to-interactive-html-report` skill.

---

Please load the `excel-to-interactive-html-report` skill first (read
`C:\Users\User\.claude\skills\excel-to-interactive-html-report\SKILL.md`),
then read the **Global Purchase Report** HTML report generator as the
reference implementation:
`C:\Users\User\Desktop\Global Purchase Report Claude\Global Purchase Report RPA\HTML_Report_Generator\build_report.js`
(and `Generate-PurchaseReportHtml.ps1` for the TEMP/`-Publish` snapshot
convention it uses). Bring the following features to **[TARGET REPORT NAME]**:

- MTC Sales Report → `C:\Users\User\Desktop\MTC Sales Report Claude\HTML_Report_Generator\build_report.js`
  already exists and already has its own on-chart data labels, a NET/NET NET
  basis toggle, cross-tab drill-down, and an English/Traditional Chinese
  language toggle (see the older reuse-prompt at
  `C:\Users\User\Desktop\html-report-features-prompt.md` for those). Treat
  this prompt as filling the **remaining gap** against GPR's feature set below
  - read both files first and skip anything MTC Sales already has.
- MTC Purchase Report → `C:\Users\User\Desktop\MTC Purchase Report Claude\`
  has no interactive HTML report yet (only `Extract-ForHtmlReport.ps1`, a
  data-prep step). This is a **full build**, not a gap-fill: use GPR's
  `build_report.js` as the architectural starting point (iframe-per-sheet
  shell, Dashboard + per-sheet detail tabs, TEMP-working-file +
  `-Publish`-snapshot convention) and implement every feature below against
  MTC Purchase Report's actual sheet/column structure - read its Excel
  deliverable first to map GPR's concepts (DB/Std Comm/Vendor/Item Code
  filters, entity breakdown) onto whatever columns MTC Purchase Report
  actually has; don't assume identical column names.

## Features to bring over (all present in GPR as of 2026-09-08)

1. **Search box on every filter dropdown** - every multi-select filter across
   every tab (Dashboard, and every per-sheet detail tab) gets a text input
   that narrows the checkbox list as you type; it only filters which options
   are *shown*, it never changes what's currently *checked*. This was GPR's
   most recent upgrade - read `build_report.js` for the exact multi-select
   widget pattern (`.ms-toggle` / `.ms-panel` / search input wiring) rather
   than re-deriving it.

2. **A dynamic, correctly-scoped product/category filter** on at least one
   detail tab - GPR's example is the Fittings tab's `ProdFamiFittingsCode`
   filter, which pulls only the right family codes for that sheet (e.g.
   searching "HD" finds HDG/HDI/HDS/HDV) rather than every code in the whole
   workbook. Identify the equivalent grouping column in the target report and
   scope its options the same way - verify with a real headless-browser pass
   that the option list is actually right for that one sheet, not leaked from
   another.

3. **Live filters computed from raw transaction data, not a pivot cache** -
   every filter (Month/DB/Vendor/Item Code/etc.) must be computed from the
   Data sheet's actual rows at page-load/filter-change time, so changing a
   filter drives both the charts AND every detail table together, in sync.

4. **Item-level collapse toggle (+/-)** on any detail tab where the same item
   can have multiple DB/Vendor rows - a +/- button collapses/expands those
   rows for one item without hiding the underlying data.

5. **Drag-to-resize columns** on every detail table, persisted per-viewer in
   `localStorage` so the layout survives a reload.

6. **Click-to-drill-down from BOTH a Dashboard category/chart AND a table
   cell** (not just a summary/pivot tab) - clicking an ITEM_CODE/DB/Vendor
   cell, or a Dashboard category slice, jumps to the Data tab pre-filtered to
   exactly those rows. Verify the exact same way as MTC Sales's drill-down
   (numeric reconciliation: filtered-row sum matches the clicked value) AND
   watch for the specific race condition GPR hit here - the destination tab
   must *pull* its filter state on load rather than the parent *pushing* it
   into an iframe that may not be ready yet.

7. **A CSV-exportable, sortable/searchable interactive table** for any
   sheet that's currently static (GPR did this for its Exchange Rate tab) -
   same DataTables treatment as the main Data tab, export reads the
   *currently filtered* rows (`table.rows({search:'applied'}).data()`), with
   a UTF-8 BOM prepended so Excel doesn't mangle non-ASCII text on open.

8. **Dashboard breakdown pie chart(s) + matching KPI cards with a weighted
   %** - GPR's example is its SKU vs Part Family breakdown (Top 10 + Other
   slices) with KPI cards split into matching sections, each showing its own
   weighted percentage. Pick the equivalent breakdown dimension(s) for the
   target report.

9. **A welcome/intro popup** on first load, dismissed and remembered via a
   versioned `localStorage` key (bump the key whenever the popup content
   changes materially, so returning viewers see the update once even if
   they'd dismissed an older version).

10. **TEMP-working-file + `-Publish`-snapshot convention** for the generator
    script itself (if not already present) - default run overwrites a fixed
    `..._TEMP.html` safe to re-run while testing; `-Publish` additionally
    writes a period-stamped snapshot into the report folder. Source-file
    auto-detection (if the generator picks its own source Excel) must sort by
    the reporting period encoded in the filename, not file modification time.

## Verification

For every feature above: verify with a real Playwright pass (headless
Chromium for general checks, real `webkit` engine for anything mobile-related
per the skill's gotchas) that actually taps the UI and reads back live DOM/JS
state - not code review alone. Re-run the drill-down and filter checks after
toggling dark/light mode and (for MTC Sales) after switching language, since
that combination is exactly how the skill's known gotchas were originally
caught. Clean up any temporary test scripts afterward and open the final
file. Reply in both English and Traditional Chinese per my standing
preference.
