# MTC Report Workflow

## 1. What the launchers do

Two menu-driven batch files provide entry points to the MTC monthly workflow:

- `Builder Data.bat` starts source checks, sales-data merge, or ACT001-to-Pur-Details conversion.
- `Builder HTML.bat` opens the browser staging tool, starts HTML generation, or runs the app-folder merge-script copy.

The launchers prompt for the workbook paths and reporting period required by each operation. Merge and conversion actions require a new output file path.

## 2. PowerShell roles

MTC PowerShell scripts are collected under `PowerShell/`.

- `PowerShell\Check-SourceTransactionCounts.ps1` checks that the three monthly source files agree on transaction row count.
- `PowerShell\Combine-Into-Master.ps1` merges the staged import into a new copy of the cleaned 95-column master, maps input fields by header, and preserves formula columns.
- `PowerShell\Convert-ACT001ToPurDetails.ps1` converts ACT001 purchase rows and appends them to Pur Details, subject to validation.
- `PowerShell\Generate-MTCSalesReportHtml.ps1` extracts the FIN_MTC report workbook data and creates the interactive HTML report.
- `PowerShell\App\Combine-Into-Master.ps1` is the merge-script copy paired with the browser staging tool.
- `PowerShell\Backup\Combine-Into-Master- BACKUP.ps1` is a backup and is not part of routine processing.

## 3. Updated SOP

The updated `MTC Sales Report Import SOP.docx` documents:

- The cleaned 95-column workbook as the monthly merge master, with the original 145-column workbook retained as the backup/reference.
- The 19 input fields that the merge script writes by header name.
- The 33 approved ignored legacy headers and the warning behavior for non-zero values.
- Protection of formula columns and the separate ACT001-to-Pur-Details process.

## 4. Current report-width limitation

The current HTML generator reads the 145-column `DATA` sheet from the FIN_MTC workbook. The cleaned 95-column MTC Sales Details master is used for the import workflow, but the generated HTML Data tab has not yet been changed to 95 columns. Update the report generator before treating a future HTML report as a 95-column output.

## 5. Validation status

Both batch menus were opened and exited successfully. This only checked launcher startup; the batch files were not used to run Excel-writing operations during this check. Follow the SOP and review script validation output when processing actual monthly workbooks.
