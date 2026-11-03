# MTC PowerShell Script List

## Workflow launchers

| Launcher | Menu actions |
|---|---|
| `Builder Data.bat` | Check monthly source transaction counts; merge staged rows using `PowerShell\Combine-Into-Master.ps1`; convert ACT001 purchase rows to Pur Details using a separate output workbook. |
| `Builder HTML.bat` | Open the browser staging tool; generate the interactive HTML report; merge staged rows using `PowerShell\App\Combine-Into-Master.ps1`. |

## PowerShell scripts

Project scripts are grouped under `PowerShell/`.

| Script | Purpose | Used by |
|---|---|---|
| `PowerShell\Check-SourceTransactionCounts.ps1` | Checks transaction row counts in IVR032, Sales & Commission, and Purchase Commission source files. Stops the user from proceeding with mismatched source exports. | `Builder Data.bat`, option 1 |
| `PowerShell\Combine-Into-Master.ps1` | Merges `MTC_Import_Rows.xlsx` into a new copy of the cleaned 95-column master. Maps the approved input fields by header and protects formula columns. Requires master, import, and output paths. | `Builder Data.bat`, option 2 |
| `PowerShell\Convert-ACT001ToPurDetails.ps1` | Converts a loaded `ACT001_MMYY` sheet, enriches it using the matching `Pur_Combined_MMYY` and Exchange Rate data, then appends rows to Pur Details. Use a separate output path. | `Builder Data.bat`, option 3 |
| `PowerShell\Generate-MTCSalesReportHtml.ps1` | Extracts data from the FIN_MTC Excel report and builds the interactive HTML report. It auto-selects a matching source report unless a source path is supplied. | `Builder HTML.bat`, option 2 |
| `PowerShell\App\Combine-Into-Master.ps1` | Separate merge-script copy paired with the browser-based staging tool. | `Builder HTML.bat`, option 3 |
| `PowerShell\Backup\Combine-Into-Master- BACKUP.ps1` | Backup copy. Not part of the normal run sequence. | Do not use for routine processing |

## Operating notes

- Use the cleaned 95-column master as the merge input. Keep the original 145-column master as a backup/reference.
- For merge and ACT001 conversion, choose a new output filename for each run. The launchers reject an output path that already exists.
- Close Excel workbooks before running Excel COM automation scripts.
- The launchers have been checked to open and exit their menus. This check did not run the Excel-writing operations.
- Current limitation: the HTML report generator still reads the 145-column `DATA` sheet in the FIN_MTC workbook. It does not yet produce a 95-column Data tab.
