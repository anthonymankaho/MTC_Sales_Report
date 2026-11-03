@echo off
setlocal
cd /d "%~dp0"
set "ROOT=%~dp0"
set "PS=powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass"

:menu
cls
echo ==========================================
echo              Builder Data
echo ==========================================
echo 1. Check monthly source transaction counts
echo 2. Merge staged import into a new 95-column master copy
echo 3. Convert ACT001 rows to Pur Details on a separate copy
echo 0. Exit
echo.
set "CHOICE="
set /p "CHOICE=Choose an action: "
if "%CHOICE%"=="1" goto counts
if "%CHOICE%"=="2" goto merge
if "%CHOICE%"=="3" goto act001
if "%CHOICE%"=="0" goto done
goto menu

:counts
%PS% -File "%ROOT%PowerShell\Check-SourceTransactionCounts.ps1"
if errorlevel 1 (echo ERROR: Source counts do not match or the check failed.) else echo Source count check completed.
goto finished

:merge
set "MASTER=%ROOT%Sales-Data-MTC-Mthly Rate - UPDATED_20261103_CLEANED_95COL.xlsx"
set "IMPORT=%ROOT%MTC_Import_Rows.xlsx"
set "OUTPUT="
echo Press Enter to accept the default master and import paths.
set /p "MASTER=Cleaned 95-column master [%MASTER%]: "
if not defined MASTER set "MASTER=%ROOT%Sales-Data-MTC-Mthly Rate - UPDATED_20261103_CLEANED_95COL.xlsx"
set /p "IMPORT=Staged import workbook [%IMPORT%]: "
if not defined IMPORT set "IMPORT=%ROOT%MTC_Import_Rows.xlsx"
set /p "OUTPUT=New output workbook path (include the reporting month): "
if not defined OUTPUT goto menu
if exist "%OUTPUT%" (echo ERROR: Output already exists. Choose a new output path.& goto finished)
%PS% -File "%ROOT%PowerShell\Combine-Into-Master.ps1" -MasterPath "%MASTER%" -ImportPath "%IMPORT%" -OutputPath "%OUTPUT%"
if errorlevel 1 (echo ERROR: Merge failed. Review the PowerShell message above.) else echo Merge completed.
goto finished

:act001
set "WORKBOOK="
set "SHEET="
set "MONTH="
set "OUTPUT="
set /p "WORKBOOK=Workbook path containing ACT001_MMYY and Pur_Combined_MMYY: "
if not defined WORKBOOK goto menu
set /p "SHEET=ACT001 sheet name (for example ACT001_0926): "
if not defined SHEET goto menu
set /p "MONTH=Reporting month YYYY-MM (for example 2026-09): "
if not defined MONTH goto menu
set /p "OUTPUT=Separate output workbook path (must be a new file): "
if not defined OUTPUT goto menu
if exist "%OUTPUT%" (echo ERROR: Output already exists. Choose a new output path.& goto finished)
%PS% -File "%ROOT%PowerShell\Convert-ACT001ToPurDetails.ps1" -WorkbookPath "%WORKBOOK%" -ACT001Sheet "%SHEET%" -ReportingMonth "%MONTH%" -OutputPath "%OUTPUT%"
if errorlevel 1 (echo ERROR: ACT001 conversion failed. Review the PowerShell message above.) else echo ACT001 conversion completed.
goto finished

:finished
echo.
echo PowerShell process finished. Check the messages above for success or errors.
pause
goto menu

:done
endlocal
exit /b 0
