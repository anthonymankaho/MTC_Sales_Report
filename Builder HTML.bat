@echo off
setlocal
cd /d "%~dp0"
set "ROOT=%~dp0"
set "PS=powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass"

:menu
cls
echo ==========================================
echo              Builder HTML
echo ==========================================
echo 1. Open the browser-based import staging tool
echo 2. Generate the interactive HTML report
echo 3. Merge staged import using PowerShell\App\Combine-Into-Master.ps1
echo 0. Exit
echo.
set "CHOICE="
set /p "CHOICE=Choose an action: "
if "%CHOICE%"=="1" goto staging
if "%CHOICE%"=="2" goto report
goto mergecheck
:mergecheck
if "%CHOICE%"=="3" goto merge
if "%CHOICE%"=="0" goto done
goto menu

:staging
start "" "%ROOT%app\index.html"
echo Opened the browser-based staging tool.
goto finished

:report
echo WARNING: Current generator reads the FIN_MTC DATA sheet (145 columns).
echo It does not yet generate a 95-column Data tab.
%PS% -File "%ROOT%PowerShell\Generate-MTCSalesReportHtml.ps1"
if errorlevel 1 (echo ERROR: HTML report generation failed. Review the PowerShell message above.) else echo HTML report generation completed.
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
%PS% -File "%ROOT%PowerShell\App\Combine-Into-Master.ps1" -MasterPath "%MASTER%" -ImportPath "%IMPORT%" -OutputPath "%OUTPUT%"
if errorlevel 1 (echo ERROR: Merge failed. Review the PowerShell message above.) else echo Merge completed.
goto finished

:finished
echo.
echo PowerShell process finished. Check the messages above for success or errors.
pause
goto menu

:done
endlocal
exit /b 0
