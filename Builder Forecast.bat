@echo off
setlocal
cd /d "%~dp0"
set "ROOT=%~dp0"
set "PS=powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass"
set "FC=%ROOT%Forecast"

:menu
cls
echo ==========================================
echo            Builder Forecast
echo   (SO/PO combine outstanding forecast)
echo ==========================================
echo 1. Run month-end forecast NOW (build + verify)
echo 2. Store Epicor credentials (one-time, for auto-fetch)
echo 3. Fetch SO/PO detail from Epicor (needs creds)
echo 4. Build forecast from a SCF017 / SC17 file
echo 5. Verify forecast vs SCF017 / SC17 (qty compare)
echo 0. Exit
echo.
set "CHOICE="
set /p "CHOICE=Choose an action: "
if "%CHOICE%"=="1" goto runnow
if "%CHOICE%"=="2" goto setcred
if "%CHOICE%"=="3" goto fetch
if "%CHOICE%"=="4" goto buildsc
if "%CHOICE%"=="5" goto verify
if "%CHOICE%"=="0" goto done
goto menu

:runnow
%PS% -File "%FC%\Run-ForecastMonthEnd.ps1" -Force
goto finished

:setcred
%PS% -File "%FC%\Run-ForecastMonthEnd.ps1" -SetCredentials
goto finished

:fetch
echo Requires EPICOR_USER / EPICOR_PASS env, or stored creds via option 2.
%PS% -Command "& { $c='%FC%\epicor_credentials.dat'; if(Test-Path $c){ $j=Get-Content $c -Raw|ConvertFrom-Json; $s=ConvertTo-SecureString $j.EncryptedPassword; $env:EPICOR_USER=$j.Email; $env:EPICOR_PASS=[Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($s)) } ; python '%FC%\fetch_mtc_details.py' }"
goto finished

:buildsc
set "SC="
set /p "SC=Path to SCF017 / SC17 xlsx: "
if not defined SC goto menu
python "%FC%\build_forecast_from_scf017.py" "%SC%"
goto finished

:verify
set "FCX="
set "SCX="
set /p "FCX=Forecast xlsx (Enter = newest MTC_FORECAST_*_UPDATED): "
set /p "SCX=SCF017 / SC17 xlsx: "
if not defined SCX goto menu
if defined FCX (python "%FC%\verify_forecast_vs_sc17.py" "%FCX%" "%SCX%") else (python "%FC%\verify_forecast_vs_sc17.py" "" "%SCX%")
goto finished

:finished
echo.
echo Finished. Check the messages above for success or errors.
pause
goto menu

:done
endlocal
exit /b 0
