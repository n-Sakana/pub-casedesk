@echo off
echo === Build + Run CaseDesk (sample) ===
powershell -ExecutionPolicy Bypass -File "%~dp0scripts\Build-Addin.ps1" -Sample
if errorlevel 1 (
    echo Build failed.
    pause
    exit /b 1
)
echo.
echo Opening sample workbook and installing add-in...
powershell -ExecutionPolicy Bypass -File "%~dp0scripts\Run-Sample.ps1"
echo.
echo Done.
