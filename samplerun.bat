@echo off
echo === CaseDesk Sample Run ===
echo.

echo [1/3] Generating sample data...
powershell -ExecutionPolicy Bypass -File "%~dp0scripts\Build-Sample.ps1" -Count 100
if errorlevel 1 (
    echo Sample data generation failed.
    pause
    exit /b 1
)
echo.

echo [2/3] Building add-in...
powershell -ExecutionPolicy Bypass -File "%~dp0scripts\Build-Addin.ps1" -Sample
if errorlevel 1 (
    echo Build failed.
    pause
    exit /b 1
)
echo.

echo [3/3] Opening Excel with sample workbook + add-in...
powershell -ExecutionPolicy Bypass -File "%~dp0scripts\Run-Sample.ps1"
echo.
echo Done. Use the CaseDesk tab (rightmost) to open the panel.
