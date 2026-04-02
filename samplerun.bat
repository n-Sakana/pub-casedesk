@echo off

echo === Build + Run CaseDesk (sample) ===
powershell -NoProfile -ExecutionPolicy Bypass -Command "if (Get-Process EXCEL -ErrorAction SilentlyContinue) { Write-Host 'Close all Excel windows before running samplerun.bat.' -ForegroundColor Red; exit 1 }"
if errorlevel 1 (
    pause
    exit /b 1
)

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

powershell -NoProfile -ExecutionPolicy Bypass -Command "$deadline=(Get-Date).AddSeconds(15); while(Get-Process EXCEL -ErrorAction SilentlyContinue){ if((Get-Date) -ge $deadline){ Write-Host 'Excel did not shut down after build.' -ForegroundColor Red; exit 1 }; Start-Sleep -Milliseconds 500 }"
if errorlevel 1 (
    pause
    exit /b 1
)
echo.

echo [3/3] Opening Excel with sample workbook + add-in...
powershell -ExecutionPolicy Bypass -File "%~dp0scripts\Run-Sample.ps1"
if errorlevel 1 (
    echo Run sample failed.
    pause
    exit /b 1
)

echo.
echo Done. Use the CaseDesk tab (rightmost) to open the panel.
