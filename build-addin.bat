@echo off

echo === Build CaseDesk Add-in ===
powershell -ExecutionPolicy Bypass -File "%~dp0scripts\Build-Addin.ps1" %*
if errorlevel 1 (
    echo Build failed.
    pause
    exit /b 1
)

echo.
echo Ready.
