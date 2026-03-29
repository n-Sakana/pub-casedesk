# Run-Sample.ps1
# xlamをアドインとして登録し、サンプル台帳を開く

$ErrorActionPreference = 'Stop'
$projectDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$xlamPath = Join-Path $projectDir 'casedesk.xlam'
$samplePath = Join-Path $projectDir 'sample\casedesk-sample.xlsx'

if (-not (Test-Path $xlamPath)) {
    Write-Host "ERROR: $xlamPath not found. Run Build-Addin.ps1 first." -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $samplePath)) {
    Write-Host "ERROR: $samplePath not found." -ForegroundColor Red
    exit 1
}

Write-Host "Starting Excel..." -ForegroundColor Cyan
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true

try {
    # Open sample workbook first (so it becomes ActiveWorkbook)
    Write-Host "  Opening sample: $samplePath"
    $sampleWb = $excel.Workbooks.Open($samplePath)

    # Install add-in
    Write-Host "  Installing add-in: $xlamPath"
    $excel.Workbooks.Open($xlamPath)

    Write-Host ''
    Write-Host 'Ready. Use CaseDesk menu > Show Panel to start.' -ForegroundColor Green

} catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
    exit 1
}

# Release COM (Excel stays open for user)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
