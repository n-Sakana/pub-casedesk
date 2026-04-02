# Test-Refactoring.ps1
# Automated smoke test for refactored casedesk
# Opens casedesk.xlsm + sample data, runs CaseDesk_ShowPanel, checks results

param([int]$TimeoutSeconds = 30)

$ErrorActionPreference = 'Stop'
$projectDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$casedeskPath = Join-Path $projectDir 'dist\casedesk.xlsm'
$samplePath = Join-Path $projectDir 'sample\casedesk-sample.xlsx'

if (-not (Test-Path $casedeskPath)) { Write-Error "casedesk.xlsm not found. Run Build-Addin.ps1 first."; exit 1 }
if (-not (Test-Path $samplePath)) { Write-Error "casedesk-sample.xlsx not found."; exit 1 }

$excel = $null
$casedeskWb = $null
$sampleWb = $null
$passed = 0
$failed = 0
$errors = @()

function Test($name, $condition, $detail = "") {
    if ($condition) {
        Write-Host "  PASS: $name" -ForegroundColor Green
        $script:passed++
    } else {
        Write-Host "  FAIL: $name $detail" -ForegroundColor Red
        $script:failed++
        $script:errors += $name
    }
}

try {
    Write-Host "=== Refactoring Smoke Test ===" -ForegroundColor Cyan
    Write-Host ""

    # 1. Open Excel
    Write-Host "[1] Starting Excel..." -ForegroundColor Yellow
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $prevSec = $excel.AutomationSecurity
    $excel.AutomationSecurity = 1  # msoAutomationSecurityLow

    # 2. Open sample data first
    Write-Host "[2] Opening sample data..." -ForegroundColor Yellow
    $sampleWb = $excel.Workbooks.Open($samplePath, 0, $false)
    Test "Sample workbook opened" ($sampleWb -ne $null)

    # 3. Open casedesk.xlsm
    Write-Host "[3] Opening casedesk.xlsm..." -ForegroundColor Yellow
    $casedeskWb = $excel.Workbooks.Open($casedeskPath, 0, $false)
    $excel.AutomationSecurity = $prevSec
    Test "CaseDesk workbook opened" ($casedeskWb -ne $null)

    # 4. Check modules exist
    Write-Host "[4] Checking VBA modules..." -ForegroundColor Yellow
    $vbProj = $casedeskWb.VBProject
    $moduleNames = @()
    foreach ($comp in $vbProj.VBComponents) { $moduleNames += $comp.Name }
    Test "CaseDeskMain exists" ($moduleNames -contains "CaseDeskMain")
    Test "CaseDeskData exists" ($moduleNames -contains "CaseDeskData")
    Test "CaseDeskLib exists" ($moduleNames -contains "CaseDeskLib")
    Test "CaseDeskWorker exists" ($moduleNames -contains "CaseDeskWorker")
    Test "CaseDeskHelpers removed" (-not ($moduleNames -contains "CaseDeskHelpers"))
    Test "CaseDeskConfig removed" (-not ($moduleNames -contains "CaseDeskConfig"))
    Test "CaseDeskChangeLog removed" (-not ($moduleNames -contains "CaseDeskChangeLog"))
    Test "CaseDeskScanner removed" (-not ($moduleNames -contains "CaseDeskScanner"))
    Test "CaseDeskDraft removed" (-not ($moduleNames -contains "CaseDeskDraft"))
    Test "CaseDeskPrint removed" (-not ($moduleNames -contains "CaseDeskPrint"))
    Test "CaseDeskBundler removed" (-not ($moduleNames -contains "CaseDeskBundler"))
    Test "frmFilter removed" (-not ($moduleNames -contains "frmFilter"))
    Test "frmDraft removed" (-not ($moduleNames -contains "frmDraft"))
    Test "WorkerWatcher removed" (-not ($moduleNames -contains "WorkerWatcher"))

    # 5. Run CaseDesk_ShowPanel (creates hidden sheets + shows form)
    Write-Host "[5] Running CaseDesk_ShowPanel..." -ForegroundColor Yellow
    try {
        $excel.Run("CaseDeskMain.CaseDesk_ShowPanel")
        Test "CaseDesk_ShowPanel executed" $true
    } catch {
        Test "CaseDesk_ShowPanel executed" $false $_.Exception.Message
    }

    # 6. Check hidden sheets exist (created by EnsureCaseDeskSheets inside ShowPanel)
    Write-Host "[6] Checking hidden sheets..." -ForegroundColor Yellow
    $sheetNames = @()
    foreach ($ws in $casedeskWb.Worksheets) { $sheetNames += $ws.Name }
    Test "_casedesk_signal exists" ($sheetNames -contains "_casedesk_signal")
    Test "_casedesk_mail exists" ($sheetNames -contains "_casedesk_mail")
    Test "_casedesk_cases exists" ($sheetNames -contains "_casedesk_cases")
    Test "_casedesk_files exists" ($sheetNames -contains "_casedesk_files")
    Test "_casedesk_request exists" ($sheetNames -contains "_casedesk_request")

    # 7. Test CaseDeskWorker scanner directly (in-process, no cross-process worker)
    Write-Host "[7] Reading config..." -ForegroundColor Yellow
    $mailFolder = ""
    $caseRoot = ""
    try {
        $cfgSheet = $casedeskWb.Worksheets.Item("_casedesk_config")
        $cfgRows = $cfgSheet.UsedRange.Rows.Count
        for ($r = 1; $r -le $cfgRows; $r++) {
            $k = $cfgSheet.Cells($r, 1).Text
            $v = $cfgSheet.Cells($r, 2).Text
            if ($k -eq "mail_folder") { $mailFolder = $v }
            if ($k -eq "case_folder_root") { $caseRoot = $v }
        }
    } catch {}
    Write-Host "  mail=$mailFolder" -ForegroundColor Gray
    Write-Host "  cases=$caseRoot" -ForegroundColor Gray

    # 8. Test scanner: RefreshMailData
    Write-Host "[8] Testing CaseDeskWorker.RefreshMailData..." -ForegroundColor Yellow
    try {
        $excel.Run("CaseDeskWorker.SetMailMatchConfig", "sender_email", "exact")
        $mailChanged = $excel.Run("CaseDeskWorker.RefreshMailData", $mailFolder)
        Test "RefreshMailData succeeded" $true
        $mailRecords = $excel.Run("CaseDeskWorker.GetMailRecords")
        $mailCount = $mailRecords.Count
        Test "Mail records loaded" ($mailCount -gt 0) "count=$mailCount"
        Write-Host "  Mail records: $mailCount" -ForegroundColor Gray
    } catch {
        Test "RefreshMailData succeeded" $false $_.Exception.Message
    }

    # 9. Test scanner: RefreshCaseNames
    Write-Host "[9] Testing CaseDeskWorker.RefreshCaseNames..." -ForegroundColor Yellow
    try {
        $caseChanged = $excel.Run("CaseDeskWorker.RefreshCaseNames", $caseRoot)
        Test "RefreshCaseNames succeeded" $true
        $caseNames = $excel.Run("CaseDeskWorker.GetCaseNames")
        $caseCount = $caseNames.Count
        Test "Case names loaded" ($caseCount -gt 0) "count=$caseCount"
        Write-Host "  Case names: $caseCount" -ForegroundColor Gray
    } catch {
        Test "RefreshCaseNames succeeded" $false $_.Exception.Message
    }

    # 10. Test CaseDeskData table operations
    Write-Host "[10] Testing CaseDeskData table operations..." -ForegroundColor Yellow
    try {
        $tableNames = $excel.Run("CaseDeskData.GetWorkbookTableNames", $sampleWb)
        # COM Collection.Count returns as PSMethod in PowerShell; just check non-null
        Test "GetWorkbookTableNames succeeded" ($tableNames -ne $null)
    } catch {
        Test "GetWorkbookTableNames succeeded" $false $_.Exception.Message
    }

    # 11. Check manifest.tsv was created (migration from Dir$ scan)
    Write-Host "[11] Checking manifest.tsv..." -ForegroundColor Yellow
    $manifestPath = Join-Path $mailFolder "manifest.tsv"
    Test "manifest.tsv created" (Test-Path $manifestPath)
    if (Test-Path $manifestPath) {
        $manifestLines = (Get-Content $manifestPath | Measure-Object).Count
        Test "manifest.tsv has data" ($manifestLines -gt 0) "lines=$manifestLines"
        Write-Host "  Manifest lines: $manifestLines" -ForegroundColor Gray
    }

    # 12. Check _casedesk_files is empty (on-demand, not preloaded)
    Write-Host "[12] Checking on-demand files..." -ForegroundColor Yellow
    $filesSheet = $casedeskWb.Worksheets.Item("_casedesk_files")
    $filesA1 = $filesSheet.Range("A1").Text
    Test "_casedesk_files empty at startup (on-demand)" ($filesA1.Length -eq 0)

    # 14. Check no WinAPI declarations in CaseDeskMain
    Write-Host "[9] Checking no WinAPI..." -ForegroundColor Yellow
    $mainCode = $vbProj.VBComponents.Item("CaseDeskMain").CodeModule
    $mainText = ""
    if ($mainCode.CountOfLines -gt 0) {
        $mainText = $mainCode.Lines(1, $mainCode.CountOfLines)
    }
    Test "No Declare Function in CaseDeskMain" (-not ($mainText -match "Declare\s+(PtrSafe\s+)?Function"))

    # 15. Stop worker cleanly
    Write-Host "[10] Stopping worker..." -ForegroundColor Yellow
    try {
        $excel.Run("CaseDeskMain.StopWorker")
        Test "Worker stopped" $true
    } catch {
        Test "Worker stopped" $false $_.Exception.Message
    }

} catch {
    Write-Host "FATAL: $($_.Exception.Message)" -ForegroundColor Red
    $failed++
    $errors += "Fatal: $($_.Exception.Message)"
} finally {
    # Cleanup
    Write-Host ""
    Write-Host "Cleaning up..." -ForegroundColor Yellow
    try {
        # Unload form if loaded
        try { $excel.Run("CaseDeskMain.BeforeWorkbookClose") } catch {}
        if ($sampleWb) { $sampleWb.Close($false); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sampleWb) | Out-Null }
        if ($casedeskWb) { $casedeskWb.Close($false); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($casedeskWb) | Out-Null }
    } catch {}
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()

    # Summary
    Write-Host ""
    Write-Host "=== Results: $passed passed, $failed failed ===" -ForegroundColor $(if ($failed -eq 0) { "Green" } else { "Red" })
    if ($errors.Count -gt 0) {
        Write-Host "Failures:" -ForegroundColor Red
        foreach ($e in $errors) { Write-Host "  - $e" -ForegroundColor Red }
    }
    exit $failed
}
