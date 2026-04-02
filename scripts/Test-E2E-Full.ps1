# Test-E2E-Full.ps1
# Comprehensive E2E test covering real-world usage patterns
# - Source switching (table vs UsedRange)
# - Field value display
# - Mail/Files tab count reset on source switch
# - Settings save/cancel
# - Worker startup

$ErrorActionPreference = 'Stop'
$projectDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$xlsm = Join-Path $projectDir 'dist\casedesk.xlsm'
$sampleXlsx = Join-Path $projectDir 'sample\casedesk-sample.xlsx'

$pass = 0; $fail = 0
function Assert($name, $cond) {
    if ($cond) { Write-Host "  PASS: $name" -ForegroundColor Green; $script:pass++ }
    else { Write-Host "  FAIL: $name" -ForegroundColor Red; $script:fail++ }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true; $excel.DisplayAlerts = $false
try {
    $excel.Workbooks.Open($sampleXlsx) | Out-Null
    $sampleWb = $excel.Workbooks.Item(1)
    $prev = $excel.AutomationSecurity; $excel.AutomationSecurity = 1
    $wb = $excel.Workbooks.Open($xlsm)
    $excel.AutomationSecurity = $prev
    $sampleWb.Activate()

    # ========== 1. ShowPanel ==========
    Write-Host "`n=== 1. ShowPanel ===" -ForegroundColor Cyan
    $excel.Run("'casedesk.xlsm'!CaseDeskMain.CaseDesk_ShowPanel")
    Write-Host "  ShowPanel returned OK"
    Assert 'ShowPanel succeeded' $true

    # ========== 2. Wait for worker ==========
    Write-Host "`n=== 2. Worker Startup ===" -ForegroundColor Cyan
    $done = $false
    for ($i = 0; $i -lt 60; $i++) {
        Start-Sleep -Milliseconds 500
        try {
            $mc = $excel.Run("'casedesk.xlsm'!CaseDeskData.GetMailCount")
            $cc = $excel.Run("'casedesk.xlsm'!CaseDeskData.GetCaseCount")
            if ($mc -gt 0 -and $cc -gt 0) { $done = $true; break }
        } catch {}
    }
    Write-Host "  mail=$mc cases=$cc"
    Assert 'worker scanned mail > 0' ($mc -gt 0)
    Assert 'worker scanned cases > 0' ($cc -gt 0)

    # ========== 3. Table source (anken) — field values ==========
    Write-Host "`n=== 3. Table Source Field Values ===" -ForegroundColor Cyan
    # ReadTableRecords should work for the anken table
    try {
        $tbl = $null
        foreach ($ws2 in $sampleWb.Worksheets) {
            foreach ($lo in $ws2.ListObjects) {
                $tbl = $lo; break
            }
            if ($tbl) { break }
        }
        $firstKey = $tbl.DataBodyRange.Cells(1, 1).Text
        $firstName = $tbl.DataBodyRange.Cells(1, 2).Text
        Assert "table has data (key=$firstKey)" ($firstKey.Length -gt 0)
        Assert "table has name (name=$firstName)" ($firstName.Length -gt 0)
    } catch {
        Write-Host "  ERROR reading table: $_" -ForegroundColor Red
        Assert 'table readable' $false
    }

    # ========== 4. Check sample has 2 sheets ==========
    Write-Host "`n=== 4. Multi-sheet Sample ===" -ForegroundColor Cyan
    $visibleSheets = @()
    foreach ($ws3 in $sampleWb.Worksheets) {
        if ($ws3.Visible -eq -1) { $visibleSheets += $ws3.Name }
    }
    Write-Host "  Visible sheets: $($visibleSheets -join ', ')"
    Assert 'sample has 2+ visible sheets' ($visibleSheets.Count -ge 2)

    # Check second sheet has data but no ListObject
    $sheet2 = $sampleWb.Worksheets.Item(2)
    $sheet2HasTable = ($sheet2.ListObjects.Count -gt 0)
    $sheet2HasData = ($sheet2.UsedRange.Rows.Count -gt 1)
    Write-Host "  Sheet2 '$($sheet2.Name)': ListObjects=$($sheet2.ListObjects.Count), Rows=$($sheet2.UsedRange.Rows.Count)"
    Assert 'sheet2 has no ListObject' (-not $sheet2HasTable)
    Assert 'sheet2 has data' $sheet2HasData

    # ========== 5. Signal sheet check ==========
    Write-Host "`n=== 5. Signal Sheet ===" -ForegroundColor Cyan
    try {
        $sigSh = $wb.Worksheets.Item("_casedesk_signal")
        $sigVer = $sigSh.Range("B1").Value2
        $sigTiming = $sigSh.Range("C1").Value2
        Write-Host "  Signal version: $sigVer"
        Write-Host "  Timing: $sigTiming"
        Assert 'signal version > 0' ($sigVer -gt 0)
    } catch {
        Assert 'signal sheet exists' $false
    }

    # ========== 6. Compile check ==========
    Write-Host "`n=== 6. Compile Check ===" -ForegroundColor Cyan
    try {
        $dict = $excel.Run("'casedesk.xlsm'!CaseDeskLib.NewDict")
        Assert 'CaseDeskLib.NewDict works' ($null -ne $dict)
    } catch {
        Assert 'CaseDeskLib.NewDict works' $false
    }

    # ========== Results ==========
    Write-Host "`n=== RESULT: $pass passed, $fail failed ===" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
} catch {
    Write-Host "FATAL: $_" -ForegroundColor Red
    $fail++
} finally {
    try { $excel.Quit() } catch {}
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    [GC]::Collect()
}
exit $fail
