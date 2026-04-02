# Test-E2E-Full.ps1
# Comprehensive E2E test — exercises all user-facing scenarios via COM
# No GUI interaction needed; calls VBA methods directly.

$ErrorActionPreference = 'Stop'
$projectDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$xlsm = Join-Path $projectDir 'dist\casedesk.xlsm'
$sampleXlsx = Join-Path $projectDir 'sample\casedesk-sample.xlsx'

$pass = 0; $fail = 0
function Assert($name, $cond) {
    if ($cond) { Write-Host "  PASS: $name" -ForegroundColor Green; $script:pass++ }
    else { Write-Host "  FAIL: $name" -ForegroundColor Red; $script:fail++ }
}

# Excel.Run wrapper
function XRun {
    $m = "'casedesk.xlsm'!" + $args[0]
    switch ($args.Count - 1) {
        0 { return $script:excel.Run($m) }
        1 { return $script:excel.Run($m, $args[1]) }
        2 { return $script:excel.Run($m, $args[1], $args[2]) }
        3 { return $script:excel.Run($m, $args[1], $args[2], $args[3]) }
        4 { return $script:excel.Run($m, $args[1], $args[2], $args[3], $args[4]) }
    }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true; $excel.DisplayAlerts = $false
try {
    $sampleWb = $excel.Workbooks.Open($sampleXlsx)
    $prev = $excel.AutomationSecurity; $excel.AutomationSecurity = 1
    $wb = $excel.Workbooks.Open($xlsm)
    $excel.AutomationSecurity = $prev
    $sampleWb.Activate()

    # ================================================================
    Write-Host "`n=== 1. Sample Data Structure ===" -ForegroundColor Cyan
    # ================================================================
    $visSheets = @()
    foreach ($ws in $sampleWb.Worksheets) { if ($ws.Visible -eq -1) { $visSheets += $ws.Name } }
    Assert 'sample has 2+ visible sheets' ($visSheets.Count -ge 2)

    $tblSheet = $sampleWb.Worksheets.Item(1)
    Assert 'sheet1 has ListObject' ($tblSheet.ListObjects.Count -gt 0)
    $tbl = $tblSheet.ListObjects.Item(1)
    $tblRowCount = $tbl.DataBodyRange.Rows.Count
    $tblName = $tbl.Name
    Assert "table '$tblName' has rows ($tblRowCount)" ($tblRowCount -gt 0)

    $sheet2 = $sampleWb.Worksheets.Item(2)
    $sheet2Name = $sheet2.Name
    Assert "sheet2 '$sheet2Name' has no ListObject" ($sheet2.ListObjects.Count -eq 0)
    $sheet2Rows = $sheet2.UsedRange.Rows.Count
    Assert "sheet2 has data ($sheet2Rows rows)" ($sheet2Rows -gt 1)

    $key1 = $tbl.DataBodyRange.Cells(1, 1).Text
    $name1 = $tbl.DataBodyRange.Cells(1, 2).Text
    Assert "first record key='$key1'" ($key1.Length -gt 0)
    Assert "first record name present" ($name1.Length -gt 0)
    $keyLast = $tbl.DataBodyRange.Cells($tblRowCount, 1).Text
    Assert "last record key='$keyLast'" ($keyLast.Length -gt 0)

    # Second sheet header row
    $s2Header1 = $sheet2.Cells.Item(1, 1).Text
    Assert "sheet2 header row has content ('$s2Header1')" ($s2Header1.Length -gt 0)
    $s2Data1 = $sheet2.Cells.Item(2, 1).Text
    Assert "sheet2 first data row has content ('$s2Data1')" ($s2Data1.Length -gt 0)

    # ================================================================
    Write-Host "`n=== 2. CaseDeskLib Config ===" -ForegroundColor Cyan
    # ================================================================
    XRun 'CaseDeskLib.EnsureConfigSheets'
    $mailDir = XRun 'CaseDeskLib.GetStr' 'mail_folder'
    $caseDir = XRun 'CaseDeskLib.GetStr' 'case_folder_root'
    Assert "mail_folder configured" ($mailDir.Length -gt 0)
    Assert "case_folder_root configured" ($caseDir.Length -gt 0)

    # ================================================================
    Write-Host "`n=== 3. ShowPanel + Worker ===" -ForegroundColor Cyan
    # ================================================================
    XRun 'CaseDeskMain.CaseDesk_ShowPanel'
    Assert 'ShowPanel succeeded' $true

    $done = $false
    for ($i = 0; $i -lt 60; $i++) {
        Start-Sleep -Milliseconds 500
        try {
            $mc = XRun 'CaseDeskData.GetMailCount'
            $cc = XRun 'CaseDeskData.GetCaseCount'
            if ($mc -gt 0 -and $cc -gt 0) { $done = $true; break }
        } catch {}
    }
    Write-Host "  mail=$mc cases=$cc"
    Assert 'worker scanned mail > 0' ($mc -gt 0)
    Assert 'worker scanned cases > 0' ($cc -gt 0)

    # Signal sheet
    $sigSh = $wb.Worksheets.Item("_casedesk_signal")
    $sigVer = $sigSh.Range("B1").Value2
    Assert "signal version > 0 (v=$sigVer)" ($sigVer -gt 0)

    # Mail/Cases hidden sheets have data
    $mailSh = $wb.Worksheets.Item("_casedesk_mail")
    $mailA1 = $mailSh.Range("A1").Value2
    Assert 'mail sheet populated' ($null -ne $mailA1 -and "$mailA1".Length -gt 0)
    $casesSh = $wb.Worksheets.Item("_casedesk_cases")
    $casesA1 = $casesSh.Range("A1").Value2
    Assert 'cases sheet populated' ($null -ne $casesA1 -and "$casesA1".Length -gt 0)

    # ================================================================
    Write-Host "`n=== 4. Source Config CRUD ===" -ForegroundColor Cyan
    # ================================================================
    # EnsureSource + Set/Get
    XRun 'CaseDeskLib.EnsureSource' $tblName
    XRun 'CaseDeskLib.SetSourceStr' $tblName 'source_sheet' $tblSheet.Name
    XRun 'CaseDeskLib.SetSourceStr' $tblName 'key_column' $tbl.ListColumns.Item(1).Name
    XRun 'CaseDeskLib.SetSourceStr' $tblName 'display_name_column' $tbl.ListColumns.Item(2).Name

    $keyCol = XRun 'CaseDeskLib.GetSourceStr' $tblName 'key_column'
    Assert "key_column = '$keyCol'" ($keyCol.Length -gt 0)
    $dispCol = XRun 'CaseDeskLib.GetSourceStr' $tblName 'display_name_column'
    Assert "display_name_column = '$dispCol'" ($dispCol.Length -gt 0)

    # Non-table source
    XRun 'CaseDeskLib.EnsureSource' $sheet2Name
    XRun 'CaseDeskLib.SetSourceStr' $sheet2Name 'source_sheet' $sheet2Name
    $srcSheet = XRun 'CaseDeskLib.GetSourceStr' $sheet2Name 'source_sheet'
    Assert "non-table source_sheet = '$srcSheet'" ($srcSheet -eq $sheet2Name)

    # ================================================================
    Write-Host "`n=== 5. SaveToSheets + Persistence ===" -ForegroundColor Cyan
    # ================================================================
    XRun 'CaseDeskLib.SaveToSheets'

    # Read _casedesk_sources sheet directly
    $srcSh = $wb.Worksheets.Item("_casedesk_sources")
    Assert 'sources sheet header' ($srcSh.Range("A1").Value2 -eq 'source_name')
    $found = $false
    $lastRow = $srcSh.UsedRange.Rows.Count
    for ($r = 2; $r -le $lastRow; $r++) {
        if ($srcSh.Cells.Item($r, 1).Value2 -eq $tblName) { $found = $true; break }
    }
    Assert "source '$tblName' persisted" $found

    # ================================================================
    Write-Host "`n=== 6. Cancel Rollback ===" -ForegroundColor Cyan
    # ================================================================
    XRun 'CaseDeskLib.SetSourceStr' $tblName 'mail_match_mode' 'domain'
    $before = XRun 'CaseDeskLib.GetSourceStr' $tblName 'mail_match_mode'
    Assert 'in-memory mutation applied' ($before -eq 'domain')

    XRun 'CaseDeskLib.LoadFromSheets'
    $after = XRun 'CaseDeskLib.GetSourceStr' $tblName 'mail_match_mode' 'exact'
    Assert "rollback discarded mutation ($after)" ($after -ne 'domain')

    # ================================================================
    Write-Host "`n=== 7. FindMailRecords ===" -ForegroundColor Cyan
    # ================================================================
    $testEmail = $mailSh.Range("B1").Value2
    if ($testEmail -and "$testEmail".Length -gt 0) {
        Write-Host "  Testing: $testEmail"
        $matchCount = 0
        try {
            $matches = XRun 'CaseDeskData.FindMailRecords' $testEmail 'sender_email' 'exact'
            $matchCount = $matches.Count
        } catch { Write-Host "  FindMailRecords error: $_" -ForegroundColor Yellow }
        Assert "FindMailRecords matched ($matchCount)" ($matchCount -gt 0)
    } else {
        Assert 'mail data available' $false
    }

    # ================================================================
    Write-Host "`n=== 8. ReadTableRecords ===" -ForegroundColor Cyan
    # ================================================================
    try {
        $foundTbl = XRun 'CaseDeskData.FindTable' $sampleWb $tblName
        Assert 'FindTable found table' ($null -ne $foundTbl)
        $records = XRun 'CaseDeskData.ReadTableRecords' $foundTbl
        $recCount = $records.Count
        Assert "ReadTableRecords: $recCount records (expect $tblRowCount)" ($recCount -eq $tblRowCount)
    } catch {
        Write-Host "  Error: $_" -ForegroundColor Yellow
        Assert 'ReadTableRecords succeeded' $false
    }

    # ================================================================
    Write-Host "`n=== 9. Field Config ===" -ForegroundColor Cyan
    # ================================================================
    # EnsureField + Get/Set
    $col1Name = $tbl.ListColumns.Item(1).Name
    XRun 'CaseDeskLib.EnsureField' $tblName $col1Name
    XRun 'CaseDeskLib.SetFieldStr' $tblName $col1Name 'role' 'case_id'
    $role = XRun 'CaseDeskLib.GetFieldStr' $tblName $col1Name 'role'
    Assert "field role = '$role'" ($role -eq 'case_id')

    XRun 'CaseDeskLib.SetFieldBool' $tblName $col1Name 'visible' $true
    $vis = XRun 'CaseDeskLib.GetFieldBool' $tblName $col1Name 'visible'
    Assert 'field visible = True' ($vis -eq $true)

    # ================================================================
    Write-Host "`n=== 10. Edge Cases ===" -ForegroundColor Cyan
    # ================================================================
    try {
        $dict = XRun 'CaseDeskLib.NewDict'
        Assert 'NewDict works' $true
    } catch { Assert 'NewDict works' $false }

    $missing = XRun 'CaseDeskLib.GetSourceStr' 'nonexistent_src' 'key_column' 'fallback'
    Assert "missing source default = '$missing'" ($missing -eq 'fallback')

    try {
        XRun 'CaseDeskLib.EnsureField' 'test_src_xyz' 'test_field'
        Assert 'EnsureField on new source OK' $true
    } catch { Assert 'EnsureField on new source OK' $false }

    # GetSourceStr with empty default
    $empty = XRun 'CaseDeskLib.GetSourceStr' 'nonexistent_src' 'anything'
    Assert "missing source empty default = '$empty'" ($empty.Length -eq 0)

    # ================================================================
    Write-Host "`n=== RESULT: $pass passed, $fail failed ===" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
} catch {
    Write-Host "FATAL: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    $fail++
} finally {
    try { $excel.Quit() } catch {}
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    [GC]::Collect()
}
exit $fail
