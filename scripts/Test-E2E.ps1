# Test-E2E.ps1
# End-to-end test: Worker scan -> FE sheet write -> FE cache load -> matching

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
$xlsmPath = Join-Path $projectDir 'dist\casedesk.xlsm'
$sampleDir = Join-Path $projectDir 'sample'

if (-not (Test-Path $xlsmPath)) { Write-Host "ERROR: casedesk.xlsm not found." -ForegroundColor Red; exit 1 }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $prevSec = $excel.AutomationSecurity
    $excel.AutomationSecurity = 1

    # Open sample first so it becomes ActiveWorkbook
    $samplePath = Join-Path $sampleDir 'casedesk-sample.xlsx'
    $sampleWb = $excel.Workbooks.Open($samplePath, 0, $false)

    $wb = $excel.Workbooks.Open($xlsmPath)
    $excel.AutomationSecurity = $prevSec

    $tempFile = Join-Path $env:TEMP "casedesk_e2e_test.txt"
    if (Test-Path $tempFile) { Remove-Item $tempFile -Force }

    $mailDir = (Join-Path $sampleDir 'mail') -replace '\\','\\'
    $casesDir = (Join-Path $sampleDir 'cases') -replace '\\','\\'

    $testCode = @"
Option Explicit

Public Sub RunE2ETests()
    Dim fnum As Integer
    fnum = FreeFile
    Open "$tempFile" For Output As #fnum

    On Error Resume Next

    Dim mailFolder As String: mailFolder = "$mailDir"
    Dim casesFolder As String: casesFolder = "$casesDir"

    ' --- Step 1: Config ---
    CaseDeskLib.EnsureConfigSheets
    Print #fnum, "INFO  | mail_folder | " & CaseDeskLib.GetStr("mail_folder")
    Print #fnum, "INFO  | case_folder_root | " & CaseDeskLib.GetStr("case_folder_root")
    Dim sources As Collection: Set sources = CaseDeskLib.GetSourceNames()
    Print #fnum, "INFO  | sources | " & sources.Count
    Dim srcName As String
    If sources.Count > 0 Then srcName = CStr(sources(1))
    Print #fnum, "INFO  | source | " & srcName
    Print #fnum, "INFO  | mail_link_column | " & CaseDeskLib.GetSourceStr(srcName, "mail_link_column")
    Print #fnum, "INFO  | folder_link_column | " & CaseDeskLib.GetSourceStr(srcName, "folder_link_column")

    ' --- Step 2: Worker scan ---
    CaseDeskWorker.SetMailMatchConfig "sender_email", "exact"
    CaseDeskWorker.RefreshMailData mailFolder
    CaseDeskWorker.RefreshCaseData casesFolder
    Dim mailRecs As Object: Set mailRecs = CaseDeskWorker.GetMailRecords()
    Dim mailIdx As Object: Set mailIdx = CaseDeskWorker.GetMailIndex()
    Dim cNames As Object: Set cNames = CaseDeskWorker.GetCaseNames()
    Print #fnum, "STEP2 | mail records | " & mailRecs.Count
    Print #fnum, "STEP2 | mail index | " & mailIdx.Count
    Print #fnum, "STEP2 | case names | " & cNames.Count

    ' --- Step 3: Ensure hidden sheets ---
    Dim shArr As Variant: shArr = Array("_casedesk_mail", "_casedesk_mail_idx", "_casedesk_cases", "_casedesk_files")
    Dim si As Long
    For si = 0 To UBound(shArr)
        Dim ws As Worksheet: Set ws = Nothing
        Set ws = ThisWorkbook.Worksheets(CStr(shArr(si)))
        If ws Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = CStr(shArr(si)): ws.Visible = xlSheetVeryHidden
        End If
    Next si

    ' --- Step 4: Write mail to sheet ---
    Set ws = ThisWorkbook.Worksheets("_casedesk_mail")
    ws.UsedRange.ClearContents
    Dim mKeys As Variant: mKeys = mailRecs.Keys
    Dim n As Long: n = mailRecs.Count
    If n > 0 Then
        Dim mD() As Variant: ReDim mD(1 To n, 1 To 11)
        Dim i As Long
        For i = 0 To n - 1
            Dim rec As Object: Set rec = mailRecs(mKeys(i))
            mD(i+1,1) = CaseDeskLib.DictStr(rec,"entry_id")
            mD(i+1,2) = CaseDeskLib.DictStr(rec,"sender_email")
            mD(i+1,3) = CaseDeskLib.DictStr(rec,"sender_name")
            mD(i+1,4) = CaseDeskLib.DictStr(rec,"subject")
            mD(i+1,5) = CaseDeskLib.DictStr(rec,"received_at")
            mD(i+1,6) = CaseDeskLib.DictStr(rec,"folder_path")
            mD(i+1,7) = CaseDeskLib.DictStr(rec,"body_path")
            mD(i+1,8) = CaseDeskLib.DictStr(rec,"msg_path")
            mD(i+1,9) = "": mD(i+1,10) = CaseDeskLib.DictStr(rec,"_mail_folder")
            mD(i+1,11) = CaseDeskLib.DictStr(rec,"body_text")
        Next i
        ws.Range("A1").Resize(n,11).Value = mD
    End If
    Print #fnum, "STEP4 | mail sheet rows | " & n

    ' Write mail index
    Set ws = ThisWorkbook.Worksheets("_casedesk_mail_idx")
    ws.UsedRange.ClearContents
    Dim oKeys As Variant: oKeys = mailIdx.Keys
    Dim tot As Long: tot = 0
    For i = 0 To UBound(oKeys): tot = tot + mailIdx(oKeys(i)).Count: Next i
    If tot > 0 Then
        Dim iD() As Variant: ReDim iD(1 To tot, 1 To 2)
        Dim ix As Long: ix = 0
        For i = 0 To UBound(oKeys)
            Dim inner As Object: Set inner = mailIdx(oKeys(i))
            Dim iK As Variant: iK = inner.Keys
            Dim j As Long
            For j = 0 To UBound(iK)
                ix = ix + 1: iD(ix,1) = CStr(oKeys(i)): iD(ix,2) = CStr(iK(j))
            Next j
        Next i
        ws.Range("A1").Resize(ix,2).Value = iD
    End If
    Print #fnum, "STEP4 | mail index rows | " & tot

    ' Write cases
    Set ws = ThisWorkbook.Worksheets("_casedesk_cases")
    ws.UsedRange.ClearContents
    Dim ck As Variant: ck = cNames.Keys
    If cNames.Count > 0 Then
        Dim cD() As Variant: ReDim cD(1 To cNames.Count, 1 To 1)
        For i = 0 To cNames.Count - 1: cD(i+1,1) = CStr(ck(i)): Next i
        ws.Range("A1").Resize(cNames.Count,1).Value = cD
    End If
    Print #fnum, "STEP4 | cases rows | " & cNames.Count

    ' NOTE: case FILES not written (WriteCaseFilesToFE is private, needs m_caseFiles)

    ' --- Step 5: FE load ---
    CaseDeskData.LoadFromLocalSheets ThisWorkbook
    Print #fnum, "STEP5 | FE mail count | " & CaseDeskData.GetMailCount()
    Print #fnum, "STEP5 | FE case count | " & CaseDeskData.GetCaseCount()

    ' --- Step 6: Match test ---
    Dim tbl As ListObject: Set tbl = CaseDeskData.FindTable(ActiveWorkbook, srcName)
    If tbl Is Nothing Then
        Print #fnum, "FAIL  | Table not found"
    Else
        Print #fnum, "STEP6 | Table found | " & tbl.Name & " (" & tbl.DataBodyRange.Rows.Count & " rows)"
        Dim mlc As String: mlc = CaseDeskLib.GetSourceStr(srcName, "mail_link_column")
        Dim flc As String: flc = CaseDeskLib.GetSourceStr(srcName, "folder_link_column")

        ' Mail match
        If Len(mlc) > 0 Then
            Dim ev As Variant: ev = tbl.DataBodyRange.Cells(1, tbl.ListColumns(mlc).Index).Value
            Print #fnum, "INFO  | Row1 email | " & CStr(ev)
            Dim mResult As Object: Set mResult = CaseDeskData.FindMailRecords(CStr(ev), "sender_email", "exact")
            Print #fnum, "STEP6 | FindMailRecords | " & mResult.Count
        Else
            Print #fnum, "WARN  | mail_link_column empty"
        End If

        ' File match
        If Len(flc) > 0 Then
            Dim cv As Variant: cv = tbl.DataBodyRange.Cells(1, tbl.ListColumns(flc).Index).Value
            Print #fnum, "INFO  | Row1 caseID | " & CStr(cv)
            Dim fResult As Object: Set fResult = CaseDeskData.FindCaseFiles(CStr(cv))
            Print #fnum, "STEP6 | FindCaseFiles | " & fResult.Count
        Else
            Print #fnum, "WARN  | folder_link_column empty"
        End If
    End If

    On Error GoTo 0
    Close #fnum
End Sub
"@

    $vbProj = $wb.VBProject
    $testComp = $vbProj.VBComponents.Add(1)
    $testComp.Name = "TestE2E"
    $testComp.CodeModule.AddFromString($testCode)

    Write-Host 'Running E2E tests...' -ForegroundColor Cyan

    # Set ActiveWorkbook to sample (g_dataWb will be set in test code)
    $sampleWb.Activate()

    try {
        $excel.Run("'$($wb.Name)'!RunE2ETests")
    } catch {
        Write-Host "Run error: $($_.Exception.Message)" -ForegroundColor Red
    }

    Write-Host ''
    Write-Host '--- E2E Test Results ---' -ForegroundColor Cyan
    if (Test-Path $tempFile) {
        Get-Content $tempFile -Encoding UTF8 | ForEach-Object {
            if ($_ -match '^PASS|^STEP') { Write-Host "  $_" -ForegroundColor Green }
            elseif ($_ -match '^FAIL') { Write-Host "  $_" -ForegroundColor Red }
            elseif ($_ -match '^WARN') { Write-Host "  $_" -ForegroundColor Yellow }
            else { Write-Host "  $_" -ForegroundColor Gray }
        }
        Remove-Item $tempFile -Force
    } else {
        Write-Host '  No results file.' -ForegroundColor Red
    }

    try { $vbProj.VBComponents.Remove($testComp) } catch {}

} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    try { $excel.Run("CaseDeskMain.BeforeWorkbookClose") } catch {}
    try { if ($sampleWb) { $sampleWb.Close($false); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sampleWb) | Out-Null } } catch {}
    try { if ($wb) { $wb.Close($false); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null } } catch {}
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
}
