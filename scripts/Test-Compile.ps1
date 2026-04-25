# Test-Compile.ps1
# 各モジュールの各プロシージャを個別にテストし、エラーを出力する

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
$xlsmPath = Join-Path $projectDir 'dist\casedesk.xlsm'

if (-not (Test-Path $xlsmPath)) {
    Write-Host "ERROR: $xlsmPath not found." -ForegroundColor Red
    exit 1
}

Write-Host "Opening $xlsmPath ..." -ForegroundColor Cyan
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
# Interactive = $false blocks MsgBox / VBE modal dialogs (including "構文エラー"
# compile error popups) that would otherwise hang the headless run.
try { $excel.Interactive = $false } catch {}
# Capture the PID of the Excel instance we just launched so we can kill it
# specifically on cleanup without touching other Excel sessions (e.g. the
# developer's own editor or a running CaseDesk worker).
$excelPid = $null
try {
    Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public class W{[DllImport("user32.dll")]public static extern int GetWindowThreadProcessId(IntPtr h,out int p);}' -ErrorAction SilentlyContinue
    $null = [W]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
} catch {}

$sampleWb = $null

try {
    $wb = $excel.Workbooks.Open($xlsmPath)
    $vbProj = $wb.VBProject

    if ($vbProj -eq $null) {
        Write-Host 'ERROR: Cannot access VBProject.' -ForegroundColor Red
        exit 1
    }

    # Open sample workbook
    $samplePath = Join-Path $projectDir 'sample\casedesk-sample.xlsx'
    if (Test-Path $samplePath) {
        $sampleWb = $excel.Workbooks.Open($samplePath)
        Write-Host "Sample workbook opened." -ForegroundColor Green
    }

    # --- Inject test module ---
    # Create a temporary test module that calls each procedure with error handling
    # and writes results to a temp file
    $tempFile = Join-Path $env:TEMP "casedesk_test_result.txt"
    if (Test-Path $tempFile) { Remove-Item $tempFile -Force }

    $testCode = @"
Option Explicit

Public Sub RunAllTests()
    Dim fnum As Integer
    fnum = FreeFile
    Open "$($tempFile -replace '\\','\\')" For Output As #fnum

    ' --- CaseDeskLib (merged from CaseDeskHelpers + CaseDeskConfig + CaseDeskChangeLog) ---
    TestCall fnum, "CaseDeskLib.NewDict", ""
    TestCall fnum, "CaseDeskLib.ParseJson", ""
    TestCall fnum, "CaseDeskLib.SafeName", ""
    TestCall fnum, "CaseDeskLib.FileExists", ""
    TestCall fnum, "CaseDeskLib.FolderExists", ""
    TestCall fnum, "CaseDeskLib.EnsureConfigSheets", ""
    TestCall fnum, "CaseDeskLib.EnsureLogSheet", ""

    ' --- CaseDeskLib Roles (spec 5.3) ---
    TestCall fnum, "CaseDeskLib.GetRoleIds", ""
    TestCall fnum, "CaseDeskLib.GetRequiredRoleIds", ""
    TestCall fnum, "CaseDeskLib.GetRoleLabel.case_id", ""
    TestCall fnum, "CaseDeskLib.GetRoleLabel.empty", ""
    TestCall fnum, "CaseDeskLib.GuessRoleFromColumnName.kenmei", ""
    TestCall fnum, "CaseDeskLib.GuessRoleFromColumnName.anken_id", ""
    TestCall fnum, "CaseDeskLib.GuessRoleFromColumnName.status", ""
    TestCall fnum, "CaseDeskLib.GuessRoleFromColumnName.noise", ""
    TestCall fnum, "CaseDeskLib.MissingRequiredRoles.empty_source", ""

    ' --- CaseDeskData ---
    TestCall fnum, "CaseDeskData.GetWorkbookTableNames", ""
    TestCall fnum, "CaseDeskData.FindTable", ""

    ' --- FieldEditor (skip - class, needs form) ---

    ' --- frmCaseDesk (instantiation test) ---
    TestFormLoad fnum, "frmCaseDesk"
    TestFormLoad fnum, "frmSettings"

    Close #fnum
End Sub

Private Sub TestCall(fnum As Integer, procName As String, note As String)
    On Error GoTo ErrHandler
    Dim result As String
    Select Case procName
        Case "CaseDeskLib.NewDict"
            Dim d As Object: Set d = CaseDeskLib.NewDict()
            result = "OK (Dict created)"
        Case "CaseDeskLib.ParseJson"
            Dim j As Object: Set j = CaseDeskLib.ParseJson("{""a"":1}")
            result = "OK (parsed)"
        Case "CaseDeskLib.SafeName"
            Dim sn As String: sn = CaseDeskLib.SafeName("test/file:name")
            result = "OK (" & sn & ")"
        Case "CaseDeskLib.FileExists"
            Dim fe As Boolean: fe = CaseDeskLib.FileExists("C:\nonexist.txt")
            result = "OK (" & fe & ")"
        Case "CaseDeskLib.FolderExists"
            Dim fde As Boolean: fde = CaseDeskLib.FolderExists("C:\")
            result = "OK (" & fde & ")"
        Case "CaseDeskLib.EnsureConfigSheets"
            CaseDeskLib.EnsureConfigSheets
            result = "OK"
        Case "CaseDeskLib.EnsureLogSheet"
            CaseDeskLib.EnsureLogSheet
            result = "OK"
        Case "CaseDeskLib.GetRoleIds"
            Dim ri As Collection: Set ri = CaseDeskLib.GetRoleIds()
            If ri.Count < 6 Then Err.Raise 5, , "too few roles: " & ri.Count
            result = "OK (count=" & ri.Count & ")"
        Case "CaseDeskLib.GetRequiredRoleIds"
            Dim rri As Collection: Set rri = CaseDeskLib.GetRequiredRoleIds()
            If rri.Count <> 2 Then Err.Raise 5, , "expected 2 required: " & rri.Count
            result = "OK (count=" & rri.Count & ")"
        Case "CaseDeskLib.GetRoleLabel.case_id"
            Dim lab As String: lab = CaseDeskLib.GetRoleLabel("case_id")
            If Len(lab) = 0 Or lab = "case_id" Then Err.Raise 5, , "no label: " & lab
            result = "OK (" & lab & ")"
        Case "CaseDeskLib.GetRoleLabel.empty"
            Dim lab2 As String: lab2 = CaseDeskLib.GetRoleLabel("")
            result = "OK (" & lab2 & ")"
        Case "CaseDeskLib.GuessRoleFromColumnName.kenmei"
            ' "件名" via ChrW — avoid `ChrW`$ because PowerShell here-string
            ' interprets `$(...)` as a subexpression and would strip the `$`.
            Dim g1 As String
            g1 = CaseDeskLib.GuessRoleFromColumnName(ChrW(20214) & ChrW(21517))
            If g1 <> "title" Then Err.Raise 5, , "expected title got: " & g1
            result = "OK (" & g1 & ")"
        Case "CaseDeskLib.GuessRoleFromColumnName.anken_id"
            ' "案件ID"
            Dim g2 As String
            g2 = CaseDeskLib.GuessRoleFromColumnName(ChrW(26696) & ChrW(20214) & "ID")
            If g2 <> "case_id" Then Err.Raise 5, , "expected case_id got: " & g2
            result = "OK (" & g2 & ")"
        Case "CaseDeskLib.GuessRoleFromColumnName.status"
            ' "状態"
            Dim g3 As String
            g3 = CaseDeskLib.GuessRoleFromColumnName(ChrW(29366) & ChrW(24907))
            If g3 <> "status" Then Err.Raise 5, , "expected status got: " & g3
            result = "OK (" & g3 & ")"
        Case "CaseDeskLib.GuessRoleFromColumnName.noise"
            ' "備考欄"
            Dim g4 As String
            g4 = CaseDeskLib.GuessRoleFromColumnName(ChrW(20633) & ChrW(32771) & ChrW(27396))
            If Len(g4) > 0 Then Err.Raise 5, , "expected empty for noise, got: " & g4
            result = "OK (empty)"
        Case "CaseDeskLib.MissingRequiredRoles.empty_source"
            Dim mrr As Collection: Set mrr = CaseDeskLib.MissingRequiredRoles("__nonexistent_source__")
            If mrr.Count <> 2 Then Err.Raise 5, , "expected 2 missing, got: " & mrr.Count
            result = "OK (missing=" & mrr.Count & ")"
        Case "CaseDeskData.GetWorkbookTableNames"
            Dim tn As Collection: Set tn = CaseDeskData.GetWorkbookTableNames(ActiveWorkbook)
            result = "OK (count=" & tn.Count & ")"
        Case "CaseDeskData.FindTable"
            Dim tbl As ListObject: Set tbl = CaseDeskData.FindTable(ActiveWorkbook, "anken")
            If tbl Is Nothing Then result = "OK (not found)" Else result = "OK (found: " & tbl.Name & ")"
        Case Else
            result = "SKIP"
    End Select
    Print #fnum, "PASS | " & procName & " | " & result
    Exit Sub
ErrHandler:
    Print #fnum, "FAIL | " & procName & " | Err " & Err.Number & ": " & Err.Description
    Resume Next
End Sub

Private Sub TestFormLoad(fnum As Integer, formName As String)
    On Error GoTo ErrHandler
    Select Case formName
        Case "frmCaseDesk"
            Dim f1 As New frmCaseDesk
            Print #fnum, "PASS | frmCaseDesk.New | OK (instantiated)"
            Unload f1
        Case "frmSettings"
            Dim f2 As New frmSettings
            Print #fnum, "PASS | frmSettings.New | OK (instantiated)"
            Unload f2
    End Select
    Exit Sub
ErrHandler:
    Print #fnum, "FAIL | " & formName & ".New | Err " & Err.Number & ": " & Err.Description
    Resume Next
End Sub
"@

    Write-Host 'Injecting test module...' -ForegroundColor Cyan
    $testComp = $vbProj.VBComponents.Add(1) # Standard module
    $testComp.Name = "TestRunner"
    $testComp.CodeModule.AddFromString($testCode)

    # Run the test
    Write-Host 'Running tests...' -ForegroundColor Cyan
    try {
        $excel.Run("'casedesk.xlsm'!RunAllTests")
    } catch {
        Write-Host "Run error: $($_.Exception.Message)" -ForegroundColor Red
        # Try without workbook prefix
        try {
            $excel.Run("RunAllTests")
        } catch {
            Write-Host "Run error (2): $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Read results
    Write-Host ''
    Write-Host '--- Test Results ---' -ForegroundColor Cyan
    if (Test-Path $tempFile) {
        $results = Get-Content $tempFile -Encoding UTF8
        foreach ($line in $results) {
            if ($line -match '^PASS') {
                Write-Host "  $line" -ForegroundColor Green
            } elseif ($line -match '^FAIL') {
                Write-Host "  $line" -ForegroundColor Red
            } else {
                Write-Host "  $line"
            }
        }
        Remove-Item $tempFile -Force
    } else {
        Write-Host '  No results file generated - test module may have failed to compile.' -ForegroundColor Red
        Write-Host '  This usually means a compile error exists in the project.' -ForegroundColor Yellow

        # Try to get compile errors by reading each module line by line
        Write-Host ''
        Write-Host '--- Checking each module for syntax errors ---' -ForegroundColor Cyan
        foreach ($comp in $vbProj.VBComponents) {
            try {
                $codeMod = $comp.CodeModule
                $lineCount = $codeMod.CountOfLines
                if ($lineCount -eq 0) { continue }
                # Try to read all lines - this alone won't detect compile errors
                # but accessing ProcOfLine can surface some issues
                Write-Host "  $($comp.Name): $lineCount lines" -ForegroundColor Gray
                for ($i = 1; $i -le $lineCount; $i++) {
                    try {
                        $null = $codeMod.Lines($i, 1)
                    } catch {
                        Write-Host "    Line $i ERROR: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            } catch {
                Write-Host "  $($comp.Name): ERROR accessing code - $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }

    # Clean up test module
    try { $vbProj.VBComponents.Remove($testComp) } catch {}

} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    try { if ($sampleWb) { $sampleWb.Close($false); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sampleWb) | Out-Null } } catch {}
    try { if ($wb) { $wb.Close($false); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null } } catch {}
    try { $excel.Quit() } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    [GC]::Collect()
    # Hard-kill as a last resort: if VBE is stuck in break-mode or a modal
    # dialog slipped past Interactive=$false, $excel.Quit() will no-op and
    # leave the process alive. Only kill the specific PID we launched — we
    # must NOT blanket-kill EXCEL.EXE because unrelated Excel sessions
    # (developer's editor, running worker) could have unsaved work.
    if ($excelPid) {
        try { Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue } catch {}
    }
}
