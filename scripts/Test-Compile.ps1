# Test-Compile.ps1
# 各モジュールの各プロシージャを個別にテストし、エラーを出力する

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
$xlsmPath = Join-Path $projectDir 'casedesk.xlsm'

if (-not (Test-Path $xlsmPath)) {
    Write-Host "ERROR: $xlsmPath not found." -ForegroundColor Red
    exit 1
}

Write-Host "Opening $xlsmPath ..." -ForegroundColor Cyan
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

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
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
}
