# Test-Manifest.ps1
# Test CaseDeskWorker manifest reading (mail + cases) with sample data

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
$xlsmPath = Join-Path $projectDir 'dist\casedesk.xlsm'
$sampleDir = Join-Path $projectDir 'sample'

if (-not (Test-Path $xlsmPath)) {
    Write-Host "ERROR: $xlsmPath not found." -ForegroundColor Red
    exit 1
}

Write-Host "Opening $xlsmPath ..." -ForegroundColor Cyan
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $prevSec = $excel.AutomationSecurity
    $excel.AutomationSecurity = 1
    $wb = $excel.Workbooks.Open($xlsmPath)
    $excel.AutomationSecurity = $prevSec
    $vbProj = $wb.VBProject

    if ($vbProj -eq $null) {
        Write-Host 'ERROR: Cannot access VBProject.' -ForegroundColor Red
        exit 1
    }

    $tempFile = Join-Path $env:TEMP "casedesk_manifest_test.txt"
    if (Test-Path $tempFile) { Remove-Item $tempFile -Force }

    $mailDir = Join-Path $sampleDir 'mail'
    $casesDir = Join-Path $sampleDir 'cases'

    $testCode = @"
Option Explicit

Public Sub RunManifestTests()
    Dim fnum As Integer
    fnum = FreeFile
    Open "$($tempFile -replace '\\','\\')" For Output As #fnum

    ' Test mail manifest reading
    Dim mailFolder As String: mailFolder = "$($mailDir -replace '\\','\\')"
    CaseDeskWorker.SetMailMatchConfig "sender_email", "exact"
    Dim mailChanged As Boolean: mailChanged = CaseDeskWorker.RefreshMailData(mailFolder)

    Dim mailRecs As Object: Set mailRecs = CaseDeskWorker.GetMailRecords()
    If mailRecs.Count > 0 Then
        Print #fnum, "PASS | RefreshMailData | " & mailRecs.Count & " mail records loaded"
    Else
        Print #fnum, "FAIL | RefreshMailData | 0 mail records"
    End If

    Dim mailIdx As Object: Set mailIdx = CaseDeskWorker.GetMailIndex()
    If mailIdx.Count > 0 Then
        Print #fnum, "PASS | MailIndex | " & mailIdx.Count & " index entries"
    Else
        Print #fnum, "FAIL | MailIndex | 0 index entries"
    End If

    ' Test case manifest reading
    Dim casesFolder As String: casesFolder = "$($casesDir -replace '\\','\\')"
    Dim caseChanged As Boolean: caseChanged = CaseDeskWorker.RefreshCaseData(casesFolder)

    Dim caseNames As Object: Set caseNames = CaseDeskWorker.GetCaseNames()
    If caseNames.Count > 0 Then
        Print #fnum, "PASS | RefreshCaseData | " & caseNames.Count & " case names loaded"
    Else
        Print #fnum, "FAIL | RefreshCaseData | 0 case names"
    End If

    Close #fnum
End Sub
"@

    Write-Host 'Injecting manifest test module...' -ForegroundColor Cyan
    $testComp = $vbProj.VBComponents.Add(1)
    $testComp.Name = "TestManifest"
    $testComp.CodeModule.AddFromString($testCode)

    Write-Host 'Running manifest tests...' -ForegroundColor Cyan
    try {
        $excel.Run("RunManifestTests")
    } catch {
        Write-Host "Run error: $($_.Exception.Message)" -ForegroundColor Red
    }

    Write-Host ''
    Write-Host '--- Manifest Test Results ---' -ForegroundColor Cyan
    $passed = 0
    $failed = 0
    if (Test-Path $tempFile) {
        $results = Get-Content $tempFile -Encoding UTF8
        foreach ($line in $results) {
            if ($line -match '^PASS') {
                Write-Host "  $line" -ForegroundColor Green
                $passed++
            } elseif ($line -match '^FAIL') {
                Write-Host "  $line" -ForegroundColor Red
                $failed++
            } else {
                Write-Host "  $line"
            }
        }
        Remove-Item $tempFile -Force
    } else {
        Write-Host '  No results file generated.' -ForegroundColor Red
        $failed = 1
    }

    try { $vbProj.VBComponents.Remove($testComp) } catch {}

    Write-Host ''
    Write-Host "=== $passed passed, $failed failed ===" -ForegroundColor $(if ($failed -eq 0) { "Green" } else { "Red" })
    exit $failed

} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
} finally {
    try { if ($wb) { $wb.Close($false); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null } } catch {}
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
}
