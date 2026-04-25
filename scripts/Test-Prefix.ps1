# Test-Prefix.ps1
# Test the new prefix system (IsHiddenField, IsReadOnlyField, StripFieldPrefix, etc.)
# and number formatting changes

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
$xlsmPath = Join-Path $projectDir 'dist\casedesk.xlsm'

if (-not (Test-Path $xlsmPath)) {
    Write-Host "ERROR: $xlsmPath not found. Run Build-Addin.ps1 first." -ForegroundColor Red
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

    $tempFile = Join-Path $env:TEMP "casedesk_prefix_test.txt"
    if (Test-Path $tempFile) { Remove-Item $tempFile -Force }

    $testCode = @"
Option Explicit

Public Sub RunPrefixTests()
    Dim fnum As Integer
    fnum = FreeFile
    Open "$($tempFile -replace '\\','\\')" For Output As #fnum

    ' === IsHiddenField ===
    AssertEqual fnum, "IsHiddenField('__foo')", CStr(CaseDeskLib.IsHiddenField("__foo")), "True"
    AssertEqual fnum, "IsHiddenField('__AB')", CStr(CaseDeskLib.IsHiddenField("__AB")), "True"
    AssertEqual fnum, "IsHiddenField('___foo')", CStr(CaseDeskLib.IsHiddenField("___foo")), "False"
    AssertEqual fnum, "IsHiddenField('_abc')", CStr(CaseDeskLib.IsHiddenField("_abc")), "False"
    AssertEqual fnum, "IsHiddenField('abc')", CStr(CaseDeskLib.IsHiddenField("abc")), "False"
    AssertEqual fnum, "IsHiddenField('_a')", CStr(CaseDeskLib.IsHiddenField("_a")), "False"
    AssertEqual fnum, "IsHiddenField('__')", CStr(CaseDeskLib.IsHiddenField("__")), "False"

    ' === IsReadOnlyField ===
    AssertEqual fnum, "IsReadOnlyField('_a_Foo')", CStr(CaseDeskLib.IsReadOnlyField("_a_Foo")), "True"
    AssertEqual fnum, "IsReadOnlyField('_ab_Foo')", CStr(CaseDeskLib.IsReadOnlyField("_ab_Foo")), "True"
    AssertEqual fnum, "IsReadOnlyField('_a')", CStr(CaseDeskLib.IsReadOnlyField("_a")), "True"
    AssertEqual fnum, "IsReadOnlyField('__foo')", CStr(CaseDeskLib.IsReadOnlyField("__foo")), "False"
    AssertEqual fnum, "IsReadOnlyField('abc')", CStr(CaseDeskLib.IsReadOnlyField("abc")), "False"
    AssertEqual fnum, "IsReadOnlyField('a_B')", CStr(CaseDeskLib.IsReadOnlyField("a_B")), "False"

    ' === StripFieldPrefix ===
    AssertEqual fnum, "StripFieldPrefix('xx_AAA')", CaseDeskLib.StripFieldPrefix("xx_AAA"), "xx_AAA"
    AssertEqual fnum, "StripFieldPrefix('_xx_AAA')", CaseDeskLib.StripFieldPrefix("_xx_AAA"), "xx_AAA"
    AssertEqual fnum, "StripFieldPrefix('__AAA')", CaseDeskLib.StripFieldPrefix("__AAA"), "AAA"
    AssertEqual fnum, "StripFieldPrefix('abc')", CaseDeskLib.StripFieldPrefix("abc"), "abc"

    ' === GetFieldGroup ===
    AssertEqual fnum, "GetFieldGroup('xx_AAA')", CaseDeskLib.GetFieldGroup("xx_AAA"), "xx"
    AssertEqual fnum, "GetFieldGroup('_xx_AAA')", CaseDeskLib.GetFieldGroup("_xx_AAA"), "xx"
    AssertEqual fnum, "GetFieldGroup('abc')", CaseDeskLib.GetFieldGroup("abc"), ""

    ' === GetFieldShortName ===
    AssertEqual fnum, "GetFieldShortName('xx_AAA')", CaseDeskLib.GetFieldShortName("xx_AAA"), "AAA"
    AssertEqual fnum, "GetFieldShortName('_xx_AAA')", CaseDeskLib.GetFieldShortName("_xx_AAA"), "AAA"
    AssertEqual fnum, "GetFieldShortName('__AAA')", CaseDeskLib.GetFieldShortName("__AAA"), "AAA"
    AssertEqual fnum, "GetFieldShortName('abc')", CaseDeskLib.GetFieldShortName("abc"), "abc"

    ' === FormatFieldValue (number with commas) ===
    AssertEqual fnum, "FormatFieldValue(1234, 'number')", CaseDeskLib.FormatFieldValue(1234, "number"), "1,234"
    AssertEqual fnum, "FormatFieldValue(1234567, 'number')", CaseDeskLib.FormatFieldValue(1234567, "number"), "1,234,567"
    AssertEqual fnum, "FormatFieldValue(1234.5, 'number')", CaseDeskLib.FormatFieldValue(1234.5, "number"), "1,234.5"
    AssertEqual fnum, "FormatFieldValue(0, 'number')", CaseDeskLib.FormatFieldValue(0, "number"), "0"
    AssertEqual fnum, "FormatFieldValue(1234, 'currency')", CaseDeskLib.FormatFieldValue(1234, "currency"), "1,234"
    ' Large number (no CLng overflow)
    AssertEqual fnum, "FormatFieldValue(9999999999, 'number')", CaseDeskLib.FormatFieldValue(CDbl(9999999999#), "number"), "9,999,999,999"
    AssertEqual fnum, "FormatFieldValue('abc', 'text')", CaseDeskLib.FormatFieldValue("abc", "text"), "abc"

    Close #fnum
End Sub

Private Sub AssertEqual(fnum As Integer, testName As String, actual As String, expected As String)
    If actual = expected Then
        Print #fnum, "PASS | " & testName & " | " & actual
    Else
        Print #fnum, "FAIL | " & testName & " | expected=" & expected & " actual=" & actual
    End If
End Sub
"@

    Write-Host 'Injecting prefix test module...' -ForegroundColor Cyan
    $testComp = $vbProj.VBComponents.Add(1)
    $testComp.Name = "TestPrefix"
    $testComp.CodeModule.AddFromString($testCode)

    Write-Host 'Running prefix tests...' -ForegroundColor Cyan
    try {
        $excel.Run("RunPrefixTests")
    } catch {
        Write-Host "Run error: $($_.Exception.Message)" -ForegroundColor Red
    }

    Write-Host ''
    Write-Host '--- Prefix Test Results ---' -ForegroundColor Cyan
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
        Write-Host '  No results file generated - test module may have failed to compile.' -ForegroundColor Red
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
