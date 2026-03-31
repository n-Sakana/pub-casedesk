# Build-Addin.ps1
# VBAソースファイルをインポートしたxlam/xlsmを自動生成する
#
# 前提条件:
#   Excel > ファイル > オプション > トラストセンター > トラストセンターの設定
#   > マクロの設定 > 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」をONにすること
#
# 使い方:
#   powershell -ExecutionPolicy Bypass -File Build-Addin.ps1
#   powershell -ExecutionPolicy Bypass -File Build-Addin.ps1 -OutputFormat xlsm

param(
    [ValidateSet('xlsm', 'xlam')]
    [string]$OutputFormat = 'xlam',
    [string]$OutputName = '',
    [switch]$Sample
)

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
$srcDir = Join-Path $projectDir 'src'

$basModules = @(
    'CaseDeskLib.bas',
    'CaseDeskData.bas',
    'CaseDeskMain.bas',
    'CaseDeskWorker.bas'
)
$clsModules = @(
    'ErrorHandler.cls',
    'FieldEditor.cls',
    'SheetWatcher.cls',
    'AppEventHandler.cls'
)
$frmModules = @(
    @{ Name = 'frmCaseDesk';       File = 'frmCaseDesk.frm' },
    @{ Name = 'frmSettings';    File = 'frmSettings.frm' }
)

# --- Helper: extract code from .cls/.frm (skip VERSION/BEGIN/END/Attribute header) ---
function Extract-VBACode($path) {
    $lines = Get-Content -Path $path -Encoding UTF8
    $codeLines = @()
    $inHeader = $true
    foreach ($line in $lines) {
        if ($inHeader) {
            if ($line -match '^Attribute VB_Exposed') { $inHeader = $false; continue }
            continue
        }
        $codeLines += $line
    }
    return ($codeLines -join "`r`n")
}

# --- Start Excel ---
Write-Host 'Starting Excel...' -ForegroundColor Cyan
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()
    $vbProj = $wb.VBProject

    # --- Check trust access ---
    if ($vbProj -eq $null) {
        Write-Host 'ERROR: VBA Project access is not trusted.' -ForegroundColor Red
        Write-Host 'Enable: Excel > Trust Center > Macro Settings > Trust access to VBA project object model' -ForegroundColor Yellow
        throw 'VBA project access denied.'
    }
    try {
        $compCount = $vbProj.VBComponents.Count
        Write-Host "  VBProject accessible ($compCount components)" -ForegroundColor Green
    } catch {
        throw 'VBA project access denied.'
    }

    # --- 1. Import .bas modules ---
    foreach ($mod in $basModules) {
        $path = Join-Path $srcDir $mod
        if (-not (Test-Path $path)) { Write-Host "  skip: $mod" -ForegroundColor Yellow; continue }
        Write-Host "  import: $mod"
        $vbProj.VBComponents.Import($path) | Out-Null
    }

    # --- 2. Create UserForms FIRST (registers MSForms reference) ---
    foreach ($frm in $frmModules) {
        $frmPath = Join-Path $srcDir $frm.File
        if (-not (Test-Path $frmPath)) { Write-Host "  skip: $($frm.File)" -ForegroundColor Yellow; continue }
        Write-Host "  create form: $($frm.Name)"
        $code = Extract-VBACode $frmPath
        $comp = $vbProj.VBComponents.Add(3) # vbext_ct_MSForm
        $comp.Name = $frm.Name
        $codeMod = $comp.CodeModule
        if ($codeMod.CountOfLines -gt 0) { $codeMod.DeleteLines(1, $codeMod.CountOfLines) }
        $codeMod.AddFromString($code)
    }

    # --- 3. Create .cls modules AFTER forms (MSForms reference now available) ---
    foreach ($mod in $clsModules) {
        $path = Join-Path $srcDir $mod
        if (-not (Test-Path $path)) { Write-Host "  skip: $mod" -ForegroundColor Yellow; continue }
        $clsName = [System.IO.Path]::GetFileNameWithoutExtension($mod)
        Write-Host "  create class: $clsName"
        $code = Extract-VBACode $path
        $comp = $vbProj.VBComponents.Add(2) # vbext_ct_ClassModule
        $comp.Name = $clsName
        $codeMod = $comp.CodeModule
        if ($codeMod.CountOfLines -gt 0) { $codeMod.DeleteLines(1, $codeMod.CountOfLines) }
        $codeMod.AddFromString($code)
    }

    # --- ThisWorkbook code (differs by format) ---
    if ($OutputFormat -eq 'xlam') {
        $thisWorkbookCode = @'
Option Explicit

Private Sub Workbook_Open()
    CaseDeskMain.InitAddin
End Sub

Private Sub Workbook_AddinInstall()
    CaseDeskMain.InitAddin
End Sub

Private Sub Workbook_AddinUninstall()
    CaseDeskMain.ShutdownAddin
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    CaseDeskMain.ShutdownAddin
    Me.Saved = True
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    Dim sn As String: sn = Sh.Name
    If Left$(sn, Len("_casedesk")) <> "_casedesk" Then Exit Sub
    Application.ScreenUpdating = False
    If CaseDeskMain.g_formLoaded Then frmCaseDesk.OnCaseDeskSheetChange sn
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub
'@
    } else {
        $thisWorkbookCode = @'
Option Explicit

Private Sub Workbook_Open()
    CaseDeskMain.InitAddin
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    CaseDeskMain.BeforeWorkbookClose
    Me.Saved = True
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    Dim sn As String: sn = Sh.Name
    If Left$(sn, Len("_casedesk")) <> "_casedesk" Then Exit Sub
    Application.ScreenUpdating = False
    If CaseDeskMain.g_formLoaded Then frmCaseDesk.OnCaseDeskSheetChange sn
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub
'@
    }

    $docComp = $vbProj.VBComponents.Item('ThisWorkbook')
    $docCode = $docComp.CodeModule
    if ($docCode.CountOfLines -gt 0) { $docCode.DeleteLines(1, $docCode.CountOfLines) }
    $docCode.AddFromString($thisWorkbookCode)

    # --- 4. Pre-create config sheets ---
    $sampleDir = Join-Path $projectDir 'sample'
    $mailDir = Join-Path $sampleDir 'mail'
    $casesDir = Join-Path $sampleDir 'cases'

    # _casedesk_config: key-value pairs
    $cfgSheet = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $cfgSheet.Name = "_casedesk_config"
    $cfgSheet.Visible = 2  # xlSheetVeryHidden
    $cfgSheet.Range("A1").Value2 = "key"
    $cfgSheet.Range("B1").Value2 = "value"
    if ($Sample) {
        $cfgSheet.Range("A2").Value2 = "mail_folder"
        $cfgSheet.Range("B2").Value2 = $mailDir
        $cfgSheet.Range("A3").Value2 = "case_folder_root"
        $cfgSheet.Range("B3").Value2 = $casesDir
    }

    # _casedesk_sources: one row per source
    $srcSheet = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $srcSheet.Name = "_casedesk_sources"
    $srcSheet.Visible = 2  # xlSheetVeryHidden
    $srcSheet.Range("A1").Value2 = "source_name"
    $srcSheet.Range("B1").Value2 = "key_column"
    $srcSheet.Range("C1").Value2 = "display_name_column"
    $srcSheet.Range("D1").Value2 = "mail_link_column"
    $srcSheet.Range("E1").Value2 = "folder_link_column"
    $srcSheet.Range("F1").Value2 = "mail_match_mode"
    if ($Sample) {
        $sampleXlsx = Join-Path $sampleDir 'casedesk-sample.xlsx'
        # Read column names from sample xlsx (no hardcoded Japanese)
        $sampleWb = $excel.Workbooks.Open($sampleXlsx, 0, $true)
        $sampleTbl = $null
        foreach ($ws in $sampleWb.Worksheets) {
            foreach ($lo in $ws.ListObjects) {
                $sampleTbl = $lo; break
            }
            if ($sampleTbl) { break }
        }
        if ($sampleTbl) {
            $srcSheet.Range("A2").Value2 = $sampleTbl.Name
            $srcSheet.Range("B2").Value2 = $sampleTbl.ListColumns(1).Name  # key
            $srcSheet.Range("C2").Value2 = $sampleTbl.ListColumns(2).Name  # display name
            # Find mail column: first column containing '@' in data
            foreach ($col in $sampleTbl.ListColumns) {
                if ($col.DataBodyRange -and $col.DataBodyRange.Cells(1,1).Text -match '@') {
                    $srcSheet.Range("D2").Value2 = $col.Name; break
                }
            }
            $srcSheet.Range("E2").Value2 = $sampleTbl.ListColumns(1).Name  # folder = key
        }
        $sampleWb.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sampleWb) | Out-Null
    }

    # _casedesk_fields: one row per source+field
    $fldSheet = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $fldSheet.Name = "_casedesk_fields"
    $fldSheet.Visible = 2  # xlSheetVeryHidden
    $fldSheet.Range("A1").Value2 = "source_name"
    $fldSheet.Range("B1").Value2 = "field_name"
    $fldSheet.Range("C1").Value2 = "type"
    $fldSheet.Range("D1").Value2 = "in_list"
    $fldSheet.Range("E1").Value2 = "editable"
    $fldSheet.Range("F1").Value2 = "multiline"

    # _casedesk_log (with ListObject table)
    $logSheet = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $logSheet.Name = "_casedesk_log"
    $logSheet.Visible = 2  # xlSheetVeryHidden
    $logSheet.Range("A1").Value2 = "timestamp"
    $logSheet.Range("B1").Value2 = "source"
    $logSheet.Range("C1").Value2 = "key"
    $logSheet.Range("D1").Value2 = "field"
    $logSheet.Range("E1").Value2 = "old_value"
    $logSheet.Range("F1").Value2 = "new_value"
    $logSheet.Range("G1").Value2 = "origin"
    $logTable = $logSheet.ListObjects.Add(1, $logSheet.Range("A1:G1"), $null, 1)  # xlSrcRange, xlYes
    $logTable.Name = "CaseDeskLog"

    if ($Sample) {
        Write-Host "  config: mail=$mailDir, cases=$casesDir" -ForegroundColor Green
    }

    # --- Save ---
    if ([string]::IsNullOrWhiteSpace($OutputName)) {
        $outputName = "casedesk.$OutputFormat"
    } else {
        $outputName = $OutputName
    }
    $outputPath = Join-Path $projectDir $outputName
    $fileFormat = if ($OutputFormat -eq 'xlam') { 55 } else { 52 }
    if (Test-Path $outputPath) { Remove-Item $outputPath -Force }
    if ($OutputFormat -eq 'xlam') {
        $wb.IsAddin = $true
    }
    $wb.SaveAs($outputPath, $fileFormat)

    Write-Host ''
    foreach ($comp in $vbProj.VBComponents) {
        $kind = switch ($comp.Type) { 1{'Module'} 2{'Class'} 3{'Form'} 100{'Document'} default{"Type$($comp.Type)"} }
        Write-Host ("  {0,-20} {1}" -f $comp.Name, $kind)
    }

} finally {
    if ($wb) { $wb.Close($false); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
}

# --- Inject customUI XML into xlam (ZIP manipulation) ---
if ($OutputFormat -eq 'xlam') {
    $customUIPath = Join-Path $srcDir 'customUI14.xml'
    if (Test-Path $customUIPath) {
        Write-Host '  Injecting customUI ribbon...'

        # xlam is a ZIP — extract, add customUI, repack
        $tempDir = Join-Path $env:TEMP "casedesk_build_$(Get-Random)"
        $zipPath = $outputPath + '.zip'

        Copy-Item $outputPath $zipPath -Force
        Expand-Archive -Path $zipPath -DestinationPath $tempDir -Force
        Remove-Item $zipPath -Force

        # Add customUI folder and XML
        $cuiDir = Join-Path $tempDir 'customUI'
        New-Item -ItemType Directory -Path $cuiDir -Force | Out-Null
        Copy-Item $customUIPath (Join-Path $cuiDir 'customUI14.xml') -Force

        # Update [Content_Types].xml — add customUI content type
        $ctPath = Join-Path $tempDir '[Content_Types].xml'
        $ctXml = [xml](Get-Content -LiteralPath $ctPath)
        $ns = $ctXml.DocumentElement.NamespaceURI
        # Check if already exists
        $existing = $ctXml.Types.Override | Where-Object { $_.PartName -eq '/customUI/customUI14.xml' }
        if (-not $existing) {
            $node = $ctXml.CreateElement('Override', $ns)
            $node.SetAttribute('PartName', '/customUI/customUI14.xml')
            $node.SetAttribute('ContentType', 'application/xml')
            $ctXml.DocumentElement.AppendChild($node) | Out-Null
            $ctXml.Save($ctPath)
        }

        # Update _rels/.rels — add relationship to customUI
        $relsPath = Join-Path $tempDir '_rels\.rels'
        $relsXml = [xml](Get-Content -LiteralPath $relsPath)
        $relsNs = $relsXml.DocumentElement.NamespaceURI
        $existing = $relsXml.Relationships.Relationship | Where-Object { $_.Target -eq 'customUI/customUI14.xml' }
        if (-not $existing) {
            $relNode = $relsXml.CreateElement('Relationship', $relsNs)
            $relNode.SetAttribute('Id', 'rIdCustomUI')
            $relNode.SetAttribute('Type', 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility')
            $relNode.SetAttribute('Target', 'customUI/customUI14.xml')
            $relsXml.DocumentElement.AppendChild($relNode) | Out-Null
            $relsXml.Save($relsPath)
        }

        # Repack as xlam
        Remove-Item $outputPath -Force
        Compress-Archive -Path (Join-Path $tempDir '*') -DestinationPath $zipPath -Force
        Move-Item $zipPath $outputPath -Force
        Remove-Item $tempDir -Recurse -Force

        Write-Host '  customUI injected.' -ForegroundColor Green
    }
}

Write-Host ''
Write-Host "Build complete: $outputPath" -ForegroundColor Green
