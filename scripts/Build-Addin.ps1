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
$distDir = Join-Path $projectDir 'dist'

if ([string]::IsNullOrWhiteSpace($OutputName)) {
    $OutputName = "casedesk.$OutputFormat"
}
$outputPath = Join-Path $distDir $OutputName

function Get-ModuleFiles {
    param([string]$Root)

    $patterns = @('*.bas', '*.cls', '*.frm')
    $items = foreach ($pattern in $patterns) {
        Get-ChildItem -Path $Root -Recurse -File -Filter $pattern | Sort-Object FullName
    }
    return @($items)
}

function Set-CodeModuleText {
    param(
        [object]$CodeModule,
        [string]$Code
    )

    if ($CodeModule.CountOfLines -gt 0) {
        $CodeModule.DeleteLines(1, $CodeModule.CountOfLines)
    }
    if (-not [string]::IsNullOrWhiteSpace($Code)) {
        $CodeModule.AddFromString($Code)
    }
}

function Release-ComObject {
    param([object]$ComObject)

    if ($null -eq $ComObject) { return }
    if ([System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

if (-not (Test-Path -LiteralPath $distDir)) {
    New-Item -ItemType Directory -Path $distDir -Force | Out-Null
}
if (Test-Path -LiteralPath $outputPath) {
    Remove-Item -LiteralPath $outputPath -Force
}

$moduleFiles = Get-ModuleFiles -Root $srcDir
if ($moduleFiles.Count -eq 0) {
    throw 'No VBA source files were found under src.'
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null
$vbProj = $null
$docComp = $null

try {
    $wb = $excel.Workbooks.Add()
    $vbProj = $wb.VBProject
    if ($null -eq $vbProj) {
        throw 'VBA project access is not trusted.'
    }
    try {
        $null = $vbProj.VBComponents.Count
    } catch {
        throw 'Enable Trust access to the VBA project object model in Excel.'
    }

    foreach ($file in $moduleFiles) {
        $vbProj.VBComponents.Import($file.FullName) | Out-Null
    }

    # --- ThisWorkbook code (differs by format) ---
    if ($OutputFormat -eq 'xlam') {
        $thisWorkbookLines = @(
            'Option Explicit',
            '',
            'Private Sub Workbook_Open()',
            '    CaseDeskMain.InitAddin',
            'End Sub',
            '',
            'Private Sub Workbook_AddinInstall()',
            '    CaseDeskMain.InitAddin',
            'End Sub',
            '',
            'Private Sub Workbook_AddinUninstall()',
            '    CaseDeskMain.ShutdownAddin',
            'End Sub',
            '',
            'Private Sub Workbook_BeforeClose(Cancel As Boolean)',
            '    On Error Resume Next',
            '    CaseDeskMain.ShutdownAddin',
            '    Me.Saved = True',
            '    On Error GoTo 0',
            'End Sub',
            '',
            'Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)',
            '    On Error Resume Next',
            '    Dim sn As String: sn = Sh.Name',
            '    If Left$(sn, Len("_casedesk")) <> "_casedesk" Then Exit Sub',
            '    Application.ScreenUpdating = False',
            '    If CaseDeskMain.g_formLoaded Then frmCaseDesk.OnCaseDeskSheetChange sn',
            '    Application.ScreenUpdating = True',
            '    On Error GoTo 0',
            'End Sub'
        )
    } else {
        $thisWorkbookLines = @(
            'Option Explicit',
            '',
            'Private Sub Workbook_Open()',
            '    CaseDeskMain.InitAddin',
            'End Sub',
            '',
            'Private Sub Workbook_BeforeClose(Cancel As Boolean)',
            '    On Error Resume Next',
            '    CaseDeskMain.BeforeWorkbookClose',
            '    Me.Saved = True',
            'End Sub',
            '',
            'Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)',
            '    On Error Resume Next',
            '    Dim sn As String: sn = Sh.Name',
            '    If Left$(sn, Len("_casedesk")) <> "_casedesk" Then Exit Sub',
            '    Application.ScreenUpdating = False',
            '    If CaseDeskMain.g_formLoaded Then frmCaseDesk.OnCaseDeskSheetChange sn',
            '    Application.ScreenUpdating = True',
            '    On Error GoTo 0',
            'End Sub'
        )
    }
    $thisWorkbookCode = [string]::Join("`r`n", $thisWorkbookLines)

    $docComp = $vbProj.VBComponents.Item('ThisWorkbook')
    Set-CodeModuleText -CodeModule $docComp.CodeModule -Code $thisWorkbookCode

    # --- Pre-create config sheets ---
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
    $srcSheet.Range("B1").Value2 = "source_sheet"
    $srcSheet.Range("C1").Value2 = "key_column"
    $srcSheet.Range("D1").Value2 = "display_name_column"
    $srcSheet.Range("E1").Value2 = "mail_link_column"
    $srcSheet.Range("F1").Value2 = "folder_link_column"
    $srcSheet.Range("G1").Value2 = "mail_match_mode"
    if ($Sample) {
        $sampleXlsx = Join-Path $sampleDir 'casedesk-sample.xlsx'
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
            $srcSheet.Range("B2").Value2 = $sampleTbl.Range.Worksheet.Name
            $srcSheet.Range("C2").Value2 = $sampleTbl.ListColumns(1).Name
            $srcSheet.Range("D2").Value2 = $sampleTbl.ListColumns(2).Name
            foreach ($col in $sampleTbl.ListColumns) {
                if ($col.DataBodyRange -and $col.DataBodyRange.Cells(1,1).Text -match '@') {
                    $srcSheet.Range("E2").Value2 = $col.Name; break
                }
            }
            $srcSheet.Range("F2").Value2 = $sampleTbl.ListColumns(1).Name
        }
        $sampleWb.Close($false)
        Release-ComObject -ComObject $sampleWb
    }

    # _casedesk_fields: one row per source+field
    $fldSheet = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $fldSheet.Name = "_casedesk_fields"
    $fldSheet.Visible = 2  # xlSheetVeryHidden
    $fldSheet.Range("A1").Value2 = "source_name"
    $fldSheet.Range("B1").Value2 = "field_name"
    $fldSheet.Range("C1").Value2 = "display_name"
    $fldSheet.Range("D1").Value2 = "type"
    $fldSheet.Range("E1").Value2 = "visible"
    $fldSheet.Range("F1").Value2 = "in_list"
    $fldSheet.Range("G1").Value2 = "editable"
    $fldSheet.Range("H1").Value2 = "multiline"
    $fldSheet.Range("I1").Value2 = "role"
    $fldSheet.Range("J1").Value2 = "sort_order"

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
    $logTable = $logSheet.ListObjects.Add(1, $logSheet.Range("A1:G1"), $null, 1)
    $logTable.Name = "CaseDeskLog"

    if ($Sample) {
        Write-Host "  config: mail=$mailDir, cases=$casesDir" -ForegroundColor Green
    }

    # --- Save ---
    $fileFormat = if ($OutputFormat -eq 'xlam') { 55 } else { 52 }
    if ($OutputFormat -eq 'xlam') {
        $wb.IsAddin = $true
    }
    $wb.SaveAs($outputPath, $fileFormat)
    Write-Host "Built add-in: $outputPath" -ForegroundColor Green
} finally {
    Release-ComObject -ComObject $docComp
    Release-ComObject -ComObject $vbProj
    if ($null -ne $wb) {
        try { $wb.Close($false) } catch {}
        Release-ComObject -ComObject $wb
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject -ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# --- Inject customUI XML into xlam (ZIP manipulation) ---
if ($OutputFormat -eq 'xlam') {
    $customUIPath = Join-Path $srcDir 'customUI14.xml'
    if (Test-Path -LiteralPath $customUIPath) {
        Write-Host 'Injecting customUI ribbon...' -ForegroundColor Cyan
        $tempDir = Join-Path $env:TEMP ("casedesk_build_" + (Get-Random))
        $zipPath = $outputPath + '.zip'

        Copy-Item -LiteralPath $outputPath -Destination $zipPath -Force
        Expand-Archive -Path $zipPath -DestinationPath $tempDir -Force
        Remove-Item -LiteralPath $zipPath -Force

        $cuiDir = Join-Path $tempDir 'customUI'
        New-Item -ItemType Directory -Path $cuiDir -Force | Out-Null
        Copy-Item -LiteralPath $customUIPath -Destination (Join-Path $cuiDir 'customUI14.xml') -Force

        $ctPath = Join-Path $tempDir '[Content_Types].xml'
        $ctXml = [xml](Get-Content -LiteralPath $ctPath)
        $ctNs = $ctXml.DocumentElement.NamespaceURI
        $existingCt = $ctXml.Types.Override | Where-Object { $_.PartName -eq '/customUI/customUI14.xml' }
        if (-not $existingCt) {
            $node = $ctXml.CreateElement('Override', $ctNs)
            $node.SetAttribute('PartName', '/customUI/customUI14.xml')
            $node.SetAttribute('ContentType', 'application/xml')
            $ctXml.DocumentElement.AppendChild($node) | Out-Null
            $ctXml.Save($ctPath)
        }

        $relsPath = Join-Path $tempDir '_rels\.rels'
        $relsXml = [xml](Get-Content -LiteralPath $relsPath)
        $relsNs = $relsXml.DocumentElement.NamespaceURI
        $existingRel = $relsXml.Relationships.Relationship | Where-Object { $_.Target -eq 'customUI/customUI14.xml' }
        if (-not $existingRel) {
            $relNode = $relsXml.CreateElement('Relationship', $relsNs)
            $relNode.SetAttribute('Id', 'rIdCustomUI')
            $relNode.SetAttribute('Type', 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility')
            $relNode.SetAttribute('Target', 'customUI/customUI14.xml')
            $relsXml.DocumentElement.AppendChild($relNode) | Out-Null
            $relsXml.Save($relsPath)
        }

        Remove-Item -LiteralPath $outputPath -Force
        Compress-Archive -Path (Join-Path $tempDir '*') -DestinationPath $zipPath -Force
        Move-Item -LiteralPath $zipPath -Destination $outputPath -Force
        Remove-Item -LiteralPath $tempDir -Recurse -Force
        Write-Host 'customUI injected.' -ForegroundColor Green
    }
}
