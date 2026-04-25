# Build-Sample.ps1
# サンプルデータ（Excel台帳 + メールアーカイブ + 案件フォルダ）を自動生成する
#
# 使い方:
#   powershell -ExecutionPolicy Bypass -File Build-Sample.ps1
#   powershell -ExecutionPolicy Bypass -File Build-Sample.ps1 -Count 500

param(
    [int]$Count = 1000
)

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
$sampleOut = Join-Path $projectDir 'sample'

# --- Data pools ---
$lastNames = @('山田','佐藤','高橋','伊藤','渡辺','中村','加藤','小林','松本','石井',
    '吉田','森','池田','橋本','藤田','前田','岡田','長谷川','村上','近藤',
    '清水','木村','林','斎藤','坂本','福田','太田','三浦','上田','西村')
$firstNames = @('太郎','花子','誠','健','健一','雅子','翔太','美咲','大輔','由美',
    '隆','直子','浩二','恵子','拓也','裕子','一郎','幸子','龍也','真由美',
    '光','和子','修','明日香','陽一','千春','正樹','亮','純','綾')
$orgPrefixes = @('北陸','東北','関東','関西','九州','中部','東海','四国','北海道','沖縄',
    '信越','山陰','山陽','首都圏','南東北','北関東','南関東','北九州','南九州','中国')
$orgSuffixes = @('地域振興協会','環境保全機構','教育推進ネットワーク','文化継承センター',
    'スポーツ振興クラブ','福祉支援機構','観光推進協議会','産業振興会',
    '子ども支援センター','まちづくり協議会','技術振興財団','農業振興協会',
    '健康推進センター','防災支援機構','国際交流協会')
$statuses = @('受付済','書類確認中','審査中','書類不備','審査完了','交付決定')
$staffNames = @('鈴木','田中','佐々木','山本','中野','井上','小川','大西')
$docNames = @('予算内訳明細','定款','事業報告書','見積書','決算報告書','役員名簿','収支計算書','組織図')
$subjects = @('交付金申請書類の送付','交付金申請について','申請書類の提出',
    '不足書類の送付','交付金に関するお問い合わせ','書類修正のお知らせ',
    '追加書類の送付','申請内容の変更について')
$domains = @('hokuriku-shinko.or.jp','kodomo-mirai.org','kankyo-suishin.or.jp',
    'sports-kitakanto.or.jp','digital-edu.net','dentou-bunka.or.jp',
    'green-energy-tohoku.co.jp','fukushi-net.or.jp','kanko-suishin.jp',
    'sangyo-shinko.or.jp','machizukuri.or.jp','nogyo-shinko.or.jp',
    'kenkou-center.or.jp','bousai-net.or.jp','kokusai-koryu.or.jp')
$fileNames = @('application.pdf','budget.xlsx','project_plan.pdf','estimate.pdf',
    'articles.pdf','approval_letter.pdf','activity_report.pdf',
    'organization_profile.pdf','reduction_note.pdf','checklist.xlsx')

$rng = [System.Random]::new(42)  # reproducible seed

function Pick($arr) { return $arr[$rng.Next($arr.Count)] }
function RandInt($lo, $hi) { return $rng.Next($lo, $hi + 1) }

# --- Helper: write UTF-8 without BOM ---
function Write-Utf8NoBom($path, $content) {
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($path, $content, $utf8)
}

function JsonEsc($s) {
    return $s.Replace('\', '\\').Replace('"', '\"').Replace("`n", '\n').Replace("`r", '')
}

# ============================================================================
# 1. Generate Excel workbook (single anken table)
# ============================================================================

Write-Host "Starting Excel..." -ForegroundColor Cyan
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()
    while ($wb.Sheets.Count -gt 1) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }
    $ws = $wb.Sheets.Item(1)
    $ws.Name = 'anken'

    # Headers — designed to demonstrate all prefix/tab features:
    #   基本_xxx    → "基本" tab (normal, editable)
    #   審査_xxx    → "審査" tab (normal, editable)
    #   _基本_登録日 → "基本" tab (readonly, _xx_AAA prefix)
    #   __内部メモ   → hidden in UI (__AAA prefix), usable as setting column
    # 12 prefix groups → 12 tabs, 36 columns total
    $headers = @(
        '案件ID',               #  1  key (no prefix)
        '基本_団体名',           #  2  基本
        '基本_代表者',           #  3
        '基本_メールアドレス',    #  4
        '基本_電話番号',         #  5
        '基本_住所',             #  6
        '基本_申請日',           #  7
        '基本_申請額',           #  8
        '審査_ステータス',       #  9  審査
        '審査_担当者',           # 10
        '審査_不足書類',         # 11
        '審査_備考',             # 12
        '審査_期限',             # 13
        '審査_スコア',           # 14
        '交付_決定日',           # 15  交付
        '交付_交付額',           # 16
        '交付_条件',             # 17
        '報告_中間報告日',       # 18  報告
        '報告_最終報告日',       # 19
        '報告_実績額',           # 20
        '報告_差額',             # 21
        '経理_伝票番号',         # 22  経理
        '経理_支払日',           # 23
        '経理_支払額',           # 24
        '連絡_最終連絡日',       # 25  連絡
        '連絡_連絡手段',         # 26
        '連絡_次回予定',         # 27
        '法務_契約番号',         # 28  法務
        '法務_契約日',           # 29
        '法務_契約状態',         # 30
        'IT_システムID',         # 31  IT
        'IT_登録状態',           # 32
        '管理_作成者',           # 33  管理
        '管理_更新者',           # 34
        '_基本_登録日',          # 35  readonly
        '__内部メモ'             # 36  hidden
    )
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $ws.Cells.Item(1, $c + 1).Value2 = $headers[$c]
    }

    # Generate rows
    Write-Host "Generating $Count rows..." -ForegroundColor Cyan
    $script:caseEmails = @{}
    $contactMethods = @('電話','メール','訪問','Web会議','FAX')
    $contractStatuses = @('有効','期限切れ','交渉中','解除')
    $itStatuses = @('登録済','未登録','申請中')

    for ($r = 1; $r -le $Count; $r++) {
        $caseId = 'R06-' + $r.ToString('000')
        $orgName = (Pick $orgPrefixes) + (Pick $orgSuffixes)
        $personLast = Pick $lastNames
        $personFirst = Pick $firstNames
        $personName = "$personLast $personFirst"
        $romanLast = [char](97 + ($rng.Next() % 26))
        $email = $romanLast + (RandInt 100 999).ToString() + '@' + (Pick $domains)
        $baseDate = [datetime]'2024-04-01'
        $applyDate = $baseDate.AddDays((RandInt 0 180))
        $amount = (RandInt 5 100) * 100000
        $status = Pick $statuses
        $staff = Pick $staffNames
        $missingDoc = if ($status -eq '書類不備') { Pick $docNames } else { '' }
        $memo = if ((RandInt 1 5) -eq 1) { "備考メモ$r`nこの案件は要確認です。" } else { '' }
        $regDate = $applyDate.AddDays(-(RandInt 1 10)).ToString('yyyy/MM/dd')
        $internalNote = "内部ID:$r priority:$(RandInt 1 5)"
        $grantDate = $applyDate.AddDays((RandInt 30 90)).ToString('yyyy/MM/dd')
        $grantAmount = [double]([Math]::Floor($amount * (RandInt 70 100) / 100))
        $midDate = $applyDate.AddDays((RandInt 60 120)).ToString('yyyy/MM/dd')
        $finalDate = $applyDate.AddDays((RandInt 150 240)).ToString('yyyy/MM/dd')
        $actualAmount = [double]([Math]::Floor($grantAmount * (RandInt 80 110) / 100))
        $payDate = $applyDate.AddDays((RandInt 90 150)).ToString('yyyy/MM/dd')
        $contactDate = $applyDate.AddDays((RandInt 10 60)).ToString('yyyy/MM/dd')
        $nextContact = $applyDate.AddDays((RandInt 70 130)).ToString('yyyy/MM/dd')
        $contractDate = $applyDate.AddDays(-(RandInt 30 365)).ToString('yyyy/MM/dd')

        $row = $r + 1
        $ws.Cells.Item($row, 1).Value2 = [string]$caseId
        $ws.Cells.Item($row, 2).Value2 = [string]$orgName
        $ws.Cells.Item($row, 3).Value2 = [string]$personName
        $ws.Cells.Item($row, 4).Value2 = [string]$email
        $ws.Cells.Item($row, 5).Value2 = '0' + (RandInt 3 9).ToString() + '-' + (RandInt 1000 9999).ToString() + '-' + (RandInt 1000 9999).ToString()
        $ws.Cells.Item($row, 6).Value2 = (Pick $orgPrefixes) + '市' + (Pick $lastNames) + '町' + (RandInt 1 20).ToString()
        $ws.Cells.Item($row, 7).Value2 = [string]$applyDate.ToString('yyyy/MM/dd')
        $ws.Cells.Item($row, 8).Value2 = [double]$amount
        $ws.Cells.Item($row, 9).Value2 = [string]$status
        $ws.Cells.Item($row, 10).Value2 = [string]$staff
        $ws.Cells.Item($row, 11).Value2 = [string]$missingDoc
        $ws.Cells.Item($row, 12).Value2 = [string]$memo
        $ws.Cells.Item($row, 13).Value2 = [string]$applyDate.AddDays((RandInt 14 30)).ToString('yyyy/MM/dd')
        $ws.Cells.Item($row, 14).Value2 = [double](RandInt 1 100)
        $ws.Cells.Item($row, 15).Value2 = [string]$grantDate
        $ws.Cells.Item($row, 16).Value2 = [double]$grantAmount
        $ws.Cells.Item($row, 17).Value2 = if ((RandInt 1 3) -eq 1) { '条件付き' } else { '' }
        $ws.Cells.Item($row, 18).Value2 = [string]$midDate
        $ws.Cells.Item($row, 19).Value2 = [string]$finalDate
        $ws.Cells.Item($row, 20).Value2 = [double]$actualAmount
        $ws.Cells.Item($row, 21).Value2 = [double]($actualAmount - $grantAmount)
        $ws.Cells.Item($row, 22).Value2 = 'D-' + $r.ToString('0000')
        $ws.Cells.Item($row, 23).Value2 = [string]$payDate
        $ws.Cells.Item($row, 24).Value2 = [double]$grantAmount
        $ws.Cells.Item($row, 25).Value2 = [string]$contactDate
        $ws.Cells.Item($row, 26).Value2 = [string](Pick $contactMethods)
        $ws.Cells.Item($row, 27).Value2 = [string]$nextContact
        $ws.Cells.Item($row, 28).Value2 = 'C-' + (RandInt 1 500).ToString('0000')
        $ws.Cells.Item($row, 29).Value2 = [string]$contractDate
        $ws.Cells.Item($row, 30).Value2 = [string](Pick $contractStatuses)
        $ws.Cells.Item($row, 31).Value2 = 'SYS-' + $r.ToString('0000')
        $ws.Cells.Item($row, 32).Value2 = [string](Pick $itStatuses)
        $ws.Cells.Item($row, 33).Value2 = [string](Pick $staffNames)
        $ws.Cells.Item($row, 34).Value2 = [string](Pick $staffNames)
        $ws.Cells.Item($row, 35).Value2 = [string]$regDate
        $ws.Cells.Item($row, 36).Value2 = [string]$internalNote

        $script:caseEmails[$caseId] = @{ Email = $email; Name = $personName }

        if ($r % 200 -eq 0) { Write-Host "  table: $r/$Count rows" }
    }

    # Format columns
    $fmtDate = 'yyyy/mm/dd'
    $fmtNum = '#,##0'
    foreach ($col in @(7, 13, 15, 18, 19, 23, 25, 27, 29, 35)) {
        $ws.Range($ws.Cells.Item(2, $col), $ws.Cells.Item($Count + 1, $col)).NumberFormat = $fmtDate
    }
    foreach ($col in @(8, 14, 16, 20, 21, 24)) {
        $ws.Range($ws.Cells.Item(2, $col), $ws.Cells.Item($Count + 1, $col)).NumberFormat = $fmtNum
    }

    # Create table
    $tblRange = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($Count + 1, $headers.Count))
    $lo = $ws.ListObjects.Add(1, $tblRange, $null, 1)
    $lo.Name = 'anken'
    $lo.TableStyle = 'TableStyleMedium2'
    $ws.Columns.AutoFit() | Out-Null

    # ================================================================
    # Sheet 2: "他社台帳" — manual mapping test target
    #   - No ListObject (UsedRange only)
    #   - No prefix naming convention
    #   - CamelCase ID column (GuessFieldRole test)
    #   - Currency column with yen format (currency type test)
    # ================================================================
    Write-Host "Generating manual-mapping sheet..." -ForegroundColor Cyan
    $ws2 = $wb.Sheets.Add([System.Type]::Missing, $wb.Sheets.Item($wb.Sheets.Count))
    $ws2.Name = '他社台帳'

    $extHeaders = @(
        'RecordId',        # CamelCase ID — should auto-guess case_id
        '件名',            # title candidate (Japanese)
        '申請者名',        # person name — no obvious role
        '連絡先',          # contact — could be mail_link
        '提出日',          # date
        '請求金額',        # currency (yen format)
        '進捗',            # status candidate (Japanese)
        'フォルダ名',      # folder — could be file_key
        '備考欄'           # freeform text (multiline)
    )
    for ($c = 0; $c -lt $extHeaders.Count; $c++) {
        $ws2.Cells.Item(1, $c + 1).Value2 = $extHeaders[$c]
        $ws2.Cells.Item(1, $c + 1).Font.Bold = $true
    }

    $extStatuses = @('未着手','対応中','確認待ち','完了','保留')
    $extCount = [Math]::Min($Count, 50)  # 50 rows is enough for manual mapping test
    for ($r = 1; $r -le $extCount; $r++) {
        $row = $r + 1
        $ws2.Cells.Item($row, 1).Value2 = 'EXT-' + $r.ToString('0000')
        $ws2.Cells.Item($row, 2).Value2 = (Pick $orgPrefixes) + '案件' + $r.ToString()
        $ws2.Cells.Item($row, 3).Value2 = (Pick $lastNames) + ' ' + (Pick $firstNames)
        $romanLast = [char](97 + ($rng.Next() % 26))
        $ws2.Cells.Item($row, 4).Value2 = $romanLast + (RandInt 100 999).ToString() + '@' + (Pick $domains)
        $baseDate = ([datetime]'2024-06-01').AddDays((RandInt 0 120))
        $ws2.Cells.Item($row, 5).Value2 = $baseDate.ToString('yyyy/MM/dd')
        $ws2.Cells.Item($row, 6).Value2 = [double]((RandInt 10 500) * 10000)
        $ws2.Cells.Item($row, 7).Value2 = Pick $extStatuses
        $ws2.Cells.Item($row, 8).Value2 = 'EXT-' + $r.ToString('0000')
        $memo = if ((RandInt 1 3) -eq 1) { "注意事項あり`n要確認" } else { '' }
        $ws2.Cells.Item($row, 9).Value2 = $memo
    }

    # Format: date and yen currency (triggers GuessFieldType "currency")
    $ws2.Range("E2:E$($extCount+1)").NumberFormat = 'yyyy/mm/dd'
    $ws2.Range("F2:F$($extCount+1)").NumberFormat = [char]0xA5 + '#,##0'
    $ws2.Columns.AutoFit() | Out-Null

    Write-Host "  manual-mapping sheet: $extCount rows (no ListObject)" -ForegroundColor Green

    # ================================================================
    # Additional sheets (30 columns each, 8 sheets = 10 total)
    # ================================================================
    Write-Host "Generating additional sheets..." -ForegroundColor Cyan

    $extraSheets = @(
        @{ Name = '顧客管理'; Table = 'kokyaku'; Headers = @(
            '顧客ID','顧客名','フリガナ','法人区分','代表者','郵便番号','住所','電話番号','FAX','メール',
            '担当営業','ランク','業種','設立日','資本金','従業員数','年商','取引開始日','最終取引日','累計取引額',
            'Webサイト','請求先住所','請求先担当','支払条件','与信限度額','取引銀行','口座番号','備考','登録日','更新日')
        },
        @{ Name = '契約管理'; Table = 'keiyaku'; Headers = @(
            '契約番号','契約名','顧客ID','契約種別','契約開始日','契約終了日','月額','年額','初期費用','支払サイクル',
            'ステータス','自動更新','担当者','担当メール','承認者','承認日','解約予告期間','SLA区分','上限ユーザ数','現ユーザ数',
            '割引率','請求先','納品先','検収条件','ライセンスキー','サポートレベル','次回更新日','備考','作成日','更新日')
        }
    )

    $extraRowCount = [Math]::Min($Count, 50)
    foreach ($sheetDef in $extraSheets) {
        $wsN = $wb.Sheets.Add([System.Type]::Missing, $wb.Sheets.Item($wb.Sheets.Count))
        $wsN.Name = $sheetDef.Name
        $hdrs = $sheetDef.Headers
        for ($c = 0; $c -lt $hdrs.Count; $c++) {
            $wsN.Cells.Item(1, $c + 1).Value2 = $hdrs[$c]
        }
        for ($r = 1; $r -le $extraRowCount; $r++) {
            $row = $r + 1
            for ($c = 0; $c -lt $hdrs.Count; $c++) {
                $wsN.Cells.Item($row, $c + 1).Value2 = "sample-$r-$($c+1)"
            }
            $wsN.Cells.Item($row, 1).Value2 = $sheetDef.Table.Substring(0,1).ToUpper() + '-' + $r.ToString('000')
        }
        $tblR = $wsN.Range($wsN.Cells.Item(1, 1), $wsN.Cells.Item($extraRowCount + 1, $hdrs.Count))
        $loN = $wsN.ListObjects.Add(1, $tblR, $null, 1)
        $loN.Name = $sheetDef.Table
        $loN.TableStyle = 'TableStyleMedium2'
        $wsN.Columns.AutoFit() | Out-Null
        Write-Host "  $($sheetDef.Name): $extraRowCount rows, $($hdrs.Count) columns" -ForegroundColor Green
    }

    # Save
    if (-not (Test-Path $sampleOut)) { New-Item -ItemType Directory -Path $sampleOut -Force | Out-Null }
    $outPath = Join-Path $sampleOut 'casedesk-sample.xlsx'
    if (Test-Path $outPath) { Remove-Item $outPath -Force }
    $wb.SaveAs($outPath, 51)

    Write-Host "Workbook saved: $outPath ($Count + $extCount rows, 2 sheets)" -ForegroundColor Green

} finally {
    if ($wb) { $wb.Close($false); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
}

# ============================================================================
# 2. Generate mail archive
# ============================================================================

Write-Host ''
Write-Host 'Creating mail archive...' -ForegroundColor Cyan

$mailOut = Join-Path $sampleOut 'mail'
if (Test-Path $mailOut) { Remove-Item $mailOut -Recurse -Force }
New-Item -ItemType Directory -Path $mailOut -Force | Out-Null

$mailNum = 0
for ($r = 1; $r -le $Count; $r++) {
    $caseId = 'R06-' + $r.ToString('000')
    $info = $script:caseEmails[$caseId]
    $mailCount = RandInt 1 3

    for ($m = 0; $m -lt $mailCount; $m++) {
        $mailNum++
        $folderName = 'mail_' + $mailNum.ToString('0000')
        $mailId = 'MAIL-' + $mailNum.ToString('0000')
        $entryId = $mailNum.ToString('00000000')

        # Use case owner for first mail, random for subsequent
        if ($m -eq 0) {
            $senderName = $info.Name
            $senderEmail = $info.Email
        } else {
            $senderName = (Pick $lastNames) + ' ' + (Pick $firstNames)
            $romanLast = [char](97 + ($rng.Next() % 26))
            $senderEmail = $romanLast + (RandInt 100 999).ToString() + '@' + (Pick $domains)
        }

        $baseDate = ([datetime]'2024-04-01').AddDays((RandInt 0 180))
        $hour = RandInt 8 17
        $minute = RandInt 0 59
        $recvAt = $baseDate.ToString('yyyy-MM-dd') + 'T' + $hour.ToString('00') + ':' + $minute.ToString('00') + ':00+09:00'
        $subj = (Pick $subjects) + "（$caseId）"

        $bodyText = "お世話になっております。`n$senderName です。`n`n案件${caseId}に関する書類をお送りいたします。"

        $attCount = RandInt 1 3
        $atts = @('application.pdf')
        for ($a = 1; $a -lt $attCount; $a++) {
            $atts += "doc_$a.pdf"
        }

        # Build JSON
        $attJson = ($atts | ForEach-Object { "{ `"path`": `"$_`" }" }) -join ', '
        $json = @"
{
  "mail_id": "$mailId",
  "entry_id": "$entryId",
  "mailbox_address": "review@example.org",
  "folder_path": "$(JsonEsc '受信トレイ/交付金申請')",
  "received_at": "$recvAt",
  "sender_name": "$(JsonEsc $senderName)",
  "sender_email": "$senderEmail",
  "subject": "$(JsonEsc $subj)",
  "body_path": "body.txt",
  "msg_path": "",
  "attachments": [$attJson]
}
"@

        $dir = Join-Path $mailOut $folderName
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Write-Utf8NoBom (Join-Path $dir 'meta.json') $json
        Write-Utf8NoBom (Join-Path $dir 'body.txt') $bodyText
        foreach ($att in $atts) {
            Write-Utf8NoBom (Join-Path $dir $att) "(sample file: $att)"
        }
    }

    if ($r % 200 -eq 0) { Write-Host "  mail: $r/$Count cases processed ($mailNum mails)" }
}
Write-Host "  mail: $mailNum folders created" -ForegroundColor Green

# ============================================================================
# 3. Generate case folders
# ============================================================================

Write-Host ''
Write-Host 'Creating case folders...' -ForegroundColor Cyan

$casesOut = Join-Path $sampleOut 'cases'
if (Test-Path $casesOut) { Remove-Item $casesOut -Recurse -Force }
New-Item -ItemType Directory -Path $casesOut -Force | Out-Null

for ($r = 1; $r -le $Count; $r++) {
    $caseId = 'R06-' + $r.ToString('000')
    $dir = Join-Path $casesOut $caseId
    New-Item -ItemType Directory -Path $dir -Force | Out-Null

    # Main files (2-5)
    $fCount = RandInt 2 5
    for ($f = 0; $f -lt $fCount; $f++) {
        $fn = $fileNames[$f]
        $fp = Join-Path $dir $fn
        Write-Utf8NoBom $fp "(sample file: $caseId/$fn)"
    }

    # Optional review subfolder (30% chance)
    if ((RandInt 1 10) -le 3) {
        $revDir = Join-Path $dir 'review'
        New-Item -ItemType Directory -Path $revDir -Force | Out-Null
        Write-Utf8NoBom (Join-Path $revDir 'checklist.txt') "審査チェックリスト`n- [ ] 申請書確認`n- [ ] 予算書確認"
        if ((RandInt 1 2) -eq 1) {
            Write-Utf8NoBom (Join-Path $revDir 'memo.txt') "審査メモ $caseId"
        }
    }

    if ($r % 200 -eq 0) { Write-Host "  cases: $r/$Count folders created" }
}
Write-Host "  cases: $Count folders created" -ForegroundColor Green

# ============================================================================
# 4. Generate manifest.csv for mail (watchbox-compatible format)
# Format: entry_id,sender_email,sender_name,subject,received_at,
#         folder_path,body_path,msg_path,attachment_paths,mail_folder,body_text
# ============================================================================

Write-Host ''
Write-Host 'Generating mail manifest.csv...' -ForegroundColor Cyan

$mailManifest = "entry_id,sender_email,sender_name,subject,received_at,folder_path,body_path,msg_path,attachment_paths,mail_folder,body_text`n"
$mailDirs = Get-ChildItem -Path $mailOut -Directory | Sort-Object Name
foreach ($dir in $mailDirs) {
    $metaPath = Join-Path $dir.FullName 'meta.json'
    if (-not (Test-Path $metaPath)) { continue }
    $metaText = [System.IO.File]::ReadAllText($metaPath, [System.Text.Encoding]::UTF8)

    # Simple JSON parse (no external dependency)
    function ExtractJsonValue($json, $key) {
        if ($json -match "`"$key`"\s*:\s*`"([^`"]*)`"") { return $matches[1] }
        return ""
    }

    $entryId = ExtractJsonValue $metaText 'entry_id'
    $senderEmail = ExtractJsonValue $metaText 'sender_email'
    $senderName = ExtractJsonValue $metaText 'sender_name'
    $subject = ExtractJsonValue $metaText 'subject'
    $receivedAt = ExtractJsonValue $metaText 'received_at'
    $folderPathVal = $dir.FullName
    $bodyPath = Join-Path $dir.FullName 'body.txt'
    $msgPath = ''
    $mailFolder = ExtractJsonValue $metaText 'folder_path'

    # Collect attachment paths (pipe-separated)
    $attPaths = @()
    $attFiles = Get-ChildItem -Path $dir.FullName -File | Where-Object { $_.Name -ne 'meta.json' -and $_.Name -ne 'body.txt' }
    foreach ($att in $attFiles) { $attPaths += $att.FullName }
    $attPathStr = $attPaths -join '|'

    # Read body text (first 500 chars, sanitize commas/newlines)
    $bodyText = ''
    if (Test-Path $bodyPath) {
        $bodyText = [System.IO.File]::ReadAllText($bodyPath, [System.Text.Encoding]::UTF8)
        if ($bodyText.Length -gt 500) { $bodyText = $bodyText.Substring(0, 500) }
        $bodyText = $bodyText.Replace(',', ' ').Replace("`r`n", ' ').Replace("`n", ' ').Replace("`r", ' ')
    }

    # Sanitize commas in fields
    $senderName = $senderName.Replace(',', ' ')
    $subject = $subject.Replace(',', ' ')

    $mailManifest += "$entryId,$senderEmail,$senderName,$subject,$receivedAt,$folderPathVal,$bodyPath,$msgPath,$attPathStr,$mailFolder,$bodyText`n"
}
Write-Utf8NoBom (Join-Path $mailOut 'manifest.csv') $mailManifest
Write-Host "  manifest.csv written ($($mailDirs.Count) entries)" -ForegroundColor Green

# ============================================================================
# 5. Generate manifest.csv for cases (watchbox-compatible format)
# Format: item_id,file_name,file_path,folder_path,relative_path,file_size,modified_at
# ============================================================================

Write-Host 'Generating cases manifest.csv...' -ForegroundColor Cyan

$caseManifest = "item_id,file_name,file_path,folder_path,relative_path,file_size,modified_at`n"
$caseFileCount = 0
$allFiles = Get-ChildItem -Path $casesOut -Recurse -File
foreach ($file in $allFiles) {
    $relativePath = $file.FullName.Substring($casesOut.Length + 1)
    # item_id = first 16 hex chars of SHA256(lowercase relative path with forward slashes)
    $normalizedPath = $relativePath.ToLower().Replace('\', '/')
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $hashBytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($normalizedPath))
    $itemId = [BitConverter]::ToString($hashBytes).Replace('-','').Substring(0,16).ToLower()

    $fileName = $file.Name
    $filePath = $file.FullName
    $folderPath = $file.DirectoryName
    $fileSize = $file.Length
    $modifiedAt = $file.LastWriteTime.ToString('yyyy-MM-ddTHH:mm:ss')

    $caseManifest += "$itemId,$fileName,$filePath,$folderPath,$relativePath,$fileSize,$modifiedAt`n"
    $caseFileCount++
}
Write-Utf8NoBom (Join-Path $casesOut 'manifest.csv') $caseManifest
Write-Host "  manifest.csv written ($caseFileCount files)" -ForegroundColor Green

Write-Host ''
Write-Host 'Sample data ready!' -ForegroundColor Green
Write-Host "  Workbook: sample\casedesk-sample.xlsx ($Count rows, 1 table)"
Write-Host "  Mail:     sample\mail\ ($mailNum folders + manifest.csv)"
Write-Host "  Cases:    sample\cases\ ($Count folders + manifest.csv)"
