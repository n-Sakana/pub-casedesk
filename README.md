# CaseDesk

Excel VBA の案件管理アドイン。開いているワークブックのテーブル (ListObject) を直接読み書きし、watchbox が生成した `manifest.csv` を読み込んでメール・案件ファイルを一画面で扱う。

## 基本原則

- 正本は Excel テーブルそのもの
- 外部データは watchbox 生成の `manifest.csv` を読む
- FE / BE は別 Excel プロセスで分離する
- 設定は Dictionary キャッシュで保持する
- 変更ログは ListObject で保持し、5000 行でローテーションする
- WinAPI は使わない

## セットアップ

### 前提条件

- Windows + Excel (Microsoft 365 / 2021 以降)
- Excel > ファイル > オプション > トラストセンター > マクロの設定 > 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を ON

### ビルド・実行

```bat
samplerun.bat            rem ビルド → sample を開く
build-addin.bat          rem casedesk.xlsm を生成
build-sample.bat         rem sample データを再生成
```

### 使い方

1. `samplerun.bat` を実行
2. `Alt+F8` → `CaseDesk_ShowPanel`
3. 左上ドロップダウンから対象テーブルを選ぶ
4. レコードを選択し、中央タブで詳細・メール・ファイルを確認する
5. 右カラムで変更ログを見る

## アーキテクチャ

```text
FE: casedesk.xlsm (ユーザーの Excel インスタンス)
  ├── frmCaseDesk        メイン UI
  ├── frmSettings        設定ダイアログ
  ├── CaseDeskMain       エントリポイント / BE 管理
  ├── CaseDeskData       FE 側キャッシュ / テーブル読み書き
  ├── CaseDeskLib        Config / ChangeLog / Utility
  ├── FieldEditor        WithEvents テキストボックス
  ├── SheetWatcher       WithEvents シート監視
  ├── ErrorHandler       エラー記録
  └── hidden sheets      _casedesk_signal, _casedesk_mail, _casedesk_cases, ...

BE: 別プロセスの Excel.Application (Visible=False)
  └── CaseDeskWorker     watchbox manifest 読み込み / ケース走査 / FE シート書き込み
```

### 通信フロー

```text
BE → FE: FE の hidden sheet へ .Value 書き込み → Workbook_SheetChange で受信
FE → BE: 起動時 / 設定変更時に Run 経由で設定更新
```

`_casedesk_request` ベースのリクエスト応答は、現行実装では使っていない。

### スイッチ式スキャンループ

```text
DoScanChunk (1秒ごと)
  ├─ TASK_MAIL   mail manifest.csv の更新確認
  ├─ TASK_CASES  case root / case manifest.csv の更新確認
  └─ TASK_WRITE  変更があれば FE hidden sheets へ反映
```

ラウンドロビンで mail / cases / write を回す。重い 5 秒ポーリングではなく、短い周期で小さく回す構成。

## モジュール一覧

本体は 9 モジュール。

| モジュール | 種別 | 役割 |
|---|---|---|
| CaseDeskMain.bas | bas | エントリポイント、BE 起動停止、PID 管理 |
| CaseDeskWorker.bas | bas | BE 側スキャン、manifest 読み込み、FE シート書き込み |
| CaseDeskData.bas | bas | FE 側キャッシュ、テーブル読み書き |
| CaseDeskLib.bas | bas | Config、ChangeLog、JSON、ファイル I/O |
| frmCaseDesk.frm | frm | 3 カラム UI、本体フォーム |
| frmSettings.frm | frm | 設定 UI |
| ErrorHandler.cls | cls | エラートレースと蓄積ログ |
| FieldEditor.cls | cls | フィールド編集イベント |
| SheetWatcher.cls | cls | テーブル変更監視 |

`frmResize` は存在しない。リサイズハンドルと左右スプリッターは `frmCaseDesk` 内蔵。

`frmCaseDeskV2` / `CaseDesk_ShowPanel2` も未実装で、`docs/frmCaseDeskV2-design.md` は構想メモ。

## データ入力

### メール

watchbox の mail profile が出力する `manifest.csv` を読む。

```text
mail_output/
  ├── manifest.csv
  └── ...mail folders...
```

フォーマットは 11 列。

```text
entry_id,sender_email,sender_name,subject,received_at,folder_path,body_path,msg_path,attachment_paths,mail_folder,body_text
```

### 案件ファイル

Case root 直下を走査するか、watchbox の folder profile が出力した `manifest.csv` を読む。現行は案件選択時オンデマンドではなく、起動後のスキャンで `m_caseFiles` に全件保持し、`_casedesk_files` に書き込む。

folder manifest の列は 7 列。

```text
item_id,file_name,file_path,folder_path,relative_path,file_size,modified_at
```

## hidden sheets

| シート | 用途 |
|---|---|
| `_casedesk_config` | 設定 KV |
| `_casedesk_sources` | ソース設定 |
| `_casedesk_fields` | フィールド設定 |
| `_casedesk_log` | 変更ログ |
| `_casedesk_signal` | バージョン / タイミング通知 |
| `_casedesk_mail` | メール一覧 |
| `_casedesk_mail_idx` | メール索引 |
| `_casedesk_cases` | 案件名一覧 |
| `_casedesk_files` | 案件ファイル一覧 |
| `_casedesk_diff` | 差分一覧 |

## テスト

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Test-ComVsTsv.ps1
powershell -ExecutionPolicy Bypass -File scripts/Test-Compile.ps1
powershell -ExecutionPolicy Bypass -File scripts/Test-CrossProcessEvents.ps1
powershell -ExecutionPolicy Bypass -File scripts/Test-FESheetChange.ps1
powershell -ExecutionPolicy Bypass -File scripts/Test-Refactoring.ps1
powershell -ExecutionPolicy Bypass -File scripts/Test-Worker.ps1
```

## ディレクトリ構成

```text
casedesk/
├── src/                     VBA ソース 9 モジュール
├── scripts/                 build / test スクリプト
├── sample/                  sample workbook / mail / cases
├── docs/
│   ├── spec.md
│   └── frmCaseDeskV2-design.md
├── samplerun.bat
├── build-addin.bat
├── build-sample.bat
└── CLAUDE.md
```
