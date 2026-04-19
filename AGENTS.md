# AGENTS.md - pub/casedesk

Codex 用の入口メモです。repo 固有の詳細は [README.md](README.md) と [CLAUDE.md](CLAUDE.md)、横断トポロジは [fin/hub/ARCHITECTURE.md](../../fin/hub/ARCHITECTURE.md) を参照してください。

## 役割

- Excel VBA の案件管理アドイン
- `watchbox` が生成した `manifest.csv` を読み、案件・メール・ファイルを 1 画面で扱う
- FE/BE を別 Excel プロセスで分離する

## runtime / 接続

- runtime: local Windows + Excel only
- data contract: `pub/watchbox` の `manifest.csv`, `log.csv`
- transport: hidden sheets への書き込み + `Workbook_SheetChange`

## まず見る場所

- `src/CaseDeskMain.bas` - エントリポイント / BE 管理
- `src/CaseDeskWorker.bas` - BE 側 scan / manifest 読み込み
- `src/CaseDeskData.bas` - FE 側キャッシュ
- `src/frmCaseDesk.frm`, `src/frmSettings.frm` - UI
- `docs/spec.md` - 詳細仕様

## 開発コマンド

```bat
samplerun.bat
build-addin.bat
build-sample.bat
```

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Test-Compile.ps1
powershell -ExecutionPolicy Bypass -File scripts/Test-Worker.ps1
```

## ガードレール

- ランタイムコードは VBA のみ。WinAPI を追加しない
- FE/BE 分離を崩さない
- `watchbox` の manifest 契約を片側だけで変えない
- hidden-sheet 名と Workbook event ベースの通信を前提にする
