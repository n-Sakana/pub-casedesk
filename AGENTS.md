# AGENTS.md - pub/casedesk

Entry-point notes for Codex. For repo-specific details see [README.md](README.md) and [CLAUDE.md](CLAUDE.md); for cross-cutting topology see [fin/hub/ARCHITECTURE.md](../../fin/hub/ARCHITECTURE.md).

## Role

- Excel VBA add-in for case management
- Reads the `manifest.csv` produced by `watchbox` and handles cases, mail, and files in a single screen
- Splits FE/BE into separate Excel processes

## runtime / connections

- runtime: local Windows + Excel only
- data contract: `manifest.csv`, `log.csv` from `pub/watchbox`
- transport: writes to hidden sheets + `Workbook_SheetChange`

## Where to look first

- `src/CaseDeskMain.bas` - entry point / BE management
- `src/CaseDeskWorker.bas` - BE-side scan / manifest loading
- `src/CaseDeskData.bas` - FE-side cache
- `src/frmCaseDesk.frm`, `src/frmSettings.frm` - UI
- `docs/spec.md` - detailed spec

## Dev commands

```bat
samplerun.bat
build-addin.bat
build-sample.bat
```

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Test-Compile.ps1
powershell -ExecutionPolicy Bypass -File scripts/Test-Worker.ps1
```

## Guardrails

- Runtime code is VBA only. Do not add WinAPI calls
- Do not break the FE/BE separation
- Do not change the `watchbox` manifest contract on only one side
- Assume hidden-sheet names and Workbook-event-based communication
