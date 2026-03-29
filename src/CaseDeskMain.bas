Attribute VB_Name = "CaseDeskMain"
Option Explicit

Public g_forceClose As Boolean
Public g_formLoaded As Boolean
Public g_workerApp As Object
Public g_workerWb As Object
Public g_appHandler As AppEventHandler
Public g_dataWb As Workbook  ' Target workbook (captured at launch)

' --- Addin Lifecycle ---

Public Sub InitAddin()
    If Not g_appHandler Is Nothing Then Exit Sub
    Set g_appHandler = New AppEventHandler
    Set g_appHandler.App = Application
End Sub

Public Sub ShutdownAddin()
    BeforeWorkbookClose
    Set g_appHandler = Nothing
End Sub

' --- Ribbon Callbacks (customUI requires control argument) ---

Public Sub Ribbon_ShowPanel(control As IRibbonControl)
    CaseDesk_ShowPanel
End Sub

Public Sub Ribbon_ShowSettings(control As IRibbonControl)
    CaseDesk_ShowSettings
End Sub

' --- Entry Points ---

Public Sub CaseDesk_ShowPanel()
    Dim eh As New ErrorHandler: eh.Enter "CaseDeskMain", "ShowPanel"
    On Error GoTo ErrHandler

    ' Capture ActiveWorkbook at launch (skip the xlam itself)
    If ActiveWorkbook Is Nothing Then
        MsgBox "No workbook is open.", vbExclamation, "CaseDesk"
        Exit Sub
    End If
    If ActiveWorkbook.FullName = ThisWorkbook.FullName Then
        MsgBox "Please activate a data workbook first.", vbExclamation, "CaseDesk"
        Exit Sub
    End If
    Set g_dataWb = ActiveWorkbook

    CaseDeskLib.EnsureConfigSheets
    CaseDeskLib.EnsureLogSheet
    EnsureCaseDeskSheets
    g_forceClose = False
    g_formLoaded = True
    frmCaseDesk.Show vbModeless
    eh.OK
    Exit Sub
ErrHandler:
    eh.Catch
End Sub

Public Sub CaseDesk_ShowSettings()
    Dim eh As New ErrorHandler: eh.Enter "CaseDeskMain", "ShowSettings"
    On Error GoTo ErrHandler
    frmSettings.Show vbModal
    eh.OK
    Exit Sub
ErrHandler:
    eh.Catch
End Sub

' --- Deferred Startup ---

Public Sub DeferredStartup()
    On Error Resume Next
    If g_formLoaded Then frmCaseDesk.DoPollCycle
    On Error GoTo 0
End Sub

' --- Workbook Close ---

Public Sub BeforeWorkbookClose()
    g_forceClose = True
    g_formLoaded = False
    StopWorker
    CaseDeskLib.SaveToSheets
    Set g_dataWb = Nothing
End Sub

' --- Cache Path ---

Private Function GetCacheRoot() As String
    GetCacheRoot = Environ$("LOCALAPPDATA") & "\CaseDesk"
End Function

' --- FE Data Sheets ---

Private Sub EnsureCaseDeskSheets()
    Dim wb As Workbook: Set wb = ThisWorkbook
    EnsureHiddenSheet wb, "_casedesk_signal"
    EnsureHiddenSheet wb, "_casedesk_mail"
    EnsureHiddenSheet wb, "_casedesk_mail_idx"
    EnsureHiddenSheet wb, "_casedesk_cases"
    EnsureHiddenSheet wb, "_casedesk_files"
    EnsureHiddenSheet wb, "_casedesk_diff"
End Sub

Private Sub EnsureHiddenSheet(wb As Workbook, shName As String)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = wb.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = shName
        ws.Visible = xlSheetVeryHidden
    End If
End Sub

' --- Worker Lifecycle ---

Public Sub StartWorker(mailFolder As String, caseRoot As String, _
                       matchField As String, matchMode As String)
    Dim eh As New ErrorHandler: eh.Enter "CaseDeskMain", "StartWorker"
    On Error GoTo ErrHandler

    If Not g_workerApp Is Nothing Then eh.OK: Exit Sub
    If Len(mailFolder) = 0 And Len(caseRoot) = 0 Then eh.OK: Exit Sub

    CleanupZombieWorker

    Dim beforePids As Object: Set beforePids = GetExcelPids()
    Set g_workerApp = CreateObject("Excel.Application")
    g_workerApp.Visible = False
    g_workerApp.DisplayAlerts = False

    Dim prevSec As Long: prevSec = g_workerApp.AutomationSecurity
    g_workerApp.AutomationSecurity = 1
    g_workerApp.Workbooks.Open ThisWorkbook.FullName, ReadOnly:=True, UpdateLinks:=0
    g_workerApp.AutomationSecurity = prevSec
    Set g_workerWb = g_workerApp.Workbooks(g_workerApp.Workbooks.Count)

    Dim cachePath As String: cachePath = GetCacheRoot()
    CaseDeskLib.EnsureFolder cachePath
    g_workerApp.Run "CaseDeskWorker.WorkerEntryPoint", mailFolder, caseRoot, matchField, matchMode, ThisWorkbook, cachePath

    WriteWorkerPid beforePids

    eh.OK: Exit Sub
ErrHandler:
    eh.Catch
    On Error Resume Next
    If Not g_workerApp Is Nothing Then g_workerApp.Quit
    Set g_workerApp = Nothing
    On Error GoTo 0
End Sub

Public Sub StopWorker()
    If g_workerApp Is Nothing Then Exit Sub
    On Error Resume Next
    Set g_workerWb = Nothing
    g_workerApp.Quit
    Set g_workerApp = Nothing
    Dim pidPath As String: pidPath = GetWorkerPidPath()
    If Len(Dir$(pidPath)) > 0 Then Kill pidPath
    On Error GoTo 0
End Sub

' --- PID Management ---

Private Function GetWorkerPidPath() As String
    GetWorkerPidPath = GetCacheRoot() & "\_worker.pid"
End Function

Private Sub WriteWorkerPid(beforePids As Object)
    On Error Resume Next
    If g_workerApp Is Nothing Then Exit Sub
    Dim afterPids As Object: Set afterPids = GetExcelPids()
    Dim pid As Long: pid = 0
    Dim k As Variant
    For Each k In afterPids.keys
        If Not beforePids.Exists(k) Then pid = CLng(k): Exit For
    Next k
    If pid = 0 Then Exit Sub
    CaseDeskLib.EnsureFolder GetCacheRoot()
    Dim pidPath As String: pidPath = GetWorkerPidPath()
    Dim f As Long: f = FreeFile
    Open pidPath For Output As #f
    Print #f, CStr(pid)
    Close #f
    On Error GoTo 0
End Sub

Private Function GetExcelPids() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Dim wmi As Object: Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Dim proc As Object
    For Each proc In wmi.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = 'EXCEL.EXE'")
        d(CStr(proc.ProcessId)) = True
    Next proc
    On Error GoTo 0
    Set GetExcelPids = d
End Function

Private Sub CleanupZombieWorker()
    On Error Resume Next
    Dim pidPath As String: pidPath = GetWorkerPidPath()
    If Len(Dir$(pidPath)) = 0 Then Exit Sub
    Dim f As Long: f = FreeFile
    Dim pidStr As String
    Open pidPath For Input As #f
    Line Input #f, pidStr
    Close #f
    If Len(pidStr) > 0 And IsNumeric(Trim$(pidStr)) Then
        Shell "cmd /c taskkill /F /PID " & Trim$(pidStr) & " >nul 2>&1", vbHide
    End If
    Kill pidPath
    On Error GoTo 0
End Sub
