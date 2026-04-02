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

Private Function GetWorkerBookPath() As String
    ' Worker cannot open the xlam (locked/IsAddin). Save a temp xlsm copy.
    Dim cachePath As String: cachePath = GetCacheRoot()
    CaseDeskLib.EnsureFolder cachePath
    Dim dest As String: dest = cachePath & "\casedesk_worker.xlsm"
    On Error Resume Next
    ' Save a non-addin copy of ThisWorkbook as xlsm format
    Dim wasAddin As Boolean: wasAddin = ThisWorkbook.IsAddin
    ThisWorkbook.IsAddin = False
    ThisWorkbook.SaveCopyAs dest
    ThisWorkbook.IsAddin = wasAddin
    On Error GoTo 0
    GetWorkerBookPath = dest
End Function

Private Sub DebugLog(cachePath As String, msg As String)
    On Error Resume Next
    Dim f As Long: f = FreeFile
    Open cachePath & "\_debug.log" For Append As #f
    Print #f, Format$(Now, "hh:nn:ss") & " " & msg
    Close #f
    On Error GoTo 0
End Sub

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
    ElseIf shName = "_casedesk_signal" Then
        ' Clear stale signal from previous session to prevent premature data load
        ws.UsedRange.ClearContents
    End If
End Sub

' --- Worker Lifecycle ---

Public Sub StartWorker(mailFolder As String, caseRoot As String, _
                       matchField As String, matchMode As String)
    If Not g_workerApp Is Nothing Then Exit Sub
    If Len(mailFolder) = 0 And Len(caseRoot) = 0 Then Exit Sub

    CleanupZombieWorker

    Dim cachePath As String: cachePath = GetCacheRoot()
    CaseDeskLib.EnsureFolder cachePath

    ' Save worker xlsm (FE must do this — it owns ThisWorkbook)
    Dim workerBookPath As String: workerBookPath = GetWorkerBookPath()

    ' Snapshot current EXCEL PIDs before launching BE
    Dim beforePids As Object: Set beforePids = GetExcelPids()
    Dim pidCsv As String: pidCsv = ""
    Dim k As Variant
    For Each k In beforePids.keys
        If Len(pidCsv) > 0 Then pidCsv = pidCsv & ","
        pidCsv = pidCsv & CStr(k)
    Next k

    ' Generate launcher VBS — runs entirely in a separate process
    Dim vbsPath As String: vbsPath = cachePath & "\_launch.vbs"
    Dim pidPath As String: pidPath = cachePath & "\_worker.pid"
    Dim feBookPath As String: feBookPath = ThisWorkbook.FullName
    Dim f As Long: f = FreeFile
    Open vbsPath For Output As #f
    Print #f, "On Error Resume Next"
    Print #f, "Dim xl"
    Print #f, "Set xl = CreateObject(""Excel.Application"")"
    Print #f, "xl.Visible = False"
    Print #f, "xl.DisplayAlerts = False"
    Print #f, ""
    Print #f, "' Find BE PID (diff against FE snapshot)"
    Print #f, "Dim before, wmi, procs, p, pid"
    Print #f, "before = ""," & pidCsv & ","""
    Print #f, "pid = 0"
    Print #f, "Set wmi = GetObject(""winmgmts:\\.\root\cimv2"")"
    Print #f, "Set procs = wmi.ExecQuery(""SELECT ProcessId FROM Win32_Process WHERE Name = 'EXCEL.EXE'"")"
    Print #f, "For Each p In procs"
    Print #f, "  If InStr(before, "","" & CStr(p.ProcessId) & "","") = 0 Then"
    Print #f, "    pid = p.ProcessId"
    Print #f, "    Exit For"
    Print #f, "  End If"
    Print #f, "Next"
    Print #f, "Dim fso: Set fso = CreateObject(""Scripting.FileSystemObject"")"
    Print #f, "If pid > 0 Then"
    Print #f, "  Dim pf: Set pf = fso.CreateTextFile(""" & Q(pidPath) & """, True)"
    Print #f, "  pf.WriteLine CStr(pid)"
    Print #f, "  pf.Close"
    Print #f, "End If"
    Print #f, ""
    Print #f, "' Open worker book and start scan"
    Print #f, "xl.AutomationSecurity = 1"
    Print #f, "Dim wb: Set wb = xl.Workbooks.Open(""" & Q(workerBookPath) & """, 0, True)"
    Print #f, "xl.AutomationSecurity = 3"
    Print #f, "Dim feWb: Set feWb = GetObject(""" & Q(feBookPath) & """)"
    Print #f, "xl.Run ""CaseDeskWorker.WorkerEntryPoint"", " & _
              """" & Q(mailFolder) & """, " & _
              """" & Q(caseRoot) & """, " & _
              """" & Q(matchField) & """, " & _
              """" & Q(matchMode) & """, " & _
              "feWb, " & _
              """" & Q(cachePath) & """"
    Close #f

    DebugLog cachePath, "Launching worker via VBS"
    Shell "wscript.exe """ & vbsPath & """", vbHide
End Sub

Private Function Q(s As String) As String
    Q = Replace(s, """", """""")
End Function

Public Sub StopWorker()
    On Error Resume Next
    ' Try COM shutdown if we still have a reference
    If Not g_workerApp Is Nothing Then
        g_workerApp.Quit
        Set g_workerApp = Nothing
        Set g_workerWb = Nothing
    End If
    ' Always attempt PID-based cleanup (covers VBS-launched workers)
    CleanupZombieWorker
    On Error GoTo 0
End Sub

' --- PID Management ---

Private Function GetWorkerPidPath() As String
    GetWorkerPidPath = GetCacheRoot() & "\_worker.pid"
End Function


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
