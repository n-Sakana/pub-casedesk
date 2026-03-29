Attribute VB_Name = "CaseDeskWorker"
Option Explicit

' ============================================================================
' CaseDeskWorker - Background scanning module
' Runs in a separate Excel.Application instance (Visible=False).
' Scans mail/case folders, writes results directly to FE's hidden sheets.
' FE's Worksheet_Change fires on write (no polling needed).
' ============================================================================

Private g_active As Boolean
Private g_scheduled As Boolean
Private g_nextYieldAt As Date

Private g_mailFolder As String
Private g_caseRoot As String
Private g_signalVersion As Long
Private g_feWb As Object  ' Reference to FE's workbook (cross-process)

' Switch-style scan: round-robin task index + per-task resume position
Private Const TASK_MAIL As Long = 0
Private Const TASK_CASES As Long = 1
Private Const TASK_WRITE As Long = 2
Private Const TASK_COUNT As Long = 3
Private g_nextTask As Long           ' Round-robin: which task to start with
Private g_mailDirty As Boolean       ' mail scan found changes
Private g_casesDirty As Boolean      ' cases scan found changes

' ============================================================================
' BE-side cache (formerly CaseDeskScanner)
' These variables are only populated in the worker process.
' ============================================================================

Private m_fso As Object
Private m_mailRecords As Object     ' Dict: folder_path -> record
Private m_mailByEntryId As Object   ' Dict: entry_id -> record
Private m_mailIndex As Object       ' Dict: normalized_key -> Dict(entry_id -> record)
Private m_mailIndexField As String
Private m_mailIndexMode As String
Private m_mailAdded As Object
Private m_mailRemoved As Object
Private m_mailDiffReady As Boolean
Private m_manifestMod As Date       ' Last known manifest.csv mod time
Private m_caseNames As Object
Private m_caseAdded As Object
Private m_caseRemoved As Object
Private m_caseDiffReady As Boolean
Private m_caseManifestMod As Date    ' Last known case manifest.csv mod time
Private m_caseFiles As Object       ' Dict: item_id -> record dict
Private m_caseFilesDirty As Boolean
Private m_cachePath As String       ' Cache root (passed from FE)

Private Sub LogProfile(msg As String)
    On Error Resume Next
    If Len(m_cachePath) = 0 Then Exit Sub
    CaseDeskLib.EnsureFolder m_cachePath
    Dim f As Long: f = FreeFile
    Open m_cachePath & "\_profile.log" For Append As #f
    Print #f, Format$(Now, "hh:nn:ss") & " " & msg
    Close #f
    On Error GoTo 0
End Sub

Private Function GetFSO() As Object
    If m_fso Is Nothing Then Set m_fso = CreateObject("Scripting.FileSystemObject")
    Set GetFSO = m_fso
End Function

Public Sub ClearCache()
    Set m_mailRecords = Nothing
    Set m_mailByEntryId = Nothing
    Set m_mailIndex = Nothing
    m_mailIndexField = ""
    m_mailIndexMode = ""
    m_mailDiffReady = False
    m_manifestMod = #1/1/1900#
    Set m_caseNames = Nothing
    m_caseDiffReady = False
    m_caseManifestMod = #1/1/1900#
    Set m_caseFiles = Nothing
    m_caseFilesDirty = False
End Sub

' ============================================================================
' BE: Mail scanning
' ============================================================================

Public Function RefreshMailData(folderPath As String) As Boolean
    Dim eh As New ErrorHandler: eh.Enter "CaseDeskWorker", "RefreshMailData"
    On Error GoTo ErrHandler
    RefreshMailData = False

    Dim manifestPath As String: manifestPath = ResolveManifestPath(folderPath)
    If Len(manifestPath) = 0 Then eh.OK: Exit Function

    Dim curMod As Date: curMod = FileDateTime(manifestPath)
    If m_mailDiffReady And curMod = m_manifestMod Then eh.OK: Exit Function
    m_manifestMod = curMod
    LoadMailFromManifest manifestPath

    If Not m_mailDiffReady Then
        Set m_mailAdded = CaseDeskLib.NewDict()
        Set m_mailRemoved = CaseDeskLib.NewDict()
        m_mailDiffReady = True
    End If

    RefreshMailData = (m_mailAdded.Count > 0 Or m_mailRemoved.Count > 0)
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

' Fast path: read manifest.csv (11 columns, comma-separated, with header)
' Format: entry_id,sender_email,sender_name,subject,received_at,
'         folder_path,body_path,msg_path,attachment_paths,mail_folder,body_text
Private Sub LoadMailFromManifest(manifestPath As String)
    On Error Resume Next
    Dim t0 As Single: t0 = Timer
    Dim content As String: content = CaseDeskLib.ReadTextFile(manifestPath)
    If Len(content) = 0 Then Exit Sub

    Dim prevRecords As Object: Set prevRecords = m_mailRecords
    Set m_mailRecords = CaseDeskLib.NewDict()
    Set m_mailByEntryId = CaseDeskLib.NewDict()
    Set m_mailIndex = CaseDeskLib.NewDict()
    Set m_mailAdded = CaseDeskLib.NewDict()
    Set m_mailRemoved = CaseDeskLib.NewDict()

    Dim lines() As String: lines = Split(content, vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim line As String: line = lines(i)
        If Len(line) = 0 Then GoTo NextManifestLine
        ' Skip header row
        If Left$(line, 8) = "entry_id" Then GoTo NextManifestLine
        Dim cols() As String: cols = Split(line, ",")
        If UBound(cols) < 9 Then GoTo NextManifestLine
        Dim eid As String: eid = cols(0)
        If Len(eid) = 0 Then GoTo NextManifestLine

        Dim rec As Object: Set rec = CaseDeskLib.NewDict()
        rec.Add "entry_id", eid
        rec.Add "sender_email", cols(1)
        rec.Add "sender_name", cols(2)
        rec.Add "subject", cols(3)
        rec.Add "received_at", cols(4)
        rec.Add "folder_path", cols(5)
        rec.Add "body_path", cols(6)
        rec.Add "msg_path", cols(7)
        ' Parse attachment_paths (pipe-separated)
        Dim attDict As Object: Set attDict = CaseDeskLib.NewDict()
        If Len(cols(8)) > 0 Then
            Dim attParts() As String: attParts = Split(cols(8), "|")
            Dim a As Long
            For a = 0 To UBound(attParts)
                If Len(attParts(a)) > 0 Then
                    Dim fn As String: fn = Mid$(attParts(a), InStrRev(attParts(a), "\") + 1)
                    attDict.Add attParts(a), fn
                End If
            Next a
        End If
        rec.Add "attachment_paths", attDict
        rec.Add "_mail_folder", cols(9)
        If UBound(cols) >= 10 Then rec.Add "body_text", cols(10)

        Set m_mailRecords(cols(9)) = rec
        Set m_mailByEntryId(eid) = rec
        AddToMailIndex rec, eid

        ' Track added (new entries not in previous cache)
        If Not prevRecords Is Nothing Then
            If Not prevRecords.Exists(cols(9)) Then
                m_mailAdded(eid) = cols(3) & " - " & cols(1)
            End If
        End If
NextManifestLine:
    Next i

    ' Track removed (entries in previous cache not in new)
    If Not prevRecords Is Nothing Then
        If prevRecords.Count > 0 Then
            Dim pKeys As Variant: pKeys = prevRecords.keys
            For i = 0 To UBound(pKeys)
                If Not m_mailRecords.Exists(pKeys(i)) Then
                    Dim remRec As Object: Set remRec = prevRecords(pKeys(i))
                    Dim remEid As String: remEid = CaseDeskLib.DictStr(remRec, "entry_id")
                    If Len(remEid) > 0 Then
                        m_mailRemoved(remEid) = CaseDeskLib.DictStr(remRec, "subject") & _
                            " - " & CaseDeskLib.DictStr(remRec, "sender_email")
                    End If
                End If
            Next i
        End If
    End If

    LogProfile "LoadMailFromManifest: " & Format$(Timer - t0, "0.000") & "s (" & m_mailRecords.Count & " records)"
    On Error GoTo 0
End Sub

Public Sub SetMailMatchConfig(field As String, mode As String)
    If field = m_mailIndexField And mode = m_mailIndexMode Then Exit Sub
    m_mailIndexField = field
    m_mailIndexMode = mode
    RebuildMailIndex
End Sub

Private Sub RebuildMailIndex()
    Set m_mailIndex = CaseDeskLib.NewDict()
    If m_mailRecords Is Nothing Then Exit Sub
    If Len(m_mailIndexField) = 0 Then Exit Sub
    If m_mailRecords.Count = 0 Then Exit Sub
    Dim items As Variant: items = m_mailRecords.Items
    Dim i As Long
    For i = 0 To UBound(items)
        Dim rec As Object: Set rec = items(i)
        Dim eid As String: eid = CaseDeskLib.DictStr(rec, "entry_id")
        If Len(eid) > 0 Then AddToMailIndex rec, eid
    Next i
End Sub

Private Sub AddToMailIndex(rec As Object, entryId As String)
    If m_mailIndex Is Nothing Then Exit Sub
    If Len(m_mailIndexField) = 0 Then Exit Sub
    If Not rec.Exists(m_mailIndexField) Then Exit Sub
    If IsNull(rec(m_mailIndexField)) Then Exit Sub
    Dim fv As String: fv = CStr(rec(m_mailIndexField))
    If Len(fv) = 0 Then Exit Sub
    Dim key As String
    If m_mailIndexMode = "domain" Then
        key = LCase$(GetDomain(fv))
    Else
        key = LCase$(fv)
    End If
    If Not m_mailIndex.Exists(key) Then m_mailIndex.Add key, CaseDeskLib.NewDict()
    Dim inner As Object: Set inner = m_mailIndex(key)
    Set inner(entryId) = rec
End Sub

Private Sub RemoveFromMailIndex(rec As Object, entryId As String)
    If m_mailIndex Is Nothing Then Exit Sub
    If Len(m_mailIndexField) = 0 Then Exit Sub
    If Not rec.Exists(m_mailIndexField) Then Exit Sub
    If IsNull(rec(m_mailIndexField)) Then Exit Sub
    Dim fv As String: fv = CStr(rec(m_mailIndexField))
    If Len(fv) = 0 Then Exit Sub
    Dim key As String
    If m_mailIndexMode = "domain" Then
        key = LCase$(GetDomain(fv))
    Else
        key = LCase$(fv)
    End If
    If m_mailIndex.Exists(key) Then
        Dim inner As Object: Set inner = m_mailIndex(key)
        If inner.Exists(entryId) Then inner.Remove entryId
        If inner.Count = 0 Then m_mailIndex.Remove key
    End If
End Sub

Private Function ResolveManifestPath(folderPath As String) As String
    ResolveManifestPath = ""
    Dim p As String
    p = folderPath & "\.manifest.csv"
    If Len(Dir$(p)) > 0 Then ResolveManifestPath = p: Exit Function
    p = folderPath & "\manifest.csv"
    If Len(Dir$(p)) > 0 Then ResolveManifestPath = p: Exit Function
End Function

Private Function GetDomain(email As String) As String
    Dim pos As Long: pos = InStr(email, "@")
    If pos > 0 Then GetDomain = Mid$(email, pos + 1) Else GetDomain = email
End Function

Public Function GetMailRecords() As Object
    If m_mailRecords Is Nothing Then Set m_mailRecords = CaseDeskLib.NewDict()
    Set GetMailRecords = m_mailRecords
End Function

Public Function GetCaseNames() As Object
    If m_caseNames Is Nothing Then Set m_caseNames = CaseDeskLib.NewDict()
    Set GetCaseNames = m_caseNames
End Function

Public Function GetMailByEntryId() As Object
    If m_mailByEntryId Is Nothing Then Set m_mailByEntryId = CaseDeskLib.NewDict()
    Set GetMailByEntryId = m_mailByEntryId
End Function

Public Function GetMailIndex() As Object
    If m_mailIndex Is Nothing Then Set m_mailIndex = CaseDeskLib.NewDict()
    Set GetMailIndex = m_mailIndex
End Function

Public Function GetMailAdded() As Object
    If m_mailAdded Is Nothing Then Set m_mailAdded = CaseDeskLib.NewDict()
    Set GetMailAdded = m_mailAdded
End Function

Public Function GetMailRemoved() As Object
    If m_mailRemoved Is Nothing Then Set m_mailRemoved = CaseDeskLib.NewDict()
    Set GetMailRemoved = m_mailRemoved
End Function

' ============================================================================
' BE: Case data from manifest.csv (generated by watchbox)
' Format: item_id,file_name,file_path,folder_path,relative_path,file_size,modified_at
' ============================================================================

Public Function RefreshCaseData(rootPath As String) As Boolean
    Dim eh As New ErrorHandler: eh.Enter "CaseDeskWorker", "RefreshCaseData"
    On Error GoTo ErrHandler
    RefreshCaseData = False

    Dim manifestPath As String: manifestPath = ResolveManifestPath(rootPath)
    If Len(manifestPath) = 0 Then eh.OK: Exit Function

    Dim curMod As Date: curMod = FileDateTime(manifestPath)
    If m_caseDiffReady And curMod = m_caseManifestMod Then eh.OK: Exit Function
    m_caseManifestMod = curMod

    ' Read manifest.csv
    Dim content As String: content = CaseDeskLib.ReadTextFile(manifestPath)
    If Len(content) = 0 Then eh.OK: Exit Function

    Dim prevNames As Object: Set prevNames = m_caseNames
    Set m_caseNames = CaseDeskLib.NewDict()
    Set m_caseFiles = CaseDeskLib.NewDict()
    Set m_caseAdded = CaseDeskLib.NewDict()
    Set m_caseRemoved = CaseDeskLib.NewDict()

    Dim lines() As String: lines = Split(content, vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim line As String: line = lines(i)
        If Len(line) = 0 Then GoTo NextCaseLine
        If Left$(line, 7) = "item_id" Then GoTo NextCaseLine
        Dim cols() As String: cols = Split(line, ",")
        If UBound(cols) < 6 Then GoTo NextCaseLine

        Dim itemId As String: itemId = cols(0)
        If Len(itemId) = 0 Then GoTo NextCaseLine

        ' Extract case name from relative_path (first path segment)
        Dim relPath As String: relPath = cols(4)
        Dim caseName As String
        Dim sepPos As Long: sepPos = InStr(relPath, "\")
        If sepPos > 0 Then
            caseName = Left$(relPath, sepPos - 1)
        Else
            caseName = ""
        End If

        ' Track case names
        If Len(caseName) > 0 And Not m_caseNames.Exists(caseName) Then
            m_caseNames(caseName) = True
        End If

        ' Store file record
        Dim rec As Object: Set rec = CaseDeskLib.NewDict()
        rec.Add "item_id", itemId
        rec.Add "file_name", cols(1)
        rec.Add "file_path", cols(2)
        rec.Add "folder_path", cols(3)
        rec.Add "relative_path", relPath
        rec.Add "file_size", cols(5)
        rec.Add "modified_at", cols(6)
        rec.Add "case_id", caseName
        Set m_caseFiles(itemId) = rec
NextCaseLine:
    Next i

    ' Compute case name diffs (skip on first run to avoid false positives)
    If m_caseDiffReady And Not prevNames Is Nothing Then
        Dim k As Variant
        For Each k In m_caseNames.keys
            If Not prevNames.Exists(k) Then m_caseAdded(CStr(k)) = True
        Next k
        For Each k In prevNames.keys
            If Not m_caseNames.Exists(k) Then m_caseRemoved(CStr(k)) = True
        Next k
    End If
    m_caseDiffReady = True

    m_caseFilesDirty = True
    RefreshCaseData = True
    LogProfile "RefreshCaseData: " & m_caseFiles.Count & " files, " & m_caseNames.Count & " cases"
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Public Function GetCaseAdded() As Object
    If m_caseAdded Is Nothing Then Set m_caseAdded = CaseDeskLib.NewDict()
    Set GetCaseAdded = m_caseAdded
End Function

Public Function GetCaseRemoved() As Object
    If m_caseRemoved Is Nothing Then Set m_caseRemoved = CaseDeskLib.NewDict()
    Set GetCaseRemoved = m_caseRemoved
End Function

' Clear diff dictionaries after they have been written to FE sheets
' Prevents stale diffs from being re-written when only one scan type triggers a signal bump
Public Sub ClearDiffs()
    Set m_mailAdded = CaseDeskLib.NewDict()
    Set m_mailRemoved = CaseDeskLib.NewDict()
    Set m_caseAdded = CaseDeskLib.NewDict()
    Set m_caseRemoved = CaseDeskLib.NewDict()
End Sub

' ============================================================================
' Entry Point
' ============================================================================

Public Sub WorkerEntryPoint(mailFolder As String, caseRoot As String, _
                            matchField As String, matchMode As String, _
                            feWorkbook As Object, cachePath As String)
    Dim eh As New ErrorHandler: eh.Enter "CaseDeskWorker", "EntryPoint"
    On Error GoTo ErrHandler

    g_mailFolder = mailFolder
    g_caseRoot = caseRoot
    Set g_feWb = feWorkbook
    m_cachePath = cachePath
    SetMailMatchConfig matchField, matchMode
    Application.EnableEvents = True

    g_signalVersion = 0
    g_active = True
    Application.OnTime Now, "CaseDeskWorker.WorkerInitialScan"

    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Initial full scan
' ============================================================================

Public Sub WorkerInitialScan()
    If Not g_active Then Exit Sub
    On Error Resume Next

    Dim t0 As Single: t0 = Timer
    If Len(g_mailFolder) > 0 Then RefreshMailData g_mailFolder
    Dim t1 As Single: t1 = Timer
    If Len(g_caseRoot) > 0 Then RefreshCaseData g_caseRoot
    Dim t2 As Single: t2 = Timer

    ' Write all data to FE sheets
    Dim tw0 As Single: tw0 = Timer
    WriteMailToFE
    Dim tw1 As Single: tw1 = Timer
    WriteMailIndexToFE
    Dim tw2 As Single: tw2 = Timer
    WriteCasesToFE
    WriteCaseFilesToFE
    Dim tw3 As Single: tw3 = Timer
    WriteDiffToFE
    ClearDiffs
    g_signalVersion = 1
    Dim tw4 As Single: tw4 = Timer
    WriteSignalToFE g_signalVersion, "scan mail=" & Format$(t1 - t0, "0.000") & _
        " case=" & Format$(t2 - t1, "0.000") & _
        " | write mail=" & Format$(tw1 - tw0, "0.000") & _
        " idx=" & Format$(tw2 - tw1, "0.000") & _
        " cases=" & Format$(tw3 - tw2, "0.000") & _
        " diff=" & Format$(tw4 - tw3, "0.000") & _
        " total=" & Format$(tw4 - tw0, "0.000")

    ' Start switch-style scan loop
    g_nextTask = TASK_MAIL
    ScheduleNextChunk
    On Error GoTo 0
End Sub

' ============================================================================
' Switch-style Scan Loop (1s chunk + 1s yield, continuous)
' ============================================================================

' DoScanChunk: process tasks within 1-second time budget, round-robin
Public Sub DoScanChunk()
    g_scheduled = False
    If Not g_active Then Exit Sub
    On Error Resume Next

    Dim t0 As Single: t0 = Timer
    Dim startTask As Long: startTask = g_nextTask

    Do
        Select Case g_nextTask
            Case TASK_MAIL
                If Len(g_mailFolder) > 0 Then
                    If RefreshMailData(g_mailFolder) Then g_mailDirty = True
                End If
            Case TASK_CASES
                If Len(g_caseRoot) > 0 Then
                    If RefreshCaseData(g_caseRoot) Then g_casesDirty = True
                End If
            Case TASK_WRITE
                If g_mailDirty Or g_casesDirty Or m_caseFilesDirty Then
                    g_signalVersion = g_signalVersion + 1
                    If g_mailDirty Then WriteMailToFE: WriteMailIndexToFE
                    If g_casesDirty Then WriteCasesToFE
                    If m_caseFilesDirty Then WriteCaseFilesToFE
                    WriteDiffToFE
                    ClearDiffs
                    WriteVersionToFE g_signalVersion
                    g_mailDirty = False
                    g_casesDirty = False
                End If
        End Select
        g_nextTask = (g_nextTask + 1) Mod TASK_COUNT
        If Timer - t0 >= 1 Then Exit Do
    Loop Until g_nextTask = startTask

    ' Schedule Yield (returns control to message loop, then next chunk)
    If g_active Then
        g_nextYieldAt = Now
        Application.OnTime g_nextYieldAt, "CaseDeskWorker.YieldCallback"
        g_scheduled = True
    End If
    On Error GoTo 0
End Sub

' YieldCallback: yield then schedule next chunk
Public Sub YieldCallback()
    g_scheduled = False
    If Not g_active Then Exit Sub
    On Error Resume Next
    If g_active Then ScheduleNextChunk
    On Error GoTo 0
End Sub

Private Sub ScheduleNextChunk()
    If g_scheduled Then Exit Sub
    On Error Resume Next
    Dim nextAt As Date: nextAt = Now + TimeSerial(0, 0, 1)
    Application.OnTime nextAt, "CaseDeskWorker.DoScanChunk"
    g_scheduled = True
    If Err.Number <> 0 Then g_scheduled = False: Err.Clear
    On Error GoTo 0
End Sub

' ============================================================================
' Config Update
' ============================================================================

Public Sub UpdateConfig(mailFolder As String, caseRoot As String, _
                        matchField As String, matchMode As String)
    Dim eh As New ErrorHandler: eh.Enter "CaseDeskWorker", "UpdateConfig"
    On Error GoTo ErrHandler

    g_mailFolder = mailFolder
    g_caseRoot = caseRoot
    ClearCache
    SetMailMatchConfig matchField, matchMode

    ' Force immediate full scan on next chunk
    g_mailDirty = False
    g_casesDirty = False
    m_mailDiffReady = False
    m_caseDiffReady = False
    g_nextTask = TASK_MAIL

    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' FE->BE Request Dispatcher (called via Workbook_SheetChange -> OnTime)
' ============================================================================



' ============================================================================
' Stop
' ============================================================================

Public Sub WorkerStop()
    g_active = False
    On Error Resume Next
    g_scheduled = False
    Set g_feWb = Nothing
    On Error GoTo 0
End Sub

' ============================================================================
' FE Sheet Writers (.Value=.Value to FE's workbook)
' ============================================================================

Private Function FESheet(shName As String) As Object
    If g_feWb Is Nothing Then Exit Function
    On Error Resume Next
    Set FESheet = g_feWb.Worksheets(shName)
    On Error GoTo 0
End Function

Private Sub WriteMailToFE()
    Dim ws As Object: Set ws = FESheet("_casedesk_mail")
    If ws Is Nothing Then Exit Sub
    ws.UsedRange.ClearContents

    Dim records As Object: Set records = GetMailRecords()
    If records Is Nothing Then Exit Sub
    If records.Count = 0 Then Exit Sub

    Dim keys As Variant: keys = records.keys
    Dim n As Long: n = UBound(keys) + 1
    Dim data() As Variant: ReDim data(1 To n, 1 To 11)
    Dim i As Long
    For i = 0 To UBound(keys)
        Dim rec As Object: Set rec = records(keys(i))
        data(i + 1, 1) = CaseDeskLib.DictStr(rec, "entry_id")
        data(i + 1, 2) = CaseDeskLib.DictStr(rec, "sender_email")
        data(i + 1, 3) = CaseDeskLib.DictStr(rec, "sender_name")
        data(i + 1, 4) = CaseDeskLib.DictStr(rec, "subject")
        data(i + 1, 5) = CaseDeskLib.DictStr(rec, "received_at")
        data(i + 1, 6) = CaseDeskLib.DictStr(rec, "folder_path")
        data(i + 1, 7) = CaseDeskLib.DictStr(rec, "body_path")
        data(i + 1, 8) = CaseDeskLib.DictStr(rec, "msg_path")
        Dim attStr As String: attStr = ""
        If rec.Exists("attachment_paths") Then
            If IsObject(rec("attachment_paths")) Then
                Dim attDict As Object: Set attDict = rec("attachment_paths")
                If attDict.Count > 0 Then
                    Dim attKeys As Variant: attKeys = attDict.keys
                    Dim attParts() As String: ReDim attParts(0 To UBound(attKeys))
                    Dim a As Long
                    For a = 0 To UBound(attKeys): attParts(a) = CStr(attKeys(a)): Next a
                    attStr = Join(attParts, "|")
                End If
            End If
        End If
        data(i + 1, 9) = attStr
        data(i + 1, 10) = CaseDeskLib.DictStr(rec, "_mail_folder")
        data(i + 1, 11) = CaseDeskLib.DictStr(rec, "body_text")
    Next i
    ws.Range("A1").Resize(n, 11).Value = data
End Sub

Private Sub WriteMailIndexToFE()
    Dim ws As Object: Set ws = FESheet("_casedesk_mail_idx")
    If ws Is Nothing Then Exit Sub
    ws.UsedRange.ClearContents

    Dim idx As Object: Set idx = GetMailIndex()
    If idx Is Nothing Then Exit Sub
    If idx.Count = 0 Then Exit Sub

    Dim outerKeys As Variant: outerKeys = idx.keys
    Dim total As Long: total = 0
    Dim i As Long, j As Long
    For i = 0 To UBound(outerKeys): total = total + idx(outerKeys(i)).Count: Next i
    If total = 0 Then Exit Sub

    Dim data() As Variant: ReDim data(1 To total, 1 To 2)
    Dim n As Long: n = 0
    For i = 0 To UBound(outerKeys)
        Dim key As String: key = CStr(outerKeys(i))
        Dim inner As Object: Set inner = idx(outerKeys(i))
        Dim innerKeys As Variant: innerKeys = inner.keys
        For j = 0 To UBound(innerKeys)
            n = n + 1
            data(n, 1) = key
            data(n, 2) = CStr(innerKeys(j))
        Next j
    Next i
    ws.Range("A1").Resize(n, 2).Value = data
End Sub

Private Sub WriteCasesToFE()
    Dim ws As Object: Set ws = FESheet("_casedesk_cases")
    If ws Is Nothing Then Exit Sub
    ws.UsedRange.ClearContents

    Dim names As Object: Set names = GetCaseNames()
    If names Is Nothing Then Exit Sub
    If names.Count = 0 Then Exit Sub

    Dim keys As Variant: keys = names.keys
    Dim n As Long: n = UBound(keys) + 1
    Dim data() As Variant: ReDim data(1 To n, 1 To 1)
    Dim i As Long
    For i = 0 To UBound(keys): data(i + 1, 1) = CStr(keys(i)): Next i
    ws.Range("A1").Resize(n, 1).Value = data
End Sub


Private Sub WriteCaseFilesToFE()
    On Error Resume Next
    Dim ws As Object: Set ws = FESheet("_casedesk_files")
    If ws Is Nothing Then Exit Sub
    ws.UsedRange.ClearContents

    If m_caseFiles Is Nothing Then m_caseFilesDirty = False: Exit Sub
    If m_caseFiles.Count = 0 Then m_caseFilesDirty = False: Exit Sub

    ' Columns: case_id, file_name, file_path, folder_path, relative_path, file_size, modified_at
    Dim keys As Variant: keys = m_caseFiles.keys
    Dim n As Long: n = m_caseFiles.Count
    Dim data() As Variant: ReDim data(1 To n, 1 To 7)
    Dim i As Long
    For i = 0 To UBound(keys)
        Dim rec As Object: Set rec = m_caseFiles(keys(i))
        data(i + 1, 1) = CaseDeskLib.DictStr(rec, "case_id")
        data(i + 1, 2) = CaseDeskLib.DictStr(rec, "file_name")
        data(i + 1, 3) = CaseDeskLib.DictStr(rec, "file_path")
        data(i + 1, 4) = CaseDeskLib.DictStr(rec, "folder_path")
        data(i + 1, 5) = CaseDeskLib.DictStr(rec, "relative_path")
        data(i + 1, 6) = CaseDeskLib.DictStr(rec, "file_size")
        data(i + 1, 7) = CaseDeskLib.DictStr(rec, "modified_at")
    Next i
    ws.Range("A1").Resize(n, 7).Value = data
    m_caseFilesDirty = False
    On Error GoTo 0
End Sub

Private Sub WriteDiffToFE()
    Dim ws As Object: Set ws = FESheet("_casedesk_diff")
    If ws Is Nothing Then Exit Sub
    ws.UsedRange.ClearContents

    Dim ma As Object: Set ma = GetMailAdded()
    Dim mr As Object: Set mr = GetMailRemoved()
    Dim ca As Object: Set ca = GetCaseAdded()
    Dim cr As Object: Set cr = GetCaseRemoved()
    Dim total As Long: total = ma.Count + mr.Count + ca.Count + cr.Count
    If total = 0 Then Exit Sub

    Dim data() As Variant: ReDim data(1 To total, 1 To 4)
    Dim n As Long: n = 0
    Dim i As Long

    If ma.Count > 0 Then
        Dim mak As Variant: mak = ma.keys
        For i = 0 To UBound(mak): n = n + 1
            data(n, 1) = "added": data(n, 2) = "mail"
            data(n, 3) = CStr(mak(i)): data(n, 4) = CStr(ma(mak(i)))
        Next i
    End If
    If mr.Count > 0 Then
        Dim mrk As Variant: mrk = mr.keys
        For i = 0 To UBound(mrk): n = n + 1
            data(n, 1) = "removed": data(n, 2) = "mail"
            data(n, 3) = CStr(mrk(i)): data(n, 4) = CStr(mr(mrk(i)))
        Next i
    End If
    If ca.Count > 0 Then
        Dim cak As Variant: cak = ca.keys
        For i = 0 To UBound(cak): n = n + 1
            data(n, 1) = "added": data(n, 2) = "case"
            data(n, 3) = CStr(cak(i)): data(n, 4) = CStr(cak(i))
        Next i
    End If
    If cr.Count > 0 Then
        Dim crk As Variant: crk = cr.keys
        For i = 0 To UBound(crk): n = n + 1
            data(n, 1) = "removed": data(n, 2) = "case"
            data(n, 3) = CStr(crk(i)): data(n, 4) = CStr(crk(i))
        Next i
    End If
    ws.Range("A1").Resize(n, 4).Value = data
End Sub

' ============================================================================
' Signal/Clock writes to FE
' ============================================================================


Private Sub WriteVersionToFE(ver As Long)
    Dim ws As Object: Set ws = FESheet("_casedesk_signal")
    If ws Is Nothing Then Exit Sub
    ws.Range("B1").Value = ver
End Sub

Private Sub WriteSignalToFE(ver As Long, timing As String)
    Dim ws As Object: Set ws = FESheet("_casedesk_signal")
    If ws Is Nothing Then Exit Sub
    ws.Range("A1").Value2 = Format$(Now, "hh:nn:ss") & " "
    ws.Range("B1").Value = ver
    ws.Range("C1").Value = timing
End Sub

