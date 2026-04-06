Attribute VB_Name = "CaseDeskData"
Option Explicit

' ============================================================================
' FE-side cache (populated from hidden sheets written by CaseDeskWorker)
' FE detects changes via Workbook_SheetChange on _casedesk_signal.
' ============================================================================

Private m_feMailRecords As Object    ' Dict: entry_id -> record Dict
Private m_feMailIndex As Object      ' Dict: normalized_key -> Dict(entry_id -> True)
Private m_feCaseNames As Object      ' Dict: folder_name -> True
Private m_feCaseFiles As Object      ' Dict: case_id -> Dict(file_path -> record Dict)

Private Function SafeStr(v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then SafeStr = "" Else SafeStr = CStr(v)
End Function

' ============================================================================
' Table Operations (FE: reads/writes the source Excel file directly)
' ============================================================================

Public Function GetWorkbookTableNames(wb As Workbook) As Collection
    Dim eh As New ErrorHandler: eh.Enter "CaseDeskData", "GetWorkbookTableNames"
    On Error GoTo ErrHandler
    Set GetWorkbookTableNames = New Collection
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible <> xlSheetVisible Then GoTo NextSheet
        Dim tbl As ListObject
        For Each tbl In ws.ListObjects
            GetWorkbookTableNames.Add tbl.Name
        Next tbl
NextSheet:
    Next ws
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Public Function FindTable(wb As Workbook, tableName As String) As ListObject
    Dim eh As New ErrorHandler: eh.Enter "CaseDeskData", "FindTable"
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim tbl As ListObject
        For Each tbl In ws.ListObjects
            If LCase$(tbl.Name) = LCase$(tableName) Then
                Set FindTable = tbl: eh.OK: Exit Function
            End If
        Next tbl
    Next ws
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Public Function ReadTableRecords(tbl As ListObject) As Object
    Dim eh As New ErrorHandler: eh.Enter "CaseDeskData", "ReadTableRecords"
    On Error GoTo ErrHandler
    Set ReadTableRecords = CaseDeskLib.NewDict()
    If tbl.DataBodyRange Is Nothing Then eh.OK: Exit Function
    Dim data As Variant: data = tbl.DataBodyRange.Value
    ' Handle single-row table (Value returns scalar or 1D array)
    If Not IsArray(data) Then
        Dim tmp As Variant
        ReDim tmp(1 To 1, 1 To tbl.ListColumns.Count)
        Dim c2 As Long
        For c2 = 1 To tbl.ListColumns.Count
            tmp(1, c2) = tbl.DataBodyRange.Cells(1, c2).Value
        Next c2
        data = tmp
    End If
    Dim nCols As Long: nCols = tbl.ListColumns.Count
    Dim colNames() As String: ReDim colNames(1 To nCols)
    Dim seenCols As Object: Set seenCols = CaseDeskLib.NewDict()
    Dim c As Long
    For c = 1 To nCols
        Dim cn As String: cn = tbl.ListColumns(c).Name
        ' Handle duplicate column names by appending suffix
        If seenCols.Exists(cn) Then
            Dim suffix As Long: suffix = 2
            Do While seenCols.Exists(cn & "_" & suffix): suffix = suffix + 1: Loop
            cn = cn & "_" & suffix
        End If
        seenCols(cn) = True
        colNames(c) = cn
    Next c
    Dim r As Long
    For r = 1 To UBound(data, 1)
        Dim rec As Object: Set rec = CaseDeskLib.NewDict()
        rec.Add "_row_index", r
        For c = 1 To nCols
            rec.Add colNames(c), data(r, c)
        Next c
        ReadTableRecords.Add CStr(r), rec
    Next r
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Public Sub WriteTableCell(tbl As ListObject, rowIndex As Long, colName As String, val As Variant)
    Dim eh As New ErrorHandler: eh.Enter "CaseDeskData", "WriteTableCell"
    On Error GoTo ErrHandler
    Dim col As ListColumn: Set col = tbl.ListColumns(colName)
    tbl.DataBodyRange.Cells(rowIndex, col.Index).Value = val
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Public Function GetTableColumnNames(tbl As ListObject) As Collection
    Dim eh As New ErrorHandler: eh.Enter "CaseDeskData", "GetTableColumnNames"
    On Error GoTo ErrHandler
    Set GetTableColumnNames = New Collection
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        GetTableColumnNames.Add col.Name
    Next col
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Public Function GetUsedRangeColumnNames(ws As Worksheet) As Collection
    Dim eh As New ErrorHandler: eh.Enter "CaseDeskData", "GetUsedRangeColumnNames"
    On Error GoTo ErrHandler
    Set GetUsedRangeColumnNames = New Collection
    If ws.UsedRange Is Nothing Then eh.OK: Exit Function
    Dim ur As Range: Set ur = ws.UsedRange
    Dim headerRow As Long: headerRow = ur.Row
    Dim startCol As Long: startCol = ur.Column
    Dim nCols As Long: nCols = ur.Columns.Count
    Dim c As Long
    For c = 0 To nCols - 1
        Dim v As Variant: v = ws.Cells(headerRow, startCol + c).Value
        If Not IsEmpty(v) And Len(CStr(v)) > 0 Then
            GetUsedRangeColumnNames.Add CStr(v)
        End If
    Next c
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

' ============================================================================
' FE: Mail/Case counts — read from FE-side Dictionary cache
' ============================================================================

Public Function GetMailCount() As Long
    GetMailCount = 0
    If Not m_feMailRecords Is Nothing Then GetMailCount = m_feMailRecords.Count
End Function

' FE: Find mail records matching keyValue via FE-side Dictionary cache
Public Function FindMailRecords(keyValue As String, matchField As String, matchMode As String) As Object
    Dim result As Object: Set result = CaseDeskLib.NewDict()
    Set FindMailRecords = result
    If Len(keyValue) = 0 Then Exit Function
    If m_feMailIndex Is Nothing Then Exit Function
    If m_feMailRecords Is Nothing Then Exit Function

    ' Build lookup keys (split ";" separated, normalize)
    Dim keyParts() As String: keyParts = Split(keyValue, ";")
    Dim kp As Long
    For kp = 0 To UBound(keyParts)
        Dim normKey As String: normKey = LCase$(Trim$(keyParts(kp)))
        If matchMode = "domain" Then normKey = LCase$(GetDomain(normKey))
        If Len(normKey) = 0 Then GoTo NextKey

        ' O(1) lookup in index
        If m_feMailIndex.Exists(normKey) Then
            Dim inner As Object: Set inner = m_feMailIndex(normKey)
            Dim eids As Variant: eids = inner.keys
            Dim j As Long
            For j = 0 To UBound(eids)
                Dim eid As String: eid = CStr(eids(j))
                If Not result.Exists(eid) And m_feMailRecords.Exists(eid) Then
                    Set result(eid) = m_feMailRecords(eid)
                End If
            Next j
        End If
NextKey:
    Next kp
End Function

Public Function GetCaseCount() As Long
    GetCaseCount = 0
    If Not m_feCaseNames Is Nothing Then GetCaseCount = m_feCaseNames.Count
End Function


Public Sub CreateCaseFolder(rootPath As String, caseId As String, displayName As String)
    Dim eh As New ErrorHandler: eh.Enter "CaseDeskData", "CreateCaseFolder"
    On Error GoTo ErrHandler
    If Len(rootPath) = 0 Or Len(caseId) = 0 Then eh.OK: Exit Sub
    Dim folderName As String
    folderName = CaseDeskLib.SafeName(caseId)
    If Len(displayName) > 0 Then folderName = folderName & "_" & CaseDeskLib.SafeName(displayName)
    CaseDeskLib.EnsureFolder rootPath & "\" & folderName
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Function GetDomain(email As String) As String
    Dim pos As Long: pos = InStr(email, "@")
    If pos > 0 Then GetDomain = Mid$(email, pos + 1) Else GetDomain = email
End Function

' Load from FE's own local sheets (no cross-process)
Public Sub LoadFromLocalSheets(wb As Workbook)
    On Error Resume Next
    Dim ws As Worksheet

    Set ws = Nothing
    Set ws = wb.Worksheets("_casedesk_mail")
    If Not ws Is Nothing Then LoadMailFromLocalSheet wb

    Set ws = Nothing
    Set ws = wb.Worksheets("_casedesk_mail_idx")
    If Not ws Is Nothing Then LoadMailIndexFromLocalSheet wb

    Set ws = Nothing
    Set ws = wb.Worksheets("_casedesk_cases")
    If Not ws Is Nothing Then LoadCasesFromLocalSheet wb

    Set ws = Nothing
    Set ws = wb.Worksheets("_casedesk_files")
    If Not ws Is Nothing Then LoadCaseFilesFromLocalSheet wb

    On Error GoTo 0
End Sub

Private Sub LoadMailFromLocalSheet(wb As Workbook)
    On Error GoTo ErrOut
    Dim ws As Worksheet: Set ws = wb.Worksheets("_casedesk_mail")
    If ws.Range("A1").Value = "" Then Exit Sub
    Dim data As Variant: data = ws.UsedRange.Value
    If IsEmpty(data) Then Exit Sub
    Dim newRecs As Object: Set newRecs = CaseDeskLib.NewDict()
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim eid As String: eid = SafeStr(data(i, 1))
        If Len(eid) = 0 Then GoTo NextLMail
        Dim rec As Object: Set rec = CaseDeskLib.NewDict()
        rec.Add "entry_id", eid
        rec.Add "sender_email", SafeStr(data(i, 2))
        rec.Add "sender_name", SafeStr(data(i, 3))
        rec.Add "subject", SafeStr(data(i, 4))
        rec.Add "received_at", SafeStr(data(i, 5))
        rec.Add "folder_path", SafeStr(data(i, 6))
        rec.Add "body_path", SafeStr(data(i, 7))
        rec.Add "msg_path", SafeStr(data(i, 8))
        Dim attDict As Object: Set attDict = CaseDeskLib.NewDict()
        Dim attStr As String: attStr = SafeStr(data(i, 9))
        If Len(attStr) > 0 Then
            Dim attParts() As String: attParts = Split(attStr, "|")
            Dim a As Long
            For a = 0 To UBound(attParts)
                If Len(attParts(a)) > 0 Then
                    Dim fn As String: fn = Mid$(attParts(a), InStrRev(attParts(a), "\") + 1)
                    attDict.Add attParts(a), fn
                End If
            Next a
        End If
        rec.Add "attachment_paths", attDict
        rec.Add "_mail_folder", SafeStr(data(i, 10))
        If UBound(data, 2) >= 11 Then rec.Add "body_text", SafeStr(data(i, 11))
        Set newRecs(eid) = rec
NextLMail:
    Next i
    Set m_feMailRecords = newRecs
    Exit Sub
ErrOut:
End Sub

Private Sub LoadMailIndexFromLocalSheet(wb As Workbook)
    On Error GoTo ErrOut
    Dim ws As Worksheet: Set ws = wb.Worksheets("_casedesk_mail_idx")
    If ws.Range("A1").Value = "" Then Exit Sub
    Dim data As Variant: data = ws.UsedRange.Value
    If IsEmpty(data) Then Exit Sub
    Dim newIdx As Object: Set newIdx = CaseDeskLib.NewDict()
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim key As String: key = SafeStr(data(i, 1))
        If Len(key) = 0 Then GoTo NextLIdx
        If Not newIdx.Exists(key) Then newIdx.Add key, CaseDeskLib.NewDict()
        Dim inner As Object: Set inner = newIdx(key)
        inner(SafeStr(data(i, 2))) = True
NextLIdx:
    Next i
    Set m_feMailIndex = newIdx
    Exit Sub
ErrOut:
End Sub

Private Sub LoadCasesFromLocalSheet(wb As Workbook)
    On Error GoTo ErrOut
    Dim ws As Worksheet: Set ws = wb.Worksheets("_casedesk_cases")
    If ws.Range("A1").Value = "" Then Exit Sub
    Dim data As Variant: data = ws.UsedRange.Value
    If IsEmpty(data) Then Exit Sub
    Dim newNames As Object: Set newNames = CaseDeskLib.NewDict()
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim nm As String: nm = SafeStr(data(i, 1))
        If Len(nm) > 0 Then newNames(nm) = True
    Next i
    Set m_feCaseNames = newNames
    Exit Sub
ErrOut:
End Sub

' Load ALL case files into Dict indexed by case_id (from _casedesk_files sheet)
Private Sub LoadCaseFilesFromLocalSheet(wb As Workbook)
    On Error GoTo ErrOut
    Dim ws As Worksheet: Set ws = wb.Worksheets("_casedesk_files")
    If ws.Range("A1").Value = "" Then Exit Sub
    Dim data As Variant: data = ws.UsedRange.Value
    If IsEmpty(data) Then Exit Sub
    If UBound(data, 2) < 7 Then Exit Sub
    Dim newFiles As Object: Set newFiles = CaseDeskLib.NewDict()
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim cid As String: cid = SafeStr(data(i, 1))
        If Len(cid) = 0 Then GoTo NextFile
        If Not newFiles.Exists(cid) Then newFiles.Add cid, CaseDeskLib.NewDict()
        Dim inner As Object: Set inner = newFiles(cid)
        Dim rec As Object: Set rec = CaseDeskLib.NewDict()
        rec.Add "case_id", cid
        rec.Add "file_name", SafeStr(data(i, 2))
        rec.Add "file_path", SafeStr(data(i, 3))
        rec.Add "folder_path", SafeStr(data(i, 4))
        rec.Add "relative_path", SafeStr(data(i, 5))
        rec.Add "file_size", SafeStr(data(i, 6))
        rec.Add "modified_at", SafeStr(data(i, 7))
        Set inner(SafeStr(data(i, 3))) = rec
NextFile:
    Next i
    Set m_feCaseFiles = newFiles
    Exit Sub
ErrOut:
End Sub

' O(1) lookup: get all files for a specific case ID
Public Function FindCaseFiles(caseId As String) As Object
    Set FindCaseFiles = CaseDeskLib.NewDict()
    If m_feCaseFiles Is Nothing Then Exit Function
    If Len(caseId) = 0 Then Exit Function
    ' Prefix match: case folder may be "R06-001" or "R06-001_Name"
    Dim keys As Variant
    If m_feCaseFiles.Count = 0 Then Exit Function
    keys = m_feCaseFiles.keys
    Dim i As Long
    For i = 0 To UBound(keys)
        Dim k As String: k = CStr(keys(i))
        Dim baseName As String: baseName = k
        Dim usPos As Long: usPos = InStr(baseName, "_")
        If usPos > 0 Then baseName = Left$(baseName, usPos - 1)
        If LCase$(baseName) = LCase$(caseId) Then
            Set FindCaseFiles = m_feCaseFiles(k)
            Exit Function
        End If
    Next i
End Function
