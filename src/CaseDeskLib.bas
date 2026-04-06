Attribute VB_Name = "CaseDeskLib"
Option Explicit

' ============================================================================
' CaseDeskLib: Merged from CaseDeskHelpers + CaseDeskConfig + CaseDeskChangeLog
' ============================================================================

' ##########################################################################
' # SECTION: General Utilities (ex-CaseDeskHelpers)
' ##########################################################################

' --- Helpers state ---
Private m_readStm As Object  ' Reusable ADODB.Stream for ReadTextFile
Private m_writeStm As Object ' Reusable ADODB.Stream for WriteTextFile

' --- Config state ---
Private Const SH_CONFIG As String = "_casedesk_config"
Private Const SH_SOURCES As String = "_casedesk_sources"
Private Const SH_FIELDS As String = "_casedesk_fields"
Private m_cfg As Object       ' Dict: key -> value
Private m_sources As Object   ' Dict: source_name -> Dict(col -> value)
Private m_fields As Object    ' Dict: "source|field" -> Dict(col -> value)
Private m_loaded As Boolean
Private m_dirty As Boolean

' --- ChangeLog state ---
Private Const LOG_SHEET As String = "_casedesk_log"
Private Const LOG_TABLE As String = "CaseDeskLog"
Private Const MAX_LOG_ROWS As Long = 5000

' ============================================================================
' JSON Parser
' ============================================================================

Public Function ParseJson(ByVal text As String) As Object
    Dim p As Long: p = 1
    SkipWS text, p
    If p > Len(text) Then Set ParseJson = NewDict(): Exit Function
    Dim ch As String: ch = Mid$(text, p, 1)
    If ch = "{" Then
        Set ParseJson = ParseObj(text, p)
    ElseIf ch = "[" Then
        Set ParseJson = ParseArr(text, p)
    Else
        Set ParseJson = NewDict()
    End If
End Function

Private Function ParseObj(ByRef s As String, ByRef p As Long) As Object
    Dim d As Object: Set d = NewDict()
    p = p + 1
    SkipWS s, p
    If p <= Len(s) Then
        If Mid$(s, p, 1) = "}" Then p = p + 1: Set ParseObj = d: Exit Function
    End If
    Do
        SkipWS s, p
        Dim k As String: k = ParseStr(s, p)
        SkipWS s, p
        If p <= Len(s) Then p = p + 1 ' skip :
        Dim v As Variant: ParseVal s, p, v
        d.Add k, v
        SkipWS s, p
        If p > Len(s) Then Exit Do
        If Mid$(s, p, 1) = "}" Then p = p + 1: Exit Do
        p = p + 1 ' skip ,
    Loop
    Set ParseObj = d
End Function

Private Function ParseArr(ByRef s As String, ByRef p As Long) As Object
    Dim c As New Collection
    p = p + 1
    SkipWS s, p
    If p <= Len(s) Then
        If Mid$(s, p, 1) = "]" Then p = p + 1: Set ParseArr = c: Exit Function
    End If
    Do
        SkipWS s, p
        Dim v As Variant: ParseVal s, p, v
        c.Add v
        SkipWS s, p
        If p > Len(s) Then Exit Do
        If Mid$(s, p, 1) = "]" Then p = p + 1: Exit Do
        p = p + 1 ' skip ,
    Loop
    Set ParseArr = c
End Function

Private Sub ParseVal(ByRef s As String, ByRef p As Long, ByRef result As Variant)
    SkipWS s, p
    If p > Len(s) Then result = Null: Exit Sub
    Dim ch As String: ch = Mid$(s, p, 1)
    Select Case ch
        Case "{":  Set result = ParseObj(s, p)
        Case "[":  Set result = ParseArr(s, p)
        Case """": result = ParseStr(s, p)
        Case "t":  result = True: p = p + 4
        Case "f":  result = False: p = p + 5
        Case "n":  result = Null: p = p + 4
        Case Else: result = ParseNum(s, p)
    End Select
End Sub

Private Function ParseStr(ByRef s As String, ByRef p As Long) As String
    p = p + 1
    Dim buf As String, start As Long: start = p
    Do While p <= Len(s)
        Dim ch As String: ch = Mid$(s, p, 1)
        If ch = """" Then
            buf = buf & Mid$(s, start, p - start)
            p = p + 1
            ParseStr = buf: Exit Function
        ElseIf ch = "\" Then
            buf = buf & Mid$(s, start, p - start)
            p = p + 1
            If p <= Len(s) Then
                Dim esc As String: esc = Mid$(s, p, 1)
                Select Case esc
                    Case """", "\", "/": buf = buf & esc
                    Case "n": buf = buf & vbLf
                    Case "r": buf = buf & vbCr
                    Case "t": buf = buf & vbTab
                    Case "u"
                        If p + 4 <= Len(s) Then
                            On Error Resume Next
                            buf = buf & ChrW$(CLng("&H" & Mid$(s, p + 1, 4)))
                            On Error GoTo 0
                            p = p + 4
                        End If
                End Select
                p = p + 1: start = p
            End If
        Else
            p = p + 1
        End If
    Loop
    ParseStr = buf & Mid$(s, start, p - start)
End Function

Private Function ParseNum(ByRef s As String, ByRef p As Long) As Double
    Dim start As Long: start = p
    If p <= Len(s) Then If Mid$(s, p, 1) = "-" Then p = p + 1
    Do While p <= Len(s)
        Dim ch As String: ch = Mid$(s, p, 1)
        If ch Like "[0-9.eE+-]" Then p = p + 1 Else Exit Do
    Loop
    On Error Resume Next
    ParseNum = CDbl(Mid$(s, start, p - start))
    On Error GoTo 0
End Function

Private Sub SkipWS(ByRef s As String, ByRef p As Long)
    Do While p <= Len(s)
        Select Case Mid$(s, p, 1)
            Case " ", vbTab, vbLf, vbCr: p = p + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

' ============================================================================
' JSON Serializer
' ============================================================================

Public Function ToJson(ByVal v As Variant, Optional ind As Long = -1) As String
    If IsObject(v) Then
        If v Is Nothing Then ToJson = "null": Exit Function
        Dim obj As Object: Set obj = v
        If TypeName(obj) = "Dictionary" Then ToJson = DictToJson(obj, ind): Exit Function
        If TypeName(obj) = "Collection" Then ToJson = CollToJson(obj, ind): Exit Function
        ToJson = "null"
    ElseIf IsNull(v) Or IsEmpty(v) Then
        ToJson = "null"
    ElseIf VarType(v) = vbString Then
        ToJson = """" & JsonEscape(CStr(v)) & """"
    ElseIf VarType(v) = vbBoolean Then
        ToJson = IIf(v, "true", "false")
    ElseIf IsNumeric(v) Then
        ToJson = CStr(v)
    Else
        ToJson = """" & JsonEscape(CStr(v)) & """"
    End If
End Function

Private Function DictToJson(d As Object, ind As Long) As String
    If d.Count = 0 Then DictToJson = "{}": Exit Function
    Dim keys() As Variant: keys = d.keys
    Dim nl As String, sp As String, ind2 As Long, csp As String
    If ind >= 0 Then nl = vbCrLf: sp = String$(ind + 2, " "): ind2 = ind + 2: csp = String$(ind, " ") Else ind2 = -1
    Dim parts() As String: ReDim parts(d.Count - 1)
    Dim i As Long
    For i = 0 To d.Count - 1
        Dim val As Variant
        If IsObject(d(keys(i))) Then Set val = d(keys(i)) Else val = d(keys(i))
        parts(i) = sp & """" & JsonEscape(CStr(keys(i))) & """:" & IIf(ind >= 0, " ", "") & ToJson(val, ind2)
    Next i
    DictToJson = "{" & nl & Join(parts, "," & nl) & nl & csp & "}"
End Function

Private Function CollToJson(c As Object, ind As Long) As String
    If c.Count = 0 Then CollToJson = "[]": Exit Function
    Dim nl As String, sp As String, ind2 As Long, csp As String
    If ind >= 0 Then nl = vbCrLf: sp = String$(ind + 2, " "): ind2 = ind + 2: csp = String$(ind, " ") Else ind2 = -1
    Dim parts() As String: ReDim parts(c.Count - 1)
    Dim i As Long
    For i = 1 To c.Count
        Dim val As Variant
        If IsObject(c(i)) Then Set val = c(i) Else val = c(i)
        parts(i - 1) = sp & ToJson(val, ind2)
    Next i
    CollToJson = "[" & nl & Join(parts, "," & nl) & nl & csp & "]"
End Function

Public Function JsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonEscape = s
End Function

' ============================================================================
' Dictionary Helpers
' ============================================================================

Public Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
End Function

Public Function DictStr(d As Object, key As String, Optional def As String = "") As String
    DictStr = def
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If IsObject(d(key)) Or IsNull(d(key)) Then Exit Function
    DictStr = CStr(d(key))
End Function

Public Function DictObj(d As Object, key As String) As Object
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If Not IsObject(d(key)) Then Exit Function
    Set DictObj = d(key)
End Function

Public Function DictBool(d As Object, key As String, Optional def As Boolean = False) As Boolean
    DictBool = def
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If IsObject(d(key)) Or IsNull(d(key)) Then Exit Function
    DictBool = CBool(d(key))
End Function

Public Function DictLng(d As Object, key As String, Optional def As Long = 0) As Long
    DictLng = def
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If IsObject(d(key)) Or IsNull(d(key)) Then Exit Function
    DictLng = CLng(d(key))
End Function

Public Sub DictPut(d As Object, key As String, val As Variant)
    If d.Exists(key) Then d.Remove key
    d.Add key, val
End Sub

' ============================================================================
' File System
' ============================================================================

Public Function ReadTextFile(path As String) As String
    On Error GoTo ErrOut
    ReadTextFile = ""
    If Len(Dir$(path)) = 0 Then Exit Function

    ' Read raw bytes with shared access (avoids file lock)
    Dim f As Long: f = FreeFile
    Open path For Binary Access Read Shared As #f
    Dim size As Long: size = LOF(f)
    If size = 0 Then Close #f: Exit Function
    Dim buf() As Byte: ReDim buf(0 To size - 1)
    Get #f, , buf
    Close #f

    ' Convert UTF-8 bytes to VBA string (reuse stream object)
    If m_readStm Is Nothing Then Set m_readStm = CreateObject("ADODB.Stream")
    m_readStm.Type = 1: m_readStm.Open: m_readStm.Write buf
    m_readStm.Position = 0: m_readStm.Type = 2: m_readStm.Charset = "UTF-8"
    ReadTextFile = m_readStm.ReadText
    m_readStm.Close
    Exit Function
ErrOut:
    ReadTextFile = ""
    On Error Resume Next: If Not m_readStm Is Nothing Then m_readStm.Close: On Error GoTo 0
End Function

Public Sub WriteTextFile(path As String, content As String)
    On Error GoTo ErrOut
    If m_writeStm Is Nothing Then Set m_writeStm = CreateObject("ADODB.Stream")
    m_writeStm.Type = 2: m_writeStm.Charset = "UTF-8"
    m_writeStm.Open: m_writeStm.WriteText content
    m_writeStm.Position = 0: m_writeStm.Type = 1: m_writeStm.Position = 3
    Dim out As Object: Set out = CreateObject("ADODB.Stream")
    out.Type = 1: out.Open
    m_writeStm.CopyTo out
    out.SaveToFile path, 2
    out.Close: m_writeStm.Close
    Exit Sub
ErrOut:
    On Error Resume Next
    If Not m_writeStm Is Nothing Then
        m_writeStm.Close
        Set m_writeStm = Nothing
    End If
    On Error GoTo 0
End Sub

Public Sub EnsureFolder(path As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path) Then Exit Sub
    Dim parent As String: parent = fso.GetParentFolderName(path)
    If Len(parent) > 0 Then
        If Not fso.FolderExists(parent) Then EnsureFolder parent
    End If
    If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub

Public Function FileExists(path As String) As Boolean
    FileExists = Len(Dir$(path)) > 0
End Function

Public Function FolderExists(path As String) As Boolean
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(path)
End Function

' ============================================================================
' String Utilities
' ============================================================================

Public Function SafeName(ByVal text As String) As String
    text = Trim$(text)
    If Len(text) = 0 Then text = "blank"
    Dim bad As Variant
    For Each bad In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        text = Replace(text, CStr(bad), "_")
    Next bad
    If Len(text) > 80 Then text = Left$(text, 80)
    SafeName = text
End Function

Public Function FormatFieldValue(val As Variant, Optional fieldType As String = "text") As String
    If IsNull(val) Or IsEmpty(val) Then FormatFieldValue = "": Exit Function
    If IsObject(val) Then
        If TypeName(val) = "Collection" Then
            Dim parts() As String: ReDim parts(val.Count - 1)
            Dim i As Long
            For i = 1 To val.Count: parts(i - 1) = CStr(val(i)): Next i
            FormatFieldValue = Join(parts, "; "): Exit Function
        End If
        FormatFieldValue = "": Exit Function
    End If
    Dim s As String: s = CStr(val)
    Select Case fieldType
        Case "date"
            If VarType(val) = vbDate Then
                FormatFieldValue = Format$(CDate(val), "yyyy/mm/dd"): Exit Function
            End If
        Case "currency", "number"
            If IsNumeric(s) Then
                Dim n As Double: n = CDbl(s)
                If n = Int(n) Then
                    FormatFieldValue = Format$(n, "#,##0")
                Else
                    FormatFieldValue = Format$(n, "#,##0.##")
                End If
                Exit Function
            End If
    End Select
    FormatFieldValue = s
End Function

' ============================================================================
' Field Name Prefix System
'   xx_AAA     -> normal: visible + editable
'   _xx_AAA    -> readonly: visible + NOT editable (single leading _)
'   __AAA      -> hidden: NOT visible in UI, but usable as setting columns
' ============================================================================

Public Function IsHiddenField(fieldName As String) As Boolean
    ' "__AAA" pattern: starts with "__" and third char is NOT "_"
    If Len(fieldName) >= 3 Then
        IsHiddenField = (Left$(fieldName, 2) = "__" And Mid$(fieldName, 3, 1) <> "_")
    End If
End Function

Public Function IsReadOnlyField(fieldName As String) As Boolean
    ' "_x..." pattern: starts with "_" but not "__" (min 2 chars)
    If Len(fieldName) >= 2 Then
        If Left$(fieldName, 1) = "_" And Mid$(fieldName, 2, 1) <> "_" Then
            IsReadOnlyField = True
        End If
    End If
End Function

Public Function StripFieldPrefix(fieldName As String) As String
    ' Remove leading _ for display: _xx_AAA -> xx_AAA, __AAA -> AAA (for settings)
    If IsReadOnlyField(fieldName) Then
        StripFieldPrefix = Mid$(fieldName, 2)
    ElseIf IsHiddenField(fieldName) Then
        StripFieldPrefix = Mid$(fieldName, 3)
    Else
        StripFieldPrefix = fieldName
    End If
End Function

Public Function GetFieldLabel(fieldName As String) As String
    GetFieldLabel = Replace(StripFieldPrefix(fieldName), "_", " ")
End Function

Public Function GetFieldGroup(fieldName As String) As String
    Dim cleaned As String: cleaned = StripFieldPrefix(fieldName)
    Dim pos As Long: pos = InStr(cleaned, "_")
    If pos > 1 And pos < Len(cleaned) Then GetFieldGroup = Left$(cleaned, pos - 1)
End Function

Public Function GetFieldShortName(fieldName As String) As String
    Dim cleaned As String: cleaned = StripFieldPrefix(fieldName)
    Dim pos As Long: pos = InStr(cleaned, "_")
    If pos > 1 And pos < Len(cleaned) Then
        GetFieldShortName = Mid$(cleaned, pos + 1)
    Else
        GetFieldShortName = cleaned
    End If
End Function

Public Function CountFieldGroups(fields As Collection) As Long
    Dim groups As Object: Set groups = NewDict()
    Dim i As Long
    For i = 1 To fields.Count
        Dim g As String: g = GetFieldGroup(CStr(fields(i)))
        If Len(g) > 0 And Not groups.Exists(g) Then groups.Add g, True
    Next i
    CountFieldGroups = groups.Count
End Function

' ##########################################################################
' # SECTION: Config Management (ex-CaseDeskConfig)
' ##########################################################################

' ============================================================================
' Init / Save
' ============================================================================

Public Sub EnsureConfigSheets()
    EnsureSheet SH_CONFIG, Array("key", "value")
    EnsureSheet SH_SOURCES, Array("source_name", "source_sheet", "key_column", "display_name_column", "mail_link_column", "folder_link_column", "mail_match_mode")
    EnsureSheet SH_FIELDS, Array("source_name", "field_name", "display_name", "type", "visible", "in_list", "editable", "multiline", "role", "sort_order")
    If Not m_loaded Then LoadFromSheets
End Sub

Public Sub SaveToSheets()
    If Not m_loaded Then Exit Sub
    If Not m_dirty Then Exit Sub
    On Error Resume Next
    SaveConfigSheet
    SaveSourcesSheet
    SaveFieldsSheet
    If Err.Number = 0 Then m_dirty = False
    On Error GoTo 0
End Sub

Public Sub LoadFromSheets()
    On Error Resume Next
    Set m_cfg = CreateObject("Scripting.Dictionary")
    m_cfg.CompareMode = vbTextCompare
    Set m_sources = CreateObject("Scripting.Dictionary")
    m_sources.CompareMode = vbTextCompare
    Set m_fields = CreateObject("Scripting.Dictionary")
    m_fields.CompareMode = vbTextCompare

    ' Load config KV
    Dim wsCfg As Worksheet: Set wsCfg = ThisWorkbook.Worksheets(SH_CONFIG)
    If Not wsCfg Is Nothing Then
        Dim r As Long
        For r = 2 To wsCfg.Cells(wsCfg.Rows.Count, 1).End(xlUp).Row
            Dim k As String: k = CStr(wsCfg.Cells(r, 1).Value)
            If Len(k) > 0 Then m_cfg(k) = CStr(wsCfg.Cells(r, 2).Value)
        Next r
    End If

    ' Load sources
    Dim wsSrc As Worksheet: Set wsSrc = ThisWorkbook.Worksheets(SH_SOURCES)
    If Not wsSrc Is Nothing Then
        Dim srcCols As Object: Set srcCols = ReadHeaderMap(wsSrc)
        For r = 2 To wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
            Dim sn As String: sn = CStr(wsSrc.Cells(r, 1).Value)
            If Len(sn) > 0 Then
                Dim sd As Object: Set sd = CreateObject("Scripting.Dictionary")
                Dim ck As Variant
                For Each ck In srcCols.keys
                    sd(CStr(ck)) = CStr(wsSrc.Cells(r, CLng(srcCols(ck))).Value)
                Next ck
                Set m_sources(sn) = sd
            End If
        Next r
    End If

    ' Load fields
    Dim wsFld As Worksheet: Set wsFld = ThisWorkbook.Worksheets(SH_FIELDS)
    If Not wsFld Is Nothing Then
        Dim fldCols As Object: Set fldCols = ReadHeaderMap(wsFld)
        For r = 2 To wsFld.Cells(wsFld.Rows.Count, 1).End(xlUp).Row
            Dim fs As String: fs = CStr(wsFld.Cells(r, 1).Value)
            Dim ff As String: ff = CStr(wsFld.Cells(r, 2).Value)
            If Len(fs) > 0 And Len(ff) > 0 Then
                Dim fk As String: fk = LCase$(fs) & "|" & LCase$(ff)
                Dim fd As Object: Set fd = CreateObject("Scripting.Dictionary")
                fd("source_name") = fs
                fd("field_name") = ff
                For Each ck In fldCols.keys
                    fd(CStr(ck)) = CStr(wsFld.Cells(r, CLng(fldCols(ck))).Value)
                Next ck
                Set m_fields(fk) = fd
            End If
        Next r
    End If

    m_loaded = True
    m_dirty = False
    On Error GoTo 0
End Sub

Private Function ReadHeaderMap(ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Dim h As String: h = CStr(ws.Cells(1, c).Value)
        If Len(h) > 0 Then d(h) = c
    Next c
    Set ReadHeaderMap = d
End Function

' ============================================================================
' Config (key-value)
' ============================================================================

Public Function GetStr(key As String, Optional def As String = "") As String
    If Not m_loaded Then EnsureConfigSheets
    GetStr = def
    If m_cfg.Exists(key) Then
        Dim v As String: v = CStr(m_cfg(key))
        If Len(v) > 0 Then GetStr = v
    End If
End Function

Public Sub SetStr(key As String, value As String)
    If Not m_loaded Then EnsureConfigSheets
    m_cfg(key) = value
    m_dirty = True
End Sub

Public Function GetLng(key As String, Optional def As Long = 0) As Long
    GetLng = def
    Dim s As String: s = GetStr(key)
    If Len(s) > 0 And IsNumeric(s) Then GetLng = CLng(s)
End Function

Public Sub SetLng(key As String, value As Long)
    SetStr key, CStr(value)
End Sub

' ============================================================================
' Sources
' ============================================================================

Public Function GetSourceNames() As Collection
    If Not m_loaded Then EnsureConfigSheets
    Set GetSourceNames = New Collection
    Dim k As Variant
    For Each k In m_sources.keys
        GetSourceNames.Add CStr(k)
    Next k
End Function

Public Function GetSourceStr(src As String, col As String, Optional def As String = "") As String
    If Not m_loaded Then EnsureConfigSheets
    GetSourceStr = def
    If Not m_sources.Exists(src) Then Exit Function
    Dim sd As Object: Set sd = m_sources(src)
    If sd.Exists(col) Then
        Dim v As String: v = CStr(sd(col))
        If Len(v) > 0 Then GetSourceStr = v
    End If
End Function

Public Sub SetSourceStr(src As String, col As String, value As String)
    If Not m_loaded Then EnsureConfigSheets
    If Not m_sources.Exists(src) Then Set m_sources(src) = CreateObject("Scripting.Dictionary")
    Dim sd As Object: Set sd = m_sources(src)
    sd(col) = value
    m_dirty = True
End Sub

Public Sub EnsureSource(src As String)
    If Not m_loaded Then EnsureConfigSheets
    If Not m_sources.Exists(src) Then
        Set m_sources(src) = CreateObject("Scripting.Dictionary")
        m_sources(src)("source_name") = src
        m_dirty = True
    End If
End Sub

Public Sub RemoveSource(src As String)
    If Not m_loaded Then EnsureConfigSheets
    If m_sources.Exists(src) Then
        m_sources.Remove src
        ' Also remove associated field entries
        Dim prefix As String: prefix = LCase$(src) & "|"
        Dim toRemove As New Collection
        Dim k As Variant
        For Each k In m_fields.keys
            If Left$(CStr(k), Len(prefix)) = prefix Then toRemove.Add CStr(k)
        Next k
        Dim i As Long
        For i = 1 To toRemove.Count: m_fields.Remove CStr(toRemove(i)): Next i
        m_dirty = True
    End If
End Sub

' ============================================================================
' Fields
' ============================================================================

Public Function GetFieldNames(src As String) As Collection
    If Not m_loaded Then EnsureConfigSheets
    Set GetFieldNames = SortedFieldNames(src, False)
End Function

Public Function GetVisibleFieldNames(src As String) As Collection
    If Not m_loaded Then EnsureConfigSheets
    Set GetVisibleFieldNames = SortedFieldNames(src, True)
End Function

Public Function GetFieldStr(src As String, fld As String, col As String, Optional def As String = "") As String
    If Not m_loaded Then EnsureConfigSheets
    GetFieldStr = def
    Dim fk As String: fk = LCase$(src) & "|" & LCase$(fld)
    If Not m_fields.Exists(fk) Then Exit Function
    Dim fd As Object: Set fd = m_fields(fk)
    If fd.Exists(col) Then
        Dim v As String: v = CStr(fd(col))
        If Len(v) > 0 Then GetFieldStr = v
    End If
End Function

Public Function GetFieldBool(src As String, fld As String, col As String, Optional def As Boolean = False) As Boolean
    GetFieldBool = def
    Dim v As String: v = GetFieldStr(src, fld, col)
    If Len(v) > 0 Then GetFieldBool = CBool(v)
End Function

Public Sub SetFieldStr(src As String, fld As String, col As String, value As String)
    If Not m_loaded Then EnsureConfigSheets
    Dim fk As String: fk = LCase$(src) & "|" & LCase$(fld)
    If Not m_fields.Exists(fk) Then
        Dim fd As Object: Set fd = CreateObject("Scripting.Dictionary")
        fd("source_name") = src
        fd("field_name") = fld
        Set m_fields(fk) = fd
    End If
    Dim d As Object: Set d = m_fields(fk)
    d(col) = value
    m_dirty = True
End Sub

Public Sub SetFieldBool(src As String, fld As String, col As String, value As Boolean)
    SetFieldStr src, fld, col, CStr(value)
End Sub

Public Sub EnsureField(src As String, fld As String)
    If Not m_loaded Then EnsureConfigSheets
    Dim fk As String: fk = LCase$(src) & "|" & LCase$(fld)
    If Not m_fields.Exists(fk) Then
        SetFieldStr src, fld, "display_name", StripFieldPrefix(fld)
        SetFieldStr src, fld, "type", "text"
        SetFieldStr src, fld, "visible", CStr(True)
        SetFieldStr src, fld, "in_list", CStr(False)
        SetFieldStr src, fld, "editable", CStr(True)
        SetFieldStr src, fld, "multiline", CStr(False)
        SetFieldStr src, fld, "role", ""
        SetFieldStr src, fld, "sort_order", "0"
    End If
End Sub

Public Function GetFieldDisplayName(src As String, fld As String) As String
    GetFieldDisplayName = GetFieldStr(src, fld, "display_name", StripFieldPrefix(fld))
End Function

' ============================================================================
' Field Settings Auto-Init
' ============================================================================

Public Sub InitFieldSettingsFromTable(src As String, tbl As ListObject)
    If Not m_loaded Then EnsureConfigSheets

    ' Build set of current table columns
    Dim currentCols As Object: Set currentCols = CreateObject("Scripting.Dictionary")
    Dim col As ListColumn
    Dim ordinal As Long: ordinal = 0
    For Each col In tbl.ListColumns
        On Error Resume Next
        Err.Clear
        ordinal = ordinal + 1
        ' Track all columns (including hidden) for stale removal
        currentCols(LCase$(col.Name)) = col.Name
        ' Skip hidden (__AAA) fields from UI field list
        If IsHiddenField(col.Name) Then GoTo NextCol
        Dim fk As String: fk = LCase$(src) & "|" & LCase$(col.Name)
        If Not m_fields.Exists(fk) Then
            EnsureField src, col.Name
            SetFieldStr src, col.Name, "type", GuessFieldType(col)
            SetFieldStr src, col.Name, "multiline", CStr(GuessMultiline(col))
            SetFieldStr src, col.Name, "display_name", StripFieldPrefix(col.Name)
            SetFieldStr src, col.Name, "visible", CStr(True)
            SetFieldStr src, col.Name, "role", ""
            SetFieldStr src, col.Name, "sort_order", CStr(ordinal)
            ' Read-only prefix (_xx_AAA) -> editable=False
            If IsReadOnlyField(col.Name) Then
                SetFieldStr src, col.Name, "editable", CStr(False)
            End If
        Else
            If Len(GetFieldStr(src, col.Name, "display_name")) = 0 Then
                SetFieldStr src, col.Name, "display_name", StripFieldPrefix(col.Name)
            End If
            If Len(GetFieldStr(src, col.Name, "visible")) = 0 Then
                SetFieldStr src, col.Name, "visible", CStr(True)
            End If
            If Len(GetFieldStr(src, col.Name, "sort_order")) = 0 Then
                SetFieldStr src, col.Name, "sort_order", CStr(ordinal)
            End If
        End If
NextCol:
        On Error GoTo 0
    Next col

    ' Remove field entries for columns that no longer exist in the table
    Dim prefix As String: prefix = LCase$(src) & "|"
    Dim toRemove As New Collection
    Dim k As Variant
    For Each k In m_fields.keys
        If Left$(CStr(k), Len(prefix)) = prefix Then
            Dim colName As String: colName = Mid$(CStr(k), Len(prefix) + 1)
            If Not currentCols.Exists(colName) Then toRemove.Add CStr(k)
        End If
    Next k
    Dim ri As Long
    For ri = 1 To toRemove.Count
        m_fields.Remove CStr(toRemove(ri))
        m_dirty = True
    Next ri
End Sub

Public Sub InitFieldSettingsFromRange(src As String, ws As Worksheet)
    If Not m_loaded Then EnsureConfigSheets
    Dim ur As Range
    ' Try to resolve src as named range or direct address before falling back to UsedRange
    On Error Resume Next
    Set ur = ws.Parent.Names(src).RefersToRange
    On Error GoTo 0
    If ur Is Nothing Then
        On Error Resume Next
        Set ur = ws.Parent.Names(ws.Name & "!" & src).RefersToRange
        On Error GoTo 0
    End If
    If ur Is Nothing Then
        On Error Resume Next
        Set ur = ws.Range(src)
        On Error GoTo 0
    End If
    If ur Is Nothing Then
        On Error Resume Next
        Set ur = ws.UsedRange
        On Error GoTo 0
    End If
    If ur Is Nothing Then Exit Sub

    Dim headerRow As Long: headerRow = ur.Row
    Dim startCol As Long: startCol = ur.Column
    Dim nCols As Long: nCols = ur.Columns.Count

    ' Build set of current columns
    Dim currentCols As Object: Set currentCols = CreateObject("Scripting.Dictionary")
    Dim ordinal As Long: ordinal = 0
    Dim c As Long
    For c = 0 To nCols - 1
        On Error Resume Next
        Err.Clear
        Dim colName As String
        Dim cv As Variant: cv = ws.Cells(headerRow, startCol + c).Value
        If IsEmpty(cv) Or Len(CStr(cv)) = 0 Then GoTo NextRangeCol
        colName = CStr(cv)
        ordinal = ordinal + 1
        currentCols(LCase$(colName)) = colName

        If IsHiddenField(colName) Then GoTo NextRangeCol
        Dim fk As String: fk = LCase$(src) & "|" & LCase$(colName)
        If Not m_fields.Exists(fk) Then
            EnsureField src, colName
            SetFieldStr src, colName, "display_name", StripFieldPrefix(colName)
            SetFieldStr src, colName, "visible", CStr(True)
            SetFieldStr src, colName, "role", ""
            SetFieldStr src, colName, "sort_order", CStr(ordinal)
            ' Guess type from cell data
            Dim guessType As String: guessType = "text"
            Dim sampleRow As Long: sampleRow = headerRow + 1
            If sampleRow <= ws.Cells(ws.Rows.Count, startCol + c).End(xlUp).Row Then
                Dim sv As Variant: sv = ws.Cells(sampleRow, startCol + c).Value
                If Not IsEmpty(sv) And Not IsNull(sv) Then
                    If VarType(sv) = vbDate Then guessType = "date"
                    If VarType(sv) = vbDouble Or VarType(sv) = vbLong Or VarType(sv) = vbInteger Then guessType = "number"
                End If
            End If
            SetFieldStr src, colName, "type", guessType
            If IsReadOnlyField(colName) Then
                SetFieldStr src, colName, "editable", CStr(False)
            End If
        Else
            If Len(GetFieldStr(src, colName, "display_name")) = 0 Then
                SetFieldStr src, colName, "display_name", StripFieldPrefix(colName)
            End If
            If Len(GetFieldStr(src, colName, "sort_order")) = 0 Then
                SetFieldStr src, colName, "sort_order", CStr(ordinal)
            End If
        End If
        On Error GoTo 0
NextRangeCol:
    Next c

    ' Remove stale field entries
    Dim prefix As String: prefix = LCase$(src) & "|"
    Dim toRemove As New Collection
    Dim k As Variant
    For Each k In m_fields.keys
        If Left$(CStr(k), Len(prefix)) = prefix Then
            Dim cn As String: cn = Mid$(CStr(k), Len(prefix) + 1)
            If Not currentCols.Exists(cn) Then toRemove.Add CStr(k)
        End If
    Next k
    Dim ri As Long
    For ri = 1 To toRemove.Count
        m_fields.Remove CStr(toRemove(ri))
        m_dirty = True
    Next ri
End Sub

Public Function DetectColumnChanges(src As String, tbl As ListObject) As String
    If Not m_loaded Then EnsureConfigSheets
    DetectColumnChanges = ""

    ' Get saved field names for this source
    Dim prefix As String: prefix = LCase$(src) & "|"
    Dim savedCols As Object: Set savedCols = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In m_fields.keys
        If Left$(CStr(k), Len(prefix)) = prefix Then
            Dim savedName As String: savedName = Mid$(CStr(k), Len(prefix) + 1)
            savedCols(savedName) = True
        End If
    Next k
    If savedCols.Count = 0 Then Exit Function ' First time, no diff

    ' Get current table columns
    Dim currentCols As Object: Set currentCols = CreateObject("Scripting.Dictionary")
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        currentCols(LCase$(col.Name)) = col.Name
    Next col

    ' Detect added columns (skip hidden __ fields)
    Dim added As String
    For Each k In currentCols.keys
        If Not savedCols.Exists(CStr(k)) Then
            If Not IsHiddenField(CStr(currentCols(k))) Then
                added = added & "  + " & CStr(currentCols(k)) & vbCrLf
            End If
        End If
    Next k

    ' Detect removed columns
    Dim removed As String
    For Each k In savedCols.keys
        If Not currentCols.Exists(CStr(k)) Then
            ' Resolve original casing from saved data
            Dim fd As Object: Set fd = m_fields(prefix & CStr(k))
            Dim origName As String: origName = DictStr(fd, "field_name", CStr(k))
            removed = removed & "  - " & origName & vbCrLf
        End If
    Next k

    If Len(added) > 0 Then DetectColumnChanges = "Added:" & vbCrLf & added
    If Len(removed) > 0 Then DetectColumnChanges = DetectColumnChanges & "Removed:" & vbCrLf & removed
End Function

Private Function SortedFieldNames(src As String, visibleOnly As Boolean) As Collection
    ' Collect matching entries into a growable array
    Dim cnt As Long: cnt = 0
    Dim arrNames() As String, arrOrders() As Long
    ReDim arrNames(0 To 15): ReDim arrOrders(0 To 15)
    Dim k As Variant
    For Each k In m_fields.keys
        If Left$(CStr(k), Len(src) + 1) = LCase$(src) & "|" Then
            Dim fd As Object: Set fd = m_fields(k)
            If Not visibleOnly Or DictBool(fd, "visible", True) Then
                If cnt > UBound(arrNames) Then
                    ReDim Preserve arrNames(0 To cnt * 2)
                    ReDim Preserve arrOrders(0 To cnt * 2)
                End If
                arrNames(cnt) = CStr(fd("field_name"))
                arrOrders(cnt) = CLng(Val(DictStr(fd, "sort_order", "0")))
                cnt = cnt + 1
            End If
        End If
    Next k

    ' Bubble sort on arrays (swappable, unlike Collection)
    Dim i As Long, j As Long
    For i = 0 To cnt - 2
        For j = i + 1 To cnt - 1
            If arrOrders(j) < arrOrders(i) Then
                Dim tmpN As String: tmpN = arrNames(i)
                Dim tmpO As Long: tmpO = arrOrders(i)
                arrNames(i) = arrNames(j): arrOrders(i) = arrOrders(j)
                arrNames(j) = tmpN: arrOrders(j) = tmpO
            End If
        Next j
    Next i

    Set SortedFieldNames = New Collection
    For i = 0 To cnt - 1
        SortedFieldNames.Add arrNames(i)
    Next i
End Function


Private Function GuessFieldType(col As ListColumn) As String
    GuessFieldType = "text"
    On Error Resume Next
    If col.DataBodyRange Is Nothing Then Exit Function
    If col.DataBodyRange.Rows.Count = 0 Then Exit Function
    Dim fmt As String: fmt = CStr(col.DataBodyRange.Cells(1, 1).NumberFormat)
    If fmt Like "*yy*" Or fmt Like "*mm*dd*" Then GuessFieldType = "date": Exit Function
    If fmt Like "*#,##0*" Or fmt Like "*" & ChrW$(165) & "*" Or fmt Like "*$*" Then GuessFieldType = "currency": Exit Function
    If fmt Like "#*" Or fmt Like "0*" Or fmt Like "*%*" Then GuessFieldType = "number": Exit Function
    Dim r As Long
    For r = 1 To Application.Min(10, col.DataBodyRange.Rows.Count)
        Dim v As Variant: v = col.DataBodyRange.Cells(r, 1).Value
        If Not IsEmpty(v) And Not IsNull(v) Then
            If VarType(v) = vbDate Then GuessFieldType = "date": Exit Function
            If VarType(v) = vbCurrency Then GuessFieldType = "currency": Exit Function
            If VarType(v) = vbDouble Or VarType(v) = vbLong Or VarType(v) = vbInteger Or _
               VarType(v) = vbSingle Then GuessFieldType = "number": Exit Function
            Exit Function
        End If
    Next r
    On Error GoTo 0
End Function

Private Function GuessMultiline(col As ListColumn) As Boolean
    GuessMultiline = False
    On Error Resume Next
    If col.DataBodyRange Is Nothing Then Exit Function
    Dim r As Long
    For r = 1 To Application.Min(10, col.DataBodyRange.Rows.Count)
        Dim v As Variant: v = col.DataBodyRange.Cells(r, 1).Value
        If Not IsEmpty(v) And Not IsNull(v) Then
            Dim s As String: s = CStr(v)
            If InStr(s, vbLf) > 0 Or InStr(s, vbCr) > 0 Or Len(s) > 30 Then
                GuessMultiline = True: Exit Function
            End If
        End If
    Next r
    On Error GoTo 0
End Function

' ============================================================================
' Settings Export / Import
' ============================================================================

Public Function ExportSettings(filePath As String) As Boolean
    If Not m_loaded Then EnsureConfigSheets
    On Error GoTo ErrOut

    Dim buf As String
    Dim srcCols As Variant: srcCols = Array("source_name", "source_sheet", "key_column", "display_name_column", "mail_link_column", "folder_link_column", "mail_match_mode")
    Dim fldCols As Variant: fldCols = Array("source_name", "field_name", "display_name", "type", "visible", "in_list", "editable", "multiline", "role", "sort_order")

    ' [config]
    buf = "[config]" & vbCrLf & "key,value" & vbCrLf
    Dim ck As Variant
    For Each ck In m_cfg.keys
        buf = buf & CsvEsc(CStr(ck)) & "," & CsvEsc(CStr(m_cfg(ck))) & vbCrLf
    Next ck

    ' [sources]
    buf = buf & vbCrLf & "[sources]" & vbCrLf & Join(srcCols, ",") & vbCrLf
    Dim sk As Variant
    For Each sk In m_sources.keys
        Dim sd As Object: Set sd = m_sources(sk)
        Dim si As Long
        For si = LBound(srcCols) To UBound(srcCols)
            If si > LBound(srcCols) Then buf = buf & ","
            buf = buf & CsvEsc(DictStr(sd, CStr(srcCols(si))))
        Next si
        buf = buf & vbCrLf
    Next sk

    ' [fields]
    buf = buf & vbCrLf & "[fields]" & vbCrLf & Join(fldCols, ",") & vbCrLf
    Dim fk As Variant
    For Each fk In m_fields.keys
        Dim fd As Object: Set fd = m_fields(fk)
        If Len(DictStr(fd, "source_name")) = 0 Or Len(DictStr(fd, "field_name")) = 0 Then GoTo NextField
        Dim fi As Long
        For fi = LBound(fldCols) To UBound(fldCols)
            If fi > LBound(fldCols) Then buf = buf & ","
            buf = buf & CsvEsc(DictStr(fd, CStr(fldCols(fi))))
        Next fi
        buf = buf & vbCrLf
NextField:
    Next fk

    WriteTextFile filePath, buf
    ExportSettings = True
    Exit Function
ErrOut:
    ExportSettings = False
End Function

Public Function ImportSettings(filePath As String) As Boolean
    On Error GoTo ErrOut
    Dim txt As String: txt = ReadTextFile(filePath)
    If Len(txt) = 0 Then ImportSettings = False: Exit Function
    If Not m_loaded Then EnsureConfigSheets

    Dim lines() As String: lines = Split(Replace$(txt, vbCr, ""), vbLf)
    Dim section As String
    Dim headers() As String
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim ln As String: ln = Trim$(lines(i))
        If Len(ln) = 0 Then GoTo NextLine

        ' Section header
        If Left$(ln, 1) = "[" Then
            section = Mid$(ln, 2, Len(ln) - 2)
            ' Next non-empty line is the header row
            Dim h As Long
            For h = i + 1 To UBound(lines)
                If Len(Trim$(lines(h))) > 0 Then
                    headers = ParseCsvLine(lines(h))
                    i = h
                    Exit For
                End If
            Next h
            GoTo NextLine
        End If

        Dim cols() As String: cols = ParseCsvLine(ln)

        Select Case section
        Case "config"
            If UBound(cols) >= 1 Then m_cfg(cols(0)) = cols(1)

        Case "sources"
            Dim srcName As String: srcName = cols(0)
            If Len(srcName) = 0 Then GoTo NextLine
            If Not m_sources.Exists(srcName) Then Set m_sources(srcName) = NewDict()
            Dim tgtSrc As Object: Set tgtSrc = m_sources(srcName)
            Dim sc As Long
            For sc = 0 To UBound(cols)
                If sc <= UBound(headers) Then tgtSrc(headers(sc)) = cols(sc)
            Next sc

        Case "fields"
            If UBound(cols) < 1 Then GoTo NextLine
            Dim fSrc As String: fSrc = cols(0)
            Dim fName As String: fName = cols(1)
            If Len(fSrc) = 0 Or Len(fName) = 0 Then GoTo NextLine
            Dim fk2 As String: fk2 = LCase$(fSrc) & "|" & LCase$(fName)
            If Not m_fields.Exists(fk2) Then
                Dim newFd As Object: Set newFd = NewDict()
                newFd("source_name") = fSrc
                newFd("field_name") = fName
                Set m_fields(fk2) = newFd
            End If
            Dim tgtFd As Object: Set tgtFd = m_fields(fk2)
            Dim fc As Long
            For fc = 0 To UBound(cols)
                If fc <= UBound(headers) Then tgtFd(headers(fc)) = cols(fc)
            Next fc
        End Select
NextLine:
    Next i

    m_dirty = True
    SaveToSheets
    ImportSettings = True
    Exit Function
ErrOut:
    ImportSettings = False
End Function

Private Function CsvEsc(v As String) As String
    If InStr(v, ",") > 0 Or InStr(v, """") > 0 Or InStr(v, vbLf) > 0 Or InStr(v, vbCr) > 0 Then
        CsvEsc = """" & Replace$(v, """", """""") & """"
    Else
        CsvEsc = v
    End If
End Function

Private Function ParseCsvLine(ln As String) As String()
    Dim result() As String
    Dim cnt As Long: cnt = 0
    Dim pos As Long: pos = 1
    Dim inQuote As Boolean
    Dim cur As String
    Do While pos <= Len(ln)
        Dim ch As String: ch = Mid$(ln, pos, 1)
        If inQuote Then
            If ch = """" Then
                If pos < Len(ln) And Mid$(ln, pos + 1, 1) = """" Then
                    cur = cur & """": pos = pos + 1
                Else
                    inQuote = False
                End If
            Else
                cur = cur & ch
            End If
        Else
            If ch = """" Then
                inQuote = True
            ElseIf ch = "," Then
                ReDim Preserve result(cnt): result(cnt) = cur: cnt = cnt + 1: cur = ""
            Else
                cur = cur & ch
            End If
        End If
        pos = pos + 1
    Loop
    ReDim Preserve result(cnt): result(cnt) = cur
    ParseCsvLine = result
End Function

' ============================================================================
' Sheet Persistence (private)
' ============================================================================

Private Sub EnsureSheet(shName As String, headers As Variant)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = shName
        ws.Visible = xlSheetVeryHidden
    End If

    Dim existing As Object: Set existing = CreateObject("Scripting.Dictionary")
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 0
    Dim c As Long
    For c = 1 To lastCol
        If Len(CStr(ws.Cells(1, c).Value)) > 0 Then existing(CStr(ws.Cells(1, c).Value)) = c
    Next c

    Dim i As Long
    For i = 0 To UBound(headers)
        If Not existing.Exists(CStr(headers(i))) Then
            ws.Cells(1, lastCol + 1).Value = CStr(headers(i))
            lastCol = lastCol + 1
        End If
    Next i
End Sub

Private Sub SaveConfigSheet()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_CONFIG)
    ' Clear existing data
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then ws.Rows("2:" & lastRow).Delete
    If m_cfg.Count = 0 Then Exit Sub
    Dim keys As Variant: keys = m_cfg.keys
    Dim data() As Variant: ReDim data(1 To m_cfg.Count, 1 To 2)
    Dim i As Long
    For i = 0 To UBound(keys)
        data(i + 1, 1) = CStr(keys(i))
        data(i + 1, 2) = CStr(m_cfg(keys(i)))
    Next i
    ws.Cells(2, 1).Resize(m_cfg.Count, 2).Value = data
End Sub

Private Sub SaveSourcesSheet()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_SOURCES)
    Dim hdr As Object: Set hdr = ReadHeaderMap(ws)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then ws.Rows("2:" & lastRow).Delete
    If m_sources.Count = 0 Then Exit Sub
    Dim srcKeys As Variant: srcKeys = m_sources.keys
    Dim nCols As Long: nCols = hdr.Count
    Dim data() As Variant: ReDim data(1 To m_sources.Count, 1 To nCols)
    Dim i As Long
    For i = 0 To UBound(srcKeys)
        Dim sd As Object: Set sd = m_sources(srcKeys(i))
        Dim hk As Variant
        For Each hk In hdr.keys
            Dim c As Long: c = hdr(hk)
            If sd.Exists(CStr(hk)) Then data(i + 1, c) = CStr(sd(CStr(hk)))
        Next hk
        data(i + 1, 1) = CStr(srcKeys(i))
    Next i
    ws.Cells(2, 1).Resize(m_sources.Count, nCols).Value = data
End Sub

Private Sub SaveFieldsSheet()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_FIELDS)
    Dim hdr As Object: Set hdr = ReadHeaderMap(ws)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then ws.Rows("2:" & lastRow).Delete
    If m_fields.Count = 0 Then Exit Sub
    Dim fKeys As Variant: fKeys = m_fields.keys
    Dim nCols As Long: nCols = hdr.Count
    Dim data() As Variant: ReDim data(1 To m_fields.Count, 1 To nCols)
    Dim i As Long
    For i = 0 To UBound(fKeys)
        Dim fd As Object: Set fd = m_fields(fKeys(i))
        Dim hk As Variant
        For Each hk In hdr.keys
            Dim ci As Long: ci = hdr(hk)
            If fd.Exists(CStr(hk)) Then data(i + 1, ci) = CStr(fd(CStr(hk)))
        Next hk
    Next i
    ws.Cells(2, 1).Resize(m_fields.Count, nCols).Value = data
End Sub

' ##########################################################################
' # SECTION: Change Log (ex-CaseDeskChangeLog)
' ##########################################################################

Private Function GetLogTable() As ListObject
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    If ws Is Nothing Then Exit Function
    Set GetLogTable = ws.ListObjects(LOG_TABLE)
    On Error GoTo 0
End Function

Public Sub EnsureLogSheet()
    Dim eh As New ErrorHandler: eh.Enter "CaseDeskLib", "EnsureLogSheet"
    On Error GoTo ErrHandler
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(LOG_SHEET)
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = LOG_SHEET
        ws.Visible = xlSheetVeryHidden
    End If
    ' Ensure ListObject exists
    If ws.ListObjects.Count = 0 Then
        Dim headers As Variant: headers = Array("timestamp", "source", "key", "field", "old_value", "new_value", "origin")
        Dim c As Long
        For c = 0 To 6: ws.Cells(1, c + 1).Value = headers(c): Next c
        ws.ListObjects.Add(xlSrcRange, ws.Range("A1:G1"), , xlYes).Name = LOG_TABLE
    End If
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Public Sub AddLogEntry(src As String, key As String, field As String, _
                       oldVal As String, newVal As String, origin As String)
    Dim eh As New ErrorHandler: eh.Enter "CaseDeskLib", "AddLogEntry"
    On Error GoTo ErrHandler
    Dim tbl As ListObject: Set tbl = GetLogTable()
    If tbl Is Nothing Then EnsureLogSheet: Set tbl = GetLogTable()
    If tbl Is Nothing Then Exit Sub

    RotateIfNeeded tbl, 1
    Dim lr As ListRow: Set lr = tbl.ListRows.Add
    lr.Range(1, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    lr.Range(1, 2).Value = src
    lr.Range(1, 3).Value = key
    lr.Range(1, 4).Value = field
    lr.Range(1, 5).Value = oldVal
    lr.Range(1, 6).Value = newVal
    lr.Range(1, 7).Value = origin
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Public Sub AddLogEntries(entries As Collection)
    If entries Is Nothing Then Exit Sub
    If entries.Count = 0 Then Exit Sub
    On Error GoTo ErrHandler
    Dim tbl As ListObject: Set tbl = GetLogTable()
    If tbl Is Nothing Then EnsureLogSheet: Set tbl = GetLogTable()
    If tbl Is Nothing Then Exit Sub

    RotateIfNeeded tbl, entries.Count

    Dim n As Long: n = entries.Count
    Dim data() As Variant: ReDim data(1 To n, 1 To 7)
    Dim ts As String: ts = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Dim i As Long
    For i = 1 To n
        Dim e As Object: Set e = entries(i)
        data(i, 1) = ts
        data(i, 2) = DictStr(e, "type")
        data(i, 3) = DictStr(e, "id")
        Dim act As String: act = DictStr(e, "action")
        If act = "added" Then data(i, 4) = "+" & DictStr(e, "type") _
        Else data(i, 4) = "-" & DictStr(e, "type")
        data(i, 5) = ""
        data(i, 6) = DictStr(e, "description")
        data(i, 7) = "external"
    Next i

    ' Add rows and batch write
    Dim startRow As Long
    If tbl.DataBodyRange Is Nothing Then
        tbl.ListRows.Add
        startRow = 1
    Else
        startRow = tbl.ListRows.Count + 1
        tbl.ListRows.Add
    End If
    ' Add remaining rows
    For i = 2 To n: tbl.ListRows.Add: Next i
    tbl.DataBodyRange.Rows(startRow).Resize(n, 7).Value = data
    Exit Sub
ErrHandler:
End Sub

Private Sub RotateIfNeeded(tbl As ListObject, addCount As Long)
    On Error Resume Next
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    Dim total As Long: total = tbl.ListRows.Count + addCount
    If total <= MAX_LOG_ROWS Then Exit Sub
    Dim delCount As Long: delCount = total - MAX_LOG_ROWS
    Dim i As Long
    For i = 1 To delCount: tbl.ListRows(1).Delete: Next i
    On Error GoTo 0
End Sub

Public Function GetRecentEntries(Optional count As Long = 200) As Collection
    Dim eh As New ErrorHandler: eh.Enter "CaseDeskLib", "GetRecentEntries"
    On Error GoTo ErrHandler
    Set GetRecentEntries = New Collection
    Dim tbl As ListObject: Set tbl = GetLogTable()
    If tbl Is Nothing Then eh.OK: Exit Function
    If tbl.DataBodyRange Is Nothing Then eh.OK: Exit Function

    Dim rowCount As Long: rowCount = tbl.ListRows.Count
    If rowCount = 0 Then eh.OK: Exit Function
    Dim startIdx As Long: startIdx = rowCount - count + 1
    If startIdx < 1 Then startIdx = 1

    Dim r As Long
    For r = rowCount To startIdx Step -1
        Dim rng As Range: Set rng = tbl.ListRows(r).Range
        Dim entry As Object: Set entry = NewDict()
        entry.Add "ts", CStr(rng(1, 1).Value)
        entry.Add "src", CStr(rng(1, 2).Value)
        entry.Add "key", CStr(rng(1, 3).Value)
        entry.Add "field", CStr(rng(1, 4).Value)
        entry.Add "old", CStr(rng(1, 5).Value)
        entry.Add "new", CStr(rng(1, 6).Value)
        entry.Add "origin", CStr(rng(1, 7).Value)
        GetRecentEntries.Add entry
    Next r
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Public Sub ClearLog()
    Dim eh As New ErrorHandler: eh.Enter "CaseDeskLib", "ClearLog"
    On Error GoTo ErrHandler
    Dim tbl As ListObject: Set tbl = GetLogTable()
    If tbl Is Nothing Then Exit Sub
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Public Function FormatLogLine(entry As Object) As String
    On Error Resume Next
    Dim ts As String: ts = DictStr(entry, "ts")
    If IsDate(ts) Then ts = Format$(CDate(ts), "hh:nn:ss")

    Dim origin As String: origin = DictStr(entry, "origin")
    Dim key As String: key = DictStr(entry, "key")
    Dim nm As String: nm = DictStr(entry, "name")
    Dim field As String: field = DictStr(entry, "field")
    Dim oldV As String: oldV = DictStr(entry, "old")
    Dim newV As String: newV = DictStr(entry, "new")

    Dim change As String
    If Len(field) > 0 Then change = field & ": "
    If Len(oldV) > 0 Or Len(newV) > 0 Then change = change & oldV & " -> " & newV

    Dim id As String: id = key
    If Len(nm) > 0 And nm <> key Then id = id & " " & nm

    FormatLogLine = ts & "  " & origin & "  " & id & "  " & change
    On Error GoTo 0
End Function
