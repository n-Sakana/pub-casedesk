VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E0D-00AA006002F3} frmSettings
   Caption         =   "Settings"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_cmbSheet As MSForms.ComboBox
Private WithEvents m_cmbTable As MSForms.ComboBox
Private m_cmbMailMatchMode As MSForms.ComboBox
Private WithEvents m_cmdBrowseMail As MSForms.CommandButton
Private WithEvents m_cmdBrowseCase As MSForms.CommandButton
Private WithEvents m_cmdSave As MSForms.CommandButton
Private WithEvents m_cmdCancel As MSForms.CommandButton
Private WithEvents m_cmdExport As MSForms.CommandButton
Private WithEvents m_cmdImport As MSForms.CommandButton
Private WithEvents m_cmdTabSource As MSForms.CommandButton
Private WithEvents m_cmdTabFields As MSForms.CommandButton

Private m_lblDataWb As MSForms.Label
Private m_txtMailFolder As MSForms.TextBox
Private m_txtCaseFolder As MSForms.TextBox
Private m_fraSource As MSForms.Frame
Private m_fraFields As MSForms.Frame

Private m_suppressEvents As Boolean
Private m_fieldRows As Object

Private Const ROW_H As Single = 24
Private Const GRID_ROW_H As Single = 22

Private Sub UserForm_Initialize()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "UserForm_Initialize"
    On Error GoTo ErrHandler
    Me.Width = 640
    Me.Height = 520
    m_suppressEvents = True
    Set m_fieldRows = CreateObject("Scripting.Dictionary")
    BuildLayout
    LoadConfig
    m_suppressEvents = False
    eh.OK: Exit Sub
ErrHandler:
    eh.Catch
End Sub

Private Sub BuildLayout()
    Dim cw As Single: cw = Me.InsideWidth
    Dim ch As Single: ch = Me.InsideHeight
    Dim M As Single: M = 6
    Dim tabH As Single: tabH = 24
    Dim btnBarH As Single: btnBarH = 30
    Dim bodyTop As Single: bodyTop = tabH
    Dim bodyH As Single: bodyH = ch - tabH - btnBarH

    ' === Tab buttons ===
    Set m_cmdTabSource = Btn(Me, "tabSrc", 0, 0, 72, tabH, "Source")
    Set m_cmdTabFields = Btn(Me, "tabFld", 72, 0, 72, tabH, "Fields")
    StyleTab m_cmdTabSource, True
    StyleTab m_cmdTabFields, False

    ' === Source page (standard Frame with caption = works) ===
    Set m_fraSource = Me.Controls.Add("Forms.Frame.1", "fraSrc")
    With m_fraSource
        .Caption = " Source / Paths "
        .Font.Name = "Meiryo UI": .Font.Size = 9
        .Left = M: .Top = bodyTop: .Width = cw - M * 2: .Height = bodyH
    End With

    Dim sy As Single: sy = 14
    Dim lblW As Single: lblW = 64
    Dim inL As Single: inL = lblW + 8
    Dim inW As Single: inW = m_fraSource.Width - inL - 10
    Dim bw As Single: bw = 24
    Dim pathW As Single: pathW = inW - bw - 4

    Lbl m_fraSource, "lblWb", 6, sy, lblW, "Workbook:"
    Set m_lblDataWb = Lbl(m_fraSource, "lblWbV", inL, sy, inW, "")
    sy = sy + ROW_H - 4

    Lbl m_fraSource, "lblSh", 6, sy, lblW, "Sheet:"
    Set m_cmbSheet = Cmb(m_fraSource, "cmbSh", inL, sy, inW)
    sy = sy + ROW_H

    Lbl m_fraSource, "lblTb", 6, sy, lblW, "Table:"
    Set m_cmbTable = Cmb(m_fraSource, "cmbTb", inL, sy, inW)
    sy = sy + ROW_H + 14

    Lbl m_fraSource, "lblMF", 6, sy, lblW, "Mail:"
    Set m_txtMailFolder = Txt(m_fraSource, "txtMF", inL, sy, pathW)
    Set m_cmdBrowseMail = Btn(m_fraSource, "btnMF", inL + pathW + 2, sy, bw, 18, "...")
    sy = sy + ROW_H

    Lbl m_fraSource, "lblCF", 6, sy, lblW, "Cases:"
    Set m_txtCaseFolder = Txt(m_fraSource, "txtCF", inL, sy, pathW)
    Set m_cmdBrowseCase = Btn(m_fraSource, "btnCF", inL + pathW + 2, sy, bw, 18, "...")
    sy = sy + ROW_H + 6

    Lbl m_fraSource, "lblMM", 6, sy, lblW, "Match:"
    Set m_cmbMailMatchMode = Cmb(m_fraSource, "cmbMM", inL, sy, 80)
    m_cmbMailMatchMode.AddItem "exact"
    m_cmbMailMatchMode.AddItem "domain"
    m_cmbMailMatchMode.ListIndex = 0

    ' === Fields page (standard Frame with scroll) ===
    Set m_fraFields = Me.Controls.Add("Forms.Frame.1", "fraFld")
    With m_fraFields
        .Caption = " Fields "
        .Font.Name = "Meiryo UI": .Font.Size = 9
        .Left = M: .Top = bodyTop: .Width = cw - M * 2: .Height = bodyH
        .ScrollBars = fmScrollBarsVertical
        .KeepScrollBarsVisible = fmScrollBarsNone
        .Visible = False
    End With

    ' === Bottom buttons ===
    Dim by As Single: by = ch - btnBarH + 4
    Set m_cmdExport = Btn(Me, "cmdExp", M, by, 60, 22, "Export...")
    Set m_cmdImport = Btn(Me, "cmdImp", M + 66, by, 60, 22, "Import...")
    Set m_cmdSave = Btn(Me, "cmdSave", cw - 132, by, 60, 22, "Save")
    Set m_cmdCancel = Btn(Me, "cmdCancel", cw - 66, by, 60, 22, "Cancel")
End Sub

Private Sub StyleTab(b As MSForms.CommandButton, active As Boolean)
    If active Then
        b.Font.Bold = True: b.BackColor = &HFFFFFF
    Else
        b.Font.Bold = False: b.BackColor = &HE0E0E0
    End If
End Sub

Private Sub m_cmdTabSource_Click()
    m_fraSource.Visible = True: m_fraFields.Visible = False
    StyleTab m_cmdTabSource, True: StyleTab m_cmdTabFields, False
End Sub

Private Sub m_cmdTabFields_Click()
    m_fraSource.Visible = False: m_fraFields.Visible = True
    StyleTab m_cmdTabSource, False: StyleTab m_cmdTabFields, True
End Sub

' --- Control factories (short names for compact layout code) ---

Private Function Lbl(c As Object, nm As String, l As Single, t As Single, w As Single, cap As String) As MSForms.Label
    Set Lbl = c.Controls.Add("Forms.Label.1", nm)
    With Lbl
        .Left = l: .Top = t + 3: .Width = w: .Height = 16: .Caption = cap
        .Font.Name = "Meiryo UI": .Font.Size = 9
    End With
End Function

Private Function Txt(c As Object, nm As String, l As Single, t As Single, w As Single) As MSForms.TextBox
    Set Txt = c.Controls.Add("Forms.TextBox.1", nm)
    With Txt
        .Left = l: .Top = t: .Width = w: .Height = 20
        .Font.Name = "Meiryo UI": .Font.Size = 9
    End With
End Function

Private Function Cmb(c As Object, nm As String, l As Single, t As Single, w As Single) As MSForms.ComboBox
    Set Cmb = c.Controls.Add("Forms.ComboBox.1", nm)
    With Cmb
        .Left = l: .Top = t: .Width = w: .Height = 20
        .Style = fmStyleDropDownList
        .Font.Name = "Meiryo UI": .Font.Size = 9
    End With
End Function

Private Function Btn(c As Object, nm As String, l As Single, t As Single, w As Single, h As Single, cap As String) As MSForms.CommandButton
    Set Btn = c.Controls.Add("Forms.CommandButton.1", nm)
    With Btn
        .Left = l: .Top = t: .Width = w: .Height = h: .Caption = cap
        .Font.Name = "Meiryo UI": .Font.Size = 9
    End With
End Function

Private Function Chk(c As Object, nm As String, l As Single, t As Single) As MSForms.CheckBox
    Set Chk = c.Controls.Add("Forms.CheckBox.1", nm)
    With Chk
        .Left = l: .Top = t: .Width = 18: .Height = 16: .Caption = ""
        .BackStyle = fmBackStyleTransparent
    End With
End Function

' ============================================================================
' Config Load
' ============================================================================

Private Sub LoadConfig()
    m_suppressEvents = True

    If Not CaseDeskMain.g_dataWb Is Nothing Then
        m_lblDataWb.Caption = CaseDeskMain.g_dataWb.Name
    Else
        m_lblDataWb.Caption = "(no workbook)"
    End If

    m_txtMailFolder.Text = CaseDeskLib.GetStr("mail_folder")
    m_txtCaseFolder.Text = CaseDeskLib.GetStr("case_folder_root")

    PopulateSheets

    Dim savedSource As String: savedSource = ""
    Dim sources As Collection: Set sources = CaseDeskLib.GetSourceNames()
    If sources.Count > 0 Then savedSource = CStr(sources(1))

    Dim savedSheet As String
    If Len(savedSource) > 0 Then savedSheet = CaseDeskLib.GetSourceStr(savedSource, "source_sheet")
    If Len(savedSheet) = 0 Then savedSheet = GuessFirstSheetName()
    SelectComboItem m_cmbSheet, savedSheet

    LoadTablesForSelectedSheet
    If Len(savedSource) > 0 Then SelectComboItem m_cmbTable, savedSource
    If m_cmbTable.ListIndex < 0 And m_cmbTable.ListCount > 0 Then m_cmbTable.ListIndex = 0

    If Len(savedSource) = 0 Then savedSource = m_cmbTable.Text
    If Len(savedSource) > 0 Then
        SelectComboItem m_cmbMailMatchMode, CaseDeskLib.GetSourceStr(savedSource, "mail_match_mode", "exact")
    End If

    BuildFieldRows
    m_suppressEvents = False
End Sub

Private Function GuessFirstSheetName() As String
    Dim wb As Workbook: Set wb = CaseDeskMain.g_dataWb
    If wb Is Nothing Then Exit Function
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then GuessFirstSheetName = ws.Name: Exit Function
    Next ws
End Function

Private Sub PopulateSheets()
    m_cmbSheet.Clear
    Dim wb As Workbook: Set wb = CaseDeskMain.g_dataWb
    If wb Is Nothing Then Exit Sub
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then m_cmbSheet.AddItem ws.Name
    Next ws
End Sub

Private Sub LoadTablesForSelectedSheet()
    m_cmbTable.Clear
    m_cmbTable.Style = fmStyleDropDownList
    Dim wb As Workbook: Set wb = CaseDeskMain.g_dataWb
    If wb Is Nothing Then Exit Sub
    If m_cmbSheet.ListIndex < 0 Then Exit Sub
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(m_cmbSheet.Text)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Dim tbl As ListObject
    For Each tbl In ws.ListObjects: m_cmbTable.AddItem tbl.Name: Next tbl
    If m_cmbTable.ListCount = 0 Then
        m_cmbTable.Style = fmStyleDropDownCombo
        On Error Resume Next
        Dim ur As Range: Set ur = ws.UsedRange
        On Error GoTo 0
        If Not ur Is Nothing Then
            If ur.Rows.Count > 1 And ur.Columns.Count > 1 Then m_cmbTable.AddItem ur.Address(False, False)
        End If
    End If
End Sub

' ============================================================================
' Field Grid
' ============================================================================

Private Sub BuildFieldRows()
    On Error GoTo FieldRowsExit
    ClearFieldRows
    If Len(m_cmbTable.Text) = 0 Then Exit Sub

    Dim wb As Workbook: Set wb = CaseDeskMain.g_dataWb
    If wb Is Nothing Then Exit Sub
    Dim src As String: src = m_cmbTable.Text

    CaseDeskLib.EnsureSource src
    CaseDeskLib.SetSourceStr src, "source_sheet", m_cmbSheet.Text

    Dim tbl As ListObject: Set tbl = CaseDeskData.FindTable(wb, src)
    If Not tbl Is Nothing Then
        Dim diffMsg As String: diffMsg = CaseDeskLib.DetectColumnChanges(src, tbl)
        CaseDeskLib.InitFieldSettingsFromTable src, tbl
        If Len(diffMsg) > 0 Then MsgBox "Column changes detected:" & vbCrLf & vbCrLf & diffMsg, vbInformation, "CaseDesk"
    Else
        Dim ws As Worksheet: Set ws = wb.Worksheets(m_cmbSheet.Text)
        If Not ws Is Nothing Then
            Dim cols As Collection: Set cols = GetColumnsFromRange(ws, src)
            InitFieldsFromColumns src, cols, ws
        End If
    End If

    SyncRolesFromSource src

    ' Column positions (proportional)
    Dim fw As Single: fw = m_fraFields.Width - 22
    Dim x1 As Single: x1 = 6                    ' Column name
    Dim x2 As Single: x2 = fw * 0.25            ' Display name
    Dim x3 As Single: x3 = fw * 0.50            ' Vis
    Dim x4 As Single: x4 = fw * 0.56            ' Edit
    Dim x5 As Single: x5 = fw * 0.62            ' Type
    Dim x6 As Single: x6 = fw * 0.78            ' Role
    Dim w2 As Single: w2 = x3 - x2 - 4
    Dim w5 As Single: w5 = x6 - x5 - 4
    Dim w6 As Single: w6 = fw - x6
    Dim y As Single: y = 18

    ' Header
    Dim lh As MSForms.Label
    Set lh = Lbl(m_fraFields, "hC", x1, 2, x2 - x1, "Column"): lh.Font.Bold = True
    Set lh = Lbl(m_fraFields, "hD", x2, 2, w2, "Display"): lh.Font.Bold = True
    Set lh = Lbl(m_fraFields, "hV", x3, 2, 36, "Vis"): lh.Font.Bold = True
    Set lh = Lbl(m_fraFields, "hE", x4, 2, 36, "Edit"): lh.Font.Bold = True
    Set lh = Lbl(m_fraFields, "hT", x5, 2, w5, "Type"): lh.Font.Bold = True
    Set lh = Lbl(m_fraFields, "hR", x6, 2, w6, "Role"): lh.Font.Bold = True

    Dim fields As Collection: Set fields = CaseDeskLib.GetFieldNames(src)
    Dim i As Long
    For i = 1 To fields.Count
        Dim fld As String: fld = CStr(fields(i))
        If CaseDeskLib.IsHiddenField(fld) Then GoTo NextField

        Dim lr As MSForms.Label
        Set lr = Lbl(m_fraFields, "r_" & i, x1, y, x2 - x1 - 2, fld)
        lr.Font.Size = 9

        Dim td As MSForms.TextBox: Set td = Txt(m_fraFields, "d_" & i, x2, y - 2, w2)
        td.Text = CaseDeskLib.GetFieldDisplayName(src, fld)
        td.Height = 20: td.Font.Size = 9

        Dim cv As MSForms.CheckBox: Set cv = Chk(m_fraFields, "v_" & i, x3, y)
        cv.Value = CaseDeskLib.GetFieldBool(src, fld, "visible", True)

        Dim ce As MSForms.CheckBox: Set ce = Chk(m_fraFields, "e_" & i, x4, y)
        ce.Value = CaseDeskLib.GetFieldBool(src, fld, "editable", True)
        If CaseDeskLib.IsReadOnlyField(fld) Then ce.Value = False

        Dim ct As MSForms.ComboBox: Set ct = Cmb(m_fraFields, "t_" & i, x5, y - 2, w5)
        ct.Height = 20: ct.Font.Size = 9
        ct.AddItem "text": ct.AddItem "multiline": ct.AddItem "number"
        ct.AddItem "currency": ct.AddItem "date": ct.AddItem "boolean"
        ct.AddItem "choice": ct.AddItem "path/url"
        Dim savedType As String: savedType = CaseDeskLib.GetFieldStr(src, fld, "type", "text")
        If CaseDeskLib.GetFieldBool(src, fld, "multiline") Then savedType = "multiline"
        SelectComboItem ct, savedType
        If ct.ListIndex < 0 Then ct.ListIndex = 0

        Dim cr As MSForms.ComboBox: Set cr = Cmb(m_fraFields, "rl_" & i, x6, y - 2, w6)
        cr.Height = 20: cr.Font.Size = 9
        cr.AddItem "": cr.AddItem "case_id": cr.AddItem "title": cr.AddItem "status"
        cr.AddItem "file_key": cr.AddItem "updated_at": cr.AddItem "mail_link"
        SelectComboItem cr, CaseDeskLib.GetFieldStr(src, fld, "role")
        If cr.ListIndex < 0 Then cr.ListIndex = 0

        Dim row As Object: Set row = CreateObject("Scripting.Dictionary")
        Set row("display") = td: Set row("visible") = cv: Set row("editable") = ce
        Set row("type") = ct: Set row("role") = cr: row("order") = CStr(i)
        Set m_fieldRows(fld) = row

        y = y + GRID_ROW_H
NextField:
    Next i

    m_fraFields.ScrollHeight = y + 4
FieldRowsExit:
    On Error GoTo 0
End Sub

Private Sub SyncRolesFromSource(src As String)
    Dim mapping(3, 1) As String
    mapping(0, 0) = "key_column":          mapping(0, 1) = "case_id"
    mapping(1, 0) = "display_name_column": mapping(1, 1) = "title"
    mapping(2, 0) = "mail_link_column":    mapping(2, 1) = "mail_link"
    mapping(3, 0) = "folder_link_column":  mapping(3, 1) = "file_key"
    Dim i As Long
    For i = 0 To 3
        Dim colName As String: colName = CaseDeskLib.GetSourceStr(src, mapping(i, 0))
        Dim roleName As String: roleName = mapping(i, 1)
        If Len(colName) = 0 Then GoTo NextMapping
        Dim alreadyAssigned As Boolean: alreadyAssigned = False
        Dim fields As Collection: Set fields = CaseDeskLib.GetFieldNames(src)
        Dim j As Long
        For j = 1 To fields.Count
            If CaseDeskLib.GetFieldStr(src, CStr(fields(j)), "role") = roleName Then alreadyAssigned = True: Exit For
        Next j
        If Not alreadyAssigned Then
            If Len(CaseDeskLib.GetFieldStr(src, colName, "field_name")) > 0 Or _
               Len(CaseDeskLib.GetFieldStr(src, colName, "type")) > 0 Then
                CaseDeskLib.SetFieldStr src, colName, "role", roleName
            End If
        End If
NextMapping:
    Next i
End Sub

Private Sub ClearFieldRows()
    Set m_fieldRows = CreateObject("Scripting.Dictionary")
    Do While m_fraFields.Controls.Count > 0: m_fraFields.Controls.Remove 0: Loop
    m_fraFields.ScrollTop = 0
End Sub

Private Sub SelectComboItem(cmb As MSForms.ComboBox, val As String)
    If Len(val) = 0 Then Exit Sub
    Dim dv As String: dv = CaseDeskLib.StripFieldPrefix(val)
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = dv Or cmb.List(i) = val Then cmb.ListIndex = i: Exit Sub
    Next i
End Sub

Private Sub InitFieldsFromColumns(src As String, cols As Collection, ws As Worksheet)
    Dim i As Long
    For i = 1 To cols.Count
        Dim cn As String: cn = CStr(cols(i))
        If CaseDeskLib.IsHiddenField(cn) Then GoTo NextInitCol
        CaseDeskLib.EnsureField src, cn
        If Len(CaseDeskLib.GetFieldStr(src, cn, "sort_order")) = 0 Or _
           CaseDeskLib.GetFieldStr(src, cn, "sort_order") = "0" Then
            CaseDeskLib.SetFieldStr src, cn, "sort_order", CStr(i)
        End If
NextInitCol:
    Next i
End Sub

Private Function GetColumnsFromRange(ws As Worksheet, rangeAddr As String) As Collection
    Set GetColumnsFromRange = New Collection
    On Error Resume Next
    Dim rng As Range: Set rng = ws.Range(rangeAddr)
    On Error GoTo 0
    If rng Is Nothing Then: On Error Resume Next: Set rng = ws.UsedRange: On Error GoTo 0
    If rng Is Nothing Then Exit Function
    If rng.Columns.Count < 1 Then Exit Function
    Dim hd As Variant
    If rng.Columns.Count = 1 Then
        ReDim hd(1 To 1, 1 To 1): hd(1, 1) = rng.Cells(1, 1).Value
    Else: hd = rng.Rows(1).Value: End If
    Dim c As Long
    For c = 1 To UBound(hd, 2)
        If Not IsEmpty(hd(1, c)) Then
            Dim v As String: v = CStr(hd(1, c))
            If Len(v) > 0 Then GetColumnsFromRange.Add v
        End If
    Next c
End Function

' ============================================================================
' Events
' ============================================================================

Private Sub m_cmbSheet_Change()
    If m_suppressEvents Then Exit Sub
    m_suppressEvents = True
    On Error GoTo Cleanup
    LoadTablesForSelectedSheet
    If m_cmbTable.ListCount > 0 Then m_cmbTable.ListIndex = 0
    BuildFieldRows
Cleanup: m_suppressEvents = False
End Sub

Private Sub m_cmbTable_Change()
    If m_suppressEvents Then Exit Sub
    m_suppressEvents = True
    On Error GoTo Cleanup
    BuildFieldRows
Cleanup: m_suppressEvents = False
End Sub

Private Sub m_cmdBrowseMail_Click()
    Dim p As String: p = BrowseFolder("Select Mail Archive folder")
    If Len(p) > 0 Then m_txtMailFolder.Text = p
End Sub

Private Sub m_cmdBrowseCase_Click()
    Dim p As String: p = BrowseFolder("Select Case Folder root")
    If Len(p) > 0 Then m_txtCaseFolder.Text = p
End Sub

Private Function BrowseFolder(title As String) As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = title: If .Show = -1 Then BrowseFolder = .SelectedItems(1)
    End With
End Function

Private Sub m_cmdSave_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "cmdSave_Click"
    On Error GoTo ErrHandler
    If Len(m_cmbTable.Text) = 0 Then MsgBox "Table selection is required.", vbExclamation, "Settings": Exit Sub

    Dim roleMap As Object: Set roleMap = CreateObject("Scripting.Dictionary")
    Dim fld As Variant
    For Each fld In m_fieldRows.Keys
        Dim rr As Object: Set rr = m_fieldRows(fld)
        Dim rn As String: rn = Trim$(CStr(rr("role").Text))
        If Len(rn) > 0 Then
            If roleMap.Exists(rn) Then
                MsgBox "Role """ & rn & """ is assigned to multiple columns.", vbExclamation, "Settings": Exit Sub
            End If
            roleMap(rn) = CStr(fld)
        End If
    Next fld

    Dim missing As String
    If Not roleMap.Exists("case_id") Then missing = missing & "  - Case ID" & vbCrLf
    If Not roleMap.Exists("title") Then missing = missing & "  - Title" & vbCrLf
    If Len(missing) > 0 Then
        MsgBox "Required roles not assigned:" & vbCrLf & vbCrLf & missing, vbExclamation, "Settings": Exit Sub
    End If

    Dim warn As String
    If Not roleMap.Exists("status") Then warn = warn & "  - Status" & vbCrLf
    If Not roleMap.Exists("file_key") Then warn = warn & "  - File Key" & vbCrLf
    If Not roleMap.Exists("updated_at") Then warn = warn & "  - Updated Date" & vbCrLf
    If Len(warn) > 0 Then
        If MsgBox("Recommended roles not assigned:" & vbCrLf & vbCrLf & warn & vbCrLf & "Continue?", vbQuestion + vbYesNo, "Settings") = vbNo Then Exit Sub
    End If

    Dim src As String: src = m_cmbTable.Text
    CaseDeskLib.EnsureSource src
    CaseDeskLib.SetStr "mail_folder", m_txtMailFolder.Text
    CaseDeskLib.SetStr "case_folder_root", m_txtCaseFolder.Text
    CaseDeskLib.SetSourceStr src, "source_sheet", m_cmbSheet.Text
    CaseDeskLib.SetSourceStr src, "mail_match_mode", m_cmbMailMatchMode.Text

    For Each fld In m_fieldRows.Keys
        Dim row As Object: Set row = m_fieldRows(fld)
        Dim ft As String: ft = CStr(row("type").Text)
        Dim fr As String: fr = Trim$(CStr(row("role").Text))
        CaseDeskLib.SetFieldStr src, CStr(fld), "display_name", Trim$(CStr(row("display").Text))
        CaseDeskLib.SetFieldBool src, CStr(fld), "visible", CBool(row("visible").Value)
        CaseDeskLib.SetFieldBool src, CStr(fld), "in_list", CBool(row("visible").Value)
        CaseDeskLib.SetFieldBool src, CStr(fld), "editable", CBool(row("editable").Value)
        CaseDeskLib.SetFieldStr src, CStr(fld), "type", ft
        CaseDeskLib.SetFieldBool src, CStr(fld), "multiline", (ft = "multiline")
        CaseDeskLib.SetFieldStr src, CStr(fld), "sort_order", CStr(row("order"))
        CaseDeskLib.SetFieldStr src, CStr(fld), "role", fr
    Next fld

    If roleMap.Exists("case_id") Then CaseDeskLib.SetSourceStr src, "key_column", CStr(roleMap("case_id"))
    If roleMap.Exists("title") Then CaseDeskLib.SetSourceStr src, "display_name_column", CStr(roleMap("title"))
    If roleMap.Exists("mail_link") Then CaseDeskLib.SetSourceStr src, "mail_link_column", CStr(roleMap("mail_link"))
    If roleMap.Exists("file_key") Then CaseDeskLib.SetSourceStr src, "folder_link_column", CStr(roleMap("file_key"))
    If Not roleMap.Exists("case_id") Then CaseDeskLib.SetSourceStr src, "key_column", ""
    If Not roleMap.Exists("title") Then CaseDeskLib.SetSourceStr src, "display_name_column", ""
    If Not roleMap.Exists("mail_link") Then CaseDeskLib.SetSourceStr src, "mail_link_column", ""
    If Not roleMap.Exists("file_key") Then CaseDeskLib.SetSourceStr src, "folder_link_column", ""

    CaseDeskLib.SaveToSheets
    Unload Me
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_cmdExport_Click()
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Export Settings": .Filters.Clear: .Filters.Add "CSV", "*.csv"
        .AllowMultiSelect = False: .InitialFileName = "casedesk-settings.csv"
        If .Show = -1 Then
            If CaseDeskLib.ExportSettings(.SelectedItems(1)) Then
                MsgBox "Settings exported.", vbInformation, "Export"
            Else: MsgBox "Export failed.", vbExclamation, "Export": End If
        End If
    End With
End Sub

Private Sub m_cmdImport_Click()
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Import Settings": .Filters.Clear: .Filters.Add "CSV", "*.csv": .AllowMultiSelect = False
        If .Show = -1 Then
            If MsgBox("Import from:" & vbCrLf & .SelectedItems(1) & vbCrLf & vbCrLf & "Overwrite current settings?", vbQuestion + vbYesNo, "Import") = vbYes Then
                If CaseDeskLib.ImportSettings(.SelectedItems(1)) Then
                    MsgBox "Imported. Reloading...", vbInformation, "Import": LoadConfig
                Else: MsgBox "Import failed.", vbExclamation, "Import": End If
            End If
        End If
    End With
End Sub

Private Sub m_cmdCancel_Click()
    CaseDeskLib.LoadFromSheets: Unload Me
End Sub
