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
Private m_cmbKeyCol As MSForms.ComboBox
Private m_cmbNameCol As MSForms.ComboBox
Private m_cmbMailCol As MSForms.ComboBox
Private m_cmbFolderCol As MSForms.ComboBox
Private m_cmbMailMatchMode As MSForms.ComboBox
Private WithEvents m_cmdBrowseMail As MSForms.CommandButton
Private WithEvents m_cmdBrowseCase As MSForms.CommandButton
Private WithEvents m_cmdSave As MSForms.CommandButton
Private WithEvents m_cmdCancel As MSForms.CommandButton

Private m_lblDataWb As MSForms.Label
Private m_txtMailFolder As MSForms.TextBox
Private m_txtCaseFolder As MSForms.TextBox
Private m_fraFields As MSForms.Frame

Private m_suppressEvents As Boolean
Private m_colDisplayToRaw As Object
Private m_fieldRows As Object

Private Const M As Long = 12
Private Const LBL_W As Single = 84
Private Const ROW_H As Single = 24
Private Const GRID_ROW_H As Single = 22

Private Sub UserForm_Initialize()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "UserForm_Initialize"
    On Error GoTo ErrHandler
    Me.Width = 860
    Me.Height = 640
    Me.BackColor = &HFFFFFF
    m_suppressEvents = True
    Set m_colDisplayToRaw = CreateObject("Scripting.Dictionary")
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
    Dim inputL As Single: inputL = M + LBL_W + 6
    Dim btnW As Single: btnW = 32
    Dim btnGap As Single: btnGap = 4
    Dim fullW As Single: fullW = cw - inputL - M          ' full width (no button)
    Dim inputW As Single: inputW = cw - inputL - M - btnW - btnGap  ' with browse button
    Dim y As Single

    ' --- Measure bottom sections first to allocate Fields height ---
    Dim bottomH As Single: bottomH = 20 + ROW_H * 2 + 8 + 36 + M  ' Paths section + buttons

    y = M
    AddSection Me, "secSrc", M, y, "Import source"
    y = y + 20

    AddLabel Me, "lblWb", M, y, LBL_W, "Workbook:"
    Set m_lblDataWb = AddLabel(Me, "lblDataWbVal", inputL, y, fullW, "")
    m_lblDataWb.ForeColor = &H404040
    y = y + ROW_H

    AddLabel Me, "lblSheet", M, y, LBL_W, "Sheet:"
    Set m_cmbSheet = AddCombo(Me, "cmbSheet", inputL, y, fullW)
    y = y + ROW_H

    AddLabel Me, "lblTable", M, y, LBL_W, "Table:"
    Set m_cmbTable = AddCombo(Me, "cmbTable", inputL, y, fullW)
    y = y + ROW_H + 6

    AddSection Me, "secLink", M, y, "Roles"
    y = y + 20

    AddLabel Me, "lblKey", M, y, LBL_W, "Case ID:"
    Set m_cmbKeyCol = AddCombo(Me, "cmbKey", inputL, y, fullW)
    y = y + ROW_H

    AddLabel Me, "lblName", M, y, LBL_W, "Title:"
    Set m_cmbNameCol = AddCombo(Me, "cmbName", inputL, y, fullW)
    y = y + ROW_H

    AddLabel Me, "lblMailFld", M, y, LBL_W, "Mail field:"
    Set m_cmbMailCol = AddCombo(Me, "cmbMailFld", inputL, y, fullW)
    y = y + ROW_H

    AddLabel Me, "lblMailMatch", M, y, LBL_W, "Mail match:"
    Set m_cmbMailMatchMode = AddCombo(Me, "cmbMailMatch", inputL, y, fullW)
    m_cmbMailMatchMode.AddItem "exact"
    m_cmbMailMatchMode.AddItem "domain"
    m_cmbMailMatchMode.ListIndex = 0
    y = y + ROW_H

    AddLabel Me, "lblFolderFld", M, y, LBL_W, "File key:"
    Set m_cmbFolderCol = AddCombo(Me, "cmbFolderFld", inputL, y, fullW)
    y = y + ROW_H + 6

    AddSection Me, "secFields", M, y, "Fields"
    y = y + 20

    Set m_fraFields = Me.Controls.Add("Forms.Frame.1", "fraFields")
    With m_fraFields
        .Left = M
        .Top = y
        .Width = cw - M * 2
        .Height = ch - y - bottomH
        .Caption = ""
        .BorderStyle = fmBorderStyleSingle
        .ScrollBars = fmScrollBarsVertical
        .KeepScrollBarsVisible = fmScrollBarsNone
        .SpecialEffect = fmSpecialEffectFlat
        .BackColor = &HFFFFFF
    End With
    y = y + m_fraFields.Height + 8

    AddSection Me, "secPath", M, y, "Paths"
    y = y + 20

    Dim btnL As Single: btnL = inputL + inputW + btnGap
    AddLabel Me, "lblMailDir", M, y, LBL_W, "Mail folder:"
    Set m_txtMailFolder = AddTextBox(Me, "txtMailDir", inputL, y, inputW)
    Set m_cmdBrowseMail = AddBtn(Me, "cmdBrMail", btnL, y, btnW, 20, "...")
    y = y + ROW_H

    AddLabel Me, "lblCaseDir", M, y, LBL_W, "Case folder:"
    Set m_txtCaseFolder = AddTextBox(Me, "txtCaseDir", inputL, y, inputW)
    Set m_cmdBrowseCase = AddBtn(Me, "cmdBrCase", btnL, y, btnW, 20, "...")

    Set m_cmdSave = AddBtn(Me, "cmdSave", cw - 170, ch - 36, 75, 26, "Save")
    Set m_cmdCancel = AddBtn(Me, "cmdCancel", cw - 84, ch - 36, 75, 26, "Cancel")
End Sub

Private Function AddSection(container As Object, nm As String, l As Single, t As Single, cap As String) As MSForms.Label
    Set AddSection = container.Controls.Add("Forms.Label.1", nm)
    With AddSection
        .Left = l: .Top = t: .Width = 200: .Height = 16
        .Caption = cap
        .Font.Name = "Meiryo UI": .Font.Size = 9: .Font.Bold = True
        .ForeColor = &H404040
    End With
End Function

Private Function AddLabel(container As Object, nm As String, l As Single, t As Single, w As Single, cap As String) As MSForms.Label
    Set AddLabel = container.Controls.Add("Forms.Label.1", nm)
    With AddLabel
        .Left = l: .Top = t + 2: .Width = w: .Height = 14
        .Caption = cap
        .Font.Name = "Meiryo UI": .Font.Size = 9
    End With
End Function

Private Function AddTextBox(container As Object, nm As String, l As Single, t As Single, w As Single) As MSForms.TextBox
    Set AddTextBox = container.Controls.Add("Forms.TextBox.1", nm)
    With AddTextBox
        .Left = l: .Top = t: .Width = w: .Height = 20
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .Font.Name = "Meiryo": .Font.Size = 9
    End With
End Function

Private Function AddCombo(container As Object, nm As String, l As Single, t As Single, w As Single) As MSForms.ComboBox
    Set AddCombo = container.Controls.Add("Forms.ComboBox.1", nm)
    With AddCombo
        .Left = l: .Top = t: .Width = w: .Height = 20
        .Style = fmStyleDropDownList
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .Font.Name = "Meiryo UI": .Font.Size = 9
    End With
End Function

Private Function AddBtn(container As Object, nm As String, l As Single, t As Single, w As Single, h As Single, cap As String) As MSForms.CommandButton
    Set AddBtn = container.Controls.Add("Forms.CommandButton.1", nm)
    With AddBtn
        .Left = l: .Top = t: .Width = w: .Height = h: .Caption = cap
        .Font.Name = "Meiryo UI": .Font.Size = 9
    End With
End Function

Private Function AddCheck(container As Object, nm As String, l As Single, t As Single, cap As String) As MSForms.CheckBox
    Set AddCheck = container.Controls.Add("Forms.CheckBox.1", nm)
    With AddCheck
        .Left = l: .Top = t: .Width = 70: .Height = 16
        .Caption = cap
        .Font.Name = "Meiryo UI": .Font.Size = 8
        .BackStyle = fmBackStyleTransparent
    End With
End Function

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
    If Len(savedSource) > 0 Then
        SelectComboItem m_cmbTable, savedSource
    End If
    If m_cmbTable.ListIndex < 0 And m_cmbTable.ListCount > 0 Then m_cmbTable.ListIndex = 0

    LoadColumns
    If Len(savedSource) = 0 Then savedSource = m_cmbTable.Text
    If Len(savedSource) > 0 Then
        SelectComboItem m_cmbKeyCol, CaseDeskLib.GetSourceStr(savedSource, "key_column")
        SelectComboItem m_cmbNameCol, CaseDeskLib.GetSourceStr(savedSource, "display_name_column")
        SelectComboItem m_cmbMailCol, CaseDeskLib.GetSourceStr(savedSource, "mail_link_column")
        SelectComboItem m_cmbMailMatchMode, CaseDeskLib.GetSourceStr(savedSource, "mail_match_mode", "exact")
        SelectComboItem m_cmbFolderCol, CaseDeskLib.GetSourceStr(savedSource, "folder_link_column")
    End If

    BuildFieldRows
    m_suppressEvents = False
End Sub

Private Function GuessFirstSheetName() As String
    Dim wb As Workbook: Set wb = CaseDeskMain.g_dataWb
    If wb Is Nothing Then Exit Function
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            GuessFirstSheetName = ws.Name
            Exit Function
        End If
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
    For Each tbl In ws.ListObjects
        m_cmbTable.AddItem tbl.Name
    Next tbl
    ' If no tables, show UsedRange address as editable suggestion
    If m_cmbTable.ListCount = 0 Then
        m_cmbTable.Style = fmStyleDropDownCombo
        On Error Resume Next
        Dim ur As Range: Set ur = ws.UsedRange
        On Error GoTo 0
        If Not ur Is Nothing Then
            If ur.Rows.Count > 1 And ur.Columns.Count > 1 Then
                m_cmbTable.AddItem ur.Address(False, False)
            End If
        End If
    End If
End Sub

Private Sub LoadColumns()
    m_cmbKeyCol.Clear
    m_cmbNameCol.Clear
    m_cmbMailCol.Clear
    m_cmbFolderCol.Clear
    m_cmbMailMatchMode.ListIndex = 0
    Set m_colDisplayToRaw = CreateObject("Scripting.Dictionary")
    If Len(m_cmbTable.Text) = 0 Then Exit Sub

    Dim wb As Workbook: Set wb = CaseDeskMain.g_dataWb
    If wb Is Nothing Then Exit Sub

    Dim cols As Collection
    Dim tbl As ListObject: Set tbl = CaseDeskData.FindTable(wb, m_cmbTable.Text)
    If Not tbl Is Nothing Then
        Set cols = CaseDeskData.GetTableColumnNames(tbl)
    Else
        ' No table found — treat as range on selected sheet
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = wb.Worksheets(m_cmbSheet.Text)
        On Error GoTo 0
        If ws Is Nothing Then Exit Sub
        Set cols = GetColumnsFromRange(ws, m_cmbTable.Text)
    End If

    Dim c As Variant
    m_cmbKeyCol.AddItem "": m_cmbNameCol.AddItem ""
    m_cmbMailCol.AddItem "": m_cmbFolderCol.AddItem ""
    For Each c In cols
        Dim rawName As String: rawName = CStr(c)
        If CaseDeskLib.IsHiddenField(rawName) Then GoTo NextCol
        Dim dispName As String: dispName = CaseDeskLib.StripFieldPrefix(rawName)
        If m_colDisplayToRaw.Exists(dispName) Then dispName = rawName
        m_colDisplayToRaw(dispName) = rawName
        m_cmbKeyCol.AddItem dispName
        m_cmbNameCol.AddItem dispName
        m_cmbMailCol.AddItem dispName
        m_cmbFolderCol.AddItem dispName
NextCol:
    Next c
End Sub

Private Sub BuildFieldRows()
    On Error Resume Next
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
        If Len(diffMsg) > 0 Then
            MsgBox "Column changes detected:" & vbCrLf & vbCrLf & diffMsg, vbInformation, "CaseDesk"
        End If
    Else
        ' No table — init from range (user-specified or UsedRange)
        Dim ws As Worksheet
        Set ws = wb.Worksheets(m_cmbSheet.Text)
        If Not ws Is Nothing Then
            Dim cols As Collection: Set cols = GetColumnsFromRange(ws, src)
            InitFieldsFromColumns src, cols, ws
        End If
    End If

    Dim xRaw As Single: xRaw = 8
    Dim xDisp As Single: xDisp = 152
    Dim xVisible As Single: xVisible = 326
    Dim xEditable As Single: xEditable = 384
    Dim xType As Single: xType = 444
    Dim xRole As Single: xRole = 564
    Dim y As Single: y = 8

    AddHeader xRaw, xDisp, xVisible, xEditable, xType, xRole
    y = y + 18

    Dim fields As Collection: Set fields = CaseDeskLib.GetFieldNames(src)
    Dim i As Long
    For i = 1 To fields.Count
        Dim fld As String: fld = CStr(fields(i))
        If CaseDeskLib.IsHiddenField(fld) Then GoTo NextField

        Dim lblRaw As MSForms.Label
        Set lblRaw = AddLabel(m_fraFields, "lblRaw_" & CStr(i), xRaw, y, 140, fld)
        lblRaw.Font.Size = 8

        Dim txtDisp As MSForms.TextBox
        Set txtDisp = AddTextBox(m_fraFields, "txtDisp_" & CStr(i), xDisp, y - 1, 168)
        txtDisp.Text = CaseDeskLib.GetFieldDisplayName(src, fld)

        Dim chkVisible As MSForms.CheckBox
        Set chkVisible = AddCheck(m_fraFields, "chkVisible_" & CStr(i), xVisible, y + 1, "")
        chkVisible.Value = CaseDeskLib.GetFieldBool(src, fld, "visible", True)

        Dim chkEditable As MSForms.CheckBox
        Set chkEditable = AddCheck(m_fraFields, "chkEditable_" & CStr(i), xEditable, y + 1, "")
        chkEditable.Value = CaseDeskLib.GetFieldBool(src, fld, "editable", True)
        If CaseDeskLib.IsReadOnlyField(fld) Then chkEditable.Value = False

        Dim cmbType As MSForms.ComboBox
        Set cmbType = AddCombo(m_fraFields, "cmbType_" & CStr(i), xType, y - 1, 116)
        cmbType.AddItem "text"
        cmbType.AddItem "multiline"
        cmbType.AddItem "number"
        cmbType.AddItem "currency"
        cmbType.AddItem "date"
        cmbType.AddItem "boolean"
        cmbType.AddItem "choice"
        cmbType.AddItem "path/url"
        Dim savedType As String: savedType = CaseDeskLib.GetFieldStr(src, fld, "type", "text")
        If CaseDeskLib.GetFieldBool(src, fld, "multiline") Then savedType = "multiline"
        SelectComboItem cmbType, savedType
        If cmbType.ListIndex < 0 Then cmbType.ListIndex = 0

        Dim cmbRole As MSForms.ComboBox
        Set cmbRole = AddCombo(m_fraFields, "cmbRole_" & CStr(i), xRole, y - 1, 120)
        cmbRole.AddItem ""
        cmbRole.AddItem "case_id"
        cmbRole.AddItem "title"
        cmbRole.AddItem "status"
        cmbRole.AddItem "file_key"
        cmbRole.AddItem "updated_at"
        cmbRole.AddItem "mail_link"
        Dim savedRole As String: savedRole = CaseDeskLib.GetFieldStr(src, fld, "role")
        SelectComboItem cmbRole, savedRole
        If cmbRole.ListIndex < 0 Then cmbRole.ListIndex = 0

        Dim row As Object: Set row = CreateObject("Scripting.Dictionary")
        Set row("display") = txtDisp
        Set row("visible") = chkVisible
        Set row("editable") = chkEditable
        Set row("type") = cmbType
        Set row("role") = cmbRole
        row("order") = CStr(i)
        m_fieldRows(fld) = row

        y = y + GRID_ROW_H
NextField:
    Next i

    m_fraFields.ScrollHeight = y + 8

    ' Sync separate role ComboBoxes from field grid roles
    SyncRoleComboBoxes
    On Error GoTo 0
End Sub

Private Sub SyncRoleComboBoxes()
    Dim fld As Variant
    For Each fld In m_fieldRows.Keys
        Dim row As Object: Set row = m_fieldRows(fld)
        Dim roleName As String: roleName = CStr(row("role").Text)
        Dim rawName As String: rawName = CStr(fld)
        Dim dispName As String: dispName = CaseDeskLib.StripFieldPrefix(rawName)
        Select Case roleName
            Case "case_id":  SelectComboItem m_cmbKeyCol, dispName
            Case "title":    SelectComboItem m_cmbNameCol, dispName
            Case "mail_link": SelectComboItem m_cmbMailCol, dispName
            Case "file_key": SelectComboItem m_cmbFolderCol, dispName
        End Select
    Next fld
End Sub

Private Sub AddHeader(xRaw As Single, xDisp As Single, xVisible As Single, xEditable As Single, xType As Single, xRole As Single)
    Dim lbl As MSForms.Label
    Set lbl = AddLabel(m_fraFields, "hdrRaw", xRaw, 8, 140, "Column")
    lbl.Font.Bold = True: lbl.Font.Size = 8
    Set lbl = AddLabel(m_fraFields, "hdrDisp", xDisp, 8, 168, "Display name")
    lbl.Font.Bold = True: lbl.Font.Size = 8
    Set lbl = AddLabel(m_fraFields, "hdrVis", xVisible, 8, 54, "Visible")
    lbl.Font.Bold = True: lbl.Font.Size = 8
    Set lbl = AddLabel(m_fraFields, "hdrEdit", xEditable, 8, 56, "Editable")
    lbl.Font.Bold = True: lbl.Font.Size = 8
    Set lbl = AddLabel(m_fraFields, "hdrType", xType, 8, 116, "Data type")
    lbl.Font.Bold = True: lbl.Font.Size = 8
    Set lbl = AddLabel(m_fraFields, "hdrRole", xRole, 8, 120, "Role")
    lbl.Font.Bold = True: lbl.Font.Size = 8
End Sub

Private Sub ClearFieldRows()
    Set m_fieldRows = CreateObject("Scripting.Dictionary")
    Do While m_fraFields.Controls.Count > 0
        m_fraFields.Controls.Remove 0
    Loop
    m_fraFields.ScrollTop = 0
End Sub

Private Sub SelectComboItem(cmb As MSForms.ComboBox, val As String)
    If Len(val) = 0 Then Exit Sub
    Dim dispVal As String: dispVal = CaseDeskLib.StripFieldPrefix(val)
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = dispVal Or cmb.List(i) = val Then cmb.ListIndex = i: Exit Sub
    Next i
End Sub

Private Function ResolveRawColName(dispName As String) As String
    ResolveRawColName = dispName
    If m_colDisplayToRaw Is Nothing Then Exit Function
    If m_colDisplayToRaw.Exists(dispName) Then ResolveRawColName = CStr(m_colDisplayToRaw(dispName))
End Function

Private Sub InitFieldsFromColumns(src As String, cols As Collection, ws As Worksheet)
    Dim i As Long
    For i = 1 To cols.Count
        Dim colName As String: colName = CStr(cols(i))
        If CaseDeskLib.IsHiddenField(colName) Then GoTo NextInitCol
        CaseDeskLib.EnsureField src, colName
        If Len(CaseDeskLib.GetFieldStr(src, colName, "sort_order")) = 0 Or _
           CaseDeskLib.GetFieldStr(src, colName, "sort_order") = "0" Then
            CaseDeskLib.SetFieldStr src, colName, "sort_order", CStr(i)
        End If
NextInitCol:
    Next i
End Sub

Private Function GetColumnsFromRange(ws As Worksheet, rangeAddr As String) As Collection
    Set GetColumnsFromRange = New Collection
    On Error Resume Next
    Dim rng As Range: Set rng = ws.Range(rangeAddr)
    On Error GoTo 0
    If rng Is Nothing Then
        ' Invalid address — fall back to UsedRange
        On Error Resume Next
        Set rng = ws.UsedRange
        On Error GoTo 0
    End If
    If rng Is Nothing Then Exit Function
    If rng.Columns.Count < 1 Then Exit Function
    ' Read entire header row at once (avoids cell-by-cell errors)
    Dim headerData As Variant
    If rng.Columns.Count = 1 Then
        ReDim headerData(1 To 1, 1 To 1)
        headerData(1, 1) = rng.Cells(1, 1).Value
    Else
        headerData = rng.Rows(1).Value
    End If
    Dim c As Long
    For c = 1 To UBound(headerData, 2)
        If Not IsEmpty(headerData(1, c)) Then
            Dim v As String: v = CStr(headerData(1, c))
            If Len(v) > 0 Then GetColumnsFromRange.Add v
        End If
    Next c
End Function

Private Sub m_cmbSheet_Change()
    If m_suppressEvents Then Exit Sub
    m_suppressEvents = True
    On Error GoTo Cleanup
    LoadTablesForSelectedSheet
    If m_cmbTable.ListCount > 0 Then m_cmbTable.ListIndex = 0
    LoadColumns
    BuildFieldRows
Cleanup:
    m_suppressEvents = False
End Sub

Private Sub m_cmbTable_Change()
    If m_suppressEvents Then Exit Sub
    m_suppressEvents = True
    On Error GoTo Cleanup
    LoadColumns
    BuildFieldRows
Cleanup:
    m_suppressEvents = False
End Sub

Private Sub m_cmdBrowseMail_Click()
    Dim path As String: path = BrowseFolder("Select Mail Archive folder")
    If Len(path) > 0 Then m_txtMailFolder.Text = path
End Sub

Private Sub m_cmdBrowseCase_Click()
    Dim path As String: path = BrowseFolder("Select Case Folder root")
    If Len(path) > 0 Then m_txtCaseFolder.Text = path
End Sub

Private Function BrowseFolder(title As String) As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = title
        If .Show = -1 Then BrowseFolder = .SelectedItems(1)
    End With
End Function

Private Sub m_cmdSave_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "cmdSave_Click"
    On Error GoTo ErrHandler

    If m_cmbTable.ListIndex < 0 Then
        MsgBox "Table selection is required.", vbExclamation, "Settings"
        Exit Sub
    End If

    ' Collect roles from field grid and validate
    Dim roleMap As Object: Set roleMap = CreateObject("Scripting.Dictionary")
    Dim fld As Variant
    For Each fld In m_fieldRows.Keys
        Dim rr As Object: Set rr = m_fieldRows(fld)
        Dim roleName As String: roleName = Trim$(CStr(rr("role").Text))
        If Len(roleName) > 0 Then
            If roleMap.Exists(roleName) Then
                MsgBox "Role """ & roleName & """ is assigned to multiple columns." & vbCrLf & _
                       "Each role can only be assigned to one column.", vbExclamation, "Settings"
                Exit Sub
            End If
            roleMap(roleName) = CStr(fld)
        End If
    Next fld

    ' Validate required roles
    Dim missing As String
    If Not roleMap.Exists("case_id") Then missing = missing & "  - Case ID" & vbCrLf
    If Not roleMap.Exists("title") Then missing = missing & "  - Title" & vbCrLf
    If Len(missing) > 0 Then
        MsgBox "Required roles are not assigned:" & vbCrLf & vbCrLf & missing & vbCrLf & _
               "Please assign these roles in the Role column.", vbExclamation, "Settings"
        Exit Sub
    End If

    ' Warn about recommended roles
    Dim warn As String
    If Not roleMap.Exists("status") Then warn = warn & "  - Status" & vbCrLf
    If Not roleMap.Exists("file_key") Then warn = warn & "  - File Key" & vbCrLf
    If Not roleMap.Exists("updated_at") Then warn = warn & "  - Updated Date" & vbCrLf
    If Len(warn) > 0 Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("The following recommended roles are not assigned:" & vbCrLf & vbCrLf & _
                      warn & vbCrLf & "Continue saving?", vbQuestion + vbYesNo, "Settings")
        If ans = vbNo Then Exit Sub
    End If

    Dim src As String: src = m_cmbTable.Text
    CaseDeskLib.EnsureSource src
    CaseDeskLib.SetStr "mail_folder", m_txtMailFolder.Text
    CaseDeskLib.SetStr "case_folder_root", m_txtCaseFolder.Text
    CaseDeskLib.SetSourceStr src, "source_sheet", m_cmbSheet.Text
    CaseDeskLib.SetSourceStr src, "mail_match_mode", m_cmbMailMatchMode.Text

    ' Save field settings from grid (including roles)
    For Each fld In m_fieldRows.Keys
        Dim row As Object: Set row = m_fieldRows(fld)
        Dim fieldType As String: fieldType = CStr(row("type").Text)
        Dim fieldRole As String: fieldRole = Trim$(CStr(row("role").Text))
        CaseDeskLib.SetFieldStr src, CStr(fld), "display_name", Trim$(CStr(row("display").Text))
        CaseDeskLib.SetFieldBool src, CStr(fld), "visible", CBool(row("visible").Value)
        CaseDeskLib.SetFieldBool src, CStr(fld), "in_list", CBool(row("visible").Value)
        CaseDeskLib.SetFieldBool src, CStr(fld), "editable", CBool(row("editable").Value)
        CaseDeskLib.SetFieldStr src, CStr(fld), "type", fieldType
        CaseDeskLib.SetFieldBool src, CStr(fld), "multiline", (fieldType = "multiline")
        CaseDeskLib.SetFieldStr src, CStr(fld), "sort_order", CStr(row("order"))
        CaseDeskLib.SetFieldStr src, CStr(fld), "role", fieldRole
    Next fld

    ' Update source columns from role assignments
    If roleMap.Exists("case_id") Then CaseDeskLib.SetSourceStr src, "key_column", CStr(roleMap("case_id"))
    If roleMap.Exists("title") Then CaseDeskLib.SetSourceStr src, "display_name_column", CStr(roleMap("title"))
    If roleMap.Exists("mail_link") Then CaseDeskLib.SetSourceStr src, "mail_link_column", CStr(roleMap("mail_link"))
    If roleMap.Exists("file_key") Then CaseDeskLib.SetSourceStr src, "folder_link_column", CStr(roleMap("file_key"))

    Unload Me
    eh.OK: Exit Sub
ErrHandler:
    eh.Catch
End Sub

Private Sub m_cmdCancel_Click()
    Unload Me
End Sub
