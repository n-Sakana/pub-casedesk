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
Private m_cmbKeyColumn As MSForms.ComboBox
Private m_cmbFileLink As MSForms.ComboBox
Private m_cmbMailLink As MSForms.ComboBox
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
    ' Widened for the Role column on the Fields tab (spec §5.3). Source tab
    ' still fits comfortably at this width.
    Me.Width = 560
    Me.Height = 440
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
    sy = sy + ROW_H + 14

    Lbl m_fraSource, "lblKC", 6, sy, lblW, "Key:"
    Set m_cmbKeyColumn = Cmb(m_fraSource, "cmbKC", inL, sy, inW)
    sy = sy + ROW_H

    Lbl m_fraSource, "lblFL", 6, sy, lblW, "File Link:"
    Set m_cmbFileLink = Cmb(m_fraSource, "cmbFL", inL, sy, inW)
    sy = sy + ROW_H

    Lbl m_fraSource, "lblML", 6, sy, lblW, "Mail Link:"
    Set m_cmbMailLink = Cmb(m_fraSource, "cmbML", inL, sy, inW)

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
    PopulateColumnCombos
    If Len(savedSource) > 0 Then
        SelectComboItem m_cmbKeyColumn, CaseDeskLib.GetSourceStr(savedSource, "key_column")
        SelectComboItem m_cmbFileLink, CaseDeskLib.GetSourceStr(savedSource, "folder_link_column")
        SelectComboItem m_cmbMailLink, CaseDeskLib.GetSourceStr(savedSource, "mail_link_column")
    End If
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

    ' 1. Tables (ListObjects) — primary option
    Dim tbl As ListObject
    For Each tbl In ws.ListObjects: m_cmbTable.AddItem tbl.Name: Next tbl

    ' 2. Named ranges on this sheet (both workbook-scope and sheet-scope)
    Dim nm As Name
    Dim nmRange As Range
    For Each nm In wb.Names
        Set nmRange = Nothing
        On Error Resume Next
        Set nmRange = nm.RefersToRange
        On Error GoTo 0
        If Not nmRange Is Nothing Then
            If nmRange.Worksheet.Name = ws.Name Then
                If nmRange.Rows.Count > 1 And nmRange.Columns.Count > 1 Then
                    ' Use local name if sheet-scoped (strip "Sheet1!" prefix)
                    Dim nmDisplay As String: nmDisplay = nm.Name
                    If InStr(nmDisplay, "!") > 0 Then nmDisplay = Mid$(nmDisplay, InStr(nmDisplay, "!") + 1)
                    ' Avoid duplicates with table names
                    If Not ComboContains(m_cmbTable, nmDisplay) Then
                        m_cmbTable.AddItem nmDisplay
                    End If
                End If
            End If
            Set nmRange = Nothing
        End If
    Next nm

    ' If we have tables or named ranges, keep dropdown-list mode
    ' Otherwise allow direct range input (freeform combo)
    If m_cmbTable.ListCount = 0 Then
        m_cmbTable.Style = fmStyleDropDownCombo
        ' Suggest UsedRange as starting point
        On Error Resume Next
        Dim ur As Range: Set ur = ws.UsedRange
        On Error GoTo 0
        If Not ur Is Nothing Then
            If ur.Rows.Count > 1 And ur.Columns.Count > 1 Then m_cmbTable.AddItem ur.Address(False, False)
        End If
    Else
        ' Even with tables/named ranges, allow direct entry as last resort
        m_cmbTable.Style = fmStyleDropDownCombo
    End If

    ' Restore saved source if it was a custom range not in the list
    Dim savedSource As String: savedSource = ""
    Dim sources As Collection: Set sources = CaseDeskLib.GetSourceNames()
    If sources.Count > 0 Then savedSource = CStr(sources(1))
    If Len(savedSource) > 0 Then
        Dim savedSheet As String: savedSheet = CaseDeskLib.GetSourceStr(savedSource, "source_sheet")
        If savedSheet = m_cmbSheet.Text Then
            If Not ComboContains(m_cmbTable, savedSource) Then
                ' Saved range not in list — add it so it persists
                m_cmbTable.AddItem savedSource
            End If
        End If
    End If
End Sub

Private Function ComboContains(cmb As MSForms.ComboBox, val As String) As Boolean
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If LCase$(cmb.List(i)) = LCase$(val) Then ComboContains = True: Exit Function
    Next i
End Function

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
        If Len(diffMsg) > 0 Then PromptColumnChanges diffMsg
    Else
        Dim ws As Worksheet: Set ws = wb.Worksheets(m_cmbSheet.Text)
        If Not ws Is Nothing Then
            Dim cols As Collection: Set cols = GetColumnsFromRange(ws, src)
            ' Range-based sources also need column-change prompts (spec §5.5).
            Dim rangeCols As Object: Set rangeCols = CreateObject("Scripting.Dictionary")
            Dim rci As Long
            For rci = 1 To cols.Count
                rangeCols(LCase$(CStr(cols(rci)))) = CStr(cols(rci))
            Next rci
            Dim rangeDiff As String: rangeDiff = CaseDeskLib.DetectColumnChangesFromMap(src, rangeCols)
            ' Delegate to CaseDeskLib for field init so role guess + type
            ' inference are consistent with the ListObject path.
            CaseDeskLib.InitFieldSettingsFromRange src, ws
            If Len(rangeDiff) > 0 Then PromptColumnChanges rangeDiff
        End If
    End If

    ' Column positions (proportional). Role column added for spec §5.3.
    Dim fw As Single: fw = m_fraFields.Width - 22
    Dim x1 As Single: x1 = 6                    ' Column name
    Dim x2 As Single: x2 = fw * 0.22            ' Display name
    Dim x3 As Single: x3 = fw * 0.44            ' Vis
    Dim x4 As Single: x4 = fw * 0.51            ' Edit
    Dim x5 As Single: x5 = fw * 0.58            ' Type
    Dim x6 As Single: x6 = fw * 0.78            ' Role
    Dim w2 As Single: w2 = x3 - x2 - 4
    Dim w5 As Single: w5 = x6 - x5 - 4
    Dim w6 As Single: w6 = fw - x6
    Dim y As Single: y = 18

    ' Header
    Dim lh As MSForms.Label
    Set lh = Lbl(m_fraFields, "hC", x1, 2, x2 - x1, "Column"): lh.Font.Bold = True
    Set lh = Lbl(m_fraFields, "hD", x2, 2, w2, "Display"): lh.Font.Bold = True
    Set lh = Lbl(m_fraFields, "hV", x3, 2, 28, "Vis"): lh.Font.Bold = True
    Set lh = Lbl(m_fraFields, "hE", x4, 2, 28, "Edit"): lh.Font.Bold = True
    Set lh = Lbl(m_fraFields, "hT", x5, 2, w5, "Type"): lh.Font.Bold = True
    Set lh = Lbl(m_fraFields, "hR", x6, 2, w6, "Role*"): lh.Font.Bold = True

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

        ' Role combo — populated from CaseDeskLib.GetRoleIds() so the enum lives
        ' in one place. Display form is "case_id — 案件ID" but we store the id.
        Dim cr As MSForms.ComboBox: Set cr = Cmb(m_fraFields, "r2_" & i, x6, y - 2, w6)
        cr.Height = 20: cr.Font.Size = 9
        PopulateRoleCombo cr
        Dim savedRole As String: savedRole = CaseDeskLib.GetFieldStr(src, fld, "role", "")
        SelectRoleById cr, savedRole

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

Private Sub PopulateColumnCombos()
    m_cmbKeyColumn.Clear: m_cmbFileLink.Clear: m_cmbMailLink.Clear
    Dim src As String: src = m_cmbTable.Text
    If Len(src) = 0 Then Exit Sub
    Dim fields As Collection: Set fields = CaseDeskLib.GetFieldNames(src)
    Dim i As Long
    For i = 1 To fields.Count
        Dim fn As String: fn = CStr(fields(i))
        m_cmbKeyColumn.AddItem fn
        m_cmbFileLink.AddItem fn
        m_cmbMailLink.AddItem fn
    Next i
End Sub

Private Sub ClearFieldRows()
    Set m_fieldRows = CreateObject("Scripting.Dictionary")
    Do While m_fraFields.Controls.Count > 0: m_fraFields.Controls.Remove 0: Loop
    m_fraFields.ScrollTop = 0
End Sub

Private Sub PopulateRoleCombo(cmb As MSForms.ComboBox)
    Dim ids As Collection: Set ids = CaseDeskLib.GetRoleIds()
    Dim i As Long
    For i = 1 To ids.Count
        Dim id As String: id = CStr(ids(i))
        Dim label As String
        If Len(id) = 0 Then
            label = "(none)"
        Else
            label = id & " — " & CaseDeskLib.GetRoleLabel(id)
        End If
        cmb.AddItem label
    Next i
    cmb.ListIndex = 0
End Sub

Private Function RoleIdFromComboText(text As String) As String
    ' UI stores "role_id — display_name"; pull back the id prefix.
    If Len(text) = 0 Then RoleIdFromComboText = "": Exit Function
    If text = "(none)" Then RoleIdFromComboText = "": Exit Function
    Dim sep As String: sep = " — "
    Dim p As Long: p = InStr(text, sep)
    If p > 0 Then
        RoleIdFromComboText = Left$(text, p - 1)
    Else
        RoleIdFromComboText = text
    End If
End Function

Private Sub SelectRoleById(cmb As MSForms.ComboBox, roleId As String)
    Dim target As String
    If Len(roleId) = 0 Then
        target = "(none)"
    Else
        target = roleId & " — " & CaseDeskLib.GetRoleLabel(roleId)
    End If
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = target Then cmb.ListIndex = i: Exit Sub
    Next i
    ' Fallback: match by id prefix only
    If Len(roleId) > 0 Then
        For i = 0 To cmb.ListCount - 1
            If Left$(cmb.List(i), Len(roleId)) = roleId Then cmb.ListIndex = i: Exit Sub
        Next i
    End If
    cmb.ListIndex = 0
End Sub

Private Sub SelectComboItem(cmb As MSForms.ComboBox, val As String)
    If Len(val) = 0 Then Exit Sub
    Dim dv As String: dv = CaseDeskLib.StripFieldPrefix(val)
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = dv Or cmb.List(i) = val Then cmb.ListIndex = i: Exit Sub
    Next i
End Sub

Private Sub PromptColumnChanges(diffMsg As String)
    ' Spec §5.5: differences must trigger a re-confirmation prompt. We surface
    ' the diff and pre-fill roles via guessing so the user only confirms the
    ' changed entries instead of remapping every column.
    MsgBox "Column changes detected — please review the Fields tab:" & vbCrLf & vbCrLf & diffMsg, _
           vbInformation, "CaseDesk"
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
    Dim rng As Range
    ' Try workbook-scope named range
    On Error Resume Next
    Set rng = ws.Parent.Names(rangeAddr).RefersToRange
    On Error GoTo 0
    ' Try sheet-scope named range (Sheet1!MyRange)
    If rng Is Nothing Then
        On Error Resume Next
        Set rng = ws.Parent.Names(ws.Name & "!" & rangeAddr).RefersToRange
        On Error GoTo 0
    End If
    ' Then try direct address
    If rng Is Nothing Then
        On Error Resume Next
        Set rng = ws.Range(rangeAddr)
        On Error GoTo 0
    End If
    If rng Is Nothing Then: On Error Resume Next: Set rng = ws.UsedRange: On Error GoTo 0
    If rng Is Nothing Then Exit Function
    If rng.Columns.Count < 1 Then Exit Function
    ' Guard against excessively large ranges
    If rng.Rows.Count > 200000 Or rng.Columns.Count > 1000 Then
        On Error Resume Next: Set rng = ws.UsedRange: On Error GoTo 0
        If rng Is Nothing Then Exit Function
    End If
    Dim c As Long
    For c = 1 To rng.Columns.Count
        Dim cell As Range: Set cell = rng.Cells(1, c)
        ' Handle merged cells: use the merge area's first cell value
        Dim v As Variant
        On Error Resume Next
        If cell.MergeCells Then
            v = cell.MergeArea.Cells(1, 1).Value
        Else
            v = cell.Value
        End If
        On Error GoTo 0
        If Not IsEmpty(v) Then
            If Len(CStr(v)) > 0 Then GetColumnsFromRange.Add CStr(v)
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
    PopulateColumnCombos
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

    Dim missing As String
    If m_cmbKeyColumn.ListIndex < 0 Then missing = missing & "  - Key Column" & vbCrLf
    If Len(missing) > 0 Then
        MsgBox "Required settings:" & vbCrLf & vbCrLf & missing, vbExclamation, "Settings": Exit Sub
    End If

    Dim src As String: src = m_cmbTable.Text

    ' Build a per-visible-field intended role map from current UI state.
    ' Validation below MUST use this (not stored m_fields values) because
    ' edits haven't been persisted yet — otherwise clearing a role in the
    ' grid would still pass validation against the stale stored role.
    Dim pendingRole As Object: Set pendingRole = CreateObject("Scripting.Dictionary")
    Dim fld1 As Variant
    For Each fld1 In m_fieldRows.keys
        Dim row1 As Object: Set row1 = m_fieldRows(fld1)
        pendingRole(CStr(fld1)) = RoleIdFromComboText(CStr(row1("role").Text))
    Next fld1

    ' Role uniqueness: check both pending visible-row assignments AND stored
    ' hidden-field assignments. Hidden fields carrying a role that the user
    ' also assigns to a visible row must be flagged, otherwise FindFieldWithRole
    ' becomes order-dependent after save.
    Dim roleSeen As Object: Set roleSeen = CreateObject("Scripting.Dictionary")
    Dim dupeRole As String
    Dim fld0 As Variant
    For Each fld0 In pendingRole.keys
        Dim roleId0 As String: roleId0 = CStr(pendingRole(fld0))
        If Len(roleId0) > 0 Then
            If roleSeen.Exists(roleId0) Then
                dupeRole = roleId0 & " (" & CaseDeskLib.GetRoleLabel(roleId0) & ")"
                Exit For
            End If
            roleSeen(roleId0) = CStr(fld0)
        End If
    Next fld0
    ' Also check: does a hidden field already carry a role that a visible row
    ' is claiming? That's a duplicate too (two owners after save).
    If Len(dupeRole) = 0 Then
        Dim rsKey As Variant
        For Each rsKey In roleSeen.keys
            Dim existingHidden As String
            existingHidden = CaseDeskLib.FindFieldWithRole(src, CStr(rsKey))
            ' FindFieldWithRole returns the first stored match. If that stored
            ' field is NOT in our pendingRole map (i.e., it's hidden), it's a
            ' conflict with the visible assignment.
            If Len(existingHidden) > 0 And Not pendingRole.Exists(existingHidden) Then
                dupeRole = CStr(rsKey) & " (" & CaseDeskLib.GetRoleLabel(CStr(rsKey)) & ")" & _
                           " — " & existingHidden & " already holds it"
                Exit For
            End If
        Next rsKey
    End If
    If Len(dupeRole) > 0 Then
        MsgBox "Role is assigned to two columns:" & vbCrLf & "  " & dupeRole & vbCrLf & vbCrLf & _
               "Each role must be unique. Set one to (none).", vbExclamation, "Settings"
        Exit Sub
    End If

    ' Required roles (case_id, title): at least one column must map to each.
    ' Check against pending visible assignments first; only fall back to the
    ' stored value IF that field is not represented in the visible grid
    ' (i.e. it's a hidden `__...` setting column). A visible row that had a
    ' role cleared must NOT be satisfied by its own stale stored value.
    Dim missingReq As String
    Dim requiredIds As Collection: Set requiredIds = CaseDeskLib.GetRequiredRoleIds()
    Dim ri As Long
    For ri = 1 To requiredIds.Count
        Dim reqId As String: reqId = CStr(requiredIds(ri))
        Dim satisfied As Boolean: satisfied = roleSeen.Exists(reqId)
        If Not satisfied Then
            Dim storedHolder As String
            storedHolder = CaseDeskLib.FindFieldWithRole(src, reqId)
            ' Only count a stored role holder if it's hidden (not in the
            ' visible grid). A visible row whose role got cleared in this
            ' session must NOT satisfy the requirement via its stale value.
            If Len(storedHolder) > 0 And Not pendingRole.Exists(storedHolder) Then
                satisfied = True
            End If
        End If
        If Not satisfied Then
            missingReq = missingReq & "  - " & reqId & " (" & CaseDeskLib.GetRoleLabel(reqId) & ")" & vbCrLf
        End If
    Next ri
    If Len(missingReq) > 0 Then
        MsgBox "Required roles are not assigned:" & vbCrLf & vbCrLf & missingReq & vbCrLf & _
               "Map each required role to a column on the Fields tab.", vbExclamation, "Settings"
        Exit Sub
    End If

    ' Remove old source entries that differ from current selection
    Dim oldSources As Collection: Set oldSources = CaseDeskLib.GetSourceNames()
    Dim os As Long
    For os = 1 To oldSources.Count
        If CStr(oldSources(os)) <> src Then CaseDeskLib.RemoveSource CStr(oldSources(os))
    Next os
    CaseDeskLib.EnsureSource src
    CaseDeskLib.SetStr "mail_folder", m_txtMailFolder.Text
    CaseDeskLib.SetStr "case_folder_root", m_txtCaseFolder.Text
    CaseDeskLib.SetSourceStr src, "source_sheet", m_cmbSheet.Text
    CaseDeskLib.SetSourceStr src, "mail_match_mode", m_cmbMailMatchMode.Text
    CaseDeskLib.SetSourceStr src, "key_column", m_cmbKeyColumn.Text
    CaseDeskLib.SetSourceStr src, "folder_link_column", m_cmbFileLink.Text
    CaseDeskLib.SetSourceStr src, "mail_link_column", m_cmbMailLink.Text

    Dim fld As Variant
    For Each fld In m_fieldRows.keys
        Dim row As Object: Set row = m_fieldRows(fld)
        Dim ft As String: ft = CStr(row("type").Text)
        Dim roleId As String: roleId = RoleIdFromComboText(CStr(row("role").Text))
        CaseDeskLib.SetFieldStr src, CStr(fld), "display_name", Trim$(CStr(row("display").Text))
        CaseDeskLib.SetFieldBool src, CStr(fld), "visible", CBool(row("visible").Value)
        CaseDeskLib.SetFieldBool src, CStr(fld), "in_list", CBool(row("visible").Value)
        CaseDeskLib.SetFieldBool src, CStr(fld), "editable", CBool(row("editable").Value)
        CaseDeskLib.SetFieldStr src, CStr(fld), "type", ft
        CaseDeskLib.SetFieldBool src, CStr(fld), "multiline", (ft = "multiline")
        CaseDeskLib.SetFieldStr src, CStr(fld), "role", roleId
        CaseDeskLib.SetFieldStr src, CStr(fld), "sort_order", CStr(row("order"))
    Next fld

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
