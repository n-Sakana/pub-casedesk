VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E0D-00AA006002F3} frmSettings
   Caption         =   "Settings"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6400
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' Controls
' ============================================================================
Private WithEvents m_cmbTable As MSForms.ComboBox
Private WithEvents m_cmbKeyCol As MSForms.ComboBox
Private WithEvents m_cmbNameCol As MSForms.ComboBox
Private WithEvents m_cmbMailCol As MSForms.ComboBox
Private WithEvents m_cmbFolderCol As MSForms.ComboBox
Private m_cmbMailMatchMode As MSForms.ComboBox
Private WithEvents m_cmdBrowseMail As MSForms.CommandButton
Private WithEvents m_cmdBrowseCase As MSForms.CommandButton
Private WithEvents m_cmdSave As MSForms.CommandButton
Private WithEvents m_cmdCancel As MSForms.CommandButton

Private m_lblDataWb As MSForms.Label
Private m_txtMailFolder As MSForms.TextBox
Private m_txtCaseFolder As MSForms.TextBox

' ============================================================================
' State
' ============================================================================
Private m_suppressEvents As Boolean

Private Const M As Long = 12
Private Const LBL_W As Single = 80
Private Const ROW_H As Single = 28

' ============================================================================
' Initialize
' ============================================================================

Private Sub UserForm_Initialize()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "UserForm_Initialize"
    On Error GoTo ErrHandler
    Me.Width = 440: Me.Height = 400
    m_suppressEvents = True
    BuildLayout
    LoadConfig
    m_suppressEvents = False
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Layout
' ============================================================================

Private Sub BuildLayout()
    Me.BackColor = &HFFFFFF
    Dim cw As Single: cw = Me.InsideWidth
    Dim ch As Single: ch = Me.InsideHeight

    Dim pw As Single: pw = cw - M * 2
    Dim inputL As Single: inputL = M + LBL_W + 4
    Dim inputW As Single: inputW = pw - inputL
    Dim y As Single

    y = M
    AddSection Me, "secSrc", M, y, "Source"
    y = y + 20

    ' Show data workbook name (read-only)
    AddLabel Me, "lblWb", M, y, LBL_W, "Workbook:"
    Set m_lblDataWb = AddLabel(Me, "lblDataWbVal", inputL, y, inputW, 14)
    m_lblDataWb.ForeColor = &H404040
    y = y + ROW_H

    AddLabel Me, "lblTable", M, y, LBL_W, "Table:"
    Set m_cmbTable = AddCombo(Me, "cmbTable", inputL, y, inputW)
    y = y + ROW_H

    AddLabel Me, "lblKey", M, y, LBL_W, "Key column:"
    Set m_cmbKeyCol = AddCombo(Me, "cmbKey", inputL, y, inputW)
    y = y + ROW_H

    AddLabel Me, "lblName", M, y, LBL_W, "Name column:"
    Set m_cmbNameCol = AddCombo(Me, "cmbName", inputL, y, inputW)
    y = y + ROW_H + 8

    AddSection Me, "secLink", M, y, "Link fields"
    y = y + 20

    AddLabel Me, "lblMailFld", M, y, LBL_W, "Mail field:"
    Set m_cmbMailCol = AddCombo(Me, "cmbMailFld", inputL, y, inputW)
    y = y + ROW_H

    AddLabel Me, "lblMailMatch", M, y, LBL_W, "Mail match:"
    Set m_cmbMailMatchMode = AddCombo(Me, "cmbMailMatch", inputL, y, inputW)
    m_cmbMailMatchMode.AddItem "exact"
    m_cmbMailMatchMode.AddItem "domain"
    m_cmbMailMatchMode.ListIndex = 0
    y = y + ROW_H

    AddLabel Me, "lblFolderFld", M, y, LBL_W, "Folder field:"
    Set m_cmbFolderCol = AddCombo(Me, "cmbFolderFld", inputL, y, inputW)
    y = y + ROW_H + 8

    AddSection Me, "secPath", M, y, "Paths"
    y = y + 20

    AddLabel Me, "lblMailDir", M, y, LBL_W, "Mail folder:"
    Set m_txtMailFolder = AddTextBox(Me, "txtMailDir", inputL, y, inputW - 36)
    Set m_cmdBrowseMail = AddBtn(Me, "cmdBrMail", cw - M - 32, y, 32, 20, "...")
    y = y + ROW_H

    AddLabel Me, "lblCaseDir", M, y, LBL_W, "Case folder:"
    Set m_txtCaseFolder = AddTextBox(Me, "txtCaseDir", inputL, y, inputW - 36)
    Set m_cmdBrowseCase = AddBtn(Me, "cmdBrCase", cw - M - 32, y, 32, 20, "...")

    ' --- Buttons ---
    Set m_cmdSave = AddBtn(Me, "cmdSave", cw - 170, ch - 36, 75, 26, "Save")
    Set m_cmdCancel = AddBtn(Me, "cmdCancel", cw - 84, ch - 36, 75, 26, "Cancel")
End Sub

' ============================================================================
' Factory helpers
' ============================================================================

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

' ============================================================================
' Config
' ============================================================================

Private Sub LoadConfig()
    m_suppressEvents = True

    ' Show data workbook name
    If Not CaseDeskMain.g_dataWb Is Nothing Then
        m_lblDataWb.Caption = CaseDeskMain.g_dataWb.Name
    Else
        m_lblDataWb.Caption = "(no workbook)"
    End If

    m_txtMailFolder.Text = CaseDeskLib.GetStr("mail_folder")
    m_txtCaseFolder.Text = CaseDeskLib.GetStr("case_folder_root")

    ' Load tables from data workbook
    LoadTables

    ' Restore selected source
    Dim sources As Collection: Set sources = CaseDeskLib.GetSourceNames()
    If sources.Count > 0 Then
        SelectComboItem m_cmbTable, CStr(sources(1))
        LoadColumns
        Dim src As String: src = CStr(sources(1))
        SelectComboItem m_cmbKeyCol, CaseDeskLib.GetSourceStr(src, "key_column")
        SelectComboItem m_cmbNameCol, CaseDeskLib.GetSourceStr(src, "display_name_column")
        SelectComboItem m_cmbMailCol, CaseDeskLib.GetSourceStr(src, "mail_link_column")
        SelectComboItem m_cmbMailMatchMode, CaseDeskLib.GetSourceStr(src, "mail_match_mode", "exact")
        SelectComboItem m_cmbFolderCol, CaseDeskLib.GetSourceStr(src, "folder_link_column")
    End If

    m_suppressEvents = False
End Sub

Private Sub LoadTables()
    m_cmbTable.Clear
    Dim wb As Workbook: Set wb = CaseDeskMain.g_dataWb
    If wb Is Nothing Then Exit Sub
    Dim names As Collection: Set names = CaseDeskData.GetWorkbookTableNames(wb)
    Dim n As Variant
    For Each n In names: m_cmbTable.AddItem CStr(n): Next n
End Sub

Private Sub LoadColumns()
    m_cmbKeyCol.Clear
    m_cmbNameCol.Clear
    m_cmbMailCol.Clear
    m_cmbFolderCol.Clear
    If m_cmbTable.ListIndex < 0 Then Exit Sub

    Dim wb As Workbook: Set wb = CaseDeskMain.g_dataWb
    If wb Is Nothing Then Exit Sub
    Dim tbl As ListObject: Set tbl = CaseDeskData.FindTable(wb, m_cmbTable.Text)
    If tbl Is Nothing Then Exit Sub

    Dim cols As Collection: Set cols = CaseDeskData.GetTableColumnNames(tbl)
    Dim c As Variant
    m_cmbKeyCol.AddItem "": m_cmbNameCol.AddItem ""
    m_cmbMailCol.AddItem "": m_cmbFolderCol.AddItem ""
    For Each c In cols
        m_cmbKeyCol.AddItem CStr(c)
        m_cmbNameCol.AddItem CStr(c)
        m_cmbMailCol.AddItem CStr(c)
        m_cmbFolderCol.AddItem CStr(c)
    Next c
End Sub

Private Sub SelectComboItem(cmb As MSForms.ComboBox, val As String)
    If Len(val) = 0 Then Exit Sub
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = val Then cmb.ListIndex = i: Exit Sub
    Next i
End Sub

' ============================================================================
' Events
' ============================================================================

Private Sub m_cmbTable_Change()
    If m_suppressEvents Then Exit Sub
    LoadColumns
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
        .title = title
        If .Show = -1 Then BrowseFolder = .SelectedItems(1)
    End With
End Function

Private Sub m_cmdSave_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "cmdSave_Click"
    On Error GoTo ErrHandler

    ' Validate required fields
    If m_cmbTable.ListIndex >= 0 Then
        If m_cmbKeyCol.ListIndex <= 0 Or m_cmbNameCol.ListIndex <= 0 Then
            MsgBox "Key column and Name column are required.", vbExclamation, "Settings"
            Exit Sub
        End If
    End If

    CaseDeskLib.SetStr "mail_folder", m_txtMailFolder.Text
    CaseDeskLib.SetStr "case_folder_root", m_txtCaseFolder.Text

    If m_cmbTable.ListIndex >= 0 Then
        Dim src As String: src = m_cmbTable.Text
        CaseDeskLib.EnsureSource src
        CaseDeskLib.SetSourceStr src, "key_column", m_cmbKeyCol.Text
        CaseDeskLib.SetSourceStr src, "display_name_column", m_cmbNameCol.Text
        If m_cmbMailCol.ListIndex > 0 Then CaseDeskLib.SetSourceStr src, "mail_link_column", m_cmbMailCol.Text
        CaseDeskLib.SetSourceStr src, "mail_match_mode", m_cmbMailMatchMode.Text
        If m_cmbFolderCol.ListIndex > 0 Then CaseDeskLib.SetSourceStr src, "folder_link_column", m_cmbFolderCol.Text

        ' Auto-detect field settings from table format
        Dim wb As Workbook: Set wb = CaseDeskMain.g_dataWb
        If Not wb Is Nothing Then
            Dim tbl As ListObject: Set tbl = CaseDeskData.FindTable(wb, src)
            If Not tbl Is Nothing Then CaseDeskLib.InitFieldSettingsFromTable src, tbl
        End If
    End If

    Unload Me
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_cmdCancel_Click()
    Unload Me
End Sub
