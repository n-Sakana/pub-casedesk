VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Excel.Application

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    If Sh.Parent.FullName <> ThisWorkbook.FullName Then Exit Sub
    Dim sn As String: sn = Sh.Name
    If Left$(sn, Len("_casedesk")) <> "_casedesk" Then Exit Sub
    Application.ScreenUpdating = False
    If CaseDeskMain.g_formLoaded Then frmCaseDesk.OnCaseDeskSheetChange sn
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    On Error Resume Next
    If Wb.FullName = ThisWorkbook.FullName Then
        CaseDeskMain.BeforeWorkbookClose
    End If
    On Error GoTo 0
End Sub
