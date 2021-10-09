VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private WithEvents app As Excel.Application
Attribute app.VB_VarHelpID = -1


Private Sub Workbook_AddinInstall()
    Set app = Excel.Application
End Sub

Private Sub Workbook_AddinUninstall()
    Set e = Nothing
End Sub

Private Sub Workbook_Open()
    Set app = Excel.Application
End Sub


Private Sub app_WorkbookAfterSave(ByVal Wb As Workbook, ByVal Success As Boolean)
    'FIXME use Wb?
    Dim Property As DocumentProperty
    On Error Resume Next
    Set Property = ActiveWorkbook.CustomDocumentProperties("vba-git")
    If Err.Number = 0 Then
        On Error GoTo 0
        Project.ExportComponentsToSourceFolder
        Debug.Print "Exported " & ActiveWorkbook.Name
    End If
End Sub
