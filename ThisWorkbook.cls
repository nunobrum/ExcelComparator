VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Deactivate()
    Call Sheet2.ReReverse
End Sub

Private Sub Workbook_Open()
    Sheet1.Cells(1, 4) = "Running on Excel Version " & Application.Version & vbCr & vbLf & Application.OperatingSystem
    Sheet2.InitVars

End Sub

Private Sub Workbook_Activate()
    Call RefreshWorkbooks
End Sub

Private Sub Workbook_WindowResize(ByVal Wn As Window)

End Sub