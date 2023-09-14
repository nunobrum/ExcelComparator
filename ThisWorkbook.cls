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
    Dim hyperlink As Variant
    Sheet1.Cells(1, 4) = "Sheet Comparator 2v7" & vbCr & vbLf & _
    "Running on Excel Version " & Application.Version & vbCr & vbLf & _
    Application.OperatingSystem
    Sheet1.Hyperlinks.Add _
                Anchor:=Sheet1.Range("D27"), _
                Address:="https://www.paypal.com/paypalme/send2NunoBrum", _
                ScreenTip:="Donations will serve to fund the development of more tools and additional features." & _
                "The remainer will be sent to help Ukranian refugees", _
                TextToDisplay:="Show your appreciation. Donate!"
                
    Sheet1.Range("D28").Font.Color = RGB(255, 217, 102)
    Sheet1.Range("D28").Font.Bold = True
         
    Sheet1.Hyperlinks.Add _
                Anchor:=Sheet1.Range("D28"), _
                Address:="mailto:me@nunobrum.com", _
                TextToDisplay:="Suggestions, Contact me at: me@nunobrum.com"
    Sheet1.Range("D29").Font.Color = RGB(255, 217, 102)
    
    Sheet1.Hyperlinks.Add _
            Anchor:=Sheet1.Range("D29"), _
            Address:="http://www.nunobrum.com/excelcomparator.html", _
            TextToDisplay:="© Nuno Brum 2020, http://www.nunobrum.com/excelcomparator.html"
    
    Sheet1.Range("D30").Font.Color = RGB(255, 217, 102)
    
    For Each hl In Sheet1.Hyperlinks
        If hl.Address = "mailto:me@nunobrum.com" Then
            hl.EmailSubject = "Suggestions%20to%20SheetComparator"
            Exit For
        End If
    Next hl
    
    Sheet2.InitVars
    'Call fillLanguagesCombo
    Call InitConfigStrings
    ' Application.Width = 1040
    ' Application.Height = 650
End Sub

Private Sub Workbook_Activate()
    Call RefreshWorkbooks
End Sub

Private Sub Workbook_WindowResize(ByVal Wn As Window)

End Sub
