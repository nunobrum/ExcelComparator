VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Dim Ori_Workbook As Workbook
Dim Rev_Workbook As Workbook

Dim Ori_Sheet As Worksheet
Dim Rev_Sheet As Worksheet

Dim Ori_Start_Col As Double
Dim Rev_Start_Col As Double

Dim Ori_Cell As Range
Dim Rev_Cell As Range

'Temp Values used for Reverse
Dim RV_Ori_Cell As Range
Dim RV_Rev_Cell As Range
Dim Ori_LeftStyle, Ori_TopStyle, Ori_RightStyle, Ori_BottomStyle As XlLineStyle
Dim Rev_LeftStyle, Rev_TopStyle, Rev_RightStyle, Rev_BottomStyle As XlLineStyle
Dim Ori_LeftWeight, Ori_TopWeight, Ori_RightWeight, Ori_BottomWeight As XlBorderWeight
Dim Rev_LeftWeight, Rev_TopWeight, Rev_RightWeight, Rev_BottomWeight As XlBorderWeight
'Dim Ori_Color, Rev_Color As Long
'Dim Ori_BckGr, Rev_BckGr As Long
'Dim Ori_Pattern, Rev_Pattern As Long



'Public Variables to control Navigation and SheetUpdate
Dim bAllowUpdates As Boolean


Sub InitVars()
    Set Ori_Workbook = Nothing
    Set Ori_Sheet = Nothing
    Set Ori_Cell = Nothing
    
    Set Rev_Workbook = Nothing
    Set Rev_Sheet = Nothing
    Set Rev_Cell = Nothing

    bAllowUpdates = True
    
End Sub

Sub ResetSheet()
    Dim Log_Sheet As Worksheet
    Set Log_Sheet = ThisWorkbook.Sheets("Log")
    'Log_Sheet.Rows(1).RowHeight = 41
    Log_Sheet.AutoFilterMode = False
    Log_Sheet.Cells.Clear
    Log_Sheet.Cells(logRowHeader, logColSyncOriSheet) = "Original Sheet"
    Log_Sheet.Cells(logRowHeader, logColSyncOriRow) = "O. Row"
    Log_Sheet.Cells(logRowHeader, logColSyncOriCol) = "O. Col"
    Log_Sheet.Cells(logRowHeader, logColSyncRevSheet) = "Revisioned Sheet"
    Log_Sheet.Cells(logRowHeader, logColSyncRevRow) = "R. Row"
    Log_Sheet.Cells(logRowHeader, logColSyncRevCol) = "R. Col"
    Log_Sheet.Cells(logRowOptions, logColSyncNavigation) = "Sync Navigation"
    Call set_YES_NO(Log_Sheet.Cells(logRowOptions, logColSyncNavigation + 1), bbSyncNavigation)
    Log_Sheet.Cells(logRowOptions, logColUpdateSheets) = "Update Sheets"
    Call set_YES_NO(Log_Sheet.Cells(logRowOptions, logColUpdateSheets + 1), bbUpdateSheets)
    
    Call SetValidation(ThisWorkbook.Sheets("Log").Cells(logRowOptions, logColUpdateSheets + 1), _
    "YES,NO", "", "Select YES if updates in this LOG should be written back to compared sheets", False)

    Call SetValidation(ThisWorkbook.Sheets("Log").Cells(logRowOptions, logColSyncNavigation + 1), _
    "YES,NO", "", "Select YES if navigation in this log should be visually indicated in the compared sheets.", False)
    
End Sub

Sub BlockUpdates()
  bAllowUpdates = False
End Sub



' This function sets the global variables associated with the Original Sheet
Sub SetOriCell(tr As Long)
    Dim aux As String
    Dim Row, col As Long
    
    On Error GoTo ERROR
    
    If Ori_Workbook Is Nothing Then
        Set Ori_Workbook = GetWorkbook( _
          ThisWorkbook.Sheets("Log").Cells(logRowHeader, logColSyncOriValue).Value)
    End If
    
    aux = ThisWorkbook.Sheets("Log").Cells(tr, logColSyncOriRow)
    If IsNumeric(aux) Then
        Row = Int(aux)
        col = ColNumber(ThisWorkbook.Sheets("Log").Cells(tr, logColSyncOriCol))
        If Not Ori_Workbook Is Nothing Then
            aux = ThisWorkbook.Sheets("Log").Cells(tr, logColSyncOriSheet)
            Set Ori_Sheet = Ori_Workbook.Sheets(aux)
            Set Ori_Cell = Ori_Sheet.Cells(Row, col)
        End If
        Exit Sub
    End If
ERROR:
    Set Ori_Cell = Nothing
End Sub

' This function sets the global variables associated with the Revision Sheet
Sub SetRevCell(tr As Long)
    Dim aux As String
    Dim Row, col As Long
    
    On Error GoTo ERROR
    If Rev_Workbook Is Nothing Then
        Set Rev_Workbook = GetWorkbook( _
          ThisWorkbook.Sheets("Log").Cells(logRowHeader, logColSyncRevValue).Value)
    End If
    
    aux = ThisWorkbook.Sheets("Log").Cells(tr, logColSyncRevRow)
    If IsNumeric(aux) Then
        Row = Int(aux)
        col = ColNumber(ThisWorkbook.Sheets("Log").Cells(tr, logColSyncRevCol))
        If Not Rev_Workbook Is Nothing Then
            aux = ThisWorkbook.Sheets("Log").Cells(tr, logColSyncRevSheet)
            Set Rev_Sheet = Rev_Workbook.Sheets(aux)
            Set Rev_Cell = Rev_Sheet.Cells(Row, col)
        End If
        Exit Sub
    End If
ERROR:
    Set Rev_Cell = Nothing
End Sub



Sub ReverseColors(Target As Range)
    Dim color_temp As Double
    With Target
        'color_temp = .Font.Color
        '.Font.Color = .Interior.Color
        .Interior.Color = vbYellow
        .Borders.Item(xlEdgeBottom).LineStyle = xlDashDot
        .Borders.Item(xlEdgeRight).LineStyle = xlDashDot
        .Borders.Item(xlEdgeTop).LineStyle = xlDashDot
        .Borders.Item(xlEdgeLeft).LineStyle = xlDashDot
        .Borders.Item(xlEdgeBottom).Weight = xlThick
        .Borders.Item(xlEdgeRight).Weight = xlThick
        .Borders.Item(xlEdgeTop).Weight = xlThick
        .Borders.Item(xlEdgeLeft).Weight = xlThick
        '''.BorderAround ColorIndex:=5, LineStyle:=xlSlantDashDot, Weight:=xlThick
    End With
End Sub

Sub ReReverse()
    On Error GoTo NOK_ORI

    If Not RV_Ori_Cell Is Nothing Then
        'RV_Ori_Cell.Font.Color = Ori_Color
        RV_Ori_Cell.Interior.Color = Ori_BckGr
        RV_Ori_Cell.Interior.Pattern = Ori_Pattern
        RV_Ori_Cell.Borders.Item(xlEdgeBottom).LineStyle = Ori_BottomStyle
        RV_Ori_Cell.Borders.Item(xlEdgeRight).LineStyle = Ori_RightStyle
        RV_Ori_Cell.Borders.Item(xlEdgeTop).LineStyle = Ori_TopStyle
        RV_Ori_Cell.Borders.Item(xlEdgeLeft).LineStyle = Ori_LeftStyle
        RV_Ori_Cell.Borders.Item(xlEdgeBottom).Weight = Ori_BottomWeight
        RV_Ori_Cell.Borders.Item(xlEdgeRight).Weight = Ori_RightWeight
        RV_Ori_Cell.Borders.Item(xlEdgeTop).Weight = Ori_TopWeight
        RV_Ori_Cell.Borders.Item(xlEdgeLeft).Weight = Ori_LeftWeight
        
    End If
NOK_ORI:
    Set RV_Ori_Cell = Nothing
    On Error GoTo NOK_REV
    If Not RV_Rev_Cell Is Nothing Then
        'RV_Rev_Cell.Font.Color = Rev_Color
        RV_Rev_Cell.Interior.Color = Rev_BckGr
        RV_Rev_Cell.Interior.Pattern = Rev_Pattern
        RV_Rev_Cell.Borders.Item(xlEdgeBottom).LineStyle = Rev_BottomStyle
        RV_Rev_Cell.Borders.Item(xlEdgeRight).LineStyle = Rev_RightStyle
        RV_Rev_Cell.Borders.Item(xlEdgeTop).LineStyle = Rev_TopStyle
        RV_Rev_Cell.Borders.Item(xlEdgeLeft).LineStyle = Rev_LeftStyle
        RV_Rev_Cell.Borders.Item(xlEdgeBottom).Weight = Rev_BottomWeight
        RV_Rev_Cell.Borders.Item(xlEdgeRight).Weight = Rev_RightWeight
        RV_Rev_Cell.Borders.Item(xlEdgeTop).Weight = Rev_TopWeight
        RV_Rev_Cell.Borders.Item(xlEdgeLeft).Weight = Rev_LeftWeight
        
    End If
NOK_REV:
    Set RV_Rev_Cell = Nothing
End Sub

Sub ZoomOriginal()
    If Ori_Cell Is Nothing Then Exit Sub
    Set RV_Ori_Cell = Ori_Cell
    Ori_Pattern = Ori_Cell.Interior.Pattern
    Ori_BckGr = Ori_Cell.Interior.Color
    'Ori_Color = Ori_Cell.Font.Color
    Ori_LeftStyle = Ori_Cell.Borders.Item(xlEdgeLeft).LineStyle
    Ori_TopStyle = Ori_Cell.Borders.Item(xlEdgeTop).LineStyle
    Ori_BottomStyle = Ori_Cell.Borders.Item(xlEdgeBottom).LineStyle
    Ori_RightStyle = Ori_Cell.Borders.Item(xlEdgeRight).LineStyle
    Ori_BottomWeight = RV_Ori_Cell.Borders.Item(xlEdgeBottom).Weight
    Ori_RightWeight = RV_Ori_Cell.Borders.Item(xlEdgeRight).Weight
    Ori_TopWeight = RV_Ori_Cell.Borders.Item(xlEdgeTop).Weight
    Ori_LeftWeight = RV_Ori_Cell.Borders.Item(xlEdgeLeft).Weight
    Call ReverseColors(Ori_Cell)
    'Call SelectOriCell
    'Application.Wait (Now + TimeValue("00:00:00.2"))
End Sub

Sub ZoomRevision()
    If Rev_Cell Is Nothing Then Exit Sub
    Set RV_Rev_Cell = Rev_Cell
    Rev_Pattern = Rev_Cell.Interior.Pattern
    Rev_BckGr = Rev_Cell.Interior.Color
    'Rev_Color = Rev_Cell.Font.Color
    Rev_LeftStyle = Rev_Cell.Borders.Item(xlEdgeLeft).LineStyle
    Rev_TopStyle = Rev_Cell.Borders.Item(xlEdgeTop).LineStyle
    Rev_BottomStyle = Rev_Cell.Borders.Item(xlEdgeBottom).LineStyle
    Rev_RightStyle = Rev_Cell.Borders.Item(xlEdgeRight).LineStyle
    Rev_BottomWeight = RV_Rev_Cell.Borders.Item(xlEdgeBottom).Weight
    Rev_RightWeight = RV_Rev_Cell.Borders.Item(xlEdgeRight).Weight
    Rev_TopWeight = RV_Rev_Cell.Borders.Item(xlEdgeTop).Weight
    Rev_LeftWeight = RV_Rev_Cell.Borders.Item(xlEdgeLeft).Weight
    Call ReverseColors(Rev_Cell)
    'Call SelectRevCell
    'Application.Wait (Now + TimeValue("00:00:00.2"))
End Sub

Private Sub Worksheet_Activate()
    'Call SetGlobals
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim sheetname As String
    Dim aux As String
    Dim TRow, Row, col As Long
    Dim Sheet As Worksheet
    Dim TCell As Range
      
    If bAllowUpdates And _
       ThisWorkbook.Sheets("Log").Cells(logRowOptions, logColUpdateSheets + 1) = "YES" Then
        If Target.Row > logRowHeaders Then
        ' Cycle through all the cells
            For Each TCell In Target
                If TCell.Column = logColSyncOriValue Then
                   Call SetOriCell(TCell.Row)
                    If Not Ori_Cell Is Nothing Then
                        Ori_Cell.Value = TCell.Value
                    End If
                End If
                If TCell.Column = logColSyncRevValue Then
                    Call SetRevCell(TCell.Row)
                    If Not Rev_Cell Is Nothing Then
                        Rev_Cell.Value = TCell.Value
                    End If
                End If
            Next TCell
        End If
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim Row, col As Double
    Dim sheetname As String
    Dim aux As String
    
    Call ReReverse ' In any case reverse any previous highlights
    If Target.Row <= logRowHeader Then
        Exit Sub
    End If
        
    If ThisWorkbook.Sheets("Log").Cells(logRowOptions, logColSyncNavigation + 1) = "YES" Then
        Call SetOriCell(Target.Row)
        Call SetRevCell(Target.Row)
            
        If Target.Column = logColSyncOriValue Then
            Call SelectOriCell
            ThisWorkbook.Activate
            ThisWorkbook.Sheets("Log").Activate
            Target.Select
            Call ZoomOriginal
        End If
        If Target.Column = logColSyncRevValue Then
            Call SelectRevCell
            ThisWorkbook.Activate
            ThisWorkbook.Sheets("Log").Activate
            Target.Select
            Call ZoomRevision
        End If
    End If
End Sub

Sub UpdateOriginal()
    If Not Ori_Cell Is Nothing And _
       Not Rev_Cell Is Nothing Then
        Ori_Cell.Value = Rev_Cell.Value
        Ori_Color = Rev_Color
        Ori_BckGr = Rev_BckGr
        Ori_Pattern = Rev_Pattern
    End If
End Sub

Sub UpdateRevision()
    If Not Ori_Cell Is Nothing And _
       Not Rev_Cell Is Nothing Then
       Rev_Cell.Value = Ori_Cell.Value
        Rev_Color = Ori_Color
        Rev_BckGr = Ori_BckGr
        Rev_Pattern = Ori_Pattern
    End If
End Sub


Sub SelectOriCell()
    Dim Ori_Cell_copy As Range
    Set Ori_Cell_copy = Ori_Cell
    'Call ReReverse
    If (Not Ori_Workbook Is Nothing) And (Not Ori_Cell Is Nothing) Then
        Ori_Workbook.Activate
        Ori_Sheet.Activate
        Ori_Cell_copy.Select
    End If
End Sub

Sub SelectRevCell()
    Dim Rev_Cell_copy As Range
    Set Rev_Cell_copy = Rev_Cell
    'Call ReReverse
    If (Not Rev_Workbook Is Nothing) And (Not Rev_Cell Is Nothing) Then
        Rev_Workbook.Activate
        Rev_Sheet.Activate
        Rev_Cell_copy.Select
    End If
End Sub




