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
Dim Ori_Color, Rev_Color As Long
Dim Ori_BckGr, Rev_BckGr As Long
Dim Ori_Pattern, Rev_Pattern As Long
Dim Ori_TaS, Rev_TaS As Long

Const use_borders = False

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
    Set Log_Sheet = ThisWorkbook.Sheets(sheetDiff)
    'Log_Sheet.Rows(1).RowHeight = 41
    Log_Sheet.AutoFilterMode = False
    Log_Sheet.Cells.Clear
    Log_Sheet.Cells(logRowHeader, logColSyncOriSheet) = ThisWorkbook.Sheets(sheetLanguage).Cells(rowOriginalSheet, colLanguage).Text
    Log_Sheet.Cells(logRowHeader, logColSyncOriRow) = "O. Row"
    Log_Sheet.Cells(logRowHeader, logColSyncOriCol) = "O. Col"
    Log_Sheet.Cells(logRowHeader, logColSyncRevSheet) = ThisWorkbook.Sheets(sheetLanguage).Cells(rowRevisedSheet, colLanguage).Text
    Log_Sheet.Cells(logRowHeader, logColSyncRevRow) = "R. Row"
    Log_Sheet.Cells(logRowHeader, logColSyncRevCol) = "R. Col"
    Log_Sheet.Cells(logRowSyncNavigation, logColSyncNavigation) = ThisWorkbook.Sheets(sheetLanguage).Cells(rowDiffZoom, colLanguage).Text
    Call set_YES_NO(Log_Sheet.Cells(logRowSyncNavigation, logColSyncNavigation + 1), True)
    Log_Sheet.Cells(logRowUpdateSheets, logColUpdateSheets) = ThisWorkbook.Sheets(sheetLanguage).Cells(rowDiffUpdate, colLanguage).Text
    Call set_YES_NO(Log_Sheet.Cells(logRowUpdateSheets, logColUpdateSheets + 1), True)
    
    Call SetValidation(ThisWorkbook.Sheets(sheetDiff).Cells(logRowUpdateSheets, logColUpdateSheets + 1), _
    OptionYes & "," & OptionNo, "", "", False)

    Call SetValidation(ThisWorkbook.Sheets(sheetDiff).Cells(logRowSyncNavigation, logColSyncNavigation + 1), _
    OptionYes & "," & OptionNo, "", "", False)
    
End Sub

Sub BlockUpdates()
  bAllowUpdates = False
End Sub



' This function sets the global variables associated with the Original Sheet
Sub SetOriCell(tr As Long)
    Dim aux As String
    Dim row, col As Long
    
    On Error GoTo ERROR
    
    If Ori_Workbook Is Nothing Then
        Set Ori_Workbook = GetWorkbook( _
          ThisWorkbook.Sheets(sheetDiff).Cells(logRowHeader, logColSyncOriValue).Value)
    End If
    
    aux = ThisWorkbook.Sheets(sheetDiff).Cells(tr, logColSyncOriRow)
    If IsNumeric(aux) Then
        row = Int(aux)
        col = ColNumber(ThisWorkbook.Sheets(sheetDiff).Cells(tr, logColSyncOriCol))
        If Not Ori_Workbook Is Nothing Then
            aux = ThisWorkbook.Sheets(sheetDiff).Cells(tr, logColSyncOriSheet)
            Set Ori_Sheet = Ori_Workbook.Sheets(aux)
            Set Ori_Cell = Ori_Sheet.Cells(row, col)
        End If
        Exit Sub
    End If
ERROR:
    Set Ori_Cell = Nothing
End Sub

' This function sets the global variables associated with the Revision Sheet
Sub SetRevCell(tr As Long)
    Dim aux As String
    Dim row, col As Long
    
    On Error GoTo ERROR
    If Rev_Workbook Is Nothing Then
        Set Rev_Workbook = GetWorkbook( _
          ThisWorkbook.Sheets(sheetDiff).Cells(logRowHeader, logColSyncRevValue).Value)
    End If
    
    aux = ThisWorkbook.Sheets(sheetDiff).Cells(tr, logColSyncRevRow)
    If IsNumeric(aux) Then
        row = Int(aux)
        col = ColNumber(ThisWorkbook.Sheets(sheetDiff).Cells(tr, logColSyncRevCol))
        If Not Rev_Workbook Is Nothing Then
            aux = ThisWorkbook.Sheets(sheetDiff).Cells(tr, logColSyncRevSheet)
            Set Rev_Sheet = Rev_Workbook.Sheets(aux)
            Set Rev_Cell = Rev_Sheet.Cells(row, col)
        End If
        Exit Sub
    End If
ERROR:
    Set Rev_Cell = Nothing
End Sub

Sub ReReverse()
    On Error GoTo NOK_ORI

    If Not RV_Ori_Cell Is Nothing Then
        If use_borders Then
            With RV_Ori_Cell.Borders
                .Item(xlEdgeBottom).LineStyle = Ori_BottomStyle
                .Item(xlEdgeRight).LineStyle = Ori_RightStyle
                .Item(xlEdgeTop).LineStyle = Ori_TopStyle
                .Item(xlEdgeLeft).LineStyle = Ori_LeftStyle
                .Item(xlEdgeBottom).Weight = Ori_BottomWeight
                .Item(xlEdgeRight).Weight = Ori_RightWeight
                .Item(xlEdgeTop).Weight = Ori_TopWeight
                .Item(xlEdgeLeft).Weight = Ori_LeftWeight
            End With
        Else
            RV_Ori_Cell.Font.Color = Ori_Color
            RV_Ori_Cell.Interior.Color = Ori_BckGr
            RV_Ori_Cell.Interior.Pattern = Ori_Pattern
            RV_Ori_Cell.Interior.TintAndShade = Ori_TaS
        End If
    End If
NOK_ORI:
    Set RV_Ori_Cell = Nothing
    On Error GoTo NOK_REV
    If Not RV_Rev_Cell Is Nothing Then
        If use_borders Then
            With RV_Rev_Cell.Borders
                .Item(xlEdgeBottom).LineStyle = Rev_BottomStyle
                .Item(xlEdgeRight).LineStyle = Rev_RightStyle
                .Item(xlEdgeTop).LineStyle = Rev_TopStyle
                .Item(xlEdgeLeft).LineStyle = Rev_LeftStyle
                .Item(xlEdgeBottom).Weight = Rev_BottomWeight
                .Item(xlEdgeRight).Weight = Rev_RightWeight
                .Item(xlEdgeTop).Weight = Rev_TopWeight
                .Item(xlEdgeLeft).Weight = Rev_LeftWeight
            End With
        Else
            RV_Rev_Cell.Font.Color = Rev_Color
            RV_Rev_Cell.Interior.Color = Rev_BckGr
            RV_Rev_Cell.Interior.Pattern = Rev_Pattern
            RV_Rev_Cell.Interior.TintAndShade = Rev_TaS
        End If
    End If
NOK_REV:
    Set RV_Rev_Cell = Nothing
End Sub

Sub HighlightOriginal()
    If Ori_Cell Is Nothing Then Exit Sub
    Set RV_Ori_Cell = Ori_Cell
    If use_borders Then
        With Ori_Cell.Borders
            Ori_LeftStyle = .Item(xlEdgeLeft).LineStyle
            Ori_TopStyle = .Item(xlEdgeTop).LineStyle
            Ori_BottomStyle = .Item(xlEdgeBottom).LineStyle
            Ori_RightStyle = .Item(xlEdgeRight).LineStyle
            Ori_BottomWeight = .Item(xlEdgeBottom).Weight
            Ori_RightWeight = .Item(xlEdgeRight).Weight
            Ori_TopWeight = .Item(xlEdgeTop).Weight
            Ori_LeftWeight = .Item(xlEdgeLeft).Weight
            .Item(xlEdgeBottom).LineStyle = xlDashDot
            .Item(xlEdgeRight).LineStyle = xlDashDot
            .Item(xlEdgeTop).LineStyle = xlDashDot
            .Item(xlEdgeLeft).LineStyle = xlDashDot
            .Item(xlEdgeBottom).Weight = xlThick
            .Item(xlEdgeRight).Weight = xlThick
            .Item(xlEdgeTop).Weight = xlThick
            .Item(xlEdgeLeft).Weight = xlThick
        End With
    Else
        Ori_Pattern = Ori_Cell.Interior.Pattern
        Ori_TaS = Ori_Cell.Interior.TintAndShade
        Ori_BckGr = Ori_Cell.Interior.Color
        Ori_Color = Ori_Cell.Font.Color
        If IsNull(Ori_Color) Then
            Ori_Color = 0
        End If
        Ori_Cell.Font.Color = Ori_BckGr
        Ori_Cell.Interior.Color = Ori_Color
    End If
    'Call BringToFrontOriCell
    'Application.Wait (Now + TimeValue("00:00:00.2"))
End Sub

Sub HighlightRevision()
    If Rev_Cell Is Nothing Then Exit Sub
    Set RV_Rev_Cell = Rev_Cell
    If use_borders Then
        With Rev_Cell.Borders
            Rev_LeftStyle = .Item(xlEdgeLeft).LineStyle
            Rev_TopStyle = .Item(xlEdgeTop).LineStyle
            Rev_BottomStyle = .Item(xlEdgeBottom).LineStyle
            Rev_RightStyle = .Item(xlEdgeRight).LineStyle
            Rev_BottomWeight = .Item(xlEdgeBottom).Weight
            Rev_RightWeight = .Item(xlEdgeRight).Weight
            Rev_TopWeight = .Item(xlEdgeTop).Weight
            Rev_LeftWeight = .Item(xlEdgeLeft).Weight
            .Item(xlEdgeBottom).LineStyle = xlDashDot
            .Item(xlEdgeRight).LineStyle = xlDashDot
            .Item(xlEdgeTop).LineStyle = xlDashDot
            .Item(xlEdgeLeft).LineStyle = xlDashDot
            .Item(xlEdgeBottom).Weight = xlThick
            .Item(xlEdgeRight).Weight = xlThick
            .Item(xlEdgeTop).Weight = xlThick
            .Item(xlEdgeLeft).Weight = xlThick
        End With
    Else
        Rev_Pattern = Rev_Cell.Interior.Pattern
        Rev_TaS = Rev_Cell.Interior.TintAndShade
        Rev_BckGr = Rev_Cell.Interior.Color
        Rev_Color = Rev_Cell.Font.Color
        If IsNull(Rev_Color) Then
            Rev_Color = 0
        End If
        Rev_Cell.Font.Color = Rev_BckGr
        Rev_Cell.Interior.Color = Rev_Color
    End If
    'Call BringToFrontRevCell
    'Application.Wait (Now + TimeValue("00:00:00.2"))
End Sub

Private Sub Worksheet_Activate()
    'Call SetGlobals
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal target As Range, Cancel As Boolean)
    If target.row <= logRowHeader Then
        Exit Sub
    End If
    Call ReReverse ' In any case reverse any previous highlights
        
    If target.Column = logColSyncOriValue Then
        Call SetOriCell(target.row)
        Call BringToFrontOriCell
        Ori_Cell.Select
    Else
        If target.Column = logColSyncRevValue Then
            Call SetRevCell(target.row)
            Call BringToFrontRevCell
            Rev_Cell.Select
        End If
    End If
End Sub

Private Sub Worksheet_Change(ByVal target As Range)
    Dim sheetname As String
    Dim aux, FormulaStr As String
    Dim TRow, row, col As Long
    Dim sheet As Worksheet
    Dim TCell As Range
    FormulaStr = ThisWorkbook.Sheets(sheetLanguage).Cells(rowOptionFormulas, colLanguage).Text
      
    If bAllowUpdates And _
       ThisWorkbook.Sheets(sheetDiff).Cells(logRowUpdateSheets, logColUpdateSheets + 1) = OptionYes Then
        If target.row > logRowHeaders Then
        ' Cycle through all the cells
            For Each TCell In target
                If TCell.Column = logColSyncOriValue Then
                    Call SetOriCell(TCell.row)
                    If Not Ori_Cell Is Nothing Then
                        If ThisWorkbook.Sheets("Config").Cells(cfgRowWhat, cfgColOriginal) = FormulasStr Then
                            Ori_Cell.Formula = StreamFormula(TCell.Formula)
                        Else
                            Ori_Cell = TCell
                        End If
                    End If
                End If
                If TCell.Column = logColSyncRevValue Then
                    Call SetRevCell(TCell.row)
                    If Not Rev_Cell Is Nothing Then
                        If ThisWorkbook.Sheets("Config").Cells(cfgRowWhat, cfgColOriginal) = FormulasStr Then
                            Rev_Cell = StreamFormula(TCell.Formula)
                        Else
                            Rev_Cell = TCell
                        End If
                    End If
                End If
            Next TCell
        End If
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal target As Range)
    
    Call ReReverse ' In any case reverse any previous highlights
    If target.row <= logRowHeader Then
        Exit Sub
    End If
        
    If ThisWorkbook.Sheets(sheetDiff).Cells(logRowSyncNavigation, logColSyncNavigation + 1) = OptionYes Then

        Call SetOriCell(target.row)
        Call SetRevCell(target.row)
            
        'If target.Column = logColSyncOriValue Then
            Call BringToFrontOriCell
            Call HighlightOriginal
            'ThisWorkbook.Activate
            'ThisWorkbook.Sheets(sheetDiff).Activate
            'target.Select
        'End If
        'If target.Column = logColSyncRevValue Then
            Call BringToFrontRevCell
            Call HighlightRevision
            ThisWorkbook.Activate
            ThisWorkbook.Sheets(sheetDiff).Activate
            'target.Select
        'End If
    End If
End Sub

Sub BringToFrontOriCell()
    If (Not Ori_Sheet Is Nothing) And (Not Ori_Cell Is Nothing) Then
        Ori_Workbook.Activate
        Ori_Sheet.Activate
        Ori_Cell.Show ' Show instead of select doesn't disrupt the Clipboard
    End If
End Sub

Sub BringToFrontRevCell()
    If (Not Rev_Workbook Is Nothing) And (Not Rev_Cell Is Nothing) Then
        Rev_Workbook.Activate
        Rev_Sheet.Activate
        Rev_Cell.Show ' Show instead of select doesn't disrupt the Clipboard
    End If
End Sub




