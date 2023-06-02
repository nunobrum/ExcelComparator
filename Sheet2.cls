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

Sub Set_Globals()
    Dim filename As String
    Dim sh, sh1 As String
    Dim x As Integer
      
    If Len(ThisWorkbook.ActiveSheet.Cells(1, 2)) <> 0 Then
        x = InStr(1, ThisWorkbook.Sheets("Report").Cells(1, 2), "]")
        If x <> 0 Then
            filename = Mid(ThisWorkbook.Sheets("Report").Cells(1, 2), 2, x - 2)

            Set Ori_Workbook = GetWorkbook(filename)
            If Not Ori_Workbook Is Nothing Then
                ' Not showng the opened workbook
                'wWorkbook.Worksheets.Visible = True
                sh = Mid(ThisWorkbook.Sheets("Report").Cells(1, 2), x + 1, 1000)
                x = InStr(1, sh, "!")
                sh1 = Left(sh, x - 1)
                Ori_Start_Col = Range(Right(sh, Len(sh) - x)).Column
                Set Ori_Sheet = Ori_Workbook.Sheets(sh1)
            End If
        End If
    End If

    If Len(ThisWorkbook.ActiveSheet.Cells(2, 2)) <> 0 Then
        x = InStr(1, ThisWorkbook.Sheets("Report").Cells(2, 2), "]")
        If x <> 0 Then
            filename = Mid(ThisWorkbook.Sheets("Report").Cells(2, 2), 2, x - 2)

            Set Rev_Workbook = GetWorkbook(filename)
            If Not Rev_Workbook Is Nothing Then
                ' Not showng the opened workbook
                'wWorkbook.Worksheets.Visible = True
                sh = Mid(ThisWorkbook.Sheets("Report").Cells(2, 2), x + 1, 1000)
                x = InStr(1, sh, "!")
                sh1 = Left(sh, x - 1)
                Rev_Start_Col = Range(Right(sh, Len(sh) - x)).Column
                Set Rev_Sheet = Rev_Workbook.Sheets(sh1)
            End If
        End If
    End If
End Sub

Sub ReverseColors(Target As Range)
    Dim color_temp As Double
    With Target
        color_temp = .Font.Color
        .Font.Color = .Interior.Color
        .Interior.Color = color_temp
    End With
End Sub

Private Sub Worksheet_Activate()
    'Call Set_Globals
    Set Ori_Cell = Nothing
    Set Rev_Cell = Nothing
End Sub

Private Sub Worksheet_Deactivate()
    Call ReReverse
    Set Ori_Cell = Nothing
    Set Rev_Cell = Nothing
End Sub


Sub ReReverse()
    If Not Ori_Cell Is Nothing Then
        Call ReverseColors(Ori_Cell)
        Set Ori_Cell = Nothing
    End If
    
    If Not Rev_Cell Is Nothing Then
        Call ReverseColors(Rev_Cell)
        Set Rev_Cell = Nothing
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim OriRow As Double
    Dim RevRow As Double
    Dim LineString As String
    
    If Target.Row < 4 Or Target.Column = 1 Then
        Exit Sub
    End If
    
    Call Set_Globals
    
    
    LineString = ActiveSheet.Cells(Target.Row, 1).Value
    
    If Left(LineString, 4) = "Line" Then
        OriRow = Int(Right(LineString, Len(LineString) - 5))
        RevRow = OriRow
    Else
        If Left(LineString, 4) = "Ori:" Then
            x = InStr(4, LineString, "=Rev:")
            If x <> 0 Then
                OriRow = Int(Mid(LineString, 5, x - 5))
                RevRow = Int(Mid(LineString, x + 5, Len(LineString)))
            End If
        Else
            Exit Sub
        End If
    End If

    If Not Ori_Workbook Is Nothing Then
        Ori_Workbook.Activate
        Ori_Sheet.Activate
        'If the reverse was done, undo it
        If Not Ori_Cell Is Nothing Then
            Call ReverseColors(Ori_Cell)
        End If
        'And set a new reverse
        Set Ori_Cell = Ori_Sheet.Cells(OriRow, Target.Column - 2 + Ori_Start_Col)
        Ori_Cell.Select
        'Application.Wait (Now + TimeValue("00:00:00.2"))
        Call ReverseColors(Ori_Cell)
    End If

    If Not Rev_Workbook Is Nothing Then
        Rev_Workbook.Activate
        Rev_Sheet.Activate
        'If the reverse was done, undo it
        If Not Rev_Cell Is Nothing Then
            Call ReverseColors(Rev_Cell)
        End If
        ' And set a new reverse
        Set Rev_Cell = Rev_Sheet.Cells(RevRow, Target.Column - 2 + Rev_Start_Col)
        Rev_Cell.Select
        Call ReverseColors(Rev_Cell)
        'Application.Wait (Now + TimeValue("00:00:00.2"))
    End If
  
    'Application.Goto (Ori_Workbook.Sheets(Ori_Sheet.Name).Cells(OriRow, Target.Column))
    'Application.Goto (Rev_Workbook.Sheets(Rev_Sheet.Name).Cells(RevRow, Target.Column))
    
    ' This condition is to solve a bug of Excel, if an activation of the current WorkBook
    ' It goes to the first sheet
    If Not (Ori_Workbook Is Nothing And Rev_Workbook Is Nothing) Then
        ThisWorkbook.Activate
        Target.Parent.Activate
    End If
End Sub

