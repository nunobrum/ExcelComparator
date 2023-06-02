VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim wWorkbook As Workbook
    Dim filename As String
    Dim cellRef As Range
    Dim WSheet As Worksheet
    Dim x, x1 As Integer

    If Target.Column = 2 Or Target.Column = 3 Then
        ' All the other columns are ignored.
        If Target.Row = 2 Then
            ' This is the filename that was changed
            filename = Target.Text
            'If filename is set by Range ignore
            If Not Left(ThisWorkbook.ActiveSheet.Cells(4, Target.Column).Formula, 1) = "=" Then
            ' Set the variable that will be passed as pointer
                Set cellRef = ThisWorkbook.ActiveSheet.Cells(3, Target.Column)
                'Open the workbook
                Set wWorkbook = GetWorkbook(filename)
                ' Not showng the opened workbook
                'wWorkbook.Worksheets.Visible = False
                If Not wWorkbook Is Nothing Then
                    ' Update the available sheets in the combobox in the validation
                    Call Reload_Sheets(wWorkbook, cellRef)
                End If
                Exit Sub
            End If
        End If
        If Target.Row = 3 Then
            ' If the sheet was set by the range ignore
            If Not Left(ThisWorkbook.ActiveSheet.Cells(4, Target.Column).Formula, 1) = "=" Then
                If ThisWorkbook.ActiveSheet.Cells(Target.Row, Target.Column).Text = "[ALL]" Then
                    ThisWorkbook.ActiveSheet.Cells(4, Target.Column) = "Auto Detect"
                Else
                    ' This is the Sheet that changed
                    filename = ThisWorkbook.ActiveSheet.Cells(2, Target.Column)
                    Set wWorkbook = GetWorkbook(filename)
                    If IsInWorkbook(Target.Value, wWorkbook) Then
                        ' Obtain The Sheet and detect area to compare
                        Set WSheet = wWorkbook.Sheets(Target.Value)
                        ThisWorkbook.ActiveSheet.Cells(4, Target.Column) = Detect_Table(WSheet)
                    End If
                End If
                Exit Sub
            End If
        End If
        If Target.Row = 4 Then
            If Left(ThisWorkbook.ActiveSheet.Cells(4, Target.Column).Formula, 1) = "=" Then
                x = InStr(1, ThisWorkbook.ActiveSheet.Cells(4, Target.Column).Formula, "!")
                If x > 0 Then
                    filename = Left(ThisWorkbook.ActiveSheet.Cells(4, Target.Column).Formula, x - 1)
                    x1 = InStr(1, filename, "]")
                    If x1 > 0 Then
                        ThisWorkbook.ActiveSheet.Cells(3, Target.Column) = Mid(filename, x1 + 1, 1000)
                        ThisWorkbook.ActiveSheet.Cells(2, Target.Column) = Mid(filename, 3, x1 - 3)
                    Else
                        ThisWorkbook.ActiveSheet.Cells(3, Target.Column) = Mid(filename, 2, 1000)
                    End If
                    ThisWorkbook.ActiveSheet.Cells(4, Target.Column) = Mid(ThisWorkbook.ActiveSheet.Cells(4, Target.Column).Formula, x + 1, 1000)
                Else
                    ThisWorkbook.ActiveSheet.Cells(4, Target.Column) = Mid(ThisWorkbook.ActiveSheet.Cells(4, Target.Column).Formula, 2, 1000)
                End If
                Exit Sub
            End If
        End If
        If Target.Row = 6 And Target.Column = 2 Then
            If StrComp(Target.Value, "Formulas", vbTextCompare) = 0 Then
                ThisWorkbook.ActiveSheet.Cells(13, 2) = "NO"
                With ThisWorkbook.ActiveSheet.Range("A7:B7").Font
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                End With
            Else
                With ThisWorkbook.ActiveSheet.Range("A7:B7").Font
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = -0.35
                End With
            End If
            Exit Sub
        End If
        If Target.Row = 13 And Target.Column = 2 Then
            If StrComp(Target.Value, "YES", vbTextCompare) = 0 And _
                StrComp(ThisWorkbook.ActiveSheet.Cells(6, 2), "Formulas") = 0 Then
                MsgBox "This is not supported when comparing Formulas"
                ThisWorkbook.ActiveSheet.Cells(13, 2) = "NO"
            End If
            Exit Sub
        End If
        If Target.Row = 9 And Target.Column = 2 Then
            If Target.Value = "None" Then
                With ThisWorkbook.ActiveSheet.Range("A10:A15,B10,B12:B15").Font
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = -0.35
                End With
            Else
                With ThisWorkbook.ActiveSheet.Range("A10:A15,B10,B12:B15").Font
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                End With
            End If
            Exit Sub
        End If
        If Target.Row = 17 And Target.Column = 2 Then
            If Target.Value = "YES" Then
                With ThisWorkbook.ActiveSheet.Range("A18:B18").Font
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                End With
            Else
                With ThisWorkbook.ActiveSheet.Range("A18:B18").Font
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = -0.35
                End With
            End If
        End If
    End If
End Sub




