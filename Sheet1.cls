VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Sub RangeEnable(rng As Range)
    With rng.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
End Sub

Sub RangeDisable(rng As Range)
    With rng.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.35
    End With
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
    Dim wWorkbook As Workbook
    Dim auxstring As String
    Dim cellRef As Range
    Dim WSheet As Worksheet
    Dim x, x1 As Integer

    If Target.Column = cfgColOriginal Or Target.Column = cfgColRevision Then
        ' All the other columns are ignored.
        If Target.Row = cfgRowFilename Then
            ' This is the auxstring that was changed
            auxstring = Target.Text
            'If auxstring is set by Range ignore
            If Not Left(ThisWorkbook.Sheets("Config").Cells(cfgRowRange, Target.Column).Formula, 1) = "=" Then
            ' Set the variable that will be passed as pointer
                Set cellRef = ThisWorkbook.Sheets("Config").Cells(cfgRowSheet, Target.Column)
                'Open the workbook
                Set wWorkbook = GetWorkbookConfig(Target.Column)
                If Not wWorkbook Is Nothing Then
                    With ThisWorkbook.Sheets("Config").Cells(cfgRowFilename, Target.Column)
                        If Not (.comment Is Nothing) Then
                            .comment.Delete
                        End If
                        .AddComment wWorkbook.FullName
                    End With
                    ' Not showng the opened workbook
                    'wWorkbook.Worksheets.Visible = False

                    ' Update the available sheets in the combobox in the validation
                    Call Reload_Sheets(wWorkbook, cellRef)
                End If
            End If
            Exit Sub
        End If
        If Target.Row = cfgRowSheet Then
            ' If the sheet was set by the range ignore
            If Not Left(ThisWorkbook.Sheets("Config").Cells(cfgRowRange, Target.Column).Formula, 1) = "=" Then
                If Target.Text = "[ALL]" Then
                    ThisWorkbook.Sheets("Config").Cells(cfgRowRange, Target.Column) = "Auto Detect"
                Else
                    ' This is the Sheet that changed
                    Set wWorkbook = GetWorkbookConfig(Target.Column)
                    If IsInWorkbook(Target.Value, wWorkbook) Then
                        ' Obtain The Sheet and detect area to compare
                        Set WSheet = wWorkbook.Sheets(Target.Value)
                        auxstring = Detect_Table(WSheet)
                        ThisWorkbook.Sheets("Config").Cells(cfgRowRange, Target.Column) = auxstring
                    End If
                End If
            End If
            Exit Sub
        End If
        If Target.Row = cfgRowRange Then
            If Left(ThisWorkbook.Sheets("Config").Cells(cfgRowRange, Target.Column).Formula, 1) = "=" Then
                x = InStr(1, ThisWorkbook.Sheets("Config").Cells(cfgRowRange, Target.Column).Formula, "!")
                If x > 0 Then
                    auxstring = Left(ThisWorkbook.Sheets("Config").Cells(cfgRowRange, Target.Column).Formula, x - 1)
                    If Left(auxstring, 3) = "='[" Then
                        auxstring = Mid(auxstring, 4, Len(auxstring) - 4)
                    Else
                        If Left(auxstring, 2) = "=[" Then
                            auxstring = Mid(auxstring, 3, Len(auxstring) - 2)
                        Else
                            ThisWorkbook.Sheets("Config").Cells(cfgRowTableHeaders, cfgColOption) = "" ' Inserted/Deleted rows
                            Exit Sub
                        End If
                    End If
                    x1 = InStr(1, auxstring, "]")
                    If x1 > 0 Then
                        ThisWorkbook.Sheets("Config").Cells(cfgRowSheet, Target.Column) = Mid(auxstring, x1 + 1, 1000)
                        ThisWorkbook.Sheets("Config").Cells(cfgRowFilename, Target.Column) = Mid(auxstring, 1, x1 - 1)
                    Else
                        ThisWorkbook.Sheets("Config").Cells(cfgRowSheet, Target.Column) = Mid(auxstring, 2, 1000)
                    End If
                    ThisWorkbook.Sheets("Config").Cells(cfgRowRange, Target.Column) = Mid(ThisWorkbook.Sheets("Config").Cells(cfgRowRange, Target.Column).Formula, x + 1, 1000)
                Else
                    ThisWorkbook.Sheets("Config").Cells(cfgRowRange, Target.Column) = Mid(ThisWorkbook.Sheets("Config").Cells(cfgRowRange, Target.Column).Formula, 2, 1000)
                End If
            End If
            'Update the remaining fields
            If Target.Value = "Auto Detect" Then
                ThisWorkbook.Sheets("Config").Cells(cfgRowTableHeaders, 2) = "NO" ' Table has headers = No
            End If
            'In any case force the reload of the Column names
            Call ReloadColumnNames(Target.Column)
            Exit Sub
        End If
        If Target.Row = cfgRowWhat And Target.Column = cfgColOption Then ' What to compare
            If StrComp(Target.Value, "Formulas", vbTextCompare) = 0 Then
                ThisWorkbook.Sheets("Config").Cells(cfgRowAnnoMergeText, cfgColOption) = "NO"
                Call RangeEnable(ThisWorkbook.Sheets("Config").Range("A7:B7"))
            Else
                Call RangeDisable(ThisWorkbook.Sheets("Config").Range("A7:B7"))
            End If
            Exit Sub
        End If
        
        If Target.Row = cfgRowTableHeaders Then ' Table has headers
            If Target.Column = cfgColOption Then
                If StrComp(Target.Value, "YES", vbTextCompare) = 0 Then
                    ' Fist validate if multiple sheets is programmed
                    If ThisWorkbook.Sheets("Config").Cells(cfgRowSheet, cfgColOriginal) = "[ALL]" And _
                       ThisWorkbook.Sheets("Config").Cells(cfgRowSheet, cfgColRevision) = "[ALL]" Then
                        ThisWorkbook.Sheets("Config").Cells(cfgRowTableHeaders, cfgColOption) = "NO"
                        MsgBox "This option is not allowed with multiple sheet compare", vbOKOnly
                    Else
                        Call RangeEnable(Target.Worksheet.Range("A10:C10"))
                        If ThisWorkbook.Sheets("Config").Cells(cfgRowRange, cfgColOriginal) <> "Auto Detect" Then
                            auxstring = Str(DetectHeaderRow(cfgColOriginal))
                            ThisWorkbook.Sheets("Config").Cells(cfgRowHeaderRow, cfgColOriginal) = auxstring
                        End If
                        If ThisWorkbook.Sheets("Config").Cells(cfgRowRange, cfgColRevision) <> "Auto Detect" Then
                            auxstring = Str(DetectHeaderRow(cfgColRevision))
                            ThisWorkbook.Sheets("Config").Cells(cfgRowHeaderRow, cfgColRevision) = auxstring
                        End If
                    End If
                Else
                    Call RangeDisable(Target.Worksheet.Range("A10:C10"))
                End If
                Call ReloadColumnNames(cfgColOriginal)
                Call ReloadColumnNames(cfgColRevision)
            End If
            Exit Sub
        End If
        
        If Target.Row = cfgRowHeaderRow Then
            Call ReloadColumnNames(Target.Column)
            Exit Sub
        End If

        If Target.Row = cfgRowAnnoMergeText And Target.Column = cfgColOption Then ' Merge Text Cells
            If StrComp(Target.Value, "YES", vbTextCompare) = 0 And _
                StrComp(ThisWorkbook.Sheets("Config").Cells(cfgRowWhat, cfgColOption), "Formulas") = 0 Then
                MsgBox "This is not supported when comparing Formulas"
                ThisWorkbook.Sheets("Config").Cells(cfgRowAnnoMergeText, cfgColOption) = "NO" 'Merge Text Cells = No
            End If
            Exit Sub
        End If
        
        If Target.Row = cfgRowPrimaryKey And Target.Column = cfgColOption Then
            If Target.Value = "NO" Then
                Call RangeDisable(ThisWorkbook.Sheets("Config").Range("A12:C12"))
            Else
                Call RangeEnable(ThisWorkbook.Sheets("Config").Range("A12:C12"))
            End If
            Exit Sub
        End If

        If Target.Row = cfgRowAnnotate And Target.Column = cfgColOption Then ' Make anotations on
            If Target.Value = "None" Then
                Call RangeDisable(ThisWorkbook.Sheets("Config").Range("A15:A20,B15,B17:B20"))
            Else
                Call RangeEnable(ThisWorkbook.Sheets("Config").Range("A15:A20,B15,B17:B20"))
            End If
            Exit Sub
        End If
        If Target.Row = cfgRowReport And Target.Column = cfgColOption Then ' Create Report
            If Target.Value = "YES" Then
                RangeEnable (ThisWorkbook.Sheets("Config").Range("A23:B23"))
            Else
                RangeDisable (ThisWorkbook.Sheets("Config").Range("A23:B23"))
            End If
        End If
    End If
End Sub




