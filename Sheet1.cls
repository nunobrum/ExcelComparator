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


Private Sub Worksheet_Change(ByVal target As Range)
    Dim wWorkbook As Workbook
    Dim auxstring As String
    Dim cellRef As Range
    Dim WSheet As Worksheet
    Dim x, x1 As Integer

    If target.Column = cfgColOriginal Or target.Column = cfgColRevision Then
        ' All the other columns are ignored.
        If target.row = cfgRowFilename Then
            ' This is the auxstring that was changed
            auxstring = target.Text
            'If auxstring is set by Range ignore
            If Not Left(ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowRange, target.Column).Formula, 1) = "=" Then
            ' Set the variable that will be passed as pointer
                Set cellRef = ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowSheet, target.Column)
                'Open the workbook
                Set wWorkbook = GetWorkbookConfig(target.Column)
                If Not wWorkbook Is Nothing Then
                    With ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowFilename, target.Column)
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
        If target.row = cfgRowSheet Then
            ' If the sheet was set by the range ignore
            If Not Left(ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowRange, target.Column).Formula, 1) = "=" Then
                If target.Text = ALLSheets Then
                    ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowRange, target.Column) = OptionAutoDetect
                Else
                    ' This is the Sheet that changed
                    Set wWorkbook = GetWorkbookConfig(target.Column)
                    If IsInWorkbook(target.Value, wWorkbook) Then
                        ' Obtain The Sheet and detect area to compare
                        Set WSheet = wWorkbook.Sheets(target.Value)
                        auxstring = Detect_Table(WSheet)
                        ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowRange, target.Column) = auxstring
                    End If
                End If
            End If
            Exit Sub
        End If
        If target.row = cfgRowRange Then
            If Left(ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowRange, target.Column).Formula, 1) = "=" Then
                x = InStr(1, ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowRange, target.Column).Formula, "!")
                If x > 0 Then
                    auxstring = Left(ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowRange, target.Column).Formula, x - 1)
                    If Left(auxstring, 3) = "='[" Then
                        auxstring = Mid(auxstring, 4, Len(auxstring) - 4)
                    Else
                        If Left(auxstring, 2) = "=[" Then
                            auxstring = Mid(auxstring, 3, Len(auxstring) - 2)
                        Else
                            ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowTableHeaders, cfgColOption) = "" ' Inserted/Deleted rows
                            Exit Sub
                        End If
                    End If
                    x1 = InStr(1, auxstring, "]")
                    If x1 > 0 Then
                        ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowSheet, target.Column) = Mid(auxstring, x1 + 1, 1000)
                        ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowFilename, target.Column) = Mid(auxstring, 1, x1 - 1)
                    Else
                        ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowSheet, target.Column) = Mid(auxstring, 2, 1000)
                    End If
                    ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowRange, target.Column) = Mid(ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowRange, target.Column).Formula, x + 1, 1000)
                Else
                    ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowRange, target.Column) = Mid(ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowRange, target.Column).Formula, 2, 1000)
                End If
            End If
            'Update the remaining fields
            If target.Value = OptionAutoDetect Then
                ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowTableHeaders, 2) = OptionNo ' Table has headers = No
            End If
            'In any case force the reload of the Column names
            Call ReloadColumnNames(target.Column)
            Exit Sub
        End If
        
        If target.row = cfgRowTableHeaders Then ' Table has headers
            If target.Column = cfgColOption Then
                If StrComp(target.Value, OptionYes, vbTextCompare) = 0 Then
                    Call RangeEnable(target.Worksheet.Range("A9:C9"))
                    If ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowRange, cfgColOriginal) <> OptionAutoDetect Then
                        auxstring = Str(DetectHeaderRow(cfgColOriginal))
                        ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowHeaderRow, cfgColOriginal) = auxstring
                    End If
                    If ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowRange, cfgColRevision) <> OptionAutoDetect Then
                        auxstring = Str(DetectHeaderRow(cfgColRevision))
                        ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowHeaderRow, cfgColRevision) = auxstring
                    End If
                Else
                    Call RangeDisable(target.Worksheet.Range("A9:C9"))
                End If
                Call ReloadColumnNames(cfgColOriginal)
                Call ReloadColumnNames(cfgColRevision)
            End If
            Exit Sub
        End If
        
        If target.row = cfgRowHeaderRow Then
            Call ReloadColumnNames(target.Column)
            Exit Sub
        End If

        If target.row = cfgRowAnnoMergeText And target.Column = cfgColOption Then ' Merge Text Cells
            If StrComp(target.Value, OptionYes, vbTextCompare) = 0 And _
                StrComp(ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowWhat, cfgColOption), "Formulas") = 0 Then
                MsgBox "This is not supported when comparing Formulas"
                ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowAnnoMergeText, cfgColOption) = OptionNo 'Merge Text Cells = No
            End If
            Exit Sub
        End If
        
        If target.row = cfgRowPrimaryKey And target.Column = cfgColOption Then
            If target.Value = OptionNo Then
                Call RangeDisable(ThisWorkbook.Sheets(sheetConfig).Range("A11:C11"))
            Else
                Call RangeEnable(ThisWorkbook.Sheets(sheetConfig).Range("A11:C11"))
            End If
            Exit Sub
        End If

        If target.row = cfgRowAnnotate And target.Column = cfgColOption Then ' Make anotations on
            If target.Value = ThisWorkbook.Sheets(sheetLanguage).Cells(rowOptionNone, colLanguage).Text Then
                Call RangeDisable(ThisWorkbook.Sheets(sheetConfig).Range("A14:A19,B14,B16:B18"))
            Else
                Call RangeEnable(ThisWorkbook.Sheets(sheetConfig).Range("A14,A16:A18,B15,B16:B18"))
                Call Worksheet_Change(ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowAnnoUseFormat, cfgColOption))
                Call Worksheet_Change(ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowAnnoUseRowFormat, cfgColOption))
            End If
            Exit Sub
        End If
        If target.row = cfgRowAnnoUseFormat And target.Column = cfgColOption Then
            If target.Value = OptionYes And _
                ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowAnnotate, cfgColOption).Text <> _
                ThisWorkbook.Sheets(sheetLanguage).Cells(rowOptionNone, colLanguage).Text Then
                Call RangeEnable(ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowAnnoCellFormat, CfgColTitle))
            Else
                Call RangeDisable(ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowAnnoCellFormat, CfgColTitle))
            End If
            Exit Sub
        End If
        
        If target.row = cfgRowAnnoUseRowFormat And target.Column = cfgColOption Then
            If target.Value = OptionYes And _
                ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowAnnotate, cfgColOption).Text <> _
                ThisWorkbook.Sheets(sheetLanguage).Cells(rowOptionNone, colLanguage).Text Then
                Call RangeEnable(ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowAnnoRowFormat, CfgColTitle))
            Else
                Call RangeDisable(ThisWorkbook.Sheets(sheetConfig).Cells(cfgRowAnnoRowFormat, CfgColTitle))
            End If
            Exit Sub
        End If
        If target.row = cfgRowReport And target.Column = cfgColOption Then ' Create Report
            If target.Value = OptionYes Then
                RangeEnable (ThisWorkbook.Sheets(sheetConfig).Range("A22:B22"))
            Else
                RangeDisable (ThisWorkbook.Sheets(sheetConfig).Range("A22:B22"))
            End If
        End If
    End If
    If target.row = cfgRowLanguage And target.Column = cfgColLanguage Then ' Language Change
        Call ResetFields
    End If
End Sub

