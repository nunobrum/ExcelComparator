Attribute VB_Name = "SheetCompare"
Option Explicit


Function ColNumber(col As String) As Integer
    If IsNumeric(col) Then
        ColNumber = Int(col)
    Else
        ColNumber = Range(col + "1").Column
        ' The reverse is done by Split(Cells(,col).address,"$")(1)
    End If
End Function



Function SelectFileWindows() As String
    Dim fDialog As Office.FileDialog

    ' Set up the File Dialog.
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
   
    With fDialog
        .Filters.Clear
        .Filters.Add "Excel files (*.xls?)", "*.xls?"
        .Title = "Select Excel File to Compare"
        
        If .Show = True Then
            SelectFileWindows = .SelectedItems.Item(1)
        Else
            SelectFileWindows = ""
        End If
    End With
End Function


Function SelectFileMac() As String
    Dim MyPath As String
    Dim MyScript As String
    Dim MyFiles As String

    MyPath = MacScript("return (path to documents folder) as String")
    'Or use MyPath = "Macintosh HD:Users:Ron:Desktop:TestFolder:"

    ' In the following statement, change true to false in the line "multiple
    ' selections allowed true" if you do not want to be able to select more
    ' than one file. Additionally, if you want to filter for multiple files, change
    ' {""com.microsoft.Excel.xls""} to
    ' {""com.microsoft.excel.xls"",""public.comma-separated-values-text""}
    ' if you want to filter on xls and csv files, for example.
    MyScript = _
    "set applescript's text item delimiters to "","" " & vbNewLine & _
               "set theFile to (choose file of type " & _
               " {""xls"",""xlsx"",""xlsm""} " & _
               "with prompt ""Please select an excel file"" default location alias """ & _
               MyPath & """) as string" & vbNewLine & _
               "return POSIX path of theFile"

    MyFiles = MacScript(MyScript)

    SelectFileMac = MyFiles
End Function

Function bIsBookOpen(ByRef szBookName As String) As Boolean
' Contributed by Rob Bovey
    On Error Resume Next
    bIsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
End Function



Function SelectFileWINorMAC() As String
    Dim Answer As String
' Test for the operating system.
    If Not Application.OperatingSystem Like "*Mac*" Then
        ' Is Windows.
        Answer = SelectFileWindows()
    Else
        ' Is a Mac and will test if running Excel 2011 or higher.
        If Val(Application.Version) > 14 Then
            Answer = SelectFileMac()
        End If
    End If
    SelectFileWINorMAC = Answer
End Function


Function GetWorkbook(filename As String) As Workbook
    Dim Aux As Workbook
    Dim wbk As Variant
    Dim wFilename As String
    Dim i As Integer
    
    On Error GoTo Error_Opening_File
    
    Set Aux = Nothing
    For i = 1 To Workbooks.Count
    Set wbk = Workbooks.Item(i)
      If IsNull(wbk) = False Then
        wFilename = wbk.FullName
        If StrComp(wFilename, filename, vbTextCompare) = 0 Then
            Set Aux = wbk
            Exit For
        End If
      End If
    Next i
    If Aux Is Nothing Then
        Set GetWorkbook = Workbooks.Open(filename)
        ThisWorkbook.Activate ' Make sure that the current file doesn't get hidden
   Else
        Set GetWorkbook = Aux
   End If
   Exit Function
Error_Opening_File:
   Set GetWorkbook = Nothing
End Function

Sub Reload_Sheets(wbk As Workbook, cell As Range)
    Dim sheet As Variant
    Dim list As String
    Dim MATCHED As Boolean
    
    list = ""
    MATCHED = False
    For Each sheet In wbk.Worksheets
        If Len(list) = 0 Then
            list = sheet.Name
        Else
            list = list + "," + sheet.Name
        End If
        If StrComp(cell.Text, sheet.Name, vbTextCompare) = 0 Then
            MATCHED = True
        End If
    Next
    If wbk.Worksheets.Count > 1 Then
        list = list + ",[ALL]"
    End If

    With cell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=list
        .ErrorTitle = "Invalid Sheet"
        .ErrorMessage = "Select from the list"
    End With
    If MATCHED = False Then
        If wbk.Worksheets.Count > 1 Then
            cell.Value = "[ALL]"
        Else
            'This is shit, but it didn't work any other way
            ' wbk.Worksheets(1).Name always gave an error
            For Each sheet In wbk.Worksheets
                list = sheet.Name
                Exit For
            Next
            list = sheet.Name
            cell.Value = list
        End If
    End If
End Sub

Function Detect_Table(ws As Worksheet) As String
    Dim x, xmin, xmax As Integer
    Dim y, ymin, ymax As Long
    Dim blankRows As Integer
    Dim blankColumns As Integer

    xmin = ws.Columns.Count
    xmax = 1
    ymax = 1
    ymin = ws.Rows.Count
    blankRows = 1

    For y = 1 To ws.Rows.Count
        blankColumns = 1
        For x = 1 To ws.Columns.Count
            If Not ws.Cells(y, x).Text = "" Then
                If x < xmin Then
                    xmin = x
                End If
                If x > xmax Then
                    xmax = x
                End If
                blankRows = 0
                blankColumns = 0
            Else
                blankColumns = blankColumns + 1
            End If
            If blankColumns > 30 Then
                Exit For
            End If
        Next x
        If blankRows > 30 Then
            Exit For
        Else
            ' The line wasn't blank
            If blankRows = 0 Then
                If ymin > y Then
                    ymin = y
                End If
                ymax = y
            End If
            blankRows = blankRows + 1
        End If
    Next y
    If xmax < xmin Then
        xmin = xmax
    End If
    If ymax < ymin Then
        ymin = ymax
    End If

    Detect_Table = Cells(ymin, xmin).Address + ":" + Cells(ymax, xmax).Address
End Function


Sub Open_OriginalWorkbook()
    Dim Original_Workbook As Workbook
    Dim filename As String
    Dim cellRef As Range
    Dim WSheet As Worksheet
    
    filename = SelectFileWINorMAC()
    If Len(filename) <> 0 Then
        ' Set te Filename on the proper cell
        ThisWorkbook.ActiveSheet.Cells(2, 2) = filename
        ' Set the variable that will be passed as pointer
        Set cellRef = ThisWorkbook.ActiveSheet.Cells(3, 2)
    End If
End Sub

Sub Open_RevisionWorkbook()
    Dim Revision_Workbook As Workbook
    Dim filename As String
    Dim cellRef As Range
    Dim WSheet As Worksheet
    
    filename = SelectFileWINorMAC()
    If Len(filename) <> 0 Then
        ThisWorkbook.ActiveSheet.Cells(2, 3) = filename
        Set cellRef = ThisWorkbook.ActiveSheet.Cells(3, 3)
    End If
End Sub

Function IsInWorkbook(sheetToBeFound As String, wbk As Workbook) As Boolean
    Dim sheet As Variant
    For Each sheet In wbk.Worksheets
        If sheetToBeFound = sheet.Name Then
            IsInWorkbook = True
            Exit Function
        End If
    Next
    IsInWorkbook = False
End Function


'For More Free Code & Ideas Visit http://OfficeTricks.com
Sub Compare_Excel_Files_WorkSheets()
    'Define Object for Excel Workbooks to Compare
    Dim sh As Integer, ShName As String
    Dim Ori_Workbook As Workbook, Rev_Workbook As Workbook
    Dim Ori_SheetName As String, Rev_SheetName As String
    Dim Ori_Range As String, Rev_Range As String
    Dim Ori_Sheet As Worksheet, Rev_Sheet As Worksheet
    Dim Cfg_Sheet As Worksheet, Rep_Sheet As Worksheet
    
    Dim sheetIndex, comparedSheets As Integer
    Dim bMultipleSheets As Boolean
    Dim iRow_Count As Double, iCol_Count As Double
    Dim Ori_iRow_Start As Double, Ori_iCol_Start As Double
    Dim Rev_iRow_Start As Double, Rev_iCol_Start As Double
    Dim Rep_iRow_Start As Double
    
    Dim iRow As Double, iCol As Double, iCol1 As Integer
    Dim File1_Path As String, File2_Path As String, Ori_Data As String, Rev_Data As String
    Dim tempWidth As Integer
    
    Dim iRepCount As Integer, iRepRow As Integer
    Dim bDoReport As Boolean, bFirstDifference As Boolean, bRowChanged As Boolean
    Dim bMakeAnnotation As Boolean
    Dim AnnotationSheet As Worksheet, Annotation_Row_Start As Integer, Annotation_Col_Start As Integer
    Dim bApplyChangeFormat As Boolean, bInsertComment As Boolean
    Dim ChangedCellFormat As Range
    Dim targetCell As Range
    Dim bTextMerge, bReportMerge As Boolean
    Dim bCompareFormulas As Boolean, bR1C1Format As Boolean
    Dim comment As String

    Dim annotateColumn As Integer

    Dim iDiffCount As Double
    
    Set Cfg_Sheet = ThisWorkbook.ActiveSheet
    'Assign the Workbook File Name along with its Path
    File1_Path = Cfg_Sheet.Cells(2, 2)
    File2_Path = Cfg_Sheet.Cells(2, 3)
    
    
    Set Rev_Workbook = GetWorkbook(File2_Path)
    Set Ori_Workbook = GetWorkbook(File1_Path)
    
    If Ori_Workbook Is Nothing Then
        MsgBox "File """ & File1_Path & """ doesn't exist"
        Exit Sub
    End If

    If Rev_Workbook Is Nothing Then
        MsgBox "File """ & File2_Path & """ doesn't exist"
        Exit Sub
    End If

    bMakeAnnotation = StrComp(Cfg_Sheet.Cells(9, 2).Value, "None", vbTextCompare) <> 0
  
    If bMakeAnnotation Then
        bApplyChangeFormat = (StrComp(Cfg_Sheet.Cells(10, 2), "YES", vbTextCompare) = 0)
        Set ChangedCellFormat = Cfg_Sheet.Cells(11, 2)
        bInsertComment = (StrComp(Cfg_Sheet.Cells(12, 2), "YES", vbTextCompare) = 0)
        bTextMerge = (StrComp(Cfg_Sheet.Cells(13, 2), "YES", vbTextCompare) = 0)
        
        ' if annotateColumn is zero, the annotation is not done
        If (StrComp(Cfg_Sheet.Cells(14, 2), "YES", vbTextCompare) = 0) Then
            annotateColumn = ColNumber(Cfg_Sheet.Cells(15, 2).Text)
        Else
            annotateColumn = 0
        End If
    End If
     
    'Comparing Values of Formulas
    ' This is done here so thata the bTextMerge can be direcly overriden
    If StrComp(Cfg_Sheet.Cells(6, 2).Value, "Formulas", vbTextCompare) = 0 Then
        bCompareFormulas = True
        bR1C1Format = (StrComp(Cfg_Sheet.Cells(7, 2), "YES", vbTextCompare) = 0)
        bTextMerge = False
    Else
        bCompareFormulas = False
    End If
     
    bDoReport = (StrComp(Cfg_Sheet.Cells(17, 2), "YES", vbTextCompare) = 0)
    If bDoReport Then
        bReportMerge = (StrComp(Cfg_Sheet.Cells(18, 2), "YES", vbTextCompare) = 0)
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "Creating the report..."


    'Deleting old Diff Sheets
    Application.DisplayAlerts = False
    Do While ThisWorkbook.Worksheets.Count > 2
        ThisWorkbook.Worksheets(3).Delete
    Loop
    Application.DisplayAlerts = True
    
    'With Ori_Workbook object, now it is possible to pull any data from it
    'Read Data From Each Sheets of Both Excel Files & Compare Data
    
    sheetIndex = 1
    comparedSheets = 0
    bMultipleSheets = False

' Cycling through Sheets
  Do While True
    If Cfg_Sheet.Cells(3, 2) = "[ALL]" Then
        If Cfg_Sheet.Cells(3, 2) = "[ALL]" Then
            bMultipleSheets = True
TRY_NEXT:
            If sheetIndex > Ori_Workbook.Sheets.Count Then
                GoTo EXIT_LOOP
            End If
            Ori_SheetName = Ori_Workbook.Sheets(sheetIndex).Name

            If IsInWorkbook(Ori_SheetName, Rev_Workbook) Then
                Rev_SheetName = Ori_SheetName
            Else
                sheetIndex = sheetIndex + 1
                GoTo TRY_NEXT
            End If
        Else
            ' One sheet to compare in each side
            If sheetIndex > 1 Then
                GoTo EXIT_LOOP
            End If
            If IsInWorkbook(Rev_SheetName, Ori_Workbook) Then
                Ori_SheetName = Rev_SheetName
            Else
                ' TODO: create an error flag so than an error is reported
                GoTo EXIT_LOOP
            End If
        End If
    Else
        ' One sheet to compare in each side
        Ori_SheetName = Cfg_Sheet.Cells(3, 2)
        If sheetIndex > 1 Then
            GoTo EXIT_LOOP
        End If
    
        If Cfg_Sheet.Cells(3, 3) = "[ALL]" Then
            If IsInWorkbook(Ori_SheetName, Rev_Workbook) Then
                Rev_SheetName = Ori_SheetName
            Else
                ' TODO: create an error flag so than an error is reported
                GoTo EXIT_LOOP
            End If
        Else
            Rev_SheetName = Cfg_Sheet.Cells(3, 3)
        End If
    End If
    sheetIndex = sheetIndex + 1

    Set Ori_Sheet = Ori_Workbook.Sheets(Ori_SheetName)
    Set Rev_Sheet = Rev_Workbook.Sheets(Rev_SheetName)
    
    ' Getting Row and Column Start
    Ori_Range = Cfg_Sheet.Cells(4, 2)
    Rev_Range = Cfg_Sheet.Cells(4, 3)

    If Ori_Range = "Auto Detect" Then
        Ori_Range = Detect_Table(Ori_Sheet)
    End If

    If bMultipleSheets Then
        ' TODO: This is assuming the Original Size, but this is not the best
        ' the best is to calculate the Super Set
        Rev_Range = Ori_Range
    Else
        If Rev_Range = "Auto Detect" Then
            Rev_Range = Detect_Table(Rev_Sheet)
        End If
    End If

    Ori_iRow_Start = Range(Ori_Range).Row
    Ori_iCol_Start = Range(Ori_Range).Column
    
    Rev_iRow_Start = Range(Rev_Range).Row
    Rev_iCol_Start = Range(Rev_Range).Column
 
    'Calculating count of rows and columns to process
    iRow_Count = Range(Ori_Range).Rows.Count
    iCol_Count = Range(Ori_Range).Columns.Count
    
    'Checking area sizes are equal
    If Range(Ori_Range).Rows.Count <> Range(Rev_Range).Rows.Count Or _
       Range(Ori_Range).Columns.Count <> Range(Rev_Range).Columns.Count Then

        'Assuming the bigger between the two sheets
        If Range(Rev_Range).Rows.Count > iRow_Count Then
            iRow_Count = Range(Rev_Range).Rows.Count
        End If
    
        If Range(Rev_Range).Columns.Count > iCol_Count Then
            iCol_Count = Range(Rev_Range).Columns.Count
        End If

        MsgBox "Ranges for Sheet '" & Ori_Sheet.Name & "' differ in size." & vbCr _
             & "Assuming the Smaller sizes" & vbCr _
             & "Comparing  " & iRow_Count & " row(s) x " & iCol_Count & " column(s)"
    End If
    
    
    'Subtracting 1 to both column and row count because counts are started from 0
    iRow_Count = iRow_Count - 1
    iCol_Count = iCol_Count - 1
    
    If StrComp(Cfg_Sheet.Cells(9, 2).Value, "Original", vbTextCompare) = 0 Then
        Set AnnotationSheet = Ori_Sheet
        Annotation_Row_Start = Ori_iRow_Start
        Annotation_Col_Start = Ori_iCol_Start
    Else
        If StrComp(Cfg_Sheet.Cells(9, 2).Value, "Revision", vbTextCompare) = 0 Then
            Set AnnotationSheet = Rev_Sheet
            Annotation_Row_Start = Rev_iRow_Start
            Annotation_Col_Start = Rev_iCol_Start
        Else
            Set AnnotationSheet = Nothing
        End If
    End If
    
    'With Ori_Workbook object, now it is possible to pull any data from it
    'Read Data From Each Sheets of Both Excel Files & Compare Data
    
    bFirstDifference = True ' This is used for Starting the Report

    For iRow = 0 To iRow_Count
        bRowChanged = False
        For iCol = 0 To iCol_Count
            If bCompareFormulas Then
                If bR1C1Format Then
                    Ori_Data = Ori_Sheet.Cells(iRow + Ori_iRow_Start, iCol + Ori_iCol_Start).FormulaR1C1Local
                    Rev_Data = Rev_Sheet.Cells(iRow + Rev_iRow_Start, iCol + Rev_iCol_Start).FormulaR1C1Local
                Else
                    Ori_Data = Ori_Sheet.Cells(iRow + Ori_iRow_Start, iCol + Ori_iCol_Start).FormulaLocal
                    Rev_Data = Rev_Sheet.Cells(iRow + Rev_iRow_Start, iCol + Rev_iCol_Start).FormulaLocal
                End If
            Else
                Ori_Data = Ori_Sheet.Cells(iRow + Ori_iRow_Start, iCol + Ori_iCol_Start).Text
                Rev_Data = Rev_Sheet.Cells(iRow + Rev_iRow_Start, iCol + Rev_iCol_Start).Text
            End If
            'Compare Data From Excel Sheets & Highlight the Mismatches
            If Ori_Data <> Rev_Data Then
                iDiffCount = iDiffCount + 1
                'Report the difference
                If bDoReport Then
                    'If it is the first difference there is some stuff to do
                    If bRowChanged = False Then
                        'create the report if needed
                        If bFirstDifference Then
                            If bMultipleSheets Then
                                Set Rep_Sheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count), Type:=xlWorksheet)
                                'Set Rep_Sheet = Workbooks.Add.Sheets(1)
                                Rep_Sheet.Name = "Diff-" & Ori_Sheet.Name
                            Else
                                Set Rep_Sheet = ThisWorkbook.Sheets("Report")
                            End If
                            'Rep_Sheet.Activate
                            Rep_Sheet.Cells.Clear ' Clear everything on the sheet

                            Rep_Sheet.Cells(1, 1).Value = "Original Range"
                            Rep_Sheet.Cells(2, 1).Value = "Revision Range"
                            Rep_Sheet.Cells(1, 2).Value = "[" & File1_Path & "]" & Ori_Sheet.Name & "!" & Ori_Range
                            Rep_Sheet.Cells(2, 2).Value = "[" & File2_Path & "]" & Rev_Sheet.Name & "!" & Rev_Range
                            Rep_Sheet.Cells(3, 1).Value = "Differences Found"
                            iRepRow = 0
                            Rep_iRow_Start = 4
                            'Format the widths to match the Original
                            For iCol1 = 0 To iCol_Count
                                tempWidth = Ori_Sheet.Columns(iCol1 + Ori_iCol_Start).ColumnWidth
                                If tempWidth > 250 Then
                                    tempWidth = 250
                                End If
                                Rep_Sheet.Columns(iCol1 + 2).ColumnWidth = tempWidth
                            Next iCol1
                            bFirstDifference = False
                        Else
                            iRepRow = iRepRow + 1
                        End If
                        If Ori_iRow_Start <> Rev_iRow_Start Then
                            Rep_Sheet.Cells(iRepRow + Rep_iRow_Start, 1).Value = "Ori:" & (iRow + Ori_iRow_Start) & "=Rev:" & (iRow + Rev_iRow_Start)
                        Else
                            Rep_Sheet.Cells(iRepRow + Rep_iRow_Start, 1).Value = "Line:" & (iRow + Ori_iRow_Start)
                        End If
                        ' Copy from Original
                        If bCompareFormulas Then
                            For iCol1 = 0 To iCol_Count
                                Rep_Sheet.Cells(iRepRow + Rep_iRow_Start, iCol1 + 2).Formula = Ori_Sheet.Cells(iRow + Ori_iRow_Start, iCol1 + Ori_iCol_Start).Formula
                            Next iCol1
                        Else
                            For iCol1 = 0 To iCol_Count
                                Rep_Sheet.Cells(iRepRow + Rep_iRow_Start, iCol1 + 2).Value = Ori_Sheet.Cells(iRow + Ori_iRow_Start, iCol1 + Ori_iCol_Start)
                            Next iCol1
                        End If
                    End If
                    ' Highlight the Mismatches
                    
                    With Rep_Sheet.Cells(iRepRow + Rep_iRow_Start, iCol + 2)
                        .Interior.Pattern = ChangedCellFormat.Interior.Pattern
                        .Interior.PatternColorIndex = ChangedCellFormat.Interior.PatternColorIndex
                        .Interior.ThemeColor = ChangedCellFormat.Interior.ThemeColor
                        .Interior.TintAndShade = ChangedCellFormat.Interior.TintAndShade
                        .Interior.PatternTintAndShade = ChangedCellFormat.Interior.PatternTintAndShade
                    End With
                    If bCompareFormulas Then
                        ' Remove the first equal
                        If Left(Ori_Data, 1) = "=" Then
                            Ori_Data = Right(Ori_Data, Len(Ori_Data) - 1)
                        End If
                        If Left(Rev_Data, 1) = "=" Then
                            Rev_Data = Right(Rev_Data, Len(Rev_Data) - 1)
                        End If
                    End If
                    If bReportMerge Then
                        Call MergeText(Ori_Data, Rev_Data, Rep_Sheet.Cells(iRepRow + Rep_iRow_Start, iCol + 2))
                    Else
                        Rep_Sheet.Cells(iRepRow + Rep_iRow_Start, iCol + 2) = _
                                       "Changed from: " + vbCr + vbLf + _
                                       Ori_Data + vbCr + vbLf + _
                                       "To: " + _
                                       vbCr + vbLf + Rev_Data

                    End If
                End If
                
                ' Annotating one of the sheets
                If bMakeAnnotation Then
                    Set targetCell = AnnotationSheet.Cells(iRow + Annotation_Row_Start, iCol + Annotation_Col_Start)
                    With targetCell
                        ' Formating Cells
                        If bApplyChangeFormat Then
                            .Interior.Pattern = ChangedCellFormat.Interior.Pattern
                            .Interior.PatternColorIndex = ChangedCellFormat.Interior.PatternColorIndex
                            .Interior.Color = ChangedCellFormat.Interior.Color
                            .Interior.TintAndShade = ChangedCellFormat.Interior.TintAndShade
                            .Font.Color = ChangedCellFormat.Font.Color
                            ' .Interior.ColorIndex = ChangedCellFormat.Interior.ColorIndex
                            ' .Font = ChangedCellFormat.Font
                            ' .Borders = ChangedCellFormat.Borders
                        End If
                        'Adding difference in comments. Option
                        If bInsertComment Then
                            If Not (.comment Is Nothing) Then
                                .comment.Delete
                            End If
                            comment = "Changed from: " + vbCr + vbLf + _
                                       Ori_Data + vbCr + vbLf + _
                                       "To: " + _
                                       vbCr + vbLf + Rev_Data
                            targetCell.AddComment comment
                            targetCell.comment.Visible = False
                            
                        End If
                    End With
                    If bTextMerge And IsNumeric(targetCell.Value) = False Then
                            '.Value = Ori_Data + Rev_Data
                            '.Characters(Start:=1, Length:=Len(Ori_Data)).Font.Strikethrough = True
                            '.Characters(Start:=Len(Ori_Data) + 1, Length:=Len(Rev_Data)).Font.Underline = xlUnderlineStyleSingle
                        Call MergeText(Ori_Data, Rev_Data, targetCell)
                    End If

                End If
                bRowChanged = True
            End If
        Next iCol
        If bRowChanged And annotateColumn <> 0 Then
            AnnotationSheet.Cells(iRow + Annotation_Row_Start, annotateColumn).Value = "Modified"
        End If

        ' Informing the user
        If (iRow Mod 30) = 0 Then
            Application.StatusBar = "Progress... " & Format(iRow / iRow_Count * 100, "##0.0")
        End If

    Next iRow
    comparedSheets = comparedSheets + 1
  Loop ' End Of cycling through sheets
EXIT_LOOP:
    '''''Process Completed
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Task Completed - " & comparedSheets & " sheets compared. " & vbCr _
            & iDiffCount & " Differences Found" & vbCr & "(c) Nuno Brum, www.nunobrum.com"
End Sub


Sub TestCharacterChange()
 With ActiveCell.Characters(Start:=1, length:=275).Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With ActiveCell.Characters(Start:=276, length:=-20).Font
        .Name = "Arial"
        .FontStyle = "Bold"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFrontNone
    End With

End Sub

Sub MergeArrays(ByRef Ori() As String, ByRef Rev() As String, ByRef Mrg() As String, ByRef Mrk() As String)

    Dim o, r, m, x, y As Integer
    Dim OriLen, RevLen As Integer
    
    Dim msg As String
    Dim i As Integer

    OriLen = UBound(Ori)
    RevLen = UBound(Rev)

    ReDim Mrg(OriLen + RevLen)
    ReDim Mrk(OriLen + RevLen)

    y = 0
    x = 1
    o = 0
    r = 0
    m = 0
    
    Do
        If Rev(r) = Ori(o) Then
            ' Matched
            Mrg(m) = Ori(o)
            Mrk(m) = "."
            m = m + 1
            o = o + 1
            r = r + 1
            msg = "M"
        Else
            If Rev(r + x) = Ori(o + y) Then
                ' somethng inserted and deleted
                For i = 0 To y - 1
                    Mrg(m) = Ori(o)
                    Mrk(m) = "X"
                    m = m + 1
                    o = o + 1
                Next i
                For i = 0 To x - 1
                    Mrg(m) = Rev(r)
                    Mrk(m) = "_"
                    m = m + 1
                    r = r + 1
                Next i
                x = 1
                y = 0
                msg = "I"
            Else
                If Ori(o + x) = Rev(r + y) Then
                    'Something Deleted :
                    For i = 0 To x - 1
                        Mrg(m) = Ori(o)
                        Mrk(m) = "X"
                        m = m + 1
                        o = o + 1
                    Next i
                    For i = 0 To y - 1
                        Mrg(m) = Rev(r)
                        Mrk(m) = "_"
                        m = m + 1
                        r = r + 1
                    Next i
                    x = 1
                    y = 0
                    msg = "D"
                Else
                    ' Keep trying
                    If y >= x Then
                        y = 0
                        x = x + 1
                   
                        If o + x >= OriLen Or r + x >= RevLen Then
                            ' Until the end of the line. At this point give up.
                            ' Note : This may be later be replaced by a TRY_MATCH_LIMIT
                            msg = "F"
                            Exit Do
                        End If
                    Else
                        y = y + 1
                        msg = "T"
                    End If
                End If
            End If
        End If
       
       
    Loop While o < OriLen And r < RevLen
    ' Now complete till the end with the remaining
    Do While o < OriLen
        Mrg(m) = Ori(o)
        Mrk(m) = "X"
        m = m + 1
        o = o + 1
    Loop
    Do While r < RevLen
        Mrg(m) = Rev(r)
        Mrk(m) = "_"
        m = m + 1
        r = r + 1
    Loop
  

End Sub


Sub MergeText(OriText As String, RevText As String, cell As Range)
    Dim Ori() As String
    Dim Rev() As String
    
    Dim Mrg() As String
    Dim Mrk() As String
    Dim msg As String
    Dim i, p, l As Integer
    Dim LO As Integer, LR As Integer

    LO = Len(OriText)
    LR = Len(RevText)
                        
    If LO = 0 Then
        cell.Value = RevText
        cell.Characters(Start:=1, length:=LR).Font.Underline = xlUnderlineStyleSingle
    Else
        If LR = 0 Then
            cell.Value = OriText
            cell.Characters(Start:=1, length:=LO).Font.Strikethrough = True
        Else
            If IsNumeric(RevText) = False Then
                Ori = Atomize(OriText, " .,;:")
                Rev = Atomize(RevText, " .,;:")
            
                Call MergeArrays(Ori, Rev, Mrg, Mrk)
            
                msg = ArrayToString(Mrg)
                p = 0
                With cell
                    .Value = msg
                    For i = 0 To UBound(Mrg)
                        l = Len(Mrg(i))
                        If Mrk(i) = "X" Then
                            .Characters(Start:=p + 1, length:=l).Font.Strikethrough = True
                        Else
                            If Mrk(i) = "_" Then
                                .Characters(Start:=p + 1, length:=l).Font.Underline = xlUnderlineStyleSingle
                            End If
                        End If
                        p = p + l
                     Next i
                End With
            Else
                cell.Value = "'" + OriText + vbCr + RevText
                cell.Characters(Start:=1, length:=LO).Font.Strikethrough = True
                cell.Characters(Start:=LO + 2, length:=LR).Font.Underline = xlUnderlineStyleSingle
            End If
        End If
    End If
End Sub

Function Atomize(ByVal inString As String, sep As String) As String()
    Dim pos, posmin, last, cnt, length As Integer
    Dim buff() As String
    Dim ch As Integer
    length = Len(inString)
    ReDim buff(length)
    cnt = 0
    pos = 1
    last = 1
    Do
        posmin = length
        For ch = 1 To Len(sep)
            pos = InStr(last, inString, Mid(sep, ch, 1))
            If pos <> 0 And pos < posmin Then
                posmin = pos
            End If
        Next ch
        buff(cnt) = Mid$(inString, last, posmin - last + 1)
        last = posmin + 1
        cnt = cnt + 1
    Loop While posmin <> length
    ReDim Preserve buff(cnt)
    Atomize = buff
End Function

Function StringToArray(inString As String) As String()
    Dim buff() As String
    Dim i As Integer
    ' ANSI Only Solution
    'buff = Split(StrConv(my_string, vbUnicode), Chr$(0))
    'ReDim Preserve buff(UBound(buff) - 1)

    ' Plain Loop
    ReDim buff(Len(inString) - 1)
    For i = 1 To Len(inString)
        buff(i - 1) = Mid$(inString, i, 1)
    Next

    StringToArray = buff
End Function


Function ArrayToString(inArray() As String) As String
    Dim TL As Integer
    Dim outString As String
    Dim i As Integer
    TL = 0
    outString = ""
    For i = 0 To UBound(inArray)
        TL = TL + Len(inArray(i))
        outString = outString + inArray(i)
    Next i
    ArrayToString = outString
End Function

Function UnicodeToArray(inString As String) As String()
    ' This string is made up of a surrogate pair (high surrogate
    ' U+D800 and low surrogate U+DC00) and a combining character
    ' sequence (the letter "a" with the combining grave accent).
    'Dim testString2 As String = ChrW(&HD800) & ChrW(&HDC00) & "a" & ChrW(&H300)

    ' Create and initialize a StringInfo object for the string.
    Dim si As New System.Globalization.StringInfo(inString)

    ' Create and populate the array.
    Dim unicodeTestArray(si.LengthInTextElements) As String
    For i As Integer = 0 To si.LengthInTextElements - 1
        unicodeTestArray(i) = si.SubstringByTextElements(i, 1)
    Next
    UnicodeToArray = unicodeTestArray

End Function

Sub ArrangeWindows()
    'Dim OriWorkbook As Workbook
    'Dim RevWorkbook As Workbook
    'Dim wndName As String
    Windows.Arrange ArrangeStyle:=xlTiled
    'ActiveWindow.WindowState = xlNormal
    'Set OriWorkbook = GetWorkbook(ActiveSheet.Cells(2, 2))
    'Set RevWorkbook = GetWorkbook(ActiveSheet.Cells(2, 3))
    'wndName = RevWorkbook.Windows(1).Caption
    'OriWorkbook.Activate
    'Application.Windows.CompareSideBySideWith (wndName)
End Sub



