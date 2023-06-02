Attribute VB_Name = "SheetCompare"
Public Const cfgColOriginal = 2
Public Const cfgColRevision = 3
Public Const cfgColOption = 2

Public Const cfgRowFilename As Long = 2
Public Const cfgRowSheet As Long = 3
Public Const cfgRowRange As Long = 4
Public Const cfgRowWhat As Long = 6
Public Const cfgRowR1C1 As Long = 7
Public Const cfgRowTableHeaders As Long = 9
Public Const cfgRowHeaderRow As Long = 10
Public Const cfgRowPrimaryKey As Long = 11
Public Const cfgRowPrimKeyCol As Long = 12
Public Const cfgRowAnnotate As Long = 14
Public Const cfgRowAnnoUseFormat As Long = 15
Public Const cfgRowAnnoCellFormat As Long = 16
Public Const cfgRowAnnoComments As Long = 17
Public Const cfgRowAnnoMergeText As Long = 18
Public Const cfgRowAnnoMark As Long = 19
Public Const cfgRowAnnoMarkColumn As Long = 20
Public Const cfgRowReport As Long = 22
Public Const cfgRowRepWithMerge As Long = 23

Public Const logRowOptions = 1
Public Const logColSyncNavigation = 1
Public Const logColUpdateSheets = 4


Public Const logRowHeader = 2

Public Const logColSyncOriSheet = 1
Public Const logColSyncOriRow = 2
Public Const logColSyncOriCol = 3
Public Const logColSyncRevSheet = 4
Public Const logColSyncRevRow = 5
Public Const logColSyncRevCol = 6
Public Const logColSyncOriValue = 7
Public Const logColSyncRevValue = 8

Public Const tokenCompareDifferent = -1
Public Const tokenCompareNoComparison = 0
Public Const tokenCompareEqual = 1
Public Const tokenCompareSlightDifferent = 2



Option Explicit

Function ColStr(col As Long) As String
    ColStr = Split(Cells(, col).Address, "$")(1)
End Function

Function ColNumber(col As String) As Integer
    If IsNumeric(col) Then
        ColNumber = Int(col)
    Else
        ColNumber = Range(col + "1").Column
    End If
End Function

Sub SetValidation(cell As Range, list As String, ERROR As String, information As String, ignoreBlank As Boolean)
    Dim bShowInfo, bShowError As Boolean
    bShowError = Len(ERROR) <> 0
    bShowInfo = Len(information) <> 0
    
    If Len(list) <> 0 Then
        With cell.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=list
            .ignoreBlank = ignoreBlank
            .InCellDropdown = True
            .InputTitle = "Information"
            .ErrorTitle = "Error"
            .InputMessage = information
            .ErrorMessage = ERROR
            .ShowInput = bShowInfo
            .ShowError = bShowError
        End With
    Else
        With cell.Validation
            .Delete
            .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertInformation
            .ignoreBlank = ignoreBlank
            .InputTitle = "Information"
            .ErrorTitle = "Error"
            .InputMessage = information
            .ErrorMessage = ERROR
            .ShowInput = bShowInfo
            .ShowError = bShowError
        End With
    End If
End Sub

' This resets everything to default values
Sub ResetFields()
    With ThisWorkbook.Sheets("Config")
        
        If Not (.Cells(cfgRowFilename, cfgColOriginal).comment Is Nothing) Then
            .Cells(cfgRowFilename, cfgColOriginal).comment.Delete
        End If
        .Cells(cfgRowFilename, cfgColOriginal) = "Original WorkBook"
        
        If Not (.Cells(cfgRowFilename, cfgColRevision).comment Is Nothing) Then
            .Cells(cfgRowFilename, cfgColRevision).comment.Delete
        End If
        .Cells(cfgRowFilename, cfgColRevision) = "Revisioned WorkBook"
        
        .Cells(cfgRowSheet, cfgColOriginal) = "[ALL]"
        .Cells(cfgRowSheet, cfgColRevision) = "[ALL]"
        .Cells(cfgRowRange, cfgColOriginal) = "Auto Detect"
        .Cells(cfgRowRange, cfgColRevision) = "Auto Detect"
        
        .Cells(cfgRowWhat, cfgColOption) = "Values"
        Call SetValidation(.Cells(cfgRowWhat, cfgColOption), "Values,Formulas", "", _
        "Select Formulas if to compare the formulas instead of Values", False)
        
        .Cells(cfgRowR1C1, cfgColOption) = "NO"
        Call SetValidation(.Cells(cfgRowR1C1, cfgColOption), "YES,NO", "", _
        "If Formulas are being compared, select YES if to use the relative reference format.", False)
        
        .Cells(cfgRowTableHeaders, cfgColOption) = "NO"
        Call SetValidation(.Cells(cfgRowTableHeaders, cfgColOption), "YES,NO", "", _
        "Select YES if the Table contains Rows", False)
        
        .Cells(cfgRowHeaderRow, cfgColOriginal) = "1"
        .Cells(cfgRowHeaderRow, cfgColRevision) = "1"
        
        .Cells(cfgRowPrimaryKey, cfgColOption) = "NO"
        Call SetValidation(.Cells(cfgRowPrimaryKey, cfgColOption), "YES,NO", "", _
        "Select YES if the table to compare has a primary key for matching corresponding lines.", False)
        
        .Cells(cfgRowPrimKeyCol, cfgColOriginal) = ""
        .Cells(cfgRowPrimKeyCol, cfgColRevision) = ""
        
        .Cells(cfgRowAnnotate, cfgColOption) = "None"
        Call SetValidation(.Cells(cfgRowAnnotate, cfgColOption), "None,Original,Revision", "", _
        "Choose if annotations are to be added to the Original or Revisioned Sheet. Default is None.", False)
        
        .Cells(cfgRowAnnoUseFormat, cfgColOption) = "YES"
        Call SetValidation(.Cells(cfgRowAnnoUseFormat, cfgColOption), "YES,NO", "", _
        "If Annotation is not None, use this field to specify if to use format in cell annotations.", False)
        
        .Cells(cfgRowAnnoCellFormat, cfgColOption) = "This is a changed Cell"
        
        .Cells(cfgRowAnnoComments, cfgColOption) = "NO"
        Call SetValidation(.Cells(cfgRowAnnoComments, cfgColOption), "YES,NO", "", _
        "", False)
        
        .Cells(cfgRowAnnoMergeText, cfgColOption) = "NO"
        Call SetValidation(.Cells(cfgRowAnnoMergeText, cfgColOption), "YES,NO", "", _
        "", False)
        .Cells(cfgRowAnnoMark, cfgColOption) = "NO"
        Call SetValidation(.Cells(cfgRowAnnoMark, cfgColOption), "YES,NO", "", _
        "", False)
        .Cells(cfgRowAnnoMarkColumn, cfgColOption) = "AA"
        .Cells(cfgRowReport, cfgColOption) = "NO"
        Call SetValidation(.Cells(cfgRowReport, cfgColOption), "YES,NO", "", _
        "", False)
        .Cells(cfgRowRepWithMerge, cfgColOption) = "NO"
        Call SetValidation(.Cells(cfgRowRepWithMerge, cfgColOption), "YES,NO", "", _
        "", False)
        
    End With
    
    Call Sheet2.ResetSheet

End Sub

Sub set_YES_NO(ByRef cell As Range, ByVal b As Boolean)
    If b Then
        cell = "YES"
    Else
        cell = "NO"
    End If
End Sub

Function SelectFileWindows() As String
    Dim fDialog As Office.FileDialog

    ' Set up the File Dialog.
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
   
    With fDialog
        .Filters.Clear
        .Filters.Add "Excel files (*.xls?)", "*.xls?"
        .title = "Select Excel File to Compare"
        
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
    Dim answer As String
    answer = ""
    ' Test for the operating system.
    If Application.OperatingSystem Like "*Mac*" And _
         Val(Application.Version) >= 14 Then
            ' Is a Mac and will test if running Excel 2011 or higher.
            answer = SelectFileMac()
     Else
        ' Is Windows or any other OS
        answer = SelectFileWindows()
    End If
    SelectFileWINorMAC = answer
End Function

Sub RefreshWorkbooks()
    Dim list As String
    Dim wrkbk As Variant

    list = ""
    For Each wrkbk In Workbooks
        If wrkbk.Name <> ThisWorkbook.Name Then
            ' This avoids adding this workbook to the list
            If Len(list) = 0 Then
                list = wrkbk.Name
            Else
                list = list + "," + wrkbk.Name
            End If
        End If
    Next wrkbk
    Call SetValidation(ThisWorkbook.Worksheets("Config").Cells(cfgRowFilename, cfgColOriginal), list, "", "", False)
    Call SetValidation(ThisWorkbook.Worksheets("Config").Cells(cfgRowFilename, cfgColRevision), list, "", "", False)
End Sub

Function GetWorkbook(filename As String) As Workbook
    Dim aux As Workbook
    Dim wbk As Variant
    Dim wFilename As String
    Dim i As Integer
    
    On Error GoTo Error_Opening_File
    
    Set aux = Nothing
    For i = 1 To Workbooks.count
    Set wbk = Workbooks.Item(i)
      If IsNull(wbk) = False Then
        wFilename = wbk.FullName
        If StrComp(wbk.FullName, filename, vbTextCompare) = 0 Or _
           StrComp(wbk.Name, filename, vbTextCompare) = 0 Then
            Set aux = wbk
            Exit For
        End If
      End If
    Next i
    If aux Is Nothing Then
        Set GetWorkbook = Workbooks.Open(filename)
        ThisWorkbook.Activate ' Make sure that the current file doesn't get hidden
   Else
        Set GetWorkbook = aux
   End If
   Exit Function
Error_Opening_File:
   Set GetWorkbook = Nothing
End Function

Function GetWorkbookConfig(col As Integer) As Workbook
    Dim FileCell As Range

    On Error GoTo Error_Opening_File
    Set FileCell = ThisWorkbook.Sheets("Config").Cells(cfgRowFilename, col)
    
    Set GetWorkbookConfig = GetWorkbook(FileCell.Text)

    If GetWorkbookConfig Is Nothing Then
        ' Try using the full path on comment
        Set GetWorkbookConfig = GetWorkbook(FileCell.comment.Text)
    End If
    Exit Function
Error_Opening_File:
   Set GetWorkbookConfig = Nothing
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
    If wbk.Worksheets.count > 1 Then
        list = list + ",[ALL]"
    End If
    
    Call SetValidation(cell, list, _
            "Invalid Sheet. Select from the list", "", False)
            
    If MATCHED = False Then
        If wbk.Worksheets.count > 1 Then
            cell.Value = "[ALL]"
        Else
            'This is shit, but it didn't work any other way
            ' wbk.WorkSheets("Config").Name always gave an error
            For Each sheet In wbk.Worksheets
                list = sheet.Name
                Exit For
            Next
            list = sheet.Name
            cell.Value = list
        End If
    End If
End Sub


Function TargetSheetRange(col As Integer) As Range
    Dim sheetname As String
    Dim auxstring As String
    Dim wWorkbook As Workbook
    Dim WSheet As Worksheet

    On Error GoTo ERROR
    Set TargetSheetRange = Nothing
    
    Set wWorkbook = GetWorkbookConfig(col)
    sheetname = ThisWorkbook.Sheets("Config").Cells(cfgRowSheet, col)
    If sheetname <> "[ALL]" And IsInWorkbook(sheetname, wWorkbook) Then
        ' Obtain The Sheet
        Set WSheet = wWorkbook.Sheets(sheetname)
        'For now assuming the first row of the range to compare
        auxstring = ThisWorkbook.Sheets("Config").Cells(cfgRowRange, col).Text
        Set TargetSheetRange = WSheet.Range(auxstring)
    End If
    Exit Function
ERROR:
    Set TargetSheetRange = Nothing
    Exit Function
End Function

Function DetectHeaderRow(col As Integer) As Long
    Dim rng As Range
    
    Set rng = TargetSheetRange(col)
    If Not rng Is Nothing Then
        DetectHeaderRow = rng.Row
    Else
        DetectHeaderRow = 1
    End If
End Function

Function FindPrimaryKey(sheet As Worksheet, col As Integer) As Long
' This function is used to retrieve the primary key colum. Returns 0 if fails
    Dim rng As Range
    Dim headerRow As Long
    Dim Key As String
    Dim cell As Range
    
    FindPrimaryKey = 0
    On Error GoTo FindPrimaryKeyEnd

    headerRow = Int(ThisWorkbook.Sheets("Config").Cells(cfgRowHeaderRow, col).Value)
    Key = ThisWorkbook.Sheets("Config").Cells(cfgRowPrimKeyCol, col).Text
    
    For Each cell In sheet.Rows(headerRow).Cells()
        If Not IsEmpty(cell) And cell.Text = Key Then
            FindPrimaryKey = cell.Column
            Exit Function
        End If
    Next cell
    
FindPrimaryKeyEnd:
End Function

Sub ReloadColumnNames(col As Integer)
    Dim rng As Range
    Dim firstValidField As String
    Dim auxstring As String
    Dim colstring As String
    Dim cell As Range
    Dim firstElement, match As Boolean
    Dim headerRow As Long
    
    ' Detecting the
    Set rng = TargetSheetRange(col)
    match = False
    
    auxstring = ""
    firstElement = True
    firstValidField = auxstring
            
    On Error GoTo ERROR
        
    If ThisWorkbook.Sheets("Config").Cells(cfgRowTableHeaders, cfgColOption) = "YES" Then
        ' Will populate the list with column names
        headerRow = Int(ThisWorkbook.Sheets("Config").Cells(cfgRowHeaderRow, col).Value)
       
        If Not rng Is Nothing Then
            For Each cell In rng.Worksheet.Rows(headerRow).Cells()
                If Not IsEmpty(cell) Then
                    If firstElement Then
                        auxstring = cell.Text
                        firstElement = False
                        firstValidField = cell.Text
                    Else
                        auxstring = auxstring + "," + cell.Text
                    End If
                    If cell.Text = ThisWorkbook.Sheets("Config").Cells(cfgRowPrimKeyCol, col).Text Then
                        match = True
                    End If
                End If
                If cell.Row > headerRow Then Exit For ' Stop condition
            Next cell
        
        End If
    Else
        For Each cell In rng
            If cell.Row = rng.Row Then ' only cycle the first row
                colstring = "Column " + ColStr(cell.Column)
                If firstElement Then
                    auxstring = colstring
                    firstElement = False
                    firstValidField = auxstring
                Else
                    auxstring = auxstring + "," + colstring
                End If
                If colstring = ThisWorkbook.Sheets("Config").Cells(cfgRowPrimKeyCol, col).Text Then
                    match = True
                End If
            Else
                Exit For
            End If
        Next cell
    End If
    
    Call SetValidation(ThisWorkbook.Sheets("Config").Cells(cfgRowPrimKeyCol, col), _
        auxstring, _
        "", "", True)
    If Not match Then
        ThisWorkbook.Sheets("Config").Cells(cfgRowPrimKeyCol, col) = firstValidField
    End If
    
ERROR:
End Sub

Function Detect_Table(ws As Worksheet) As String
    Dim x, xmin, xmax As Integer
    Dim y, ymin, ymax As Long
    Dim blankRows As Integer
    Dim blankColumns As Integer

    xmin = ws.Columns.count
    xmax = 1
    ymax = 1
    ymin = ws.Rows.count
    blankRows = 1

    For y = 1 To ws.Rows.count
        blankColumns = 1
        For x = 1 To ws.Columns.count
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
    Dim filename As String
    
    filename = SelectFileWINorMAC()
    If Len(filename) <> 0 Then
        ' Set the Filename on the proper cell. The rest will be automatically triggered
        ThisWorkbook.ActiveSheet.Cells(cfgRowFilename, cfgColOriginal) = filename
    End If
End Sub

Sub Open_RevisionWorkbook()
    Dim filename As String
    
    filename = SelectFileWINorMAC()
    If Len(filename) <> 0 Then
        ' Set tje filename on the proper cell. The rest will be automatically triggered
        ThisWorkbook.ActiveSheet.Cells(cfgRowFilename, cfgColRevision) = filename
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


Function ColumnMatch(OriSheet As Worksheet, OriRange As String, RevSheet As Worksheet, RevRange As String, ByRef OriCols() As Long, ByRef RevCols() As Long) As Long
    Dim count As Long
    Dim oCol, rCol As Long
    Dim OCell, RCell As Range
    Dim OHeaderRow As Long, RHeaderRow As Long
    Dim OColumns, RColumns As Long
    Dim auxRange As Range

    count = 0
        
    OHeaderRow = ThisWorkbook.Sheets("Config").Cells(cfgRowHeaderRow, cfgColOriginal)
    RHeaderRow = ThisWorkbook.Sheets("Config").Cells(cfgRowHeaderRow, cfgColRevision)
    
    Set auxRange = Range(OriRange)
    OColumns = auxRange.Columns.count + auxRange.Column
    Set auxRange = Range(RevRange)
    RColumns = auxRange.Columns.count + auxRange.Column
    
    For Each OCell In OriSheet.Rows(OHeaderRow).Cells
        If OCell.Column > OColumns Then Exit For
        If Len(OCell.Text) > 0 Then
            For Each RCell In RevSheet.Rows(RHeaderRow).Cells
                If OCell.Column > RColumns Then Exit For
                If OCell.Text = RCell.Text Then
                    ' TODO: Also control that the Revisioned column wasn't already added
                    count = count + 1
                    ReDim Preserve OriCols(count)
                    ReDim Preserve RevCols(count)
                    OriCols(count - 1) = OCell.Column
                    RevCols(count - 1) = RCell.Column
                    Exit For
                End If
            Next RCell
        End If
    Next OCell
    ReDim Preserve OriCols(count)
    ReDim Preserve RevCols(count)
    ColumnMatch = count
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
    
    Dim Rep_Workbook As Workbook
    
    Dim Log_Sheet As Worksheet
    
    Dim sheetIndex, comparedSheets, reportedSheets As Integer
    Dim bMultipleSheets As Boolean
    Dim oRow_Count, rRow_Count As Long
    Dim oCols() As Long
    Dim rCols() As Long
    Dim iCol_Count As Double, oCol As Long, rCol As Long
    Dim Ori_iRow_Start As Long, Ori_iCol_Start As Long
    Dim Rev_iRow_Start As Long, Rev_iCol_Start As Long
    Dim Rep_iRow_Start, log_iRow As Long

    ' Backup Global Variables
    Dim bbUpdateSheets, bbSyncNavigation As Boolean
    
    Dim oRow, rRow, rrRow, passedPtr, passedCount, passedRowsSize As Long
    passedRowsSize = 10
    Dim passedRows() As Long
    ReDim passedRows(passedRowsSize)
    
    passedCount = 0 ' Initializes with a zero count
    
    
    Dim iCol As Double, iCol1 As Integer
    Dim Ori_Data As String, Rev_Data As String
    Dim tempWidth As Integer
    
    Dim iRepCount As Integer, iRepRow As Integer
    Dim bDoReport As Boolean, bFirstDifference As Boolean, bRowChanged As Boolean
    Dim bMakeAnnotation As Boolean, bPrimaryKey As Boolean
    Dim oPrimaryKeyCol, rPrimaryKeyCol As Long

    Dim AnnotationSheet As Integer
    Dim bApplyChangeFormat As Boolean, bInsertComment As Boolean
    Dim ChangedCellFormat As Range
    Dim targetCell As Range
    Dim bTextMerge, bReportMerge As Boolean
    Dim bCompareFormulas As Boolean, bR1C1Format As Boolean, bHasHeaders As Boolean
    Dim comment As String

    Dim annotateColumn As Integer

    Dim iDiffCount As Double
    
    Set Cfg_Sheet = ThisWorkbook.Sheets("Config")
    'Assign the Workbook File Name along with its Path
    Set Ori_Workbook = GetWorkbookConfig(cfgColOriginal) ' OriginalWorkbook Filename
    Set Rev_Workbook = GetWorkbookConfig(cfgColRevision) ' Revision Workbook Filename
    
    If Ori_Workbook Is Nothing Then
        ' Try using the full path on comment
        
        MsgBox "File """ & Cfg_Sheet.Cells(cfgRowFilename, cfgColOriginal) & """ doesn't exist"
        Exit Sub
    End If

    If Rev_Workbook Is Nothing Then
        MsgBox "File """ & Cfg_Sheet.Cells(cfgRowFilename, cfgColRevision) & """ doesn't exist"
        Exit Sub
    End If
    
    
    bMakeAnnotation = StrComp(Cfg_Sheet.Cells(cfgRowAnnotate, cfgColOption).Value, "None", vbTextCompare) <> 0 ' Make annotation on
    Set ChangedCellFormat = Cfg_Sheet.Cells(cfgRowAnnoCellFormat, cfgColOption) ' Changed Cell Format
    
    If bMakeAnnotation Then
        bApplyChangeFormat = (StrComp(Cfg_Sheet.Cells(cfgRowAnnoUseFormat, cfgColOption), "YES", vbTextCompare) = 0) ' Use Format to identify changes
        bInsertComment = (StrComp(Cfg_Sheet.Cells(cfgRowAnnoComments, cfgColOption), "YES", vbTextCompare) = 0)     ' Insert difference in comments
        bTextMerge = (StrComp(Cfg_Sheet.Cells(cfgRowAnnoMergeText, cfgColOption), "YES", vbTextCompare) = 0) ' Use Merge Text in the differences
        
        ' if annotateColumn is zero, the annotation is not done
        If (StrComp(Cfg_Sheet.Cells(cfgRowAnnoMark, cfgColOption), "YES", vbTextCompare) = 0) Then  ' Mark modified Rows
            annotateColumn = ColNumber(Cfg_Sheet.Cells(cfgRowAnnoMarkColumn, cfgColOption).Text)   ' use column
        Else
            annotateColumn = 0
        End If
    End If
     
    'Comparing Values of Formulas
    ' This is done here so thata the bTextMerge can be direcly overriden
    If StrComp(Cfg_Sheet.Cells(cfgRowWhat, cfgColOption).Value, "Formulas", vbTextCompare) = 0 Then  ' What to compare
        bCompareFormulas = True
        bR1C1Format = (StrComp(Cfg_Sheet.Cells(cfgRowR1C1, cfgColOption), "YES", vbTextCompare) = 0)  ' Use R1C1 Format
        bTextMerge = False
    Else
        bCompareFormulas = False
    End If
     
    bDoReport = (StrComp(Cfg_Sheet.Cells(cfgRowReport, cfgColOption), "YES", vbTextCompare) = 0) ' Create Report
    If bDoReport Then
        bReportMerge = (StrComp(Cfg_Sheet.Cells(cfgRowRepWithMerge, cfgColOption), "YES", vbTextCompare) = 0)  ' Use merge in the Report
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "Creating the report..."


    Set Rep_Workbook = Nothing
    Application.DisplayAlerts = True
    
    'With Ori_Workbook object, now it is possible to pull any data from it
    'Read Data From Each Sheets of Both Excel Files & Compare Data
    
    sheetIndex = 1  ' Index of the Sheet being compared
    reportedSheets = 0
    comparedSheets = 0
    iDiffCount = 0
    Set Log_Sheet = ThisWorkbook.Sheets("Diff")
    bbSyncNavigation = "YES" = Log_Sheet.Cells(logRowOptions, logColSyncNavigation + 1).Value ' Backup Value
    bbUpdateSheets = "YES" = Log_Sheet.Cells(logRowOptions, logColUpdateSheets + 1).Value ' Backup Value
    ' This blocks the updates due to changes on the Report.
    Call Sheet2.BlockUpdates
    Call Sheet2.ResetSheet
    Log_Sheet.Cells(logRowHeader, logColSyncOriValue) = Ori_Workbook.FullName
    Log_Sheet.Cells(logRowHeader, logColSyncRevValue) = Rev_Workbook.FullName
    bMultipleSheets = False
    
    bPrimaryKey = Cfg_Sheet.Cells(cfgRowPrimaryKey, cfgColOption).Value = "YES"
    bHasHeaders = Cfg_Sheet.Cells(cfgRowTableHeaders, cfgColOption).Value = "YES"
        
    ' Cycling through Sheets
    Do While True
        If Cfg_Sheet.Cells(cfgRowSheet, cfgColOriginal) = "[ALL]" Then
            If Cfg_Sheet.Cells(cfgRowSheet, cfgColRevision) = "[ALL]" Then
                bMultipleSheets = True
TRY_NEXT:
                If sheetIndex > Ori_Workbook.Sheets.count Then
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
            Ori_SheetName = Cfg_Sheet.Cells(cfgRowSheet, cfgColOriginal)
            If sheetIndex > 1 Then
                GoTo EXIT_LOOP
            End If
        
            If Cfg_Sheet.Cells(cfgRowSheet, cfgColRevision) = "[ALL]" Then
                If IsInWorkbook(Ori_SheetName, Rev_Workbook) Then
                    Rev_SheetName = Ori_SheetName
                Else
                    ' TODO: create an error flag so than an error is reported
                    GoTo EXIT_LOOP
                End If
            Else
                Rev_SheetName = Cfg_Sheet.Cells(cfgRowSheet, cfgColRevision)
            End If
        End If
        sheetIndex = sheetIndex + 1

        Set Ori_Sheet = Ori_Workbook.Sheets(Ori_SheetName)
        Set Rev_Sheet = Rev_Workbook.Sheets(Rev_SheetName)
        
        ' Getting Row and Column Start
        Ori_Range = Cfg_Sheet.Cells(cfgRowRange, cfgColOriginal) ' Original Range to compare
        Rev_Range = Cfg_Sheet.Cells(cfgRowRange, cfgColRevision) ' Revision Range to compare
    
        If Ori_Range = "Auto Detect" Then
            Ori_Range = Detect_Table(Ori_Sheet)
        End If
        
        If Rev_Range = "Auto Detect" Then
            Rev_Range = Detect_Table(Rev_Sheet)
        End If
            
        Ori_iRow_Start = Range(Ori_Range).Row ' Setting the row Start
        Rev_iRow_Start = Range(Rev_Range).Row ' Setting the row Start
        oRow_Count = Range(Ori_Range).Rows.count
        rRow_Count = Range(Rev_Range).Rows.count
        
        If bMultipleSheets And bPrimaryKey = False Then ' Calculates the highest row count
            If oRow_Count < rRow_Count Then
                rRow_Count = oRow_Count
            Else
                oRow_Count = rRow_Count
            End If
            
        End If
        
            
        If bHasHeaders Then ' Table has headers
            iCol_Count = ColumnMatch(Ori_Sheet, Ori_Range, Rev_Sheet, Rev_Range, oCols, rCols)
            Ori_iCol_Start = 0 ' this will not be used
            Rev_iCol_Start = 0 ' this will not be used
            
        Else
            Ori_iCol_Start = Range(Ori_Range).Column
            Rev_iCol_Start = Range(Rev_Range).Column

            'Calculating count of rows and columns to process
            iCol_Count = Range(Ori_Range).Columns.count
            
            'Checking column sizes are equal
            If Range(Ori_Range).Columns.count <> Range(Rev_Range).Columns.count Then
        
                'Assuming the smaller between the two sheets
                If Range(Rev_Range).Columns.count < iCol_Count Then
                    iCol_Count = Range(Rev_Range).Columns.count
                End If
        
                MsgBox "Ranges for Sheet '" & Ori_Sheet.Name & "' differ in size." & vbCr _
                     & "Assuming the Smaller sizes" & vbCr _
                     & "Comparing  " & iCol_Count & " column(s)"
            End If
        End If
    
        ' if is empty primary key is not used
        If bPrimaryKey Then
            If Cfg_Sheet.Cells(cfgRowTableHeaders, cfgColOption).Value = "YES" Then ' table has headers
                oPrimaryKeyCol = FindPrimaryKey(Ori_Sheet, cfgColOriginal)
                rPrimaryKeyCol = FindPrimaryKey(Rev_Sheet, cfgColRevision)
                If oPrimaryKeyCol = 0 Or rPrimaryKeyCol = 0 Then
                    bPrimaryKey = False
                    MsgBox ("Failed to find primary key. Continuing without primary key")
                End If
            Else
                oPrimaryKeyCol = ColNumber(Mid(Cfg_Sheet.Cells(cfgRowPrimKeyCol, cfgColOriginal).Value, Len("Column "), 3))
                rPrimaryKeyCol = ColNumber(Mid(Cfg_Sheet.Cells(cfgRowPrimKeyCol, cfgColRevision).Value, Len("Column "), 3))
            End If
        
        End If
        
        If StrComp(Cfg_Sheet.Cells(cfgRowAnnotate, cfgColOption).Value, "Original", vbTextCompare) = 0 Then
            AnnotationSheet = 1 ' Original

        Else
            If StrComp(Cfg_Sheet.Cells(cfgRowAnnotate, cfgColOption).Value, "Revision", vbTextCompare) = 0 Then
                AnnotationSheet = 2 ' Revision

            Else
                AnnotationSheet = 0 ' None
            End If
        End If
        
        'With Ori_Workbook object, now it is possible to pull any data from it
        'Read Data From Each Sheets of Both Excel Files & Compare Data
        
        bFirstDifference = True ' This is used for Starting the Report
        
        oRow = 0
        rRow = 0
        rrRow = 0
        Do While oRow < oRow_Count Or bPrimaryKey = True
                                      ' This has a different stop condition
            If bPrimaryKey Then 'Find the corresponding revisioned Row
                Ori_Data = Ori_Sheet.Cells(oRow + Ori_iRow_Start, oPrimaryKeyCol).Text
                ' First search on the passedRevisioned
                passedPtr = 0
                rRow = -1 ' Flag for match not found
                Do While passedPtr < passedCount
                    Rev_Data = Rev_Sheet.Cells(passedRows(passedPtr) + Rev_iRow_Start, rPrimaryKeyCol).Text
                     If Rev_Data = Ori_Data Then
                        rRow = passedRows(passedPtr)
                        ' continue till the end making a shift of the end of the array to optimise space
                        passedCount = passedCount - 1
                        Do While passedPtr < passedCount
                            passedRows(passedPtr) = passedRows(passedPtr + 1)
                            passedPtr = passedPtr + 1
                        Loop
                     End If
                     passedPtr = passedPtr + 1
                Loop
                
                If rRow = -1 Then ' If it was not found
                    ' Will continue searching from  rRow forward till a first match is found. misses will be added to the passedRows array
                    Do While rrRow < rRow_Count
                        Rev_Data = Rev_Sheet.Cells(rrRow + Rev_iRow_Start, rPrimaryKeyCol).Text
                        If Rev_Data = Ori_Data Then ' Simply exit here and let the cycle continue
                            rRow = rrRow
                            rrRow = rrRow + 1 ' This is done because this cell was covered
                            Exit Do
                        Else
                            ' In this case add the rRow to the list of passedRows array
                            If passedCount >= passedRowsSize Then
                                passedRowsSize = passedRowsSize + 10
                                ReDim Preserve passedRows(passedRowsSize)
                            End If
                            passedRows(passedCount) = rrRow
                            passedCount = passedCount + 1
                        End If
                        rrRow = rrRow + 1
                    Loop
                    If rRow = -1 Then ' If it was not found again, check the stop condition
                        If oRow >= oRow_Count Then
                            'No more original sheet to show : Empty the passed Rows list
                            If passedCount > 0 Then
                                rRow = passedRows(0)
                                'Delete the first element and decrease the count
                                ' This implementation takes more time, but I don't need to declare another variable. I'm lazy
                                passedCount = passedCount - 1
                                passedPtr = 0
                                Do While passedPtr < passedCount
                                    passedRows(passedPtr) = passedRows(passedPtr + 1)
                                    passedPtr = passedPtr + 1
                                Loop
                            Else
                                Exit Do ' Finished comparison task
                            End If
                        End If
                    End If
                End If
            End If
            
            bRowChanged = False

            For iCol = 0 To iCol_Count - 1
                
                If bHasHeaders Then ' Will index the list of columns to compare
                    oCol = oCols(iCol)
                    rCol = rCols(iCol)
                Else ' Otherwise it scans sequentialy begining on
                    oCol = iCol + Ori_iCol_Start
                    rCol = iCol + Rev_iCol_Start
                End If
                
                If rRow = -1 Then
                    Rev_Data = ""
                Else
                    If bCompareFormulas Then
                        If Rev_Sheet.Cells(rRow + Rev_iRow_Start, rCol).HasFormula Then
                            If bR1C1Format Then
                                Rev_Data = Rev_Sheet.Cells(rRow + Rev_iRow_Start, rCol).FormulaR1C1Local
                            Else
                                Rev_Data = Rev_Sheet.Cells(rRow + Rev_iRow_Start, rCol).FormulaLocal
                            End If
                        Else
                            Rev_Data = Rev_Sheet.Cells(rRow + Rev_iRow_Start, rCol).Text
                        End If
                    Else
                        Rev_Data = Rev_Sheet.Cells(rRow + Rev_iRow_Start, rCol).Text
                    End If
                End If
                If oRow >= oRow_Count Then
                    Ori_Data = ""
                Else
                    If bCompareFormulas Then
                        If Ori_Sheet.Cells(oRow + Ori_iRow_Start, oCol).HasFormula Then
                            If bR1C1Format Then
                                Ori_Data = Ori_Sheet.Cells(oRow + Ori_iRow_Start, oCol).FormulaR1C1Local
                            Else
                                Ori_Data = Ori_Sheet.Cells(oRow + Ori_iRow_Start, oCol).FormulaLocal
                            End If
                        Else
                            Ori_Data = Ori_Sheet.Cells(oRow + Ori_iRow_Start, oCol).Text
                        End If
                    Else
                        Ori_Data = Ori_Sheet.Cells(oRow + Ori_iRow_Start, oCol).Text
                    End If
                End If
                'Compare Data From Excel Sheets & Highlight the Mismatches
                If Ori_Data <> Rev_Data Then
                    iDiffCount = iDiffCount + 1
                    'Avoid problems with formulas
                    If bCompareFormulas Then
                        If Mid(Ori_Data, 1, 1) = "=" Then
                            Ori_Data = "'" & Ori_Data
                         End If
                        If Mid(Rev_Data, 1, 1) = "=" Then
                            Rev_Data = "'" & Rev_Data
                         End If
                    End If
                    'Log the difference
                    Log_Sheet.Cells(iDiffCount + logRowHeader, logColSyncOriSheet).Value = Ori_Sheet.Name
                    Log_Sheet.Cells(iDiffCount + logRowHeader, logColSyncRevSheet).Value = Rev_Sheet.Name
                    If oRow >= oRow_Count Then
                        Log_Sheet.Cells(iDiffCount + logRowHeader, logColSyncOriRow).Value = "New"
                    Else
                        Log_Sheet.Cells(iDiffCount + logRowHeader, logColSyncOriRow).Value = oRow + Ori_iRow_Start
                        Log_Sheet.Cells(iDiffCount + logRowHeader, logColSyncOriCol).Value = ColStr(oCol)
                        Log_Sheet.Cells(iDiffCount + logRowHeader, logColSyncOriValue).Value = Ori_Data
                    
                    End If
                    If rRow = -1 Then
                        Log_Sheet.Cells(iDiffCount + logRowHeader, logColSyncRevRow).Value = "Deleted"
                    Else
                        Log_Sheet.Cells(iDiffCount + logRowHeader, logColSyncRevRow).Value = rRow + Rev_iRow_Start
                        Log_Sheet.Cells(iDiffCount + logRowHeader, logColSyncRevCol).Value = ColStr(rCol)
                        Log_Sheet.Cells(iDiffCount + logRowHeader, logColSyncRevValue).Value = Rev_Data
                    End If
                    'Log_Sheet.Cells(iDiffCount + logRowHeader, 9).Value = Ori_Sheet.Cells(iRow + Ori_iRow_Start, oCol).Address
                    'Log_Sheet.Cells(iDiffCount + logRowHeader, 10).Value = Rev_Sheet.Cells(iRow + Rev_iRow_Start, rCol).Address
                    
                    'Report the difference
                    If bDoReport Then
                        'If it is the first difference there is some stuff to do
                        If bRowChanged = False Then
                            'create the report if needed
                            If bFirstDifference Then
                                reportedSheets = reportedSheets + 1
                                If Rep_Workbook Is Nothing Then
                                    Set Rep_Workbook = Application.Workbooks.Add() ' Creates a new Workbook
                                End If
                                If reportedSheets <= Rep_Workbook.Worksheets.count Then
                                    Set Rep_Sheet = Rep_Workbook.Worksheets(reportedSheets)
                                Else
                                    Set Rep_Sheet = Rep_Workbook.Worksheets.Add(After:=Rep_Workbook.Worksheets(Rep_Workbook.Worksheets.count), Type:=xlWorksheet)
                                End If
                                'Set Rep_Sheet = Workbooks.Add.Sheets("Config")
                                Rep_Sheet.Name = "Diff-" & Ori_Sheet.Name
                        
                                'Rep_Sheet.Activate
                                Rep_Sheet.Cells.Clear ' Clear everything on the sheet
    
                                Rep_Sheet.Cells(1, 1).Value = "Original Range"
                                Rep_Sheet.Cells(2, 1).Value = "Revision Range"
                                Rep_Sheet.Cells(1, 2).Value = "[" & Ori_Workbook.Name & "]" & Ori_Sheet.Name & "!" & Ori_Range
                                Rep_Sheet.Cells(2, 2).Value = "[" & Rev_Workbook.Name & "]" & Rev_Sheet.Name & "!" & Rev_Range
                                Rep_Sheet.Cells(3, 1).Value = "Differences Found"
                                iRepRow = 0
                                Rep_iRow_Start = 4
                                'Format the widths to match the Original
                                For iCol1 = 0 To iCol_Count - 1
                                    If bHasHeaders Then ' Will index the list of columns to compare
                                        tempWidth = Ori_Sheet.Columns(oCols(iCol1)).ColumnWidth
                                    Else ' Otherwise it scans sequentialy begining on
                                        tempWidth = Ori_Sheet.Columns(iCol1 + Ori_iCol_Start).ColumnWidth
                                    End If
                                    If tempWidth > 250 Then
                                            tempWidth = 250
                                    End If
                                    Rep_Sheet.Columns(iCol1 + 2).ColumnWidth = tempWidth
                                Next iCol1
                                bFirstDifference = False
                            Else
                                iRepRow = iRepRow + 1
                            End If
                            
                            Rep_Sheet.Cells(iRepRow + Rep_iRow_Start, 1).Value = "Ori:" & (oRow + Ori_iRow_Start) & "=Rev:" & (rRow + Rev_iRow_Start)
                            
                            ' Copy from Original
                            If bCompareFormulas Then
                                For iCol1 = 0 To iCol_Count - 1
                                    If bHasHeaders Then
                                        Rep_Sheet.Cells(iRepRow + Rep_iRow_Start, oCols(iCol1) + 2).Formula = Ori_Sheet.Cells(oRow + Ori_iRow_Start, oCols(iCol1)).Formula
                                    Else
                                        Rep_Sheet.Cells(iRepRow + Rep_iRow_Start, iCol1 + 2).Formula = Ori_Sheet.Cells(oRow + Ori_iRow_Start, iCol1 + Ori_iCol_Start).Formula
                                    End If
                                Next iCol1
                            Else
                                For iCol1 = 0 To iCol_Count - 1
                                    If bHasHeaders Then
                                        Rep_Sheet.Cells(iRepRow + Rep_iRow_Start, oCols(iCol1) + 2).Formula = Ori_Sheet.Cells(oRow + Ori_iRow_Start, oCols(iCol1))
                                    Else
                                        Rep_Sheet.Cells(iRepRow + Rep_iRow_Start, iCol1 + 2).Value = Ori_Sheet.Cells(oRow + Ori_iRow_Start, iCol1 + Ori_iCol_Start)
                                    End If
                                Next iCol1
                            End If
                        End If
                        ' Highlight the Mismatches
                        If bHasHeaders Then
                            With Rep_Sheet.Cells(iRepRow + Rep_iRow_Start, oCols(iCol) + 2)
                                .Interior.Pattern = ChangedCellFormat.Interior.Pattern
                                .Interior.PatternColorIndex = ChangedCellFormat.Interior.PatternColorIndex
                                .Interior.ThemeColor = ChangedCellFormat.Interior.ThemeColor
                                .Interior.TintAndShade = ChangedCellFormat.Interior.TintAndShade
                                .Interior.PatternTintAndShade = ChangedCellFormat.Interior.PatternTintAndShade
                            End With
                        Else
                            With Rep_Sheet.Cells(iRepRow + Rep_iRow_Start, iCol + 2)
                                .Interior.Pattern = ChangedCellFormat.Interior.Pattern
                                .Interior.PatternColorIndex = ChangedCellFormat.Interior.PatternColorIndex
                                .Interior.ThemeColor = ChangedCellFormat.Interior.ThemeColor
                                .Interior.TintAndShade = ChangedCellFormat.Interior.TintAndShade
                                .Interior.PatternTintAndShade = ChangedCellFormat.Interior.PatternTintAndShade
                            End With
                        End If
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
                            If bHasHeaders Then
                                Call MergeText(Ori_Data, Rev_Data, Rep_Sheet.Cells(iRepRow + Rep_iRow_Start, oCols(iCol) + 2))
                            Else
                                Call MergeText(Ori_Data, Rev_Data, Rep_Sheet.Cells(iRepRow + Rep_iRow_Start, iCol + 2))
                            End If
                        Else
                            If bHasHeaders Then
                                Rep_Sheet.Cells(iRepRow + Rep_iRow_Start, oCols(iCol) + 2) = _
                                           "Changed from: " + vbCr + vbLf + _
                                           Ori_Data + vbCr + vbLf + _
                                           "To: " + _
                                           vbCr + vbLf + Rev_Data
                            Else
                                Rep_Sheet.Cells(iRepRow + Rep_iRow_Start, iCol + 2) = _
                                           "Changed from: " + vbCr + vbLf + _
                                           Ori_Data + vbCr + vbLf + _
                                           "To: " + _
                                           vbCr + vbLf + Rev_Data
                            End If
                        End If
                    End If
                    
                    ' Annotating one of the sheets
                    If bMakeAnnotation Then
                        If AnnotationSheet = 1 Then 'Original
                            If oRow <> -1 Then
                                Set targetCell = Ori_Sheet.Cells(oRow + Ori_iRow_Start, oCol + Ori_iCol_Start)
                            End If
                        Else ' Revision
                            If rRow <> -1 Then
                                Set targetCell = Rev_Sheet.Cells(rRow + Rev_iRow_Start, rCol + Rev_iCol_Start)
                            End If
                        End If
                        
                    
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
                                If Ori_Data = "" Then
                                    comment = "Added:" + vbCr + vbLf + Rev_Data
                                Else
                                    If Rev_Data = "" Then
                                        comment = "Deleted"
                                    Else
                                        comment = "Changed from: " + vbCr + vbLf + _
                                               Ori_Data + vbCr + vbLf + _
                                               "To: " + _
                                               vbCr + vbLf + Rev_Data
                                    End If
                                End If
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
                If AnnotationSheet = 1 Then ' Original
                    Ori_Sheet.Cells(oRow + Ori_iRow_Start, annotateColumn).Value = "Modified"
                Else ' Revision
                    Rev_Sheet.Cells(oRow + Rev_iRow_Start, annotateColumn).Value = "Modified"
                End If
            End If
    
            ' Informing the user
            If (oRow Mod 30) = 0 Then
                Application.StatusBar = "Progress (sheet:" & Ori_SheetName & " row:" & oRow & ")"
            End If
            
            oRow = oRow + 1
            rRow = rRow + 1
            
        Loop ' end of cycling rows
    
        comparedSheets = comparedSheets + 1
        
    Loop ' End Of cycling through sheets
EXIT_LOOP:
    
    '''''Process Completed
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    ' Now formatting the Report Sheet
    ThisWorkbook.Sheets("Diff").Activate
    
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.AutoFilter
    Columns("G:H").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 2
    End With
    ActiveWindow.FreezePanes = True
    Range("G3").Select
    
    Call Sheet2.InitVars
    Call set_YES_NO(Log_Sheet.Cells(logRowOptions, logColSyncNavigation + 1), bbSyncNavigation) ' Restore Value
    Call set_YES_NO(Log_Sheet.Cells(logRowOptions, logColUpdateSheets + 1), bbUpdateSheets) ' Restore Value
    Call ArrangeWindows
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

Function CompareTokens(Ori As String, Rev As String) As Integer
    Dim o, r, x, y As Integer
    Dim OriLen, RevLen As Integer
    Dim matches As Integer

    If Rev = Ori Then ' Exact Match
        CompareTokens = tokenCompareEqual ' Signals an exact match
        Exit Function
    Else
        OriLen = Len(Ori)
        RevLen = Len(Rev)
        If Abs(OriLen - RevLen) < 2 And OriLen > 2 Then
            o = 1
            r = 1
            x = 0
            y = 0
            matches = 0
            Do While o + x < OriLen And r + x < RevLen
                If Mid(Ori, o + x, 1) = Mid(Rev, r + y, 1) Then
                    o = o + x + 1
                    r = r + y + 1
                    matches = matches + 1
                    x = 0
                    y = 0
                Else
                    If Mid(Ori, o + y, 1) = Mid(Rev, r + x, 1) Then
                        o = o + y + 1
                        r = r + x + 1
                        matches = matches + 1
                        x = 0
                        y = 0
                    Else
                        y = y + 1
                        ' Keep trying
                        If y >= x Then
                            y = 0
                            x = x + 1
                        End If
                    End If
                End If
            Loop
            If (matches + 2) >= OriLen Then
                CompareTokens = tokenCompareSlightDifferent ' Still an acceptable match
            Else
                CompareTokens = tokenCompareDifferent ' Consider a different word
            End If
        Else
            CompareTokens = tokenCompareDifferent ' Signals a completely different
        End If
    End If
End Function

Sub MergeTokens(Ori As String, Rev As String, ByRef Mrg As String, ByRef mrk As String)
    Dim o, r, x, y, i As Integer
    Dim OriLen, RevLen As Integer

    OriLen = Len(Ori)
    RevLen = Len(Rev)
    o = 1
    r = 1
    x = 0
    y = 0

    Mrg = ""
    mrk = ""
   
    Do While o + x <= OriLen And r + x <= RevLen
        If Mid(Ori, o + y, 1) = Mid(Rev, r + x, 1) Then
            'Copy Ori
            If y > 0 Then
                Mrg = Mrg + Mid(Ori, o, y)
                For i = 0 To y - 1
                    mrk = mrk + "X"
                Next i
            End If
            If x > 0 Then
                Mrg = Mrg + Mid(Rev, r, x)
                For i = 0 To x - 1
                    mrk = mrk + "_"
                Next i
            End If
            Mrg = Mrg + Mid(Ori, o + y, 1)
            mrk = mrk + " "
            o = o + x + 1
            r = r + y + 1
            x = 0
            y = 0
        Else
            If Mid(Ori, o + x, 1) = Mid(Rev, r + y, 1) Then
                If x > 0 Then
                    Mrg = Mrg + Mid(Ori, o, x)
                    For i = 0 To x - 1
                        mrk = mrk + "X"
                    Next i
                End If
                If y > 0 Then
                    Mrg = Mrg + Mid(Rev, r, y)
                    For i = 0 To y - 1
                        mrk = mrk + "_"
                    Next i
                End If
                Mrg = Mrg + Mid(Ori, o + x, 1)
                mrk = mrk + " "
                o = o + x + 1
                r = r + y + 1
                y = y + 1
                x = 0
                y = 0
            Else
                y = y + 1
                ' Keep trying
                If y >= x Then
                    y = 0
                    x = x + 1
                End If
            End If
        End If
    Loop

    ' Now complete till the end with the remaining
    Do While o <= OriLen
        Mrg = Mrg + Mid(Ori, o, 1)
        mrk = mrk + "X"
        o = o + 1
    Loop
    Do While r <= RevLen
        Mrg = Mrg + Mid(Rev, r, 1)
        mrk = mrk + "_"
        r = r + 1
    Loop

End Sub


Sub MergeArrays(ByRef Ori() As String, ByRef Rev() As String, ByRef Mrg() As String, ByRef mrk() As String)

    Dim o, r, m, x, y As Integer
    Dim OriLen, RevLen As Integer
    Dim inSync As Boolean
    Dim compResult As Integer
    
    Dim i As Integer

    OriLen = UBound(Ori)
    RevLen = UBound(Rev)

    ReDim Mrg(OriLen + RevLen)
    ReDim mrk(OriLen + RevLen)

    o = 0
    r = 0
    m = 0
    y = 0
    x = 0
   
    Do
        compResult = tokenCompareNoComparison 'This means no comparison was done
        If r + x < RevLen And o + y < OriLen Then
            compResult = CompareTokens(Rev(r + x), Ori(o + y))
            If compResult > tokenCompareDifferent Then GoTo MATCHED
        End If
        If o + x < OriLen And r + y < RevLen Then
            compResult = CompareTokens(Ori(o + x), Rev(r + y))
            If compResult > tokenCompareDifferent Then
                ' Revert Y and X'
                i = x
                x = y
                y = i
            End If
        End If
    
MATCHED:
        If compResult > tokenCompareNoComparison Then
                ' somethng inserted and deleted
                For i = 0 To y - 1
                    Mrg(m) = Ori(o)
                    mrk(m) = "X"
                    m = m + 1
                    o = o + 1
                Next i
                For i = 0 To x - 1
                    Mrg(m) = Rev(r)
                    mrk(m) = "_"
                    m = m + 1
                    r = r + 1
                Next i
                ' Matched
                If compResult = tokenCompareEqual Then
                    Mrg(m) = Ori(o)
                    mrk(m) = "."
                Else
                    Call MergeTokens(Ori(o), Rev(r), Mrg(m), mrk(m))
                End If
                m = m + 1
                o = o + 1
                r = r + 1
                x = 0
                y = 0
        Else
            ' Keep trying
            y = y + 1
            If y >= x Then
                y = 0
                x = x + 1
            End If
        End If
    Loop While compResult <> tokenCompareNoComparison

    ' Now complete till the end with the remaining
    Do While o < OriLen
        Mrg(m) = Ori(o)
        mrk(m) = "X"
        m = m + 1
        o = o + 1
    Loop
    Do While r < RevLen
        Mrg(m) = Rev(r)
        mrk(m) = "_"
        m = m + 1
        r = r + 1
    Loop
    ReDim Preserve Mrg(m)
    ReDim Preserve mrk(m)

End Sub


Sub MergeText(OriText As String, RevText As String, cell As Range)
    Dim Ori() As String
    Dim Rev() As String
    
    Dim Mrg() As String
    Dim mrk() As String
    Dim msg As String
    Dim i, j, p, l As Integer
    Dim LO As Integer, LR As Integer
    Dim m As String

    LO = Len(OriText)
    LR = Len(RevText)
                        
    If LO = 0 Then ' If the original text is empty
        ' Insert the revisioned text underlined
        cell.Value = RevText
        cell.Characters(Start:=1, length:=LR).Font.Underline = xlUnderlineStyleSingle
    Else
        If LR = 0 Then  'If the Revisioned text is empty
            ' Insert the original text underlined
            cell.Value = OriText
            cell.Characters(Start:=1, length:=LO).Font.Strikethrough = True
        Else
            ' Otherwise a comparison is made
            If IsNumeric(RevText) = False Then
                Ori = Atomize(OriText, " .,;:")
                Rev = Atomize(RevText, " .,;:")
            
                Call MergeArrays(Ori, Rev, Mrg, mrk)
            
                msg = ArrayToString(Mrg)
                p = 1
                With cell
                    .Value = msg
                    For i = 0 To UBound(Mrg)
                        l = Len(Mrg(i))
                        m = mrk(i)
                        If Len(m) > 1 Then ' This a a merged word
                            For j = 1 To l
                                If Mid(m, j, 1) = "X" Then
                                    .Characters(Start:=p, length:=1).Font.Strikethrough = True
                                Else
                                    If Mid(m, j, 1) = "_" Then
                                        .Characters(Start:=p, length:=1).Font.Underline = xlUnderlineStyleSingle
                                    End If
                                End If
                                p = p + 1
                            Next j
                        Else
                            If m = "X" Then
                                .Characters(Start:=p, length:=l).Font.Strikethrough = True
                            Else
                                If m = "_" Then
                                    .Characters(Start:=p, length:=l).Font.Underline = xlUnderlineStyleSingle
                                End If
                            End If
                            p = p + l
                        End If
                     Next i
                End With
            Else
                cell.Value = "'" + OriText + vbCr + vbLf + RevText
                cell.Characters(Start:=1, length:=LO).Font.Strikethrough = True
                cell.Characters(Start:=LO + 3, length:=LR).Font.Underline = xlUnderlineStyleSingle
            End If
        End If
    End If
End Sub

Function Atomize(ByVal inString As String, sep As String) As String()
    ' This function will divide a string into an array of substrings,
    ' using the characters in sep argument as separators
    Dim pos, posmin, last, cnt, length As Integer
    Dim buff() As String
    Dim ch As Integer
    Dim sp As String
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
                sp = Mid(sep, ch, 1)
            End If
        Next ch
        buff(cnt) = Mid$(inString, last, posmin - last + 1) ' Adds the word
        cnt = cnt + 1
        last = posmin + 1
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
    Next i
    UnicodeToArray = unicodeTestArray

End Function

Function StreamFormula(inString As String) As String
    Dim outString, ch As String
    Dim i As Integer
    Dim inText As Boolean
    inText = False
    If Left(inString, 1) = "=" Then
        For i = 1 To Len(inString)
            ch = Mid(inString, i, 1)
            If ch = """" Then
                inText = Not inText
            End If
            If Not inText And ch = ";" Then
                outString = outString & "," ' Replaces ; by ,
            Else
                outString = outString & ch
            End If
        Next i
        StreamFormula = outString
    Else
        StreamFormula = inString ' Don't do anything
    End If
End Function
Sub ArrangeWindows()
    'Dim OriWorkbook As Workbook
    'Dim RevWorkbook As Workbook
    'Dim wndName As String
    Windows.Arrange ArrangeStyle:=xlTiled
    'ActiveWindow.WindowState = xlNormal
    'Set OriWorkbook = GetWorkbook(ActiveSheet.Cells(cfgRowFilename, cfgColOriginal))
    'Set RevWorkbook = GetWorkbook(ActiveSheet.Cells(cfgRowFilename, cfgColRevision))
    'wndName = RevWorkbook.Windows(1).Caption
    'OriWorkbook.Activate
    'Application.Windows.CompareSideBySideWith (wndName)
End Sub

