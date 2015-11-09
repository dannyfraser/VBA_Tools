Attribute VB_Name = "CommonExcelFunctions"
Option Explicit

Public Sub SetNumberOfWorksheets(NewWorkbook As Workbook, Optional n As Integer = 3)

    'This will bring the number of worksheets in a workbook to the number specified.
    'It is best used on creating a new workbook
    Quiet NewWorkbook
    With NewWorkbook
        Do While .Worksheets.Count <> n
            If .Worksheets.Count > n Then
                .Worksheets(.Worksheets.Count).Delete
            ElseIf .Worksheets.Count < n Then
                .Worksheets.Add after:=.Worksheets(.Worksheets.Count)
            End If
        Loop
    End With
    UnQuiet NewWorkbook
    
End Sub


Public Function OpenWorkbook(Title As String, Optional Filter As String = "All Files,.*") As Workbook
    'This opens a workbook with the prompt and filter specified in the parameters
    'Use in code to set workbook objects, i.e.
    'Set wbkFoo = OpenWorkbook("Open a file", "Comma Separated Value Files,*.csv")

    Dim Filename As String
    
    Filename = GetOpenFilename(Filter, , Title)
    Set OpenWorkbook = Workbooks.Open(Filename)

End Function


Public Sub DeleteSheet(SheetToDelete As Worksheet)

    'Deletes a worksheet, suppressing any alerts
    
    Quiet SheetToDelete.Parent
    SheetToDelete.Delete
    UnQuiet SheetToDelete.Parent
    
End Sub


Public Sub CloseAndDelete(WorkbookToClose As Workbook)
    
    'This will close and delete a file, if it is open in Excel as a workbook.
    'the file need not be a .xls file - this will work on .csv and .txt files opened in Excel.

    Quiet WorkbookToClose
    
    Dim Filename As String
    Filename = WorkbookToClose.FullName
    
    WorkbookToClose.Saved = True
    UnQuiet WorkbookToClose
    WorkbookToClose.Close False
    Set WorkbookToClose = Nothing
    
    Files.DeleteFile Filename
    
End Sub


Public Function SheetExists(SheetName As String, HoldingWorkbook As Workbook) As Boolean

    'Returns True if a sheet with sheetName exists in the specified workbook

    On Error GoTo fail
    
    Dim CheckSheet As Worksheet
    Set CheckSheet = HoldingWorkbook.Worksheets(SheetName)
    
    SheetExists = True
    
    Exit Function

fail:
    SheetExists = False
    
End Function


Public Function WorkbookExists(WorkbookName As String) As Boolean

    'Returns True if a workbook with the name specified is open

    On Error GoTo fail
    
    Dim Test As Workbook
    Set Test = Workbooks(WorkbookName)
    
    WorkbookExists = True
    
    Exit Function
    
fail:
    WorkbookExists = False
    
End Function


Public Function NamedRangeExists(TargetSheet As Worksheet, RangeName As String) As Boolean

    'Checks if a specific named range exists in a worksheet
    
    On Error GoTo fail
    
    Dim TargetRange As Range
    Set TargetRange = TargetSheet.Range(RangeName)
    
    NamedRangeExists = True
    
    Exit Function
    
fail:
    NamedRangeExists = False

End Function



Public Function GetValue(SearchSheet As Worksheet, Row As Long, HeaderText As String, Optional HeaderRow As Integer = 1, Optional DefaultColumn As Integer = 0) As Variant
    
    Dim Column As Integer
    Column = GetColumn(SearchSheet, HeaderText, HeaderRow, DefaultColumn)
    
    If Not Column = DefaultColumn Then
        GetValue = SearchSheet.Cells(Row, Column)
    Else
        GetValue = ""
    End If
    
End Function


Public Function GetColumn(SearchSheet As Worksheet, HeaderText As String, Optional HeaderRow As Integer = 1, Optional DefaultValue As Integer = 0) As Integer
    
    Dim Find As Range
    Set Find = SearchSheet.Rows(HeaderRow).Find(what:=HeaderText, LookIn:=xlValues, lookat:=xlWhole)
    
    If Not Find Is Nothing Then
        GetColumn = Find.Column
    Else
        GetColumn = DefaultValue
    End If
    
End Function



Public Function IsWorkBookOpen(Filename As String) As Boolean

    'Pass the full workbook path as the input parameter.
    'Returns True if the workbook is already open
    
    
    On Error GoTo fail
    
    Dim FileNumber As Long
    FileNumber = FreeFile(0)
    Open Filename For Input Lock Read As #FileNumber
    Close FileNumber
    
fail:

    Dim ErrorNumber As Integer
    ErrorNumber = Err.Number
    
    Err.Clear
    
    Select Case ErrorNumber
        Case 0
            IsWorkBookOpen = False
        Case 70
            IsWorkBookOpen = True
        Case Else
            Err.Raise ErrorNumber
    End Select
    
End Function



Public Sub QuickSort(ByRef ArrayToSort As Variant, LowBound As Long, HighBound As Long)

    'sorts an array recursively using the QuickSort algorithm
    'From http://en.allexperts.com/q/Visual-Basic-1048/string-manipulation.htm
    'via http://stackoverflow.com/questions/152319/vba-array-sort-function

    Dim TempLowBound As Long
    TempLowBound = LowBound
    
    Dim TempHighBound As Long
    TempHighBound = HighBound

    Dim Pivot   As Variant
    Pivot = ArrayToSort((LowBound + HighBound) \ 2)

    Do While (TempLowBound <= TempHighBound)

        Do While (ArrayToSort(TempLowBound) < Pivot And TempLowBound < HighBound)
            TempLowBound = TempLowBound + 1
        Loop

        Do While (Pivot < ArrayToSort(TempHighBound) And TempHighBound > LowBound)
            TempHighBound = TempHighBound - 1
        Loop

        If (TempLowBound <= TempHighBound) Then
            Dim TempSwap As Variant
            TempSwap = ArrayToSort(TempLowBound)
            ArrayToSort(TempLowBound) = ArrayToSort(TempHighBound)
            ArrayToSort(TempHighBound) = TempSwap
            TempLowBound = TempLowBound + 1
            TempHighBound = TempHighBound - 1
        End If

    Loop

    If (LowBound < TempHighBound) Then QuickSort ArrayToSort, LowBound, TempHighBound
    If (TempLowBound < HighBound) Then QuickSort ArrayToSort, TempLowBound, HighBound

End Sub


Public Function RangeToStringArray(ByVal TargetRange As Range) As String()
       
    Dim StringArray() As String
    ReDim StringArray(TargetRange.Cells.Count)
    
    If TargetRange.Rows.Count = 1 Then
        
        'Turn a row of cells into an array of column elements
        Dim Column As Integer
        For Column = LBound(StringArray) To UBound(StringArray)
            StringArray(Column) = CStr(TargetRange.Cells(1, Column + 1))
        Next Column
        
    Else
    
        'Turn a column of cells into an array of row elements
        Dim Row As Long
        For Row = LBound(StringArray) To UBound(StringArray)
            StringArray(Row) = CStr(TargetRange.Cells(Row + 1, 1))
        Next Row
        
    End If
    
    RangeToStringArray = StringArray
    
End Function

Public Sub Quiet(Workbook As Workbook)
    Workbook.Application.ScreenUpdating = False
    Workbook.Application.DisplayAlerts = False
End Sub
Public Sub UnQuiet(Workbook As Workbook)
    Workbook.Application.ScreenUpdating = True
    Workbook.Application.DisplayAlerts = True
End Sub

Public Sub Unfilter(Worksheet As Worksheet)
    Worksheet.Cells.AutoFilter
End Sub

Public Sub CopyResults(Worksheet As Worksheet, Records As DAO.Recordset)

    Dim f As DAO.Field
    For Each f In Records.Fields
        Worksheet.Cells(1, f.OrdinalPosition + 1) = f.Name
    Next f
    
    Worksheet.Cells(2, 1).CopyFromRecordset Records
    
    Worksheet.Columns.AutoFit

End Sub
