VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private this As ExcelHandlerFields
Private Type ExcelHandlerFields
    Application As Excel.Application
End Type

Private Property Let ExcelApp(e As Excel.Application)
    Set this.Application = e
End Property
Property Get ExcelApp() As Excel.Application
    Set ExcelApp = this.Application
End Property

Private Sub Class_Initialize()
    ExcelApp = New Excel.Application
End Sub

Function OpenWorkbook(WorkbookPath As String) As Excel.Workbook
    Set OpenWorkbook = ExcelApp.Workbooks.Open(WorkbookPath, False)
End Function

Private Sub Class_Terminate()
    
    Dim wb As Workbook
    For Each wb In ExcelApp.Workbooks
        wb.Saved = True
        wb.Close False
    Next wb
    
    On Error Resume Next
    ExcelApp.Quit
    ExcelApp = Nothing
    On Error GoTo 0
    
End Sub
