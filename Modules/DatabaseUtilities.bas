Attribute VB_Name = "DatabaseUtilities"
Option Explicit
Option Compare Database

Private Const DB_TYPE_FORM As Long = -32768
Private Const DB_TYPE_QUERY As Integer = 5
Private Const DB_TYPE_TABLE_LOCAL As Integer = 1
Private Const DB_TYPE_TABLE_LINKED As Integer = 6

Public Function TableExists(TableName As String, Optional Path As String) As Boolean
    
    Dim DB As DAO.Database
    If Path = vbNullString Then
        Set DB = CurrentDb
    Else
        Set DB = OpenDatabase(Path)
    End If
    
    TableExists = Exists(DB.TableDefs, TableName)
    
    If Path <> vbNullString Then DB.Close
    
End Function


Public Function QueryExists(QueryName As String, Optional Path As String) As Boolean
   
    Dim DB As DAO.Database
    If Path = vbNullString Then
        Set DB = CurrentDb
    Else
        Set DB = OpenDatabase(Path)
    End If
    
    QueryExists = Exists(DB.QueryDefs, QueryName)
    
    If Path <> vbNullString Then DB.Close
        
End Function


Public Function FormExists(FormName As String, Optional Path As String) As Boolean

    Dim DB As DAO.Database
    If Path = vbNullString Then
        Set DB = CurrentDb
    Else
        Set DB = OpenDatabase(Path)
    End If

    Dim r As DAO.Recordset
    Set r = DB.OpenRecordset("SELECT Name, Type FROM MSysObjects WHERE Type = " & DB_TYPE_FORM & " AND Name = " & Sanitize(FormName))
    
    FormExists = Not (r.EOF And r.BOF)
    
    r.Close
    Set r = Nothing
    
    If Path <> vbNullString Then DB.Close
    
End Function


Public Function FormIsLoaded(FormName As String) As Boolean
        
    On Error GoTo FormIsLoaded_Error
    
    Dim Form As AccessObject
    
    FormIsLoaded = False
    
    Set Form = CurrentProject.AllForms(FormName)
    If Form.IsLoaded Then
        If Form.CurrentView <> acCurViewDesign Then
            FormIsLoaded = True
        End If
    End If
    
FormIsLoaded_Exit:
                    
    Exit Function

FormIsLoaded_Error:
    
    ErrorLogger.LogError Err.Number, Err.Description, "FormIsLoaded"
    Resume FormIsLoaded_Exit
        
End Function


Public Function ModuleExists(ModuleName As String) As Boolean
    ModuleExists = Exists(Application.VBE.ActiveVBProject.VBComponents, ModuleName)
End Function


Public Function ControlExists(Form As Form, ControlName As String) As Boolean
    ControlExists = Exists(Form.Controls, ControlName)
End Function


Public Function GetControl(Form As Form, ControlName As String) As Control

    If ControlExists(Form, ControlName) Then
        Set GetControl = Form.Controls(ControlName)
    Else
        Set GetControl = Nothing
    End If
        
End Function


Public Sub CloseAllForms(Optional Exception As String)
        
    Dim SQL As String
    SQL = "SELECT Name FROM MSysObjects WHERE Type=" & DB_TYPE_FORM & _
            "AND LEFT(Name,3)='frm' AND Name <> '" & Exception & "'"
    
    Dim AccessObjects As DAO.Recordset
    Set AccessObjects = CurrentDb.OpenRecordset(SQL, dbOpenDynaset)
    
    Do Until AccessObjects.EOF
        If CurrentProject.AllForms(AccessObjects!Name).IsLoaded Then
            DoCmd.Close acForm, AccessObjects!Name, acSavePrompt
        End If
        AccessObjects.MoveNext
    Loop
    
    AccessObjects.Close
    Set AccessObjects = Nothing
        
End Sub


Public Sub CloseAllTables(DB As DAO.Database, Optional Exception As String)
        
    Dim SQL As String
    SQL = "SELECT Name FROM MSysObjects WHERE (Type=" & DB_TYPE_TABLE_LOCAL & " OR Type=" & DB_TYPE_TABLE_LINKED & ")" & _
            "AND Name <> '" & Exception & "'"
    
    Dim AccessObjects As DAO.Recordset
    Set AccessObjects = DB.OpenRecordset(SQL, dbOpenDynaset)
    
    Do Until AccessObjects.EOF
        On Error Resume Next
        DoCmd.Close acTable, AccessObjects!Name, acSaveNo
        AccessObjects.MoveNext
    Loop
    
    AccessObjects.Close
    Set AccessObjects = Nothing
        
End Sub


Public Sub CloseAllQueries(DB As DAO.Database, Optional Exception As String)
        
    Dim SQL As String
    SQL = "SELECT Name FROM MSysObjects WHERE (Type=" & DB_TYPE_QUERY & ")" & _
            "AND Name <> '" & Exception & "'"
    
    Dim AccessObjects As DAO.Recordset
    Set AccessObjects = DB.OpenRecordset(SQL, dbOpenDynaset)
    
    Do Until AccessObjects.EOF
        On Error Resume Next
        DoCmd.Close acTable, AccessObjects!Name, acSaveNo
        AccessObjects.MoveNext
    Loop
    
    AccessObjects.Close
    Set AccessObjects = Nothing
        
End Sub


Public Sub UpdateStatusBar(Optional Message As String, Optional PercentageDone As Integer)
   
    If Message = vbNullString Then
    
        If PercentageDone = 0 Then
            SysCmd acSysCmdClearStatus
        Else
            SysCmd acSysCmdUpdateMeter, PercentageDone
        End If
            
    Else
    
        If PercentageDone = 0 Then
            SysCmd acSysCmdSetStatus, Message
        Else
            SysCmd acSysCmdInitMeter, Message, 100
            SysCmd acSysCmdUpdateMeter, PercentageDone
        End If
            
    End If

End Sub

Public Function HasProperty(Target As Object, PropertyName As String) As Boolean
    HasProperty = Exists(Target.Properties, PropertyName)
End Function


Public Sub CreateLogDatabase(Location As String)

    On Error GoTo CreateAccessFile_Error
    
    If Location = vbNullString Then
        Exit Sub
    End If
    
    'Set application properties
    SetProperty ErrorLogFilePath, Location
    SetProperty LogFilePath, Location
    
    'Prepare new file
    Files.DeleteFile Location
    Files.CreateFolder Left(Location, InStrRev(Location, "\"))
    
    Dim DB As DAO.Database
    Set DB = CreateDatabase(Location, dbLangGeneral, dbVersion40)
    
    'Create tables in log database
    Dim ActionLogger As New ActionLogger
    ActionLogger.LogFilePath = Location
    ActionLogger.CreateTable
    
    ErrorLogger.LogFilePath = Location
    ErrorLogger.CreateTable
    
CreateAccessFile_Exit:
    
    On Error Resume Next
    DB.Close
    Exit Sub

CreateAccessFile_Error:
    MsgBox Err.Number, Err.Description, "CreateAccessFile"
    Resume CreateAccessFile_Exit

End Sub


Public Function GetForeignSysProperty(PropertyName As String, ForeignPath As String) As String
       
    On Error GoTo GetForeignSysProperty_Error
    
    If ForeignPath <> vbNullString Then
        
        Dim ForeignDB As DAO.Database
        Set ForeignDB = DAO.OpenDatabase(ForeignPath)
        
        If HasProperty(ForeignDB, PropertyName) Then
            GetForeignSysProperty = ForeignDB.Properties(PropertyName)
        End If
        
    End If
    
GetForeignSysProperty_Exit:
    
    If ForeignPath <> vbNullString Then ForeignDB.Close
    Exit Function

GetForeignSysProperty_Error:
    ErrorLogger.LogError Err.Number, Err.Description, "GetProperty"
    Resume GetForeignSysProperty_Exit
    
End Function


Public Sub LinkTable(TableName As String, Path As String)

    On Error GoTo LinkTable_Error

    Dim DB As DAO.Database
    Dim LinkedTable As DAO.TableDef

    Set DB = CurrentDb

    If TableExists(TableName) Then
        DB.TableDefs.Delete (TableName)
    End If

    Set LinkedTable = CurrentDb.CreateTableDef(TableName)

    With LinkedTable
        .Connect = ";DATABASE=" & Path
        .SourceTableName = TableName
    End With

    DB.TableDefs.Append LinkedTable

LinkTable_Exit:

    Exit Sub

LinkTable_Error:
    ErrorLogger.LogError Err.Number, Err.Description, "LinkTable"
    Resume LinkTable_Exit

End Sub

Public Sub DeleteTable(DB As Database, TableName As String)

    On Error GoTo DeleteTable_Error

    If TableExists(TableName, DB.Name) Then
        DB.TableDefs.Delete TableName
    End If

DeleteTable_Exit:

    Exit Sub

DeleteTable_Error:
    ErrorLogger.LogError Err.Number, Err.Description, "DeleteTable"
    Resume DeleteTable_Exit

End Sub

Public Sub CopyTable(TableName As String, TableAlias As String, Optional DatabasePath As String)

    On Error GoTo CopyTable_Error

    Dim DB As Database

    If DatabasePath = vbNullString Then
        DatabasePath = CurrentProject.Path
    End If

    If TableExists(TableAlias, DatabasePath) Then
        Set DB = OpenDatabase(DatabasePath)
        DB.TableDefs.Delete TableAlias
        DB.Close
        Set DB = Nothing
    End If

    DoCmd.TransferDatabase acImport, "Microsoft Access", DatabasePath, acTable, TableName, TableName & "_Temp", False
    DoCmd.TransferDatabase acExport, "Microsoft Access", DatabasePath, acTable, TableName & "_Temp", TableAlias, False
    DoCmd.DeleteObject acTable, TableName & "_Temp"

CopyTable_Exit:

    Exit Sub

CopyTable_Error:
    ErrorLogger.LogError Err.Number, Err.Description, "CopyTable"
    Resume CopyTable_Exit

End Sub

Public Sub DeleteLinkedTables()

    On Error GoTo DeleteLinkedTables_Error

    Dim LinkedTable As TableDef

    For Each LinkedTable In CurrentDb.TableDefs
        If LinkedTable.Connect <> vbNullString Then
            DeleteTable CurrentDb, LinkedTable.Name
        End If
    Next

DeleteLinkedTables_Exit:

    Exit Sub

DeleteLinkedTables_Error:
    ErrorLogger.LogError Err.Number, Err.Description, "DeleteLinkedTables"
    Resume DeleteLinkedTables_Exit

End Sub


Public Sub AutoCompactAndRepiar()
    
    'Remember that auto-compact can be a death sentence due to backups!
    'Only use this if you DEFINITELY have a network backup of the application (ideally in a VCS...)
    
    If Files.FileSize(Application.CurrentProject.FullName) > 25 Then
        Application.SetOption ("Auto Compact"), 1   'Compact Application
    Else
        Application.SetOption ("Auto Compact"), 0   'Don't Compact Application
    End If
    
End Sub


Public Sub CreateTestQuery(SQL As String, Optional QueryName As String)

    If QueryName = vbNullString Then
        QueryName = "qtmp_TestingQuery"
    End If
    
    If QueryExists(QueryName) Then
        DoCmd.DeleteObject acQuery, QueryName
    End If
        
    Dim NewQuery As DAO.QueryDef
    Set NewQuery = New DAO.QueryDef
    NewQuery.Name = QueryName
    NewQuery.SQL = SQL
    
    CurrentDb.QueryDefs.Append NewQuery
    CurrentDb.QueryDefs.Refresh

End Sub

Public Sub ShowDatabaseWindow()
    DoCmd.SelectObject acTable, , True
End Sub

Public Sub HideDatabaseWindow()
    DoCmd.SelectObject acTable, , True
    DoCmd.RunCommand acCmdWindowHide
End Sub


Public Sub RunScriptsFromFile(Filename As String, Database As Database)

    On Error GoTo ErrorHandler

    Dim FSO As New FileSystemObject
    
    DBEngine.BeginTrans
    
    Dim SQLCommand
    For Each SQLCommand In Split(FSO.OpenTextFile(Filename).ReadAll, ";")
        If Trim(SQLCommand) <> vbNullString Then
            Debug.Print SQLCommand
            Database.Execute SQLCommand, dbFailOnError
        End If
    Next SQLCommand
    
    DBEngine.CommitTrans
    
    Exit Sub
    
ErrorHandler:
    DBEngine.Rollback
    
End Sub


Public Function LowestOf(ParamArray Inputs()) As Variant

    Dim Low As Variant
    Low = Null

    Dim i As Long
    For i = LBound(Inputs) To UBound(Inputs)
        If IsNumeric(Inputs(i)) Or IsDate(Inputs(i)) Then
        
            If Inputs(i) < Low Then
                Low = Inputs(i)
            End If
            
        End If
    Next i
    
    LowestOf = Low

End Function
Public Function HighestOf(ParamArray Inputs()) As Variant
    
    Dim High As Variant
    High = Null

    Dim i As Long
    For i = LBound(Inputs) To UBound(Inputs)
        If IsNumeric(Inputs(i)) Or IsDate(Inputs(i)) Then
        
            If Inputs(i) > High Then
                High = Inputs(i)
            End If
            
        End If
    Next i
    
    HighestOf = High
    
End Function

Public Sub SelectAllInList(ListBox As ListBox)
    Dim i As Long
    For i = 0 To ListBox.ListCount - 1
        ListBox.Selected(i) = True
    Next i
End Sub
Public Sub ClearAllInList(ListBox As ListBox)
    Dim i As Long
    For i = 0 To ListBox.ListCount - 1
        ListBox.Selected(i) = False
    Next i
End Sub

Public Sub ListIndexes()
    Dim t As TableDef
    For Each t In CurrentDb.TableDefs
        If Not (t.Name Like "MSys*") Then
        Dim i As Index
        For Each i In t.Indexes
            Debug.Print t.Name, i.Name, i.Fields, i.Primary, i.Unique, i.IgnoreNulls
        Next i
        End If
    Next t
    
End Sub
