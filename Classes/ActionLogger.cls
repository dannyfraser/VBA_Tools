VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Create a new ActionLogger object for each action to be logged (alternatively, create as a global object and reuse)

Private m_LogType As String
Private m_LogPath As String
Private m_Timer As CodeProfiler


Private this As EventDetails
Private Type EventDetails
    ActionlogId As Long
    Username As String
    LogonServer As String
    ComputerName As String
    ApplicationName As String
    ActionName As String
    ActionDetail As String
    StartTime As Date
    EndTime As Date
    ElapsedSeconds As Double
    RecordCount As Long
End Type

Private Const ACTION_LOG_TABLE_NAME As String = "ActionLog"

Property Let LogFilePath(p As String)
    m_LogPath = p
End Property
Private Property Get LogFilePath() As String
    LogFilePath = m_LogPath
End Property

Private Property Let LogFileType(t As String)
    m_LogType = t
End Property
Private Property Get LogFileType() As String
    LogFileType = m_LogType
End Property
    


Private Sub Class_Initialize()
    
    LogFileType = GetProperty(LogReportingType)
    LogFilePath = GetProperty(PropertyLookup.LogFilePath)
    
    If LogFileType = "Database" And LogFilePath = vbNullString Then
        LogFilePath = CurrentProject.FullName
    End If
    
    Set m_Timer = New CodeProfiler
    
End Sub

Private Sub Class_Terminate()
    Set m_Timer = Nothing
End Sub

Public Sub Setup()

    If m_LogType = "Database" Then
        
        If Len(m_LogPath) > 0 Then
        
            If Dir(m_LogPath) = vbNullString Then
                CreateDatabase m_LogPath, dbLangGeneral
            End If
            
            If Not (TableExists(ACTION_LOG_TABLE_NAME, m_LogPath)) Then
                Me.CreateTable
            End If
            
        End If
        
    End If
        
End Sub

Private Sub SetEventDetails(Action As String, Optional ActionDetail As String, Optional RecordCount As Long)

    With this
        .ActionName = Action
        .ActionDetail = ActionDetail
        .Username = Environ("username")
        .LogonServer = Environ("logonserver")
        .ComputerName = Environ("computername")
        .ApplicationName = GetApplicationName
        If Not IsMissing(RecordCount) Then
            .RecordCount = RecordCount
        End If
    End With
    
End Sub

Sub LogEvent(Action As String, Optional ActionDetail As String, Optional RecordCount As Long)

    this.StartTime = Now()
    SetEventDetails Action, ActionDetail, RecordCount
    WriteToLog

End Sub


Private Sub WriteToLog()
    
    Select Case LogFileType
        Case "Database"
            LogToDatabase
        Case "File"
            LogToFile
        Case Else
            'Do nothing
    End Select
    
End Sub


Sub StartTimedEvent(Action As String, Optional ActionDetail As String)

    m_Timer.StartProfiling
    this.StartTime = Now()
    SetEventDetails "[START]  " & Action, ActionDetail
    WriteToLog
    
End Sub


Sub CloseTimedEvent(Optional RecordCount As Long)

    this.ElapsedSeconds = m_Timer.GetRunTime
    this.EndTime = Now
    this.RecordCount = RecordCount
    this.ActionName = Replace(this.ActionName, "[START]", "[FINISH]")
    WriteToLog
    
End Sub

Sub CreateTable()
    
    If m_LogType = "File" Then Exit Sub
    
    Dim LogDB As DAO.Database
    Set LogDB = OpenDatabase(LogFilePath)
    
    If Not DatabaseUtilities.TableExists(ACTION_LOG_TABLE_NAME, LogFilePath) Then
    
        LogDB.Execute _
            " CREATE TABLE " & ACTION_LOG_TABLE_NAME _
            & " (ActionLogId COUNTER (1,1) NOT NULL," _
            & " Action MEMO," _
            & " ActionDetail MEMO," _
            & " TimeStampStart DATE," _
            & " TimeStampEnd DATE," _
            & " ElapsedSeconds DOUBLE," _
            & " RecordCount LONG," _
            & " User TEXT(20)," _
            & " LogOnServer TEXT(25)," _
            & " Computer TEXT(25)," _
            & " Application TEXT(255))", _
            dbFailOnError
            
    End If
    
End Sub


Private Sub LogToDatabase()
    
    Dim LogDB As DAO.Database
    Set LogDB = OpenDatabase(LogFilePath)
    
    Dim NewLogEntry As Boolean
    NewLogEntry = (this.ActionlogId = 0)
    
    Dim ActionLog As DAO.Recordset
    Set ActionLog = LogDB.OpenRecordset( _
        "SELECT * FROM " & ACTION_LOG_TABLE_NAME & _
        " WHERE ActionLogID = " & IIf(NewLogEntry, 1, this.ActionlogId), _
        dbOpenDynaset)
    
    'Add Action Log Record
    If NewLogEntry Then
        ActionLog.AddNew
    Else
        ActionLog.Edit
    End If
    
    'Add Action Log Data
    With this
    
        ActionLog!Action = .ActionName
        ActionLog!ActionDetail = .ActionDetail
        ActionLog!TimeStampStart = .StartTime
        ActionLog!TimeStampEnd = .EndTime
        ActionLog!ElapsedSeconds = .ElapsedSeconds
        ActionLog!User = .Username
        ActionLog!LogonServer = .LogonServer
        ActionLog!Computer = .ComputerName
        ActionLog!Application = .ApplicationName
        
        If Not IsNull(.RecordCount) Then
            ActionLog!RecordCount = .RecordCount
        End If
        
    End With

    this.ActionlogId = ActionLog!ActionlogId
    ActionLog.Update
    
    ActionLog.Close
    LogDB.Close
    
End Sub


Private Sub LogToFile()

    Dim f As Integer
    f = FreeFile(0)
    
    Open LogFilePath For Append As #f
    
    Dim ActionText As String
    With this
        If Left(this.ActionName, Len("[FINISH]")) = "[FINISH]" Then
            ActionText = .EndTime & vbTab
        Else
            ActionText = .StartTime & vbTab
        End If
        ActionText = ActionText & .ActionName & vbTab
        ActionText = ActionText & .ActionDetail & vbTab
        If IsNull(.RecordCount) Then
            ActionText = ActionText & vbTab
        Else
            ActionText = ActionText & .RecordCount & " Records Affected" & vbTab
        End If
        ActionText = ActionText & .Username & vbTab
        ActionText = ActionText & .LogonServer & vbTab
        ActionText = ActionText & .ComputerName & vbTab
        ActionText = ActionText & .ApplicationName
    End With
    
    Print #f, ActionText
    Close #f
    
End Sub
