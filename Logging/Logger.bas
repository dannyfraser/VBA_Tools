Attribute VB_Name = "Logger"
Option Compare Database
Option Explicit

Private Path As String
Private LogEvents As New Collection
Dim fso As New FileSystemObject

Public Property Let LogFilePath(ByVal s As String)
    Path = s
    If Len(Dir(s)) = 0 Then
        fso.CreateTextFile s
    End If
End Property
Private Property Get LogFilePath() As String
    LogFilePath = Path
End Property

Public Sub Log(ByVal EventName As String, Severity As LogSeverity)
    LogEvents.Add CreateLogEvent(EventName, Severity), EventName
    WriteLog EventName
End Sub
Public Sub CloseLogEvent(ByVal EventName As String, Optional RecordsAffected As Long)
    If Not IsMissing(RecordsAffected) Then
        LogEvents(EventName).RecordsAffected = RecordsAffected
    End If
    WriteLog EventName, True
    LogEvents.Remove EventName
End Sub

Private Function CreateLogEvent(ByVal EventName As String, Severity As LogSeverity) As LogEvent
    Set CreateLogEvent = New LogEvent
    CreateLogEvent.Name = EventName
    CreateLogEvent.Severity = Severity
End Function

Private Sub WriteLog(ByVal EventName As String, Optional CloseEvent As Boolean = False)
    Dim LogFileStream As scripting.TextStream
    Set LogFileStream = fso.GetFile(LogFilePath).OpenAsTextStream(ForAppending)
    If CloseEvent Then
        LogFileStream.WriteLine LogEvents(EventName).CloseLogLine
    Else
        LogFileStream.WriteLine LogEvents(EventName).StartLogLine
    End If
End Sub

Public Function GetEventDuration(ByVal EventName As String) As Double
    GetEventDuration = LogEvents(EventName).Duration
End Function
