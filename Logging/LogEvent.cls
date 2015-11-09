VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const LOG_DELIMITER As String = vbTab

Public Enum LogSeverity
    InformationLevel
    DetailedLevel
    WarningLevel
    ErrorLevel
End Enum

Private Type LogEventFields
    Name As String
    StartTime As Date
    Duration As Double
    Severity As LogSeverity
    RecordsAffected As Long
    RecordsWereAffected As Boolean
    EventTimer As Timer
End Type
Private this As LogEventFields

Private Sub Class_Initialize()
    Set this.EventTimer = New Timer
    this.EventTimer.Start
    this.StartTime = Now
    this.RecordsWereAffected = False
End Sub

Public Property Let Name(n As String)
    this.Name = n
End Property
Private Property Get Name() As String
    Name = this.Name
End Property

Private Property Get StartTime() As Date
    StartTime = this.StartTime
End Property

Public Property Get Duration() As Double
    Duration = Round(this.EventTimer.GetRunTime, 3)
End Property

Public Property Let Severity(s As LogSeverity)
    this.Severity = s
End Property
Public Property Get Severity() As LogSeverity
    Severity = this.Severity
End Property

Public Property Let RecordsAffected(r As Long)
    this.RecordsAffected = r
    this.RecordsWereAffected = True
End Property
Private Property Get RecordsAffected() As Long
    RecordsAffected = this.RecordsAffected
End Property

Private Property Get RecordsWereAffected() As Boolean
    RecordsWereAffected = this.RecordsWereAffected
End Property

Public Function CloseLogLine() As String
    CloseLogLine = _
        Join( _
            Array(SeverityLabel, StartTime, "END " & Name, Duration, IIf(RecordsAffected, RecordsAffected, vbNullString)), _
        LOG_DELIMITER)
End Function

Public Function StartLogLine()
    StartLogLine = _
        Join( _
            Array(SeverityLabel, StartTime, Name), _
        LOG_DELIMITER)
End Function

Private Function SeverityLabel() As String

    Select Case Severity
        Case Is = LogSeverity.DetailedLevel
            SeverityLabel = "DETAIL"
        Case Is = LogSeverity.InformationLevel
            SeverityLabel = "INFORMATION"
        Case Is = LogSeverity.WarningLevel
            SeverityLabel = "WARNING"
        Case Is = LogSeverity.ErrorLevel
            SeverityLabel = "ERROR"
    End Select

End Function
