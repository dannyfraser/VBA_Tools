VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This class does not need to be declared and created - it exists as a predeclared object.
'It can be instantiated again, but there is little need to do so.

'Attribute VB_PredeclaredId = True

Private m_ErrorLogType As String
Private m_ErrorLogPath As String
Private Const ERROR_LOG_TABLE_NAME As String = "ErrorLog"

Private this As ErrorDetails
Private Type ErrorDetails
    Description As String
    Source As String
    Time As Date
    User As String
    Server As String
    Computer As String
    AppName As String
End Type


Property Let LogFilePath(p As String)
    m_ErrorLogPath = p
End Property
Property Get LogFilePath() As String
    LogFilePath = m_ErrorLogPath
End Property

Private Property Let LogType(t As String)
    m_ErrorLogType = t
End Property
Private Property Get LogType() As String
    LogType = m_ErrorLogType
End Property

Private Sub Class_Initialize()

    LogType = GetProperty(ErrorLogReportingType)
    LogFilePath = GetProperty(ErrorLogFilePath)
    
    If LogType = "Database" And LogFilePath = vbNullString Then
        LogFilePath = CurrentProject.FullName
    End If
    
End Sub

Private Property Get Silent() As Boolean
    Silent = GetProperty(SilentError)
End Property


Sub LogError(Number As Long, _
                Description As String, _
                Procedure As String, _
                Optional Form As String = vbNullString)

    On Error GoTo ErrorLog_Error

    'Get Data
    this.Description = Sanitize(Description)
    this.Source = CurrentObjectName
    this.Time = Now()
    this.User = Environ("USERNAME")
    this.Server = Environ("LOGONSERVER")
    this.Computer = Environ("COMPUTERNAME")
    this.AppName = GetProperty(AppName)
    
    If LogType = "Database" Then
        LogErrorToDatabase Number, Procedure, Form
    Else
        LogErrorToFile Number, Procedure, Form
    End If


    If Not Silent Then

        'Report Error
        Dim MsgBoxMessage As String
        MsgBoxMessage = ErrorText(Number, this.Description)
        
        'Report in Message Box
        If MsgBox(MsgBoxMessage, vbYesNo + vbDefaultButton2, Procedure) = vbYes Then
        
            Dim MsgBoxMessageFull As String
            MsgBoxMessageFull = FullErrorText(MsgBoxMessage, Form, Procedure)
        
            MsgBox MsgBoxMessageFull, vbOKOnly, Procedure
            
        End If
        
        'Report By Email
        If GetProperty(ErrorLogEmailError) Then

            'Report Errors to me
            If Number <> 70 Then
                Email.SendMail GetProperty(ErrorLogEmailRecipients), "Error in PPD", MsgBoxMessageFull, "GGM PPD"
            End If

        End If

    End If

ErrorLog_Exit:

    Exit Sub

ErrorLog_Error:

    MsgBox "Error in error logging procedure. Fatal Error." & vbCrLf & vbCrLf & Err.Description, vbOKOnly, "Error Log Error"
    End
    Resume ErrorLog_Exit

End Sub

Private Sub LogErrorToDatabase(Number As Long, Procedure As String, Form As String)
    'Build SQL
    Dim SQL As String
    SQL = vbNullString
    SQL = SQL & "INSERT INTO " & ERROR_LOG_TABLE_NAME & "("
    SQL = SQL & "ErrNumber, "
    SQL = SQL & "ErrDescription, "
    SQL = SQL & "CurrentObject, "
    SQL = SQL & "Form, "
    SQL = SQL & "ErrProcedure, "
    SQL = SQL & "ErrTime, "
    SQL = SQL & "User, "
    SQL = SQL & "LogOn, "
    SQL = SQL & "Computer, "
    SQL = SQL & "Application) "
    SQL = SQL & "SELECT "
    SQL = SQL & Sanitize(Number) & ",  "
    SQL = SQL & Sanitize(this.Description) & ",  "
    SQL = SQL & Sanitize(this.Source) & ",  "
    SQL = SQL & Sanitize(Form) & ",  "
    SQL = SQL & Sanitize(Procedure) & ", "
    SQL = SQL & Sanitize(this.Time) & ", "
    SQL = SQL & Sanitize(this.User) & ",  "
    SQL = SQL & Sanitize(this.Server) & ",  "
    SQL = SQL & Sanitize(this.Computer) & ",  "
    SQL = SQL & Sanitize(this.AppName)


    Dim ErrorDB As DAO.Database
    Set ErrorDB = OpenDatabase(LogFilePath)
                
    ErrorDB.Execute SQL, dbFailOnError
    ErrorDB.Close
    
End Sub


Private Sub LogErrorToFile(Number As Long, Procedure As String, Form As String)
    'Log Error (Text File)
    Dim Times As New Timer
    WriteToErrorLog _
        Times.ExactTimeAsString & vbTab & _
        Number & vbTab & _
        this.Description & vbTab & _
        this.Source & vbTab & _
        Form & vbTab & _
        Procedure & vbTab & _
        this.User & vbTab & _
        this.Server & vbTab & _
        this.Computer & vbTab & _
        this.AppName
End Sub

Private Function ErrorText(Number As Long, Description As String) As String
    ErrorText = _
        "An unexpected error has occurred" & vbCrLf & vbCrLf & _
        "ErrNumber: " & Number & vbCrLf & _
        "ErrDescription: " & Description & vbCrLf & _
        "CurrentObject: " & this.Source & vbCrLf
End Function

Private Function FullErrorText(MessageStart As String, Form As String, Procedure As String)
    
    FullErrorText = MessageStart
    
    FullErrorText = FullErrorText & vbCrLf & "Would you like to see full this?"
    
    FullErrorText = FullErrorText & "Form: " & Form & vbCrLf & _
        "ErrProcedure: " & Procedure & vbCrLf & _
        "ErrTime: " & this.Time & vbCrLf & _
        "User: " & this.User & vbCrLf & _
        "strLogOn: " & this.Server & vbCrLf & _
        "Computer: " & this.Computer & vbCrLf & _
        "Application: " & this.AppName & vbCrLf

End Function

Private Sub WriteToErrorLog(ErrorText As String)

    On Error GoTo WriteToErrorLog_Error

    Dim TextFile As Integer
    Dim ActionLogLocation As String

    TextFile = FreeFile()
    ActionLogLocation = Strings.AddPathSeparator(CurrentProject.Path) & GetProperty(ErrorLogFilePath)

    If Files.FileExists(ActionLogLocation) Then
        Open ActionLogLocation For Append As TextFile
        Print #TextFile, ErrorText
        Close #TextFile
    End If

WriteToErrorLog_Exit:

    Exit Sub

WriteToErrorLog_Error:
    ErrorLogger.LogError Err.Number, Err.Description, "WriteToErrorLog"
    Resume WriteToErrorLog_Exit

End Sub

Sub CreateTable()

    If m_ErrorLogType = "File" Then Exit Sub
    
    If Not DatabaseUtilities.TableExists(ERROR_LOG_TABLE_NAME, LogFilePath) Then
    
        Dim ErrorDB As DAO.Database
        Set ErrorDB = OpenDatabase(LogFilePath)

        ErrorDB.Execute _
            " CREATE TABLE " & ERROR_LOG_TABLE_NAME _
            & " (ErrorId COUNTER(1,1) NOT NULL," _
            & " ErrNumber LONG," _
            & " ErrDescription MEMO," _
            & " CurrentObject TEXT(50)," _
            & " Form TEXT(50)," _
            & " ErrProcedure TEXT(50)," _
            & " ErrTime DATE," _
            & " User TEXT(20)," _
            & " LogOnServer TEXT(25)," _
            & " Computer TEXT(25)," _
            & " Application TEXT(255))", _
            dbFailOnError
            
    End If

End Sub

