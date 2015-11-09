Attribute VB_Name = "SanitizeInput"
Option Compare Database
Option Explicit

Private Const DATE_FORMAT = "YYYY-MM-DD HH:NN:SS"

Public Function Sanitize(ByVal InputData As Variant) As String

    If IsDate(InputData) Then
        Sanitize = SanitizeDate(CDate(InputData))
    Else
        Sanitize = SanitizeString(CStr(InputData))
    End If
    
End Function

Private Function SanitizeString(ByVal InputString As String) As String
    
    If Not StringIsClean(InputString) Then
    
        If InStr(2, InputString, "'") > 0 Then
            InputString = """" & InputString & """"
        Else
            InputString = "'" & InputString & "'"
        End If
        
    End If
    
    SanitizeString = InputString

End Function

Private Function SanitizeDate(InputDate As Date) As String
    SanitizeDate = "'" & Format(InputDate, DATE_FORMAT) & "'"
End Function

Private Function StringIsClean(InputString As String) As Boolean

    StringIsClean = _
        (Left(InputString, 1) = "'" And Right(InputString, 1) = "'") _
        Or (Left(InputString, 1) = """" And Right(InputString, 1) = """")

End Function
