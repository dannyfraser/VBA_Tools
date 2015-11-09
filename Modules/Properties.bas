Attribute VB_Name = "Properties"
Option Compare Database
Option Explicit

'The enum must match the Properties table Id
Public Enum PropertyLookup
    [_Invalid] = -1
    [_First] = 1
    LogFilePath = 1
    LogReportingType = 2
    ErrorLogEmailError = 3
    ErrorLogEmailRecipients = 4
    ErrorLogFilePath = 5
    ErrorLogReportingType = 6
    AppName = 7
    LivePath = 8
    ReleaseVersion = 9
    SilentError = 10
    [_Last] = 10
End Enum

Public Function GetProperty(PropertyId As PropertyLookup) As String
    If RegKeyLookup(PropertyId) <> vbNullString Then
        GetProperty = Trim(GetSetting(GetApplicationName, "User Settings", RegKeyLookup(PropertyId), vbNullString))
    Else
        GetProperty = Trim(CurrentDb.Containers("Databases").Documents("UserDefined").Properties(StringLookup(PropertyId)))
    End If
End Function

Public Sub SetProperty(PropertyId As PropertyLookup, PropertyValue As Variant)
    If RegKeyLookup(PropertyId) <> vbNullString Then
        SaveSetting GetApplicationName, "User Settings", RegKeyLookup(PropertyId), Trim(CStr(PropertyValue))
    Else
        CurrentDb.Containers("Databases").Documents("UserDefined").Properties(StringLookup(PropertyId)) = Trim(PropertyValue)
    End If
End Sub

Public Function StringLookup(PropertyId As PropertyLookup) As String
    StringLookup = DLookupStringWrapper("PropertyName", "Properties", "PropertyId = " & PropertyId, vbNullString)
End Function
Public Function RegKeyLookup(PropertyId As PropertyLookup) As String
    RegKeyLookup = DLookupStringWrapper("PropertyRegKey", "Properties", "PropertyId = " & PropertyId, vbNullString)
End Function

Public Function GetApplicationName() As String
    GetApplicationName = GetProperty(AppName)
End Function


Public Sub SaveAllUserSettings()
       
    Dim rs As New RecordsetWrapper
    rs.OpenRecordset "SELECT PropertyId, PropertyRegKey, DefaultValue FROM Properties"
    
    Do While Not rs.EOF
        
        Dim PropertyName As String
        PropertyName = Nz(rs!PropertyRegKey, vbNullString)
        
        Dim PropertyValue As Variant
        PropertyValue = Nz(rs!DefaultValue, vbNullString)
        
        If Len(PropertyName) = 0 Then
            'save locally
            If GetProperty(rs!PropertyId) = vbNullString Then
                SetProperty rs!PropertyId, PropertyValue
            End If
        
        Else
            'Save to registry
            If GetSetting(GetApplicationName, "User Settings", PropertyName) = vbNullString Then
                SaveSetting GetApplicationName, "User Settings", PropertyName, PropertyValue
            End If
        End If
        
        rs.MoveNext
    
    Loop
    
End Sub
