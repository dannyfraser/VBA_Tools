Attribute VB_Name = "Distribution"
Option Compare Database
Option Explicit

Public Function AutoExec_Startup()

    SetDevelopmentOptions

    Properties.SaveAllUserSettings

    If UpdateAvailable And Not (OpenedViaCitrix) Then
        MsgBox "NEW VERSION AVAILABLE AT:" & vbNewLine & GetProperty(LivePath)
        SelectFile GetProperty(LivePath)
        DoCmd.Quit acQuitSaveNone
    End If
    
End Function

Public Sub SetDevelopmentOptions()

    If IsCompiled Then
        PrepareForDistribution
    Else
        PrepareForDevelopment
    End If

End Sub


Private Sub PrepareForDistribution()
    
    Dim ToolBarIndex As Integer
    For ToolBarIndex = 1 To CommandBars.Count
        CommandBars(ToolBarIndex).Enabled = False
    Next ToolBarIndex

    If CDbl(Application.Version) > 11 Then
        DoCmd.ShowToolbar "Ribbon", acToolbarNo
    End If
    
    'Application Properties
    ChangeApplicationProperty "StartupShowDBWindow", dbBoolean, False
    ChangeApplicationProperty "StartupShowStatusBar", dbBoolean, False
    ChangeApplicationProperty "AllowBuiltinToolbars", dbBoolean, False
    ChangeApplicationProperty "AllowFullMenus", dbBoolean, False
    ChangeApplicationProperty "AllowBreakIntoCode", dbBoolean, False
    ChangeApplicationProperty "AllowSpecialKeys", dbBoolean, False
    ChangeApplicationProperty "AllowBypassKey", dbBoolean, False
    
    DatabaseUtilities.HideDatabaseWindow
        
End Sub

Private Sub PrepareForDevelopment()
        
    Dim ToolBarIndex As Integer
    For ToolBarIndex = 1 To CommandBars.Count
        CommandBars(ToolBarIndex).Enabled = True
    Next ToolBarIndex

    If CDbl(Application.Version) > 11 Then
        DoCmd.ShowToolbar "Ribbon", acToolbarYes
    End If
    
    'Application Properties
    ChangeApplicationProperty "StartupShowDBWindow", dbBoolean, True
    ChangeApplicationProperty "StartupShowStatusBar", dbBoolean, True
    ChangeApplicationProperty "AllowBuiltinToolbars", dbBoolean, True
    ChangeApplicationProperty "AllowFullMenus", dbBoolean, True
    ChangeApplicationProperty "AllowBreakIntoCode", dbBoolean, True
    ChangeApplicationProperty "AllowSpecialKeys", dbBoolean, True
    ChangeApplicationProperty "AllowBypassKey", dbBoolean, True
      
End Sub


Public Function IsCompiled() As Boolean
    If Exists(CurrentDb.Properties, "MDE") Then
        IsCompiled = (CurrentDb.Properties("MDE") = "T")
    End If
End Function

Private Sub ChangeApplicationProperty(PropertyName As String, PropertyType As DAO.DataTypeEnum, PropertyValue As Variant)

    If Exists(CurrentDb.Properties, PropertyName) Then
        CurrentDb.Properties(PropertyName) = PropertyValue
    Else
        Dim NewProperty As DAO.Property
        Set NewProperty = CurrentDb.CreateProperty(PropertyName, PropertyType, PropertyValue)
        CurrentDb.Properties.Append NewProperty
    End If
    
End Sub


Private Function UpdateAvailable() As Boolean

    UpdateAvailable = False
    
    If HasProperty(CurrentDb, StringLookup(ReleaseVersion)) And GetProperty(PropertyLookup.LivePath) <> vbNullString Then

        Dim LocalVersion As Double
        LocalVersion = GetProperty(ReleaseVersion)

        Dim LiveVersion As Double
        LiveVersion = GetLiveVersion(GetProperty(PropertyLookup.LivePath))
        
        If LiveVersion > LocalVersion Then
            UpdateAvailable = True
        End If
        
    End If
    
End Function

Private Function GetLiveVersion(LivePath As String) As Double

    If Files.FileExists(LivePath) Then
        GetLiveVersion = OpenDatabase(LivePath).Containers("Databases").Documents("UserDefined").Properties(StringLookup(ReleaseVersion))
    Else
        GetLiveVersion = 0
    End If

End Function

Public Function IncrementVersion()
    SetProperty ReleaseVersion, GetProperty(ReleaseVersion) + 1
End Function

Public Function OpenedViaCitrix() As Boolean
    OpenedViaCitrix = (InStr(Environ("computername"), "CTX") > 0)
End Function
