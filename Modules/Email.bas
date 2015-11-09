Attribute VB_Name = "Email"
Option Explicit
Option Compare Database

Public Sub SendMail(EmailAddress As String, SubjectLine As String, BodyText As String, SenderName As String)

    With New CDO.Message
        Set .Configuration = CreateConfiguration
        .To = EmailAddress
        .CC = vbNullString
        .BCC = vbNullString
        .FROM = SenderEmailAddress(SenderName)
        .Subject = SubjectLine
        .TextBody = BodyText
        .Send
    End With
    
End Sub

Private Function SenderEmailAddress(Sender As String) As String
    SenderEmailAddress = """" & Sender & """ <NO-REPLY-GPRO_" & Replace(Sender, " ", "_") & "@woodmac.com>"
End Function

Private Function CreateConfiguration() As CDO.Configuration
    
    Set CreateConfiguration = New CDO.Configuration
    
    CreateConfiguration.Load -1    ' CDO Source Defaults
    With CreateConfiguration.Fields
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "wmetech.woodmac.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        .Update
    End With
    
End Function
