VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordsetWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_rs As DAO.Recordset


Public Function GetRecordsetClone(rs As DAO.Recordset) As DAO.Recordset
    If Not m_rs Is Nothing Then
        Debug.Assert False ' This is only designed to be used once
    Else
        Set m_rs = rs.Clone
        Set GetRecordsetClone = m_rs
    End If
End Function


Sub OpenRecordset(Domain As String, _
                    Optional Criteria As String = vbNullString, _
                    Optional OrderBy As String = vbNullString, _
                    Optional RecordsetType As DAO.RecordsetTypeEnum = dbOpenDynaset, _
                    Optional RecordsetOptions As DAO.RecordsetOptionEnum _
                    )
    
    
    If Not m_rs Is Nothing Then
        ' Close the recordset so it can be re-used
        CloseRecordset
    End If
    
    Dim SQL As String
    SQL = "SELECT * FROM (" & Domain & ")"
    
    If Criteria <> vbNullString Then
        SQL = SQL & " WHERE " & Criteria
    End If
    
    If OrderBy <> "" Then
        SQL = SQL & " ORDER BY " & OrderBy
    End If
    
    On Error GoTo ErrorHandler
    Set m_rs = CurrentDb.OpenRecordset(SQL, RecordsetType, RecordsetOptions)

Done:
    Exit Sub
ErrorHandler:
    ' verify the private Recordset object was not set
    Debug.Assert m_rs Is Nothing
    
End Sub


Sub Delete()
    m_rs.Delete
End Sub


Sub AddNew()
    m_rs.AddNew
End Sub


Sub Edit()
    m_rs.Edit
End Sub


Sub Update()
    m_rs.Update
End Sub


Sub MoveNext()
    m_rs.MoveNext
End Sub

Sub MovePrevious()
    m_rs.MovePrevious
End Sub

Sub MoveLast()
    m_rs.MoveLast
End Sub

Sub MoveFirst()
    m_rs.MoveFirst
End Sub

Function EOF() As Boolean
    EOF = m_rs.EOF
End Function

Function BOF() As Boolean
    BOF = m_rs.BOF
End Function

Function RecordCount() As Long
    RecordCount = m_rs.RecordCount
End Function


Sub CloseRecordset()

    m_rs.Close
    Set m_rs = Nothing

End Sub


Public Property Get Recordset() As DAO.Recordset
Attribute Recordset.VB_UserMemId = 0
'    Attribute Recordset.VB_UserMemId = 0
    Set Recordset = m_rs
End Property


Private Sub Class_Terminate()
    If Not m_rs Is Nothing Then
        m_rs.Close
        Set m_rs = Nothing
    End If
End Sub

