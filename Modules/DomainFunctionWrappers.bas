Attribute VB_Name = "DomainFunctionWrappers"
Option Compare Database
Option Explicit

Private Enum DomainFunctionWrapperEnum
    DLookup_Wrapper
    DCount_Wrapper
    DSum_Wrapper
    DMax_Wrapper
    DMin_Wrapper
    DAvg_Wrapper
End Enum

Private Function DomainFunctionWrapper(DomainFunction As DomainFunctionWrapperEnum, _
                                    Expr As String, _
                                    Domain As String, _
                                    Optional Criteria As String) As Variant
    On Error GoTo ErrorHandler
   
    Select Case DomainFunction
    Case DLookup_Wrapper
        DomainFunctionWrapper = DLookup(Expr, Domain, Criteria)
    Case DCount_Wrapper
        DomainFunctionWrapper = DCount(Expr, Domain, Criteria)
    Case DSum_Wrapper
        DomainFunctionWrapper = DSum(Expr, Domain, Criteria)
    Case DMax_Wrapper
        DomainFunctionWrapper = DMax(Expr, Domain, Criteria)
    Case DMin_Wrapper
        DomainFunctionWrapper = DMin(Expr, Domain, Criteria)
    Case DSum_Wrapper
        DomainFunctionWrapper = DSum(Expr, Domain, Criteria)
    Case DAvg_Wrapper
        DomainFunctionWrapper = DAvg(Expr, Domain, Criteria)
    Case Else
        ' Unexpected DomainFunction argument
        Debug.Assert False
    End Select

Done:
    Exit Function

ErrorHandler:
    Debug.Print Err.Number & " - " & Err.Description

End Function


'--------------------------------------------------------
' DLookupWrapper is just like DLookup only it will trap errors.
'--------------------------------------------------------
Public Function DLookupWrapper(Expr As String, Domain As String, Optional Criteria As String) As Variant
    DLookupWrapper = DomainFunctionWrapper(DLookup_Wrapper, Expr, Domain, Criteria)
End Function


'--------------------------------------------------------
' DLookupStringWrapper is just like DLookup wrapped in an Nz
' This will always return a String.
'--------------------------------------------------------
Public Function DLookupStringWrapper(Expr As String, Domain As String, Optional Criteria As String, Optional ValueIfNull As String = vbNullString) As String
    DLookupStringWrapper = Nz(DLookupWrapper(Expr, Domain, Criteria), ValueIfNull)
End Function


'--------------------------------------------------------
' DLookupNumberWrapper is just like DLookup wrapped in
' an Nz that defaults to 0.
'--------------------------------------------------------
Public Function DLookupNumberWrapper(Expr As String, Domain As String, Optional Criteria As String, Optional ValueIfNull = 0) As Variant
    DLookupNumberWrapper = Nz(DLookupWrapper(Expr, Domain, Criteria), ValueIfNull)
End Function


'--------------------------------------------------------
' DCountWrapper is just like DCount only it will trap errors.
'--------------------------------------------------------
Public Function DCountWrapper(Expr As String, Domain As String, Optional Criteria As String) As Long
    DCountWrapper = DomainFunctionWrapper(DCount_Wrapper, Expr, Domain, Criteria)
End Function


'--------------------------------------------------------
' DMaxWrapper is just like DMax only it will trap errors.
'--------------------------------------------------------
Public Function DMaxWrapper(Expr As String, Domain As String, Optional Criteria As String) As Long
    DMaxWrapper = Nz(DomainFunctionWrapper(DMax_Wrapper, Expr, Domain, Criteria), 0)
End Function


'--------------------------------------------------------
' DMinWrapper is just like DMin only it will trap errors.
'--------------------------------------------------------
Public Function DMinWrapper(Expr As String, Domain As String, Optional Criteria As String) As Long
    DMinWrapper = DomainFunctionWrapper(DMin_Wrapper, Expr, Domain, Criteria)
End Function


'--------------------------------------------------------
' DSumWrapper is just like DSum only it will trap errors.
'--------------------------------------------------------
Public Function DSumWrapper(Expr As String, Domain As String, Optional Criteria As String) As Long
    DSumWrapper = DomainFunctionWrapper(DSum_Wrapper, Expr, Domain, Criteria)
End Function


'--------------------------------------------------------
' DAvgWrapper is just like DAvg only it will trap errors.
'--------------------------------------------------------
Public Function DAvgWrapper(Expr As String, Domain As String, Optional Criteria As String) As Long
    DAvgWrapper = DomainFunctionWrapper(DAvg_Wrapper, Expr, Domain, Criteria)
End Function

