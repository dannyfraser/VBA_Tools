Attribute VB_Name = "GenericFunctions"
Option Explicit

Public Function Exists(Collection As Object, Key As String) As Boolean
    
    'Collection is declared as an Object to allow iteration over
    'DAO Properties collections/TableDefs etc. as well as VBA Collection objects
    
    On Error GoTo DoesNotExist
    Dim s As String
    s = TypeName(Collection(Key))
    Exists = True
    
    Exit Function
    
DoesNotExist:
    Exists = False
    
End Function


Public Function IsArrayInitialized(ArrayToCheck As Variant) As Boolean
    
    On Error GoTo IsArrayInitialized_Exit
    
    If UBound(ArrayToCheck) > -1 Then
        IsArrayInitialized = True
    End If
    
    Exit Function
    
IsArrayInitialized_Exit:
    IsArrayInitialized = False

End Function


Public Function AddToStringList(List As String, NewItem As String) As String

    If Len(List) > 0 Then
        AddToStringList = List & "," & Sanitize(NewItem)
    Else
        AddToStringList = List & Sanitize(NewItem)
    End If

End Function

Public Function AddToNumberList(List As String, NewItem As Variant) As String

    If Len(List) > 0 Then
        AddToNumberList = List & "," & CStr(NewItem)
    Else
        AddToNumberList = List & CStr(NewItem)
    End If
    
End Function

Public Sub EnableXPath(XMLDoc As DOMDocument)
    XMLDoc.SetProperty "SelectionLanguage", "XPath"
End Sub
