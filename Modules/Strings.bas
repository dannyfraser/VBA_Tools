Attribute VB_Name = "Strings"
Option Explicit


Public Function Concat(FirstPart As String, SecondPart As String) As String

    'Concatenates two strings by pre-allocating space.
    'Faster than adding to a string using '&'

    Dim NewLength As Long
    NewLength = Len(FirstPart) + Len(SecondPart)
    
    Concat = Space(NewLength)
    
    Mid(Concat, 1, Len(FirstPart)) = FirstPart
    Mid(Concat, Len(FirstPart) + 1, Len(SecondPart)) = SecondPart
    
End Function


Public Function RemoveIllegalChars(ByVal Name As String) As String

    'This will make a valid filename from an otherwise invalid string.
    'The original string remains intact (unless specified outside of this function) as it is passed ByVal

    Dim IllegalChars() As Variant
    IllegalChars = Array("<", ">", "|", "/", "*", "\", "?", """", ":")
    
    Dim Char As Variant
    For Each Char In IllegalChars
        Name = Replace(Name, Char, vbNullString, 1)
    Next Char
    
    RemoveIllegalChars = Name
        
End Function


Public Function IsAlphaNumeric(Check As String) As Boolean

    Dim re As RegExp
    
    Set re = New RegExp
    
    re.IgnoreCase = True
    re.Pattern = "^[A-Z0-9]*$"
        
    IsAlphaNumeric = re.Test(Check)
    
End Function


Public Function ReverseString(Text As String) As String

    ReverseString = Space(Len(Text))
    
    Dim CharPosition As Integer
    For CharPosition = Len(Text) To 1 Step -1
        Mid(ReverseString, CharPosition, 1) = Mid(Text, Len(Text) - CharPosition + 1, 1)
    Next CharPosition

End Function


Public Function IsValidEmailAddress(Address As String) As Boolean

    With New RegExp
        .IgnoreCase = True
        .Pattern = "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}$"
        IsValidEmailAddress = .Test(Address)
    End With

End Function


Public Function AddPathSeparator(Path As String) As String

    If Right$(Path, 1) <> "\" Then
        AddPathSeparator = Path & "\"
    Else
        AddPathSeparator = Path
    End If

End Function


Public Function RemovePathSeparator(Path As String) As String

    If Right$(Path, 1) = "\" Then
        RemovePathSeparator = Left(Path, Len(Path) - 1)
    Else
        RemovePathSeparator = Path
    End If
   
End Function


Public Function StripBrackets(Text As String) As String

    If Len(Replace(Text, "(", vbNullString)) <> Len(Replace(Text, ")", vbNullString)) Then
        Exit Function
    End If
    
    Dim LastOpenBracket As Integer
    LastOpenBracket = InStrRev(Text, "(")
    
    If LastOpenBracket <> 0 Then
    
        Dim NextCloseBracket As Integer
        NextCloseBracket = InStr(LastOpenBracket, Text, ")")
        
        If NextCloseBracket = 0 Then
            Exit Function
        End If
        
        Dim StrippedText As String
        StrippedText = RTrim(Left(Text, LastOpenBracket - 1)) & Right(Text, Len(Text) - NextCloseBracket)
        
        StripBrackets = StripBrackets(StrippedText)
        
    Else
    
        StripBrackets = Text
    
    End If

End Function

Public Function RemoveDoubleSpaces(ByVal StringToAmend As String) As String

    Do While InStr(1, StringToAmend, "  ") > 0
        StringToAmend = Replace(StringToAmend, "  ", " ")
    Loop

    RemoveDoubleSpaces = StringToAmend
    
End Function
