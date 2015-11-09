Attribute VB_Name = "Reference"
Option Explicit

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
   "RegOpenKeyExA" (ByVal Key As Long, ByVal SubKey As String, _
   ByVal Options As Long, ByVal Desired As Long, Result As Long) _
   As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal Key As Long, ByVal ValueName As String, _
   ByVal Reserved As Long, RegType As Long, _
   ByVal Path As String, Data As Long) As Long
                                                                                   
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Key As Long) As Long

Private Const REG_SZ As Long = 1
Private Const KEY_ALL_ACCESS = &H3F
Private Const HKEY_LOCAL_MACHINE = &H80000002

Public Sub ListReferences()
     
    'Adapted from "www.ozgrid.com/forum/showthread.php?t=22483"
    
    Dim VBProj As VBIDE.VBProject
    Set VBProj = VBE.ActiveVBProject
    
    Dim Reference As VBIDE.Reference
    For Each Reference In VBProj.References
        With Reference
            Debug.Print .Name
            Debug.Print .Description
            Debug.Print .GUID
            Debug.Print .Major
            Debug.Print .Minor
            Debug.Print .FullPath
            Debug.Print vbNullString
        End With
    Next
    
End Sub

Public Sub AddReferenceFromGuid(GUID As String, Major As Long, Minor As Long)
     
    'Adapted from "www.ozgrid.com/forum/showthread.php?t=22483"

    On Error GoTo ExitSub
    
    VBE.ActiveVBProject.References.AddFromGuid GUID, Major, Minor

ExitSub:
End Sub

Public Sub AddReferenceFromFile(FilePath As String)
     
    'Adapted from "www.ozgrid.com/forum/showthread.php?t=22483"

    On Error GoTo ExitSub
    
    VBE.ActiveVBProject.References.AddFromFile FilePath
    
ExitSub:
End Sub

Public Sub RemoveReference(ReferenceName As String)
     
    'Adapted from "www.ozgrid.com/forum/showthread.php?t=22483"

    On Error GoTo ExitSub
    
    Dim Reference As VBIDE.Reference
    For Each Reference In VBE.ActiveVBProject.References
        If Reference.Name = ReferenceName Then
            References.Remove Reference
        End If
    Next

ExitSub:
End Sub

Public Sub InstallExcelReference()
    
    Dim bolReferenceFound As Boolean
    bolReferenceFound = False
    
    Dim objReference As VBIDE.Reference
    For Each objReference In VBE.ActiveVBProject.References
        If objReference.Name = "Excel" Then
            bolReferenceFound = True
            Exit Sub
        End If
    Next
    
    If Not bolReferenceFound Then
        VBE.ActiveVBProject.References.AddFromFile GetExcelInstallPath
    End If

End Sub


Public Sub PrintAllEnvirons()

    Dim i As Integer
    i = 1
    Do While Environ(i) <> vbNullString
        Debug.Print Environ(i)
        i = i + 1
    Loop

End Sub


Private Function GetExcelInstallPath() As String

    Dim Key As Long
    Dim RetVal As Long
    Dim CLSID As String
    Dim Path As String
    Dim n As Long

   'First, get the clsid from the progid from the registry key:
   'HKEY_LOCAL_MACHINE\Software\Classes\<PROGID>\CLSID
   RetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Classes\Excel.Application\CLSID", 0&, KEY_ALL_ACCESS, Key)
   If RetVal = 0 Then
      RetVal = RegQueryValueEx(Key, "", 0&, REG_SZ, "", n)
      CLSID = Space(n)
      RetVal = RegQueryValueEx(Key, "", 0&, REG_SZ, CLSID, n)
      CLSID = Left(CLSID, n - 1)  'drop null-terminator
      RegCloseKey Key
   End If
   
   'Now that we have the CLSID, locate the server path at
   'HKEY_LOCAL_MACHINE\Software\Classes\CLSID\{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxx}\LocalServer32

    RetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Classes\CLSID\" & CLSID & "\LocalServer32", 0&, KEY_ALL_ACCESS, Key)
     If RetVal = 0 Then
      
          RetVal = RegQueryValueEx(Key, "", 0&, REG_SZ, "", n)
          Path = Space(n)
    
          RetVal = RegQueryValueEx(Key, "", 0&, REG_SZ, Path, n)
          Path = Left(Path, n - 1)
          RegCloseKey Key
    
    End If
    
    GetExcelInstallPath = Trim(Left(Path, InStr(Path, "/") - 1))

End Function

