Attribute VB_Name = "Files"
Option Explicit

Public Enum DriveType
    Unknown = 0
    Absent = 1
    Removable = 2
    Fixed = 3
    Remote = 4
    CDRom = 5
    RamDisk = 6
End Enum

'Used to get UNC path
Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" _
        (ByVal LocalName As String, ByVal RemoteName As String, RemoteName As Long) As Long
        
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" _
        (ByVal nDrive As String) As Long

Private FSO As New Scripting.FileSystemObject


Public Function FileExists(Location As String) As Boolean
    FileExists = FSO.FileExists(Location)
End Function


Public Function FolderExists(Location As String) As Boolean
    FolderExists = FSO.FolderExists(Location)
End Function


Public Sub DeleteFile(Location As String)
    If FileExists(Location) Then
       Kill Location
    End If
End Sub


Public Sub DeleteFolder(Location As String)
    If FolderExists(Location) Then
       FSO.DeleteFolder Location
    End If
End Sub


Public Sub CopyFile(SourceLocation As String, DestinationLocation As String)
    If FileExists(SourceLocation) Then
       FileCopy SourceLocation, DestinationLocation
    End If
End Sub


Public Sub CopyFolder(SourceLocation As String, DestinationLocation As String)
    
    If FolderExists(SourceLocation) Then
        If Not FolderExists(DestinationLocation) Then
            CreateFolder DestinationLocation
        End If
        
        Dim f As Folder
        Set f = FSO.GetFolder(SourceLocation)
        
        If f.SubFolders.Count = 0 Then
            FSO.CopyFolder Strings.RemovePathSeparator(Trim(SourceLocation)), Strings.RemovePathSeparator(Trim(DestinationLocation)), True
        Else
            FSO.CopyFolder Trim(SourceLocation) & "*.*", Trim(DestinationLocation), True
        End If
    End If
    
End Sub


Public Sub CreateFolder(FolderLocation As String)
    
    Dim PathParts() As String
    PathParts = Split(FolderLocation, "\")
    
    Dim PathPart As Variant
    Dim FullPath As String
    For Each PathPart In PathParts
        If Len(PathPart) > 0 Then
            FullPath = FullPath & PathPart & "\"
            If Not FolderExists(FullPath) Then
                MkDir FullPath
            End If
        End If
    Next PathPart

End Sub


Public Function GetFilePath(FilePath As String) As String

    'Returns a directory from a filepath
    GetFilePath = Left$(FilePath, InStrRev(FilePath, "\"))
    
End Function


Public Function GetFileNameFromPath(Path As String) As String

    'Returns a file name from a filepath - like Dir but will return the filename even for a non-existent path

    GetFileNameFromPath = vbNullString
    
    Dim PathArray() As String
    PathArray = Split(Path, "\")
    
    If PathArray(UBound(PathArray)) = vbNullString Then
        GetFileNameFromPath = PathArray(UBound(PathArray) - 1)
    Else
        GetFileNameFromPath = PathArray(UBound(PathArray))
    End If
    
End Function


Public Function StripExtension(Filename As String) As String
    If InStrRev(GetFileNameFromPath(Filename), ".") = 0 Then
        StripExtension = Filename
    Else
        StripExtension = Left(GetFileNameFromPath(Filename), InStrRev(GetFileNameFromPath(Filename), ".") - 1)
    End If
End Function


Public Function FileList(Path As String, Optional Filter As String = "*.*") As Collection

    'Returns a collection of Scripting.File objects keyed by path

    Dim StartingFolder As Scripting.Folder
    Set StartingFolder = FSO.GetFolder(Path)
    
    Set FileList = New Collection
    RecursiveGetFiles StartingFolder, FileList, Filter
    
End Function


Private Function RecursiveGetFiles(StartingFolder As Scripting.Folder, ByRef FullFileList As Collection, Optional Filter As String = "*.*")
    
    Dim File As Scripting.File
    For Each File In StartingFolder.Files
        If File.Name Like Filter Then
            FullFileList.Add File, File.Path
        End If
    Next File
    
    Dim SubFolder As Scripting.Folder
    For Each SubFolder In StartingFolder.SubFolders
        RecursiveGetFiles SubFolder, FullFileList
    Next SubFolder
        
End Function


Public Function GetFileOwner(FilePath As String) As String
    
    'Returns the name of the person who currently has ownership of a file
    
    Dim File As Integer
    File = FreeFile(0)
    
    Open FilePath For Binary As #File
    
    Dim FileText As String
    FileText = Space(LOF(File))
    Get File, , FileText
    Close #File
    
    Dim SpaceFlag As String, SpacePos As Long
    SpaceFlag = Space(2)
    SpacePos = InStr(1, FileText, SpaceFlag)
    
    Dim NullFlag As String, NameStart As Long
    NullFlag = vbNullChar & vbNullChar
    NameStart = InStrRev(FileText, NullFlag, SpacePos) + Len(NullFlag)
    
    Dim INameLen As Byte
    INameLen = Asc(Mid(FileText, NameStart - 3, 1))
    
    GetFileOwner = Mid(FileText, NameStart, INameLen)

End Function

Public Function GetUNCPath(DriveLetter As String) As String

    On Local Error GoTo GetUNCPath_Err

    Const ERROR_BAD_DEVICE = 1200&
    Const ERROR_CONNECTION_UNAVAIL = 1201&
    Const ERROR_EXTENDED_ERROR = 1208&
    Const ERROR_MORE_DATA = 234
    Const ERROR_NOT_SUPPORTED = 50&
    Const ERROR_NO_NET_OR_BAD_PATH = 1203&
    Const ERROR_NO_NETWORK = 1222&
    Const ERROR_NOT_CONNECTED = 2250&
    Const NO_ERROR = 0
    
    Const INVALID_CHAR As Integer = 0

    Dim LocalName As String
    LocalName = Replace(UCase(DriveLetter), "\", vbNullString)
    If InStr(1, LocalName, ":") = 0 Then
        LocalName = LocalName & ":"
    End If

    Dim RemoteName As String
    RemoteName = Space(256)

    Dim RemoteNameLength As Long
    RemoteNameLength = Len(RemoteName)

    Dim ConnType As Long
    ConnType = WNetGetConnection(LocalName, RemoteName, RemoteNameLength)

    Dim Msg As String
    Select Case ConnType
        Case ERROR_BAD_DEVICE
            Msg = "Error: Bad Device"
        Case ERROR_CONNECTION_UNAVAIL
            Msg = "Error: Connection Un-Available"
        Case ERROR_EXTENDED_ERROR
            Msg = "Error: Extended Error"
        Case ERROR_MORE_DATA
            Msg = "Error: More Data"
        Case ERROR_NOT_SUPPORTED
            Msg = "Error: Feature not Supported"
        Case ERROR_NO_NET_OR_BAD_PATH
            Msg = "Error: No Network Available or Bad Path"
        Case ERROR_NO_NETWORK
            Msg = "Error: No Network Available"
        Case ERROR_NOT_CONNECTED
            Msg = "Error: Not Connected"
        Case NO_ERROR
            ' all is successful...
    End Select

    If ConnType <> NO_ERROR Then Err.Raise (ConnType)

    GetUNCPath = Replace(Trim(Left$(RemoteName, RemoteNameLength)), Chr(INVALID_CHAR), vbNullString)
    Exit Function

GetUNCPath_Err:
    GetUNCPath = Msg

End Function


Public Function IsUNCPath(Path As String) As Boolean

    IsUNCPath = (Left$(Path, 2) = "\\")

End Function

Public Function IsLocalPath(Path As String) As Boolean
'    Dim s As String
'    If IsUNCPath(Path) Then
'        IsLocalPath = False
'        Exit Function
'    Else
'        s = GetUNCPath(Path)
'        IsLocalPath = (s = "Error: Bad Device")
'    End If
    IsLocalPath = Not (Left(ConvertToUNCPath(Path), 2) = "\\")
End Function


Public Function ConvertToUNCPath(Path As String) As String

    If IsUNCPath(Path) Or DriveType(Left(Path, 1)) = Fixed Then
            
        ConvertToUNCPath = Path
            
    Else
        
        Dim FirstSep As Integer
        FirstSep = InStr(1, Path, "\")
        
        Dim PathWithoutDrive As String
        PathWithoutDrive = Right(Path, Len(Path) - FirstSep)
        
        ConvertToUNCPath = GetUNCPath(Left(Path, 1)) & "\" & PathWithoutDrive
            
    End If
    
End Function


Public Function GetOpenFilename(Optional Title As String = "Select File To Open", _
    Optional Filter As Variant = Empty, _
    Optional FilterIndex As Variant = Empty, _
    Optional InitialPath As String = vbNullString) As String
    
    On Error GoTo GetOpenFileName_Error
    
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = Title
        
        If IsArray(Filter) And IsArray(FilterIndex) Then

            Dim i As Integer
            For i = LBound(Filter) To UBound(Filter)
                .Filters.Add Filter(i), FilterIndex(i)
            Next i
            
        ElseIf Filter <> vbNullString And FilterIndex <> vbNullString Then
        
            .Filters.Add Filter, FilterIndex
            
        Else
        
            .Filters.Add "All Files", "*.*", 1
            
        End If
        
        .InitialFileName = IIf(InitialPath = vbNullString, CurrentProject.Path, InitialPath)
        
        .Show
        
        GetOpenFilename = .SelectedItems(1)
    
    End With
    
GetOpenFileName_Exit:

    Exit Function

GetOpenFileName_Error:

    GetOpenFilename = "False"
    GoTo GetOpenFileName_Exit
    
End Function


Public Function GetFolderName(Optional Title As String = "Select Folder", Optional InitialPath As String = vbNullString) As String

    On Error GoTo NothingSelected

    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = Title
        .InitialFileName = IIf(InitialPath = vbNullString, CurrentProject.Path, InitialPath)
        .Show
        GetFolderName = .SelectedItems(.SelectedItems.Count)
    End With
    
    Exit Function
NothingSelected:
    GetFolderName = "False"

End Function


Public Function GetSaveAsFilename(Optional Title As String = "Save File As...", Optional InitialFileName As String = vbNullString) As String

    With Application.FileDialog(msoFileDialogSaveAs)
        
        .AllowMultiSelect = False
        .Title = Title
        .InitialFileName = InitialFileName
        .Show
        If .SelectedItems.Count = 0 Then
            GetSaveAsFilename = "False"
        Else
            GetSaveAsFilename = .SelectedItems(.SelectedItems.Count)
        End If

    End With

End Function


Public Function FileSize(Path As String) As Double
    
    Dim File As Object
    Set File = FSO.GetFile(Path)

    FileSize = File.Size / 10 ^ 6

End Function


Public Sub CreateTextFile(Location As String, Optional TextToWrite As String)

    CreateFolder (Left(Location, InStrRev(Location, "\")))
    
    Dim t As TextStream
    Set t = FSO.CreateTextFile(Location, True)
    
    If TextToWrite <> vbNullString Then
        t.Write TextToWrite
    End If
    
    t.Close

End Sub

Public Sub ReplaceTextInFile(Path As String, OldText As String, NewText As String)

    Dim OldFile As TextStream
    Set OldFile = FSO.OpenTextFile(Path, ForReading, False, TristateUseDefault)
    
    Dim NewFileContent As String
    NewFileContent = Replace(OldFile.ReadAll, OldText, NewText)
    
    OldFile.Close
    Kill Path
    
    Dim NewFile As TextStream
    Set NewFile = FSO.CreateTextFile(Path, True)
    
    NewFile.Write NewFileContent
    NewFile.Close

End Sub


Public Function FileLocked(FilePath As String) As Boolean

    On Error GoTo Locked
    
    Dim FileNumber As Integer
    FileNumber = FreeFile(0)
    
    Open FilePath For Binary Access Read Write Lock Read Write As #FileNumber
    Close #FileNumber
    FileLocked = False
   
    Exit Function
   
Locked:
    FileLocked = True
   
End Function

Public Function LastModified(FilePath As String) As Date
    LastModified = FileDateTime(FilePath)
End Function

Public Sub SelectFile(FilePath As String)
    If FileExists(FilePath) Then
        Shell "explorer.exe /select , " & FilePath, vbNormalFocus
    End If
End Sub

Public Function GetUniqueSaveFileName(Optional Title As String = "Save File As...", Optional InitialFileName As String = vbNullString) As String
    
    Do While True
        
        Dim InitialFilePath As String
        If GetUniqueSaveFileName = vbNullString Then
            InitialFilePath = InitialFileName
        Else
            InitialFilePath = GetUniqueSaveFileName
        End If
        
        GetUniqueSaveFileName = _
            GetSaveAsFilename( _
                Title, _
                InitialFilePath _
            )
        If GetUniqueSaveFileName = "False" Then Exit Function
        
        If Files.FileLocked(GetUniqueSaveFileName) Then
            MsgBox "File currently in use. Please use a different file name.", vbInformation, "Error"
        Else
            Exit Do
        End If
        
    Loop
    
End Function


Public Function DriveType(DriveLetter As String) As DriveType
    DriveType = GetDriveType(Left(DriveLetter, 1) & ":")
End Function
