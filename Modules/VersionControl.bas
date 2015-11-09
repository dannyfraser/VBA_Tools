Attribute VB_Name = "VersionControl"
Option Explicit

Public Sub VersionControl()
    
    SaveCodeModules VBE.ActiveVBProject
    SaveTableSchemas
    SaveQueries
    SaveMacros
    
End Sub


Private Sub SaveCodeModules(VersionProject As VBProject)

    Dim VBC As VBComponent
    For Each VBC In VersionProject.VBComponents
    
        If VBC.CodeModule.CountOfLines > 0 Then
        
            Dim ModuleName As String
            Dim ModuleType As vbext_ComponentType
        
            ModuleName = VBC.CodeModule.Name
            ModuleType = VBC.Type
            
            If ModuleName Like "Form_*" Then
                Application.SaveAsText _
                    acForm, _
                    Right(ModuleName, Len(ModuleName) - Len("Form_")), _
                    GetOutputFolder(GetSubFolderName(ModuleType)) & ModuleName & ".frm"
            Else
                VBC.Export _
                    GetOutputFolder(GetSubFolderName(ModuleType)) & ModuleName & GetExtension(ModuleType)
            End If
            
        End If
                
    Next VBC

End Sub


Private Function GetExtension(ModuleType As vbext_ComponentType) As String

    Select Case ModuleType
    
        Case Is = vbext_ct_StdModule
            GetExtension = ".bas"
        
        Case Is = vbext_ct_ClassModule, vbext_ct_Document
            GetExtension = ".cls"
        
        Case Is = vbext_ct_MSForm
            GetExtension = ".frm"
        
        Case Is = vbext_ct_ActiveXDesigner
            GetExtension = ".axd"
    
    End Select

End Function

Private Function GetSubFolderName(ModuleType As vbext_ComponentType) As String

    Select Case ModuleType
    
        Case Is = vbext_ct_StdModule
            GetSubFolderName = "Modules"
        
        Case Is = vbext_ct_ClassModule
            GetSubFolderName = "Class Modules"
            
        Case Is = vbext_ct_Document, vbext_ct_MSForm
            GetSubFolderName = "Forms"
        
        Case Is = vbext_ct_ActiveXDesigner
            GetSubFolderName = "Other"
    
    End Select

End Function


Private Function VersionPath() As String
    
    VersionPath = Left(CurrentDb.Name, InStrRev(CurrentDb.Name, "\")) & "Components\"
    
    If Not Files.FolderExists(VersionPath) Then
        Files.CreateFolder (VersionPath)
    End If
    
End Function

Private Function GetOutputFolder(SubFolderName As String) As String
    GetOutputFolder = VersionPath & SubFolderName & "\"
    If Not Files.FolderExists(GetOutputFolder) Then
        Files.CreateFolder GetOutputFolder
    End If
End Function


Private Sub SaveTableSchemas()

    Dim OutputPath As String
    OutputPath = GetOutputFolder("Table Schemas")

    Dim t As TableDef
    For Each t In CurrentDb.TableDefs
        If Not (t.Name Like "MSys*") Then
            
            ExportXML _
                objecttype:=acExportTable, _
                DataSource:=t.Name, _
                schematarget:=OutputPath & t.Name & ".xsd"
                
        End If
    Next t
    
End Sub

Private Sub SaveQueries()

    Dim OutputPath As String
    OutputPath = GetOutputFolder("Queries")
    
    Dim q As QueryDef
    For Each q In CurrentDb.QueryDefs
    
        If Left(q.Name, 1) <> "~" Then
            Files.CreateTextFile OutputPath & q.Name & ".qry", q.SQL
        End If
        
    Next q

End Sub

Private Sub SaveMacros()
    
    Dim OutputPath As String
    OutputPath = GetOutputFolder("Macros")
    
    Dim DB As Database
    Set DB = CurrentDb
    
    Dim Macro As DAO.Document
    For Each Macro In DB.Containers("Scripts").Documents
        
        SaveAsText _
            objecttype:=acMacro, _
            objectname:=Macro.Name, _
            Filename:=OutputPath & Macro.Name & ".mcr"
            
    Next Macro
    
End Sub

