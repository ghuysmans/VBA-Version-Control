Attribute VB_Name = "Project"
''
' This module is created to easily export and import VBA code
' into and from `./src` directory and run basic git commands.
'
' This is needed as Git can't read Excel files directly, but can read the
' source files that are exported.
'
' The export should be ran either on every save or on certain events.
' Note, can't use other modules with importing as they will be deleted and
' won't be able to be accessed.
'
' @author Robert Todar <robert@roberttodar.com>
' @status Development - Exporting seems to be working fine right now,
'                       but still having issues with importing. Looks Like
'                       I can import everything but Sheets, ThisWorkbook, and .frx files.
'                       Maybe create a solution to store just the text of code for sheets
'                       and workbook vs the code export.
' @ref {Microsoft Visual Basic For Applications Extensibility 5.3} VBComponets
' @ref {Microsoft Visual Basic For Applications Extensibility 5.3} VBComponet
' @ref {Microsoft Scripting Runtime} Scripting.FileSystemObject
''
Option Explicit

Private Type StoredReference
    Guid As String
    Major As Integer
    Minor As Integer
    FullPath As String
    Description As String
End Type

Private Const UNTOUCHED_COMPONENTS As String = _
    "{00020430-0000-0000-C000-000000000046}    OLE Automation" & _
    "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}    Microsoft Office 16.0 Object Library" & _
    "{0D452EE1-E08F-101A-852E-02608C4D0BB4}    Microsoft Forms 2.0 Object Library"

Public Doc As Object

' Directory where all source code will be stored. `./src`
Public Property Get SourceDirectory() As String
    SourceDirectory = joinPaths(Doc.path, "src")
End Property

' Helper function to run scripts from the root directory.
Public Function Bash(script As String, Optional keepCommandWindowOpen As Boolean = False) As Double
    ' cmd.exe Opens the command prompt.
    ' /S      Modifies the treatment of string after /C or /K (see below)
    ' /C      Carries out the command specified by string and then terminates
    ' /K      Carries out the command specified by string but remains
    ' cd      Change directory to the root directory.
    Bash = Shell("cmd.exe /S /" & IIf(keepCommandWindowOpen, "K", "C") & " cd " & Doc.path & " && " & script)
End Function

' Initiates a new Git Project in the current folder.
' Safe to run even if project is initialized.
Public Sub InitializeProject()
    Dim fso As New Scripting.FileSystemObject
    ' Create a default .gitignore file if it doesn't exist already
    ' @see https://git-scm.com/docs/gitignore
    Dim gitignorePath As String
    gitignorePath = joinPaths(Doc.path, ".gitignore")
    If Not fso.FileExists(gitignorePath) Then
        With fso.OpenTextFile(gitignorePath, ForWriting, True)
            .WriteLine ("# Packages")
            .WriteLine ("node_modules")
            .WriteBlankLines 1
            .WriteLine ("# Excel's Backup copies")
            .Write ("~$*.xl*")
            .Close
        End With
    End If
    
    ' Initialize git (safe even if it already exists)
    ' @see https://git-scm.com/docs/git-init
    Bash script:="git init", keepCommandWindowOpen:=False

    On Error Resume Next 'make this idempotent
    Doc.CustomDocumentProperties.add _
        Name:="vba-git", _
        LinkToContent:=False, _
        Type:=msoPropertyTypeBoolean, _
        Value:=True
End Sub

' Get the file extension for a VBComponent. That is the component name and the proper extension.
Private Function getVBComponentFilename(ByRef component As VBComponent) As String
    Select Case component.Type
        Case vbext_ComponentType.vbext_ct_ClassModule
            getVBComponentFilename = component.Name & ".cls"
            
        Case vbext_ComponentType.vbext_ct_StdModule
            getVBComponentFilename = component.Name & ".bas"
            
        Case vbext_ComponentType.vbext_ct_MSForm
            getVBComponentFilename = component.Name & ".frm"
            
        Case vbext_ComponentType.vbext_ct_Document
            getVBComponentFilename = component.Name & ".cls"
            
        Case Else
            ' @TODO: Need to think of possible throwing an error?
            ' Is it possible to get something else?? I don't think so
            ' Will need to double check this.
            Debug.Print "Unknown component"
    End Select
End Function

' Check to see if component exits in this current Project
Private Function componentExists(ByVal FileName As String) As Boolean
    Dim index As Long
    For index = 1 To Doc.VBProject.VBComponents.count
        Dim component As VBComponent
        Set component = Doc.VBProject.VBComponents(index)
        
        If getVBComponentFilename(component) = FileName Then
            componentExists = True
            Exit Function
        End If
    Next index
End Function

' Export all modules in this current workbook into a src dir
Public Sub ExportComponentsToSourceFolder()
    ' Make sure the source directory exists before adding to it.
    Dim fso As New Scripting.FileSystemObject
    If Not fso.FolderExists(SourceDirectory) Then
        fso.CreateFolder SourceDirectory
    Else
        Dim file As file
        For Each file In fso.GetFolder(SourceDirectory).Files
            file.Delete
        Next file
    End If
    
    ' Loop each component within this project and export to source directory.
    Dim index As Long
    For index = 1 To Doc.VBProject.VBComponents.count
        Dim component As VBComponent
        Set component = Doc.VBProject.VBComponents(index)
        
        ' Export component to the source directory using the components name and file extension.
        component.Export joinPaths(SourceDirectory, getVBComponentFilename(component))
    Next index
End Sub

' Import source code from the source Directory.
' This works by first deleting all current components,
' then importing all the components from the source directory.
'
' @status Testing && Development
' @warn This will cause files to overwrite that already exists.
' @warn This will also remove files not found in the source component.
Public Sub DangerouslyImportComponentsFromSourceFolder()
    If MsgBox("Are you sure you want to import from source folder? There is no going back!!!", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    Dim fso As New Scripting.FileSystemObject
    
    ' Remove current components to make room for the imported ones.
    Dim file As file
    For Each file In fso.GetFolder(SourceDirectory).Files
        ' If the component already, it needs to be deleted in order to
        ' import the file, otherwise an error is thrown.
        If componentExists(file.Name) And file.Name <> "Project.bas" Then
            Dim component As VBComponent
            Set component = Doc.VBProject.VBComponents.Item(fso.GetBaseName(file.Name))
            
            ' Unable to remove document type components (Sheets, workbook)
            If component.Type <> vbext_ct_Document Then
                ' This removes the component but doesn't from memory until
                ' after all code execution has completed.
                Doc.VBProject.VBComponents.Remove component
            End If
        End If
    Next file
    
    ' After all code is finished executing, the components removed above will
    ' finally be removed from memory.
    Application.OnTime Now, "saftleyImportAfterCleanup"
End Sub

Private Sub saftleyImportAfterCleanup()
    Dim fso As New Scripting.FileSystemObject

    Dim file As file
    For Each file In fso.GetFolder(SourceDirectory).Files
        If Not componentExists(file.Name) And fso.GetExtensionName(file.Name) <> "frx" Then
            ' Safe to import the source file as there are no conflicts of names.
            Doc.VBProject.VBComponents.Import joinPaths(SourceDirectory, file.Name)
        End If
    Next file
End Sub

' Converts the VBComponent enum to a string representation of type of component.
Private Function getVBComponentTypeName(ByRef component As VBComponent) As String
    Select Case component.Type
        Case vbext_ComponentType.vbext_ct_ClassModule
            getVBComponentTypeName = "Class Module"
            
        Case vbext_ComponentType.vbext_ct_StdModule
            getVBComponentTypeName = "Module"
            
        Case vbext_ComponentType.vbext_ct_MSForm
            getVBComponentTypeName = "Form"
            
        Case vbext_ComponentType.vbext_ct_Document
            getVBComponentTypeName = "Document"
            
        Case Else
            ' All components should be accounted for, this is just in case ;)
            Debug.Print "Unknown type: " & component.Type
    End Select
End Function

' Prints out details about a specific VBComponent. Used for
' @status Production
Private Function getComponentDetails(ByRef component As VBComponent) As String
    getComponentDetails = component.Name & vbTab _
                          & getVBComponentTypeName(component) & vbTab _
                          & getVBComponentFilename(component)
End Function

' Prints out details about all VBComponents in the current project
' @status Production
Public Property Get ComponentsDetails() As String
    Dim index As Long
    For index = 1 To Doc.VBProject.VBComponents.count
        Dim component As VBComponent
        Set component = Doc.VBProject.VBComponents(index)
        
        ComponentsDetails = ComponentsDetails & getComponentDetails(component) & vbNewLine
    Next index
End Property

' Prints out details about all VBComponents in the current project
' @status Development
Private Sub printDiffFromSourceFolder()
    Dim index As Long
    For index = 1 To Doc.VBProject.VBComponents.count
        Dim component As VBComponent
        Set component = Doc.VBProject.VBComponents(index)
        
        Debug.Print getVBComponentFilename(component)
    Next index
End Sub

' Helper function to join paths...
Private Function joinPaths(ParamArray paths() As Variant) As String
    Dim fso As New Scripting.FileSystemObject
    Dim index As Long
    For index = LBound(paths) To UBound(paths)
        joinPaths = fso.BuildPath(joinPaths, Replace(paths(index), "/", "\"))
    Next
End Function

Private Property Get referenceFile() As String
    referenceFile = joinPaths(SourceDirectory, "References.lst")
End Property

Private Sub ExportReferencesToSourceFolder()
    Dim r As Reference
    Dim f As Integer
    f = FreeFile
    Open referenceFile For Output As #f
    For Each r In Doc.VBProject.References
        If Not r.BuiltIn And InStr(UNTOUCHED_COMPONENTS, r.Guid) = 0 Then
            'Debug.Print r.Guid, r.Description
            Write #f, r.Guid, r.Major, r.Minor, r.FullPath, r.Description
        End If
    Next r
    Close #f
    'Shell "notepad.exe " & FileName, vbNormalFocus 'FIXME
End Sub

Private Sub DangerouslyImportReferencesFromSourceFolder()
    If MsgBox("Are you sure you want to import from source folder? There is no going back!!!", vbYesNo) = vbNo Then
        Exit Sub
    End If

    Dim f As Integer
    f = FreeFile
    Dim a() As StoredReference
    ReDim a(7)
    Dim i As Integer
    Open referenceFile For Input Shared As #f
    While Not EOF(f)
        Input #f, a(i).Guid, a(i).Major, a(i).Minor, a(i).FullPath, a(i).Description
        i = i + 1
        If i > UBound(a) Then ReDim Preserve a(2 * i - 1)
    Wend
    Close #f
    'TODO sync?
    Dim refs As References
    Set refs = Doc.VBProject.References
    Dim r As Reference
    For Each r In refs
        If Not r.BuiltIn And InStr(Untouched, r.Guid) = 0 Then _
            refs.Remove r
    Next r
    While i > 0
        i = i - 1
        refs.AddFromGuid a(i).Guid, a(i).Major, a(i).Minor
        'FIXME use FullPath?
    Wend
End Sub
