Attribute VB_Name = "Module1"
Sub ImportAllModulesFromFolder()
    Dim fso As Object
    Dim folderPicker As FileDialog
    Dim importFolder As String
    Dim rootFolder As Object

    ' Prompt for folder selection
    Set folderPicker = Application.FileDialog(msoFileDialogFolderPicker)
    folderPicker.Title = "Select the root folder of your macro repository"
    
    If folderPicker.Show <> -1 Then Exit Sub
    importFolder = folderPicker.SelectedItems(1)

    ' Set reference to FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rootFolder = fso.GetFolder(importFolder)

    ' Recursively import and rename .bas files
    Call ImportFromFolder(fso, rootFolder)

    MsgBox "All .bas modules imported and renamed!", vbInformation
End Sub

Private Sub ImportFromFolder(ByVal fso As Object, ByVal folder As Object)
    Dim file As Object, subFolder As Object
    Dim vbComp As Object
    Dim fileName As String, baseName As String

    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "bas" Then
            Set vbComp = ThisWorkbook.VBProject.VBComponents.Import(file.Path)
            baseName = fso.GetBaseName(file.Name)

            ' Rename the module if it's a standard module (1 = vbext_ct_StdModule)
            If vbComp.Type = 1 Then
                On Error Resume Next
                vbComp.Name = baseName
                On Error GoTo 0
            End If
        End If
    Next

    For Each subFolder In folder.SubFolders
        ImportFromFolder fso, subFolder
    Next
End Sub

