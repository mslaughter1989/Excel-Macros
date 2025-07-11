Attribute VB_Name = "Exporting_ExportToFolders"
Sub ExportModulesToFolders()
    Dim vbComp As Object
    Dim folderPicker As FileDialog
    Dim exportRoot As String
    Dim exportPath As String
    Dim baseName As String
    Dim prefix As String

    ' Prompt user to select export root folder
    Set folderPicker = Application.FileDialog(msoFileDialogFolderPicker)
    With folderPicker
        .Title = "Select export root folder for VBA modules"
        If .Show <> -1 Then Exit Sub
        exportRoot = .SelectedItems(1)
    End With

    ' Loop through all VBA components
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        ' Only export standard modules (Type 1)
        If vbComp.Type = 1 Then
            baseName = vbComp.Name
            prefix = ParseModulePrefix(baseName)

            ' Build and create folder if needed
            exportPath = exportRoot & "\" & prefix
            If Dir(exportPath, vbDirectory) = "" Then MkDir exportPath

            ' Export the module
            vbComp.Export exportPath & "\" & baseName & ".bas"
        End If
    Next

    MsgBox "Modules exported successfully into grouped folders!", vbInformation
End Sub

Private Function ParseModulePrefix(moduleName As String) As String
    ' Extract prefix before first underscore (if any)
    If InStr(moduleName, "_") > 0 Then
        ParseModulePrefix = Left(moduleName, InStr(moduleName, "_") - 1)
    Else
        ParseModulePrefix = "Uncategorized"
    End If
End Function

