Attribute VB_Name = "AutoProcessAndSaveFiles_ExcelDr"

Sub AutoProcessAndSaveFiles_ExcelDriven_Main()
    Dim wb As Workbook, ws As Worksheet
    Dim fileName As String, matchFound As Boolean
    Dim sftpData As Variant, item As Variant
    Dim groupName As String, filePattern As String, savePath As String, fileDateFormat As String
    Dim savedFiles As String, skippedFiles As String, fileDate As Date
    Dim folderPath As String, file As String, fileFullPath As String

    folderPath = Environ("USERPROFILE") & "\Downloads"
    savedFiles = "": skippedFiles = ""

    sftpData = LoadSFTPDataFromExcel() ' Load group data from SFTPfiles.xlsx

    file = Dir(folderPath & "*.csv")
    Do While file <> ""
        fileFullPath = folderPath & file
        Set wb = Workbooks.Open(fileFullPath)
        fileName = wb.Name
        matchFound = False

        For Each item In sftpData
            groupName = item(0): filePattern = item(1): savePath = item(2)
            fileDateFormat = ExtractDateFormatFromPattern(filePattern)

            If ValidateFileAgainstPattern(fileName, filePattern) Then
                matchFound = True
                fileDate = ExtractDateFromFileName(fileName, fileDateFormat)
                If fileDate = 0 Then
                    skippedFiles = skippedFiles & fileName & " ✗ Date extraction failed (" & fileDateFormat & ")" & vbCrLf
                    wb.Close False: Exit For
                End If

                Dim finalSavePath As String
                finalSavePath = BuildFinalSavePath(savePath, fileDate)
                CreateFullPath finalSavePath

                If FileExists(finalSavePath & "" & fileName) Then
                    skippedFiles = skippedFiles & fileName & " ✗ Duplicate found in " & finalSavePath & vbCrLf
                Else
                    wb.SaveAs fileName:=finalSavePath & "" & fileName, fileFormat:=xlCSV
                    savedFiles = savedFiles & fileName & " → " & finalSavePath & vbCrLf
                End If
                wb.Close False
                Exit For
            End If
        Next item

        If Not matchFound Then
            skippedFiles = skippedFiles & fileName & " ✗ No matching pattern" & vbCrLf
            wb.Close False
        End If

        file = Dir
    Loop

    MsgBox "Macro Complete!" & vbCrLf & String(40, "-") & vbCrLf & _
           "Saved Files:" & vbCrLf & savedFiles & vbCrLf & _
           "Skipped Files:" & vbCrLf & skippedFiles, vbInformation
End Sub
