Attribute VB_Name = "CopilotAutoSaveandclean"
Sub AutoProcessAndSaveFiles()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fileName As String
    Dim filePath As String
    Dim savePath As String
    Dim fileDate As Date
    Dim monthFolder As String
    Dim sftpFile As Workbook
    Dim sftpSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim groupName As String
    Dim filePattern As String
    Dim dateStr As String
    Dim regex As Object
    Dim match As Object
    
    ' Open the SFTPfiles.xlsx file
    Set sftpFile = Workbooks.Open("C:\Users\MichaelSlaughter\AppData\Roaming\Microsoft\Excel\XLSTART\SFTPfiles.xlsx")
    Set sftpSheet = sftpFile.Sheets(1)
    
    ' Get the last row of the SFTPfiles sheet
    lastRow = sftpSheet.Cells(sftpSheet.Rows.count, "A").End(xlUp).Row
    
    ' Create a regex object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Iterate over each file in the current directory
    filePath = ThisWorkbook.Path & "\"
    fileName = Dir(filePath & "*.csv")
    
    Do While fileName <> ""
        ' Iterate over each row in the SFTPfiles sheet
        For i = 2 To lastRow
            groupName = sftpSheet.Cells(i, 1).Value
            filePattern = sftpSheet.Cells(i, 2).Value
            savePath = sftpSheet.Cells(i, 3).Value
            
            ' Set the regex pattern based on the file pattern
            regex.pattern = Replace(Replace(Replace(filePattern, "mmddyyyy", "\d{8}"), "mmddyy", "\d{6}"), "yyyymmdd", "\d{8}")
            regex.IgnoreCase = True
            regex.Global = False
            
            ' Check if the file name matches the file pattern
            If regex.Test(fileName) Then
                ' Extract the date from the file name based on the pattern
                Set match = regex.Execute(fileName)(0)
                dateStr = match.Value
                
                If InStr(filePattern, "yyyymmdd") > 0 Then
                    fileDate = DateSerial(Left(dateStr, 4), Mid(dateStr, 5, 2), Right(dateStr, 2))
                ElseIf InStr(filePattern, "mmddyyyy") > 0 Then
                    fileDate = DateSerial(Right(dateStr, 4), Left(dateStr, 2), Mid(dateStr, 3, 2))
                ElseIf InStr(filePattern, "mmddyy") > 0 Then
                    fileDate = DateSerial(Right(dateStr, 2), Left(dateStr, 2), Mid(dateStr, 3, 2))
                End If
                
                ' Create the month folder
                monthFolder = Format(fileDate, "mm") & Format(fileDate, "mmm") & Format(fileDate, "yy")
                If Dir(savePath & "\" & monthFolder, vbDirectory) = "" Then
                    MkDir savePath & "\" & monthFolder
                End If
                
                ' Save the file to the final path
                If Dir(savePath & "\" & monthFolder & "\" & fileName) = "" Then
                    Name filePath & fileName As savePath & "\" & monthFolder & "\" & fileName
                Else
                    MsgBox "Duplicate file: " & fileName & " already exists in " & monthFolder
                End If
                
                Exit For
            End If
        Next i
        
        fileName = Dir
    Loop
    
    ' Close the SFTPfiles.xlsx file
    sftpFile.Close False
End Sub

