Attribute VB_Name = "SaveSFTP"
Sub SaveCopyOfActiveCSVToMonthlyFolder()
    Dim filePath As String
    Dim fileName As String
    Dim fileDate As String
    Dim fileMonth As String
    Dim fileYear As String
    Dim folderName As String
    Dim fullPath As String
    Dim basePath As String
    Dim monthNames As Variant
    Dim regex As Object
    Dim matches As Object

    ' Get the full path and file name of the active workbook
    filePath = ActiveWorkbook.FullName
    fileName = ActiveWorkbook.Name

    ' Remove extension from file name
    If InStr(fileName, ".") > 0 Then
        fileName = Left(fileName, InStrRev(fileName, ".") - 1)
    End If

    ' Debug: Show parsed file name
    MsgBox "Parsed file name (no extension): " & fileName

    ' Use RegExp to find the 8-digit date
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = "\d{8}"  ' Match any 8-digit number
    regex.Global = False
    regex.IgnoreCase = True
    regex.MultiLine = False

    If regex.Test(fileName) Then
        Set matches = regex.Execute(fileName)
        fileDate = matches(0)
    Else
        MsgBox "Could not find a valid 8-digit date (mmddyyyy) in the file name: " & fileName
        Exit Sub
    End If

    ' Extract month and year from the date part
    fileMonth = Left(fileDate, 2)
    fileYear = Right(fileDate, 2)

    ' Month names array
    monthNames = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _
                       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    ' Construct folder name like 08Aug25
    folderName = fileMonth & monthNames(CInt(fileMonth) - 1) & fileYear

    ' Base path
    basePath = "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\21st Century Equipment\"

    ' Full path
    fullPath = basePath & folderName & "\"

    ' Create folder if it doesn't exist
    If Dir(fullPath, vbDirectory) = "" Then
        MkDir fullPath
    End If

    ' Save a copy of the current file to the new location
    ActiveWorkbook.SaveCopyAs fullPath & ActiveWorkbook.Name

    MsgBox "A copy of the file was saved to: " & fullPath & ActiveWorkbook.Name
End Sub


