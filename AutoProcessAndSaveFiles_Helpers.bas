Attribute VB_Name = "AutoProcessAndSaveFiles_Helpers"

Function LoadSFTPDataFromExcel() As Variant
    Dim ws As Worksheet, wb As Workbook, lastRow As Long
    Dim arr(), i As Long

    Set wb = Workbooks.Open(Environ("USERPROFILE") & "\AppData\Roaming\Microsoft\Excel\XLSTART\SFTPfiles.xlsx")
    Set ws = wb.Sheets("Sheet1")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

    ReDim arr(1 To lastRow - 1, 1 To 3)
    For i = 2 To lastRow
        arr(i - 1, 1) = ws.Cells(i, 1).Value  ' Group Name
        arr(i - 1, 2) = ws.Cells(i, 2).Value  ' File Name Formatting
        arr(i - 1, 3) = ws.Cells(i, 3).Value  ' Save Path
    Next i

    wb.Close False
    LoadSFTPDataFromExcel = arr
End Function

Function ValidateFileAgainstPattern(fileName As String, filePattern As String) As Boolean
    Dim basePattern As String
    basePattern = Split(filePattern, "_")(0)
    ValidateFileAgainstPattern = (InStr(fileName, basePattern) > 0)
End Function

Function ExtractDateFormatFromPattern(filePattern As String) As String
    If InStr(filePattern, "yyyymmdd") > 0 Then
        ExtractDateFormatFromPattern = "yyyymmdd"
    ElseIf InStr(filePattern, "mmddyyyy") > 0 Then
        ExtractDateFormatFromPattern = "mmddyyyy"
    ElseIf InStr(filePattern, "mmddyy") > 0 Then
        ExtractDateFormatFromPattern = "mmddyy"
    Else
        ExtractDateFormatFromPattern = ""
    End If
End Function

Function ExtractDateFromFileName(fileName As String, dateFormat As String) As Date
    Dim regEx As Object, matches As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = False

    Select Case dateFormat
        Case "yyyymmdd": regEx.pattern = "\d{8}"
        Case "mmddyyyy": regEx.pattern = "\d{8}"
        Case "mmddyy": regEx.pattern = "\d{6}"
        Case Else: ExtractDateFromFileName = 0: Exit Function
    End Select

    If regEx.Test(fileName) Then
        Set matches = regEx.Execute(fileName)
        Dim dtStr As String: dtStr = matches(0)
        On Error GoTo Fail
        Select Case dateFormat
            Case "yyyymmdd": ExtractDateFromFileName = DateSerial(Left(dtStr, 4), Mid(dtStr, 5, 2), Right(dtStr, 2))
            Case "mmddyyyy": ExtractDateFromFileName = DateSerial(Right(dtStr, 4), Left(dtStr, 2), Mid(dtStr, 3, 2))
            Case "mmddyy": ExtractDateFromFileName = DateSerial(2000 + Right(dtStr, 2), Left(dtStr, 2), Mid(dtStr, 3, 2))
        End Select
        Exit Function
    End If

Fail:
    ExtractDateFromFileName = 0
End Function

Function BuildFinalSavePath(basePath As String, fileDate As Date) As String
    BuildFinalSavePath = basePath & "\" & Format(fileDate, "mmMM") & Right(Year(fileDate), 2)
End Function

Sub CreateFullPath(fullPath As String)
    Dim fso As Object, pathParts() As String, curPath As String
    Dim i As Integer
    Set fso = CreateObject("Scripting.FileSystemObject")

    pathParts = Split(fullPath, "\")
    curPath = pathParts(0)
    For i = 1 To UBound(pathParts)
        curPath = curPath & "\" & pathParts(i)
        If Not fso.FolderExists(curPath) Then fso.CreateFolder curPath
    Next i
End Sub

Function FileExists(filePath As String) As Boolean
    FileExists = (Dir(filePath) <> "")
End Function
