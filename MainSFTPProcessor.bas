Attribute VB_Name = "MainSFTPProcessor"
Sub ProcessAndSaveAllWorkbooks()
    Dim wb As Workbook, ws As Worksheet
    Dim fileName As String, fileDate As String
    Dim fileMonth As String, fileYear As String
    Dim folderName As String, fullPath As String
    Dim monthNames As Variant
    Dim regex As Object, matches As Object
    Dim mappings As Object, key As Variant
    Dim logPath As String, logFile As Integer
    Dim zipFormatted As Boolean, apexApplied As Boolean
    Dim headerCell As Range, cleanHeader As String
    Dim keywordList As Variant, keyword As Variant
    Dim lastCol As Long, lastRow As Long
    Dim dict As Object, cell As Range, rngP As Range
    Dim i As Long, j As Long
    Dim msg As String

    ' Load mappings from SFTPMappings module
    Set mappings = LoadMappings()

    ' Month names
    monthNames = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _
                       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    ' Prepare log file
    logPath = ThisWorkbook.Path & "\SaveLog.txt"
    logFile = FreeFile
    Open logPath For Output As #logFile
    Print #logFile, "Save Log - " & Now
    Print #logFile, String(60, "-")

    ' Keywords for ZIP formatting
    keywordList = Array("zip", "zipcode", "zip code", "postalcode", "postal code")

    ' Loop through all open workbooks
    For Each wb In Application.Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            Set ws = wb.Sheets(1)
            fileName = wb.Name
            If InStr(fileName, ".") > 0 Then
                fileName = Left(fileName, InStrRev(fileName, ".") - 1)
            End If

            zipFormatted = False
            apexApplied = False

            ' === ZIP Formatting ===
            lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
            For j = 1 To lastCol
                Set headerCell = ws.Cells(1, j)
                cleanHeader = LCase(Replace(Replace(Replace(Trim(headerCell.Value), "_", ""), "-", ""), " ", ""))
                For Each keyword In keywordList
                    If InStr(cleanHeader, Replace(LCase(keyword), " ", "")) > 0 Then
                        ws.Columns(j).NumberFormat = "00000"
                        zipFormatted = True
                        Exit For
                    End If
                Next keyword
            Next j

            ' === APEX Logic ===
            If InStr(1, UCase(wb.Name), "APEX") > 0 Then
                apexApplied = True
                Set rngP = ws.Range("P2:P" & ws.Cells(ws.Rows.count, "P").End(xlUp).Row)
                Set dict = CreateObject("Scripting.Dictionary")
                For Each cell In rngP
                    If Not dict.Exists(cell.Value) Then
                        dict.Add cell.Value, 1
                    Else
                        dict(cell.Value) = dict(cell.Value) + 1
                    End If
                Next cell
                For i = ws.Cells(ws.Rows.count, "P").End(xlUp).Row To 2 Step -1
                    If dict.Exists(ws.Cells(i, "P").Value) And dict(ws.Cells(i, "P").Value) > 1 _
                    And ws.Cells(i, "N").Value <> "" Then
                        ws.Rows(i).Delete
                    End If
                Next i
                Set dict = CreateObject("Scripting.Dictionary")
                lastRow = ws.Cells(ws.Rows.count, "P").End(xlUp).Row
                For Each cell In ws.Range("P2:P" & lastRow)
                    If Not dict.Exists(cell.Value) Then
                        dict.Add cell.Value, cell.Row
                    Else
                        If ws.Cells(dict(cell.Value), "M").Value < ws.Cells(cell.Row, "M").Value Then
                            ws.Rows(dict(cell.Value)).Delete
                            dict(cell.Value) = cell.Row
                        Else
                            ws.Rows(cell.Row).Delete
                        End If
                    End If
                Next cell
            End If

            ' === Save Logic ===
            Set regex = CreateObject("VBScript.RegExp")
            regex.pattern = "\d{8}"
            regex.Global = False
            regex.IgnoreCase = True
            If regex.Test(fileName) Then
                Set matches = regex.Execute(fileName)
                fileDate = matches(0)
                fileMonth = Left(fileDate, 2)
                fileYear = Right(fileDate, 2)
                For Each key In mappings.Keys
                    If InStr(fileName, Left(key, InStr(key, "_mm") - 1)) > 0 Then
                        folderName = fileMonth & monthNames(CInt(fileMonth) - 1) & fileYear
                        fullPath = mappings(key) & "\" & folderName & "\"
                        If Dir(fullPath, vbDirectory) = "" Then MkDirRecursive fullPath
                        wb.SaveCopyAs fullPath & wb.Name
                        Print #logFile, "✔ Saved: " & wb.Name & " → " & fullPath
                        Exit For
                    End If
                Next key
            Else
                Print #logFile, "✘ Skipped (no date): " & wb.Name
            End If
        End If
    Next wb

    Close #logFile
    MsgBox "Processing complete. See SaveLog.txt for details.", vbInformation
End Sub

Sub MkDirRecursive(ByVal fullPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(fullPath) Then fso.CreateFolder fullPath
End Sub
