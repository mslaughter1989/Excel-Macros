Attribute VB_Name = "APCgptMAIN"
Sub AutoProcessAndSaveFiles()
    Dim wb As Workbook, ws As Worksheet
    Dim rngP As Range, cell As Range
    Dim dict As Object
    Dim i As Long, j As Long
    Dim lastRow As Long, lastCol As Long
    Dim headerCell As Range, cleanHeader As String
    Dim zipColFound As Boolean, apexLogicApplied As Boolean
    Dim keywordList As Variant, keyword As Variant
    Dim zipFormattedList As String, apexAppliedList As String, untouchedList As String
    Dim fileName As String
    Dim sftpData As Variant, item As Variant
    Call LoadSFTPData(sftpData)
    Dim groupName As String, filePattern As String, savePath As String
    Dim matchFound As Boolean, fileDate As Date, dtString As String
    Dim monthFolder As String
    Dim regex As Object, matches As Object
    Dim tempName As String, tempPath As String, finalPath As String
    zipFormattedList = ""
    apexAppliedList = ""
    untouchedList = ""

    ' Embedded SFTP data
sftpData = CombineArrays(sftpDataPart1, sftpDataPart2, sftpDataPart3)



End Sub

