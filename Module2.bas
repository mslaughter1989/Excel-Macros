Attribute VB_Name = "Module2"
Sub CreateFullPath(fullPath As String)
    Dim fso As Object, pathParts() As String
    Dim curPath As String
    Dim i As Integer
    Set fso = CreateObject("Scripting.FileSystemObject")

    pathParts = Split(fullPath, "\")
    curPath = pathParts(0)

    For i = 1 To UBound(pathParts)
        curPath = curPath & "\" & pathParts(i)
        If Not fso.FolderExists(curPath) Then fso.CreateFolder curPath
    Next i
End Sub

Function CombineArrays(ParamArray arrays() As Variant) As Variant
    Dim totalLength As Long
    Dim i As Long, j As Long, index As Long
    Dim result() As Variant

    ' Count total length
    For i = LBound(arrays) To UBound(arrays)
        totalLength = totalLength + UBound(arrays(i)) - LBound(arrays(i)) + 1
    Next i

    ReDim result(0 To totalLength - 1)
    index = 0

    ' Copy elements
    For i = LBound(arrays) To UBound(arrays)
        For j = LBound(arrays(i)) To UBound(arrays(i))
            result(index) = arrays(i)(j)
            index = index + 1
        Next j
    Next i

    CombineArrays = result
End Function
