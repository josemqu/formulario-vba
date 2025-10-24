Attribute VB_Name = "modTable"
Option Explicit

Public Function EnsureTable(WS As Worksheet, tableName As String, Headers As Variant) As ListObject
    Dim lo As ListObject
    On Error Resume Next
    Set lo = WS.ListObjects(tableName)
    On Error GoTo 0
    If lo Is Nothing Then
        Dim lastCol As Long, i As Long
        lastCol = UBound(Headers) - LBound(Headers) + 1
        Dim arr()
        ReDim arr(1 To 1, 1 To lastCol)
        For i = 1 To lastCol
            arr(1, i) = Headers(LBound(Headers) + i - 1)
        Next i
        WS.Cells(1, 1).Resize(1, lastCol).value = arr
        Set lo = WS.ListObjects.Add(xlSrcRange, WS.Range(WS.Cells(1, 1), WS.Cells(1, lastCol)), , xlYes)
        lo.name = tableName
    End If
    Set EnsureTable = lo
End Function

Public Function ColumnIndex(lo As ListObject, header As String) As Long
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).name, header, vbTextCompare) = 0 Then
            ColumnIndex = i
            Exit Function
        End If
    Next i
    Err.Raise 5, , "Header not found: " & header
End Function

Public Function Nz(v As Variant, Optional defaultValue As Variant) As Variant
    If IsMissing(defaultValue) Then defaultValue = ""
    If IsError(v) Then
        Nz = defaultValue
    ElseIf IsEmpty(v) Then
        Nz = defaultValue
    ElseIf VarType(v) = vbString And v = "" Then
        Nz = defaultValue
    Else
        Nz = v
    End If
End Function

Public Function FindRowBy(lo As ListObject, keyColumn As String, keyValue As String) As ListRow
    Dim iCol As Long: iCol = ColumnIndex(lo, keyColumn)
    Dim rw As ListRow
    For Each rw In lo.ListRows
        If CStr(Nz(rw.Range.Cells(1, iCol).value, "")) = CStr(keyValue) Then
            Set FindRowBy = rw
            Exit Function
        End If
    Next rw
End Function

Public Sub Upsert(lo As ListObject, keyColumn As String, data As Object)
    Dim rw As ListRow
    Set rw = FindRowBy(lo, keyColumn, CStr(data(keyColumn)))
    If rw Is Nothing Then
        Set rw = lo.ListRows.Add
    End If
    Dim col As ListColumn
    For Each col In lo.ListColumns
        If data.Exists(col.name) Then
            rw.Range(1, col.Index).value = data(col.name)
        End If
    Next col
End Sub
