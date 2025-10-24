VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Repositorio generico sobre ListObjects basado en Dictionary (late bound)

Public Function Save(WS As Worksheet, tableName As String, keyColumn As String, Headers As Variant, data As Object) As String
    Dim lo As ListObject
    Set lo = EnsureTable(WS, tableName, Headers)
    If Not data.Exists(keyColumn) Or LenB(CStr(data(keyColumn))) = 0 Then
        data(keyColumn) = NewUUID()
    End If
    Upsert lo, keyColumn, data
    Save = CStr(data(keyColumn))
End Function

Public Function Find(WS As Worksheet, tableName As String, keyColumn As String, keyValue As String) As Object
    Dim lo As ListObject
    Set lo = WS.ListObjects(tableName)
    If lo Is Nothing Then Exit Function

    Dim rw As ListRow
    Set rw = FindRowBy(lo, keyColumn, keyValue)
    If rw Is Nothing Then Exit Function

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim col As ListColumn
    For Each col In lo.ListColumns
        dict(col.name) = rw.Range(1, col.Index).value
    Next col
    Set Find = dict
End Function

Public Sub Delete(WS As Worksheet, tableName As String, keyColumn As String, keyValue As String)
    Dim lo As ListObject
    Set lo = WS.ListObjects(tableName)
    If lo Is Nothing Then Exit Sub
    Dim rw As ListRow
    Set rw = FindRowBy(lo, keyColumn, keyValue)
    If Not rw Is Nothing Then rw.Delete
End Sub

