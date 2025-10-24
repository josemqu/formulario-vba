Attribute VB_Name = "modCatalogos"
Option Explicit

Public Function RangeByName(name As String) As Range
    Set RangeByName = ThisWorkbook.Names(name).RefersToRange
End Function

Public Sub LoadCatalogToCombo(target As Object, catRange As Range)
    Dim arr
    arr = catRange.value
    target.Clear
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        If LenB(CStr(arr(i, 1))) > 0 Then target.AddItem CStr(arr(i, 1))
    Next i
End Sub
