Attribute VB_Name = "modValidation"
Option Explicit

Public Function ValidateEnum(value As String, allowed As Variant, fieldName As String) As String
    Dim i As Long
    For i = LBound(allowed) To UBound(allowed)
        If StrComp(value, allowed(i), vbTextCompare) = 0 Then Exit Function
    Next i
    ValidateEnum = fieldName & " inválido: " & value
End Function

Public Function HasErrors(errors As Collection) As Boolean
    HasErrors = (errors.Count > 0)
End Function
