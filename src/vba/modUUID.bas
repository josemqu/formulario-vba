Attribute VB_Name = "modUUID"
Option Explicit

Public Function NewUUID() As String
    Dim guid As String
    guid = CreateObject("Scriptlet.TypeLib").guid
    NewUUID = Mid$(guid, 2, Len(guid) - 2)
End Function

Public Function NowIso() As String
    NowIso = Format$(Now, "yyyy-mm-dd\Thh:nn:ss")
End Function

Public Function UserNameOrDefault() As String
    On Error Resume Next
    UserNameOrDefault = Environ$("Username")
    If Len(UserNameOrDefault) = 0 Then UserNameOrDefault = "usuario"
End Function
