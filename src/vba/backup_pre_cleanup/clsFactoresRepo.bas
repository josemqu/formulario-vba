Attribute VB_Name = "clsFactoresRepo"
Option Explicit

Private Function Headers() As Variant
    Headers = Array( _
        "id_factor", "id_incidente", "tipo_superficie", "posee_banquina", "tipo_ruta", "densidad_trafico", _
        "condicion_ruta", "iluminacion_ruta", "senalizacion_ruta", "geometria_ruta", "condiciones_climaticas", "rango_temperaturas")
End Function

Private Function WS() As Worksheet
    Set WS = ThisWorkbook.Worksheets("Factores")
End Function

Public Function SaveEntity(ByVal ent As clsFactoresExternos) As String
    Dim repo As New clsRepository
    Dim d As Object: Set d = ent.ToDict
    If LenB(CStr(ent.id_factor)) = 0 Then
        d("id_factor") = NewUUID()
    End If
    SaveEntity = repo.Save(WS, "tbFactores", "id_factor", Headers, d)
End Function

Public Function FindById(id As String) As clsFactoresExternos
    Dim repo As New clsRepository
    Dim d As Object
    Set d = repo.Find(WS, "tbFactores", "id_factor", id)
    If Not d Is Nothing Then
        Dim e As New clsFactoresExternos
        e.FromDict d
        Set FindById = e
    End If
End Function
