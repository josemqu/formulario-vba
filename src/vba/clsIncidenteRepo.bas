Attribute VB_Name = "clsIncidenteRepo"
Option Explicit

Private Function Headers() As Variant
    Headers = Array( _
        "id_incidente", "fecha_hora_ocurrencia", "pais", "provincia", "localidad_zona", "coordenadas_geograficas", _
        "lugar_especifico", "uo_incidente", "uo_accidentado", "descripcion_esv", _
        "denuncia_policial", "examen_alcoholemia", "examen_sustancias", "entrevistas_testigos", _
        "accion_inmediata", "consecuencias_seguridad", "fecha_hora_reporte", _
        "cantidad_personas", "cantidad_vehiculos", "clase_evento", "tipo_colision", "nivel_severidad", "clasificacion_esv", _
        "creado_por", "creado_en", "actualizado_por", "actualizado_en")
End Function

Private Function WS() As Worksheet
    Set WS = ThisWorkbook.Worksheets("Incidentes")
End Function

Public Function SaveEntity(ByVal ent As clsIncidente) As String
    Dim repo As New clsRepository
    Dim id As String
    Dim d As Object: Set d = ent.ToDict
    If LenB(CStr(ent.id_incidente)) = 0 Then
        d("id_incidente") = NewUUID()
        d("creado_por") = UserNameOrDefault()
        d("creado_en") = NowIso()
    Else
        d("actualizado_por") = UserNameOrDefault()
        d("actualizado_en") = NowIso()
    End If
    id = repo.Save(WS, "tbIncidente", "id_incidente", Headers, d)
    SaveEntity = id
End Function

Public Function FindById(id As String) As clsIncidente
    Dim repo As New clsRepository
    Dim d As Object
    Set d = repo.Find(WS, "tbIncidente", "id_incidente", id)
    If Not d Is Nothing Then
        Dim e As New clsIncidente
        e.FromDict d
        Set FindById = e
    End If
End Function
