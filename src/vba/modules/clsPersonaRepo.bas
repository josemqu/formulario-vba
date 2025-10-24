Attribute VB_Name = "clsPersonaRepo"
Option Explicit

Private Function Headers() As Variant
    Headers = Array( _
        "id_persona", "id_incidente", "nombre_persona", "apellido_persona", "edad_persona", _
        "tipo_persona", "rol_persona", "antiguedad_persona", "tarea_operativa", "turno_operativo", _
        "tipo_danio_persona", "dias_perdidos", "atencion_medica", "in_itinere", _
        "tipo_afectacion", "parte_afectada")
End Function

Private Function WS() As Worksheet
    Set WS = ThisWorkbook.Worksheets("Personas")
End Function

Public Function SaveEntity(ByVal ent As clsPersona) As String
    Dim repo As New clsRepository
    Dim d As Object: Set d = ent.ToDict
    If LenB(CStr(ent.id_persona)) = 0 Then
        d("id_persona") = NewUUID()
    End If
    SaveEntity = repo.Save(WS, "tbPersona", "id_persona", Headers, d)
End Function

Public Function FindById(id As String) As clsPersona
    Dim repo As New clsRepository
    Dim d As Object
    Set d = repo.Find(WS, "tbPersona", "id_persona", id)
    If Not d Is Nothing Then
        Dim e As New clsPersona
        e.FromDict d
        Set FindById = e
    End If
End Function
