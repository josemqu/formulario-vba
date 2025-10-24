VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsIncidente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Campos de Incidente (alineados a tbIncidente)
Public id_incidente As String
Public fecha_hora_ocurrencia As Variant
Public pais As String
Public provincia As String
Public localidad_zona As String
Public coordenadas_geograficas As String
Public lugar_especifico As String
Public uo_incidente As String
Public uo_accidentado As String
Public descripcion_esv As String
Public denuncia_policial As String
Public examen_alcoholemia As String
Public examen_sustancias As String
Public entrevistas_testigos As String
Public accion_inmediata As String
Public consecuencias_seguridad As String
Public fecha_hora_reporte As Variant
Public cantidad_personas As Variant
Public cantidad_vehiculos As Variant
Public clase_evento As String
Public tipo_colision As String
Public nivel_severidad As String
Public clasificacion_esv As String
Public creado_por As String
Public creado_en As Variant
Public actualizado_por As String
Public actualizado_en As Variant

Public Function ToDict() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("id_incidente") = id_incidente
    d("fecha_hora_ocurrencia") = fecha_hora_ocurrencia
    d("pais") = pais
    d("provincia") = provincia
    d("localidad_zona") = localidad_zona
    d("coordenadas_geograficas") = coordenadas_geograficas
    d("lugar_especifico") = lugar_especifico
    d("uo_incidente") = uo_incidente
    d("uo_accidentado") = uo_accidentado
    d("descripcion_esv") = descripcion_esv
    d("denuncia_policial") = denuncia_policial
    d("examen_alcoholemia") = examen_alcoholemia
    d("examen_sustancias") = examen_sustancias
    d("entrevistas_testigos") = entrevistas_testigos
    d("accion_inmediata") = accion_inmediata
    d("consecuencias_seguridad") = consecuencias_seguridad
    d("fecha_hora_reporte") = fecha_hora_reporte
    d("cantidad_personas") = cantidad_personas
    d("cantidad_vehiculos") = cantidad_vehiculos
    d("clase_evento") = clase_evento
    d("tipo_colision") = tipo_colision
    d("nivel_severidad") = nivel_severidad
    d("clasificacion_esv") = clasificacion_esv
    d("creado_por") = creado_por
    d("creado_en") = creado_en
    d("actualizado_por") = actualizado_por
    d("actualizado_en") = actualizado_en
    Set ToDict = d
End Function

Public Sub FromDict(d As Object)
    id_incidente = CStr(Nz(d("id_incidente"), ""))
    fecha_hora_ocurrencia = Nz(d("fecha_hora_ocurrencia"), Empty)
    pais = CStr(Nz(d("pais"), ""))
    provincia = CStr(Nz(d("provincia"), ""))
    localidad_zona = CStr(Nz(d("localidad_zona"), ""))
    coordenadas_geograficas = CStr(Nz(d("coordenadas_geograficas"), ""))
    lugar_especifico = CStr(Nz(d("lugar_especifico"), ""))
    uo_incidente = CStr(Nz(d("uo_incidente"), ""))
    uo_accidentado = CStr(Nz(d("uo_accidentado"), ""))
    descripcion_esv = CStr(Nz(d("descripcion_esv"), ""))
    denuncia_policial = CStr(Nz(d("denuncia_policial"), ""))
    examen_alcoholemia = CStr(Nz(d("examen_alcoholemia"), ""))
    examen_sustancias = CStr(Nz(d("examen_sustancias"), ""))
    entrevistas_testigos = CStr(Nz(d("entrevistas_testigos"), ""))
    accion_inmediata = CStr(Nz(d("accion_inmediata"), ""))
    consecuencias_seguridad = CStr(Nz(d("consecuencias_seguridad"), ""))
    fecha_hora_reporte = Nz(d("fecha_hora_reporte"), Empty)
    cantidad_personas = Nz(d("cantidad_personas"), Empty)
    cantidad_vehiculos = Nz(d("cantidad_vehiculos"), Empty)
    clase_evento = CStr(Nz(d("clase_evento"), ""))
    tipo_colision = CStr(Nz(d("tipo_colision"), ""))
    nivel_severidad = CStr(Nz(d("nivel_severidad"), ""))
    clasificacion_esv = CStr(Nz(d("clasificacion_esv"), ""))
    creado_por = CStr(Nz(d("creado_por"), ""))
    creado_en = Nz(d("creado_en"), Empty)
    actualizado_por = CStr(Nz(d("actualizado_por"), ""))
    actualizado_en = Nz(d("actualizado_en"), Empty)
End Sub

