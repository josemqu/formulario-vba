VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id_persona As String
Public id_incidente As String
Public nombre_persona As String
Public apellido_persona As String
Public edad_persona As Variant
Public tipo_persona As String
Public rol_persona As String
Public antiguedad_persona As String
Public tarea_operativa As String
Public turno_operativo As String
Public tipo_danio_persona As String
Public dias_perdidos As Variant
Public atencion_medica As String
Public in_itinere As String
Public tipo_afectacion As String
Public parte_afectada As String

Public Function ToDict() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("id_persona") = id_persona
    d("id_incidente") = id_incidente
    d("nombre_persona") = nombre_persona
    d("apellido_persona") = apellido_persona
    d("edad_persona") = edad_persona
    d("tipo_persona") = tipo_persona
    d("rol_persona") = rol_persona
    d("antiguedad_persona") = antiguedad_persona
    d("tarea_operativa") = tarea_operativa
    d("turno_operativo") = turno_operativo
    d("tipo_danio_persona") = tipo_danio_persona
    d("dias_perdidos") = dias_perdidos
    d("atencion_medica") = atencion_medica
    d("in_itinere") = in_itinere
    d("tipo_afectacion") = tipo_afectacion
    d("parte_afectada") = parte_afectada
    Set ToDict = d
End Function

Public Sub FromDict(d As Object)
    id_persona = CStr(Nz(d("id_persona"), ""))
    id_incidente = CStr(Nz(d("id_incidente"), ""))
    nombre_persona = CStr(Nz(d("nombre_persona"), ""))
    apellido_persona = CStr(Nz(d("apellido_persona"), ""))
    edad_persona = Nz(d("edad_persona"), Empty)
    tipo_persona = CStr(Nz(d("tipo_persona"), ""))
    rol_persona = CStr(Nz(d("rol_persona"), ""))
    antiguedad_persona = CStr(Nz(d("antiguedad_persona"), ""))
    tarea_operativa = CStr(Nz(d("tarea_operativa"), ""))
    turno_operativo = CStr(Nz(d("turno_operativo"), ""))
    tipo_danio_persona = CStr(Nz(d("tipo_danio_persona"), ""))
    dias_perdidos = Nz(d("dias_perdidos"), Empty)
    atencion_medica = CStr(Nz(d("atencion_medica"), ""))
    in_itinere = CStr(Nz(d("in_itinere"), ""))
    tipo_afectacion = CStr(Nz(d("tipo_afectacion"), ""))
    parte_afectada = CStr(Nz(d("parte_afectada"), ""))
End Sub

