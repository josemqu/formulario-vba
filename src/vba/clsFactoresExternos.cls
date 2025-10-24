VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFactoresExternos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id_factor As String
Public id_incidente As String
Public tipo_superficie As String
Public posee_banquina As String
Public tipo_ruta As String
Public densidad_trafico As String
Public condicion_ruta As String
Public iluminacion_ruta As String
Public senalizacion_ruta As String
Public geometria_ruta As String
Public condiciones_climaticas As String
Public rango_temperaturas As String

Public Function ToDict() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("id_factor") = id_factor
    d("id_incidente") = id_incidente
    d("tipo_superficie") = tipo_superficie
    d("posee_banquina") = posee_banquina
    d("tipo_ruta") = tipo_ruta
    d("densidad_trafico") = densidad_trafico
    d("condicion_ruta") = condicion_ruta
    d("iluminacion_ruta") = iluminacion_ruta
    d("senalizacion_ruta") = senalizacion_ruta
    d("geometria_ruta") = geometria_ruta
    d("condiciones_climaticas") = condiciones_climaticas
    d("rango_temperaturas") = rango_temperaturas
    Set ToDict = d
End Function

Public Sub FromDict(d As Object)
    id_factor = CStr(Nz(d("id_factor"), ""))
    id_incidente = CStr(Nz(d("id_incidente"), ""))
    tipo_superficie = CStr(Nz(d("tipo_superficie"), ""))
    posee_banquina = CStr(Nz(d("posee_banquina"), ""))
    tipo_ruta = CStr(Nz(d("tipo_ruta"), ""))
    densidad_trafico = CStr(Nz(d("densidad_trafico"), ""))
    condicion_ruta = CStr(Nz(d("condicion_ruta"), ""))
    iluminacion_ruta = CStr(Nz(d("iluminacion_ruta"), ""))
    senalizacion_ruta = CStr(Nz(d("senalizacion_ruta"), ""))
    geometria_ruta = CStr(Nz(d("geometria_ruta"), ""))
    condiciones_climaticas = CStr(Nz(d("condiciones_climaticas"), ""))
    rango_temperaturas = CStr(Nz(d("rango_temperaturas"), ""))
End Sub

