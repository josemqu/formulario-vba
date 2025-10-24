VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVehiculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id_vehiculo As String
Public id_incidente As String
Public tipo_vehiculo As String
Public duenio_vehiculo As String
Public uso_vehiculo As String
Public posee_patente As String
Public numero_patente As String
Public anio_fabricacion_vehiculo As String
Public tarea_vehiculo As String
Public tipo_danio_vehiculo As String
Public cinturon_seguridad As String
Public cabina_cuchetas As String
Public airbags As String
Public gestion_flotas As String
Public token_conductor As String
Public marca_dispositivo As String
Public deteccion_fatiga As String
Public camara_trasera As String
Public limitador_velocidad As String
Public camara_delantera As String
Public camara_punto_ciego As String
Public camara_360 As String
Public espejo_punto_ciego As String
Public alarma_marcha_atras As String
Public sistema_frenos As String
Public monitoreo_neumaticos As String
Public proteccion_lateral As String
Public proteccion_trasera As String
Public acondicionador_cabina As String
Public calefaccion_cabina As String
Public manos_libres_cabina As String
Public kit_alcoholemia As String
Public kit_emergencia As String
Public epps_vehiculo As String
Public observaciones_vehiculo As String
Public creado_por As String
Public creado_en As Variant
Public actualizado_por As String
Public actualizado_en As Variant

Public Function ToDict() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("id_vehiculo") = id_vehiculo
    d("id_incidente") = id_incidente
    d("tipo_vehiculo") = tipo_vehiculo
    d("duenio_vehiculo") = duenio_vehiculo
    d("uso_vehiculo") = uso_vehiculo
    d("posee_patente") = posee_patente
    d("numero_patente") = numero_patente
    d("anio_fabricacion_vehiculo") = anio_fabricacion_vehiculo
    d("tarea_vehiculo") = tarea_vehiculo
    d("tipo_danio_vehiculo") = tipo_danio_vehiculo
    d("cinturon_seguridad") = cinturon_seguridad
    d("cabina_cuchetas") = cabina_cuchetas
    d("airbags") = airbags
    d("gestion_flotas") = gestion_flotas
    d("token_conductor") = token_conductor
    d("marca_dispositivo") = marca_dispositivo
    d("deteccion_fatiga") = deteccion_fatiga
    d("camara_trasera") = camara_trasera
    d("limitador_velocidad") = limitador_velocidad
    d("camara_delantera") = camara_delantera
    d("camara_punto_ciego") = camara_punto_ciego
    d("camara_360") = camara_360
    d("espejo_punto_ciego") = espejo_punto_ciego
    d("alarma_marcha_atras") = alarma_marcha_atras
    d("sistema_frenos") = sistema_frenos
    d("monitoreo_neumaticos") = monitoreo_neumaticos
    d("proteccion_lateral") = proteccion_lateral
    d("proteccion_trasera") = proteccion_trasera
    d("acondicionador_cabina") = acondicionador_cabina
    d("calefaccion_cabina") = calefaccion_cabina
    d("manos_libres_cabina") = manos_libres_cabina
    d("kit_alcoholemia") = kit_alcoholemia
    d("kit_emergencia") = kit_emergencia
    d("epps_vehiculo") = epps_vehiculo
    d("observaciones_vehiculo") = observaciones_vehiculo
    d("creado_por") = creado_por
    d("creado_en") = creado_en
    d("actualizado_por") = actualizado_por
    d("actualizado_en") = actualizado_en
    Set ToDict = d
End Function

Public Sub FromDict(d As Object)
    id_vehiculo = CStr(Nz(d("id_vehiculo"), ""))
    id_incidente = CStr(Nz(d("id_incidente"), ""))
    tipo_vehiculo = CStr(Nz(d("tipo_vehiculo"), ""))
    duenio_vehiculo = CStr(Nz(d("duenio_vehiculo"), ""))
    uso_vehiculo = CStr(Nz(d("uso_vehiculo"), ""))
    posee_patente = CStr(Nz(d("posee_patente"), ""))
    numero_patente = CStr(Nz(d("numero_patente"), ""))
    anio_fabricacion_vehiculo = CStr(Nz(d("anio_fabricacion_vehiculo"), ""))
    tarea_vehiculo = CStr(Nz(d("tarea_vehiculo"), ""))
    tipo_danio_vehiculo = CStr(Nz(d("tipo_danio_vehiculo"), ""))
    cinturon_seguridad = CStr(Nz(d("cinturon_seguridad"), ""))
    cabina_cuchetas = CStr(Nz(d("cabina_cuchetas"), ""))
    airbags = CStr(Nz(d("airbags"), ""))
    gestion_flotas = CStr(Nz(d("gestion_flotas"), ""))
    token_conductor = CStr(Nz(d("token_conductor"), ""))
    marca_dispositivo = CStr(Nz(d("marca_dispositivo"), ""))
    deteccion_fatiga = CStr(Nz(d("deteccion_fatiga"), ""))
    camara_trasera = CStr(Nz(d("camara_trasera"), ""))
    limitador_velocidad = CStr(Nz(d("limitador_velocidad"), ""))
    camara_delantera = CStr(Nz(d("camara_delantera"), ""))
    camara_punto_ciego = CStr(Nz(d("camara_punto_ciego"), ""))
    camara_360 = CStr(Nz(d("camara_360"), ""))
    espejo_punto_ciego = CStr(Nz(d("espejo_punto_ciego"), ""))
    alarma_marcha_atras = CStr(Nz(d("alarma_marcha_atras"), ""))
    sistema_frenos = CStr(Nz(d("sistema_frenos"), ""))
    monitoreo_neumaticos = CStr(Nz(d("monitoreo_neumaticos"), ""))
    proteccion_lateral = CStr(Nz(d("proteccion_lateral"), ""))
    proteccion_trasera = CStr(Nz(d("proteccion_trasera"), ""))
    acondicionador_cabina = CStr(Nz(d("acondicionador_cabina"), ""))
    calefaccion_cabina = CStr(Nz(d("calefaccion_cabina"), ""))
    manos_libres_cabina = CStr(Nz(d("manos_libres_cabina"), ""))
    kit_alcoholemia = CStr(Nz(d("kit_alcoholemia"), ""))
    kit_emergencia = CStr(Nz(d("kit_emergencia"), ""))
    epps_vehiculo = CStr(Nz(d("epps_vehiculo"), ""))
    observaciones_vehiculo = CStr(Nz(d("observaciones_vehiculo"), ""))
    creado_por = CStr(Nz(d("creado_por"), ""))
    creado_en = Nz(d("creado_en"), Empty)
    actualizado_por = CStr(Nz(d("actualizado_por"), ""))
    actualizado_en = Nz(d("actualizado_en"), Empty)
End Sub

