Attribute VB_Name = "modSetup"
Option Explicit

Private Function EnsureSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.name = sheetName
    End If
End Function

Public Sub SetupESVWorkbook()
    Dim wsI As Worksheet, wsP As Worksheet, wsV As Worksheet, wsF As Worksheet, wsC As Worksheet
    Set wsI = EnsureSheet("Incidentes")
    Set wsP = EnsureSheet("Personas")
    Set wsV = EnsureSheet("Vehiculos")
    Set wsF = EnsureSheet("Factores")
    Set wsC = EnsureSheet("Catalogos")

    Dim hInc, hPer, hVeh, hFac

    hInc = Array( _
        "id_incidente", "fecha_hora_ocurrencia", "pais", "provincia", "localidad_zona", "coordenadas_geograficas", _
        "lugar_especifico", "uo_incidente", "uo_accidentado", "descripcion_esv", _
        "denuncia_policial", "examen_alcoholemia", "examen_sustancias", "entrevistas_testigos", _
        "accion_inmediata", "consecuencias_seguridad", "fecha_hora_reporte", _
        "cantidad_personas", "cantidad_vehiculos", "clase_evento", "tipo_colision", "nivel_severidad", "clasificacion_esv", _
        "creado_por", "creado_en", "actualizado_por", "actualizado_en")

    hPer = Array( _
        "id_persona", "id_incidente", "nombre_persona", "apellido_persona", "edad_persona", _
        "tipo_persona", "rol_persona", "antiguedad_persona", "tarea_operativa", "turno_operativo", _
        "tipo_danio_persona", "dias_perdidos", "atencion_medica", "in_itinere", _
        "tipo_afectacion", "parte_afectada")

    hVeh = Array( _
        "id_vehiculo", "id_incidente", "tipo_vehiculo", "duenio_vehiculo", "uso_vehiculo", _
        "posee_patente", "numero_patente", "anio_fabricacion_vehiculo", "tarea_vehiculo", "tipo_danio_vehiculo", _
        "cinturon_seguridad", "cabina_cuchetas", "airbags", "gestion_flotas", "token_conductor", _
        "marca_dispositivo", "deteccion_fatiga", "camara_trasera", "limitador_velocidad", "camara_delantera", _
        "camara_punto_ciego", "camara_360", "espejo_punto_ciego", "alarma_marcha_atras", "sistema_frenos", _
        "monitoreo_neumaticos", "proteccion_lateral", "proteccion_trasera", "acondicionador_cabina", "calefaccion_cabina", _
        "manos_libres_cabina", "kit_alcoholemia", "kit_emergencia", "epps_vehiculo", _
        "observaciones_vehiculo", "creado_por", "creado_en", "actualizado_por", "actualizado_en")

    hFac = Array( _
        "id_factor", "id_incidente", "tipo_superficie", "posee_banquina", "tipo_ruta", "densidad_trafico", _
        "condicion_ruta", "iluminacion_ruta", "senalizacion_ruta", "geometria_ruta", "condiciones_climaticas", "rango_temperaturas")

    Dim loI As ListObject, loP As ListObject, loV As ListObject, loF As ListObject
    Set loI = EnsureTable(wsI, "tbIncidente", hInc)
    Set loP = EnsureTable(wsP, "tbPersona", hPer)
    Set loV = EnsureTable(wsV, "tbVehiculo", hVeh)
    Set loF = EnsureTable(wsF, "tbFactores", hFac)

    SetupCatalogos wsC

    MsgBox "Estructura creada/actualizada.", vbInformation
End Sub

Private Sub SetupCatalogos(WS As Worksheet)
    Dim cats As Variant

    ' Catálogo simple SI/NO/NA
    WS.Range("A1").value = "cat_si_no_na"
    WS.Range("A2:A4").value = Application.WorksheetFunction.Transpose(Array("SI", "NO", "NA"))
    AddOrUpdateName "cat_si_no_na", WS.Range("A2:A4")
    AddOrUpdateName "CAT_SI_NO_NA", WS.Range("A2:A4")

    ' Tipo de vehículo (placeholder, se puede ampliar)
    WS.Range("C1").value = "cat_tipo_vehiculo"
    WS.Range("C2:C9").value = Application.WorksheetFunction.Transpose(Array( _
        "Bicicleta", "Moto", "Ciclomotor", "Autom" & ChrW(243) & "vil", "Pickup", "Cami" & ChrW(243) & "n chasis", "Cami" & ChrW(243) & "n con Cisterna", ChrW(211) & "mnibus"))
    AddOrUpdateName "cat_tipo_vehiculo", WS.Range("C2:C9")
    AddOrUpdateName "CAT_TIPO_VEHICULO", WS.Range("C2:C9")

    ' Dueño de vehículo
    WS.Range("E1").value = "cat_duenio_vehiculo"
    WS.Range("E2:E4").value = Application.WorksheetFunction.Transpose(Array("Propio", "Contratista", "Tercero"))
    AddOrUpdateName "cat_duenio_vehiculo", WS.Range("E2:E4")
    AddOrUpdateName "CAT_DUENIO_VEHICULO", WS.Range("E2:E4")

    ' Uso del vehículo
    WS.Range("G1").value = "cat_uso_vehiculo"
    WS.Range("G2:G6").value = Application.WorksheetFunction.Transpose(Array("Comercial", "Particular", "Otro", "No se sabe", "NA"))
    AddOrUpdateName "cat_uso_vehiculo", WS.Range("G2:G6")
    AddOrUpdateName "CAT_USO_VEHICULO", WS.Range("G2:G6")
End Sub

Private Sub AddOrUpdateName(nameText As String, refersToRng As Range)
    On Error Resume Next
    Dim nm As name
    Set nm = ThisWorkbook.Names(nameText)
    On Error GoTo 0
    If nm Is Nothing Then
        ThisWorkbook.Names.Add name:=nameText, RefersTo:=refersToRng
    Else
        nm.RefersTo = refersToRng
    End If
End Sub
