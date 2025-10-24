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
    ' Comunes
    EnsureCatalog WS, "A", "cat_si_no_na", Array("SI", "NO", "NA")

    ' Vehículos (defaults de ejemplo)
    EnsureCatalog WS, "C", "cat_tipo_vehiculo", Array( _
        "Bicicleta", "Moto", "Ciclomotor", "Autom" & ChrW(243) & "vil", "Pickup", _
        "Cami" & ChrW(243) & "n chasis", "Cami" & ChrW(243) & "n con Cisterna", ChrW(211) & "mnibus")
    EnsureCatalog WS, "E", "cat_duenio_vehiculo", Array("Propio", "Contratista", "Tercero")
    EnsureCatalog WS, "G", "cat_uso_vehiculo", Array("Comercial", "Particular", "Otro", "No se sabe", "NA")

    ' Incidente (placeholders para carga manual)
    EnsureCatalog WS, "I", "cat_pais"
    EnsureCatalog WS, "K", "cat_provincia"
    EnsureCatalog WS, "M", "cat_localidad_zona"
    EnsureCatalog WS, "O", "cat_uo_incidente"
    EnsureCatalog WS, "Q", "cat_uo_accidentado"
    EnsureCatalog WS, "S", "cat_clase_evento"
    EnsureCatalog WS, "U", "cat_tipo_colision"
    EnsureCatalog WS, "W", "cat_nivel_severidad"
    EnsureCatalog WS, "Y", "cat_clasificacion_esv"

    ' Personas (placeholders)
    EnsureCatalog WS, "AA", "cat_tipo_persona"
    EnsureCatalog WS, "AC", "cat_rol_persona"
    EnsureCatalog WS, "AE", "cat_antiguedad_persona"
    EnsureCatalog WS, "AG", "cat_tarea_operativa"
    EnsureCatalog WS, "AI", "cat_turno_operativo"
    EnsureCatalog WS, "AK", "cat_tipo_danio_persona"
    EnsureCatalog WS, "AM", "cat_tipo_afectacion"
    EnsureCatalog WS, "AO", "cat_parte_afectada"

    ' Vehículo adicionales (placeholders)
    EnsureCatalog WS, "AQ", "cat_tarea_vehiculo"
    EnsureCatalog WS, "AS", "cat_tipo_danio_vehiculo"

    ' Factores (placeholders)
    EnsureCatalog WS, "AU", "cat_tipo_superficie"
    EnsureCatalog WS, "AW", "cat_tipo_ruta"
    EnsureCatalog WS, "AY", "cat_densidad_trafico"
    EnsureCatalog WS, "BA", "cat_condicion_ruta"
    EnsureCatalog WS, "BC", "cat_iluminacion_ruta"
    EnsureCatalog WS, "BE", "cat_senalizacion_ruta"
    EnsureCatalog WS, "BG", "cat_geometria_ruta"
    EnsureCatalog WS, "BI", "cat_condiciones_climaticas"
    EnsureCatalog WS, "BK", "cat_rango_temperaturas"
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

Private Sub EnsureCatalog(WS As Worksheet, colLetter As String, header As String, Optional defaults As Variant)
    Dim hdrCell As Range, firstData As Range, lastCell As Range, dataRng As Range
    Set hdrCell = WS.Range(colLetter & "1")
    hdrCell.value = header

    Set firstData = WS.Range(colLetter & "2")
    ' Buscar último valor en la columna
    Set lastCell = WS.Cells(WS.Rows.Count, hdrCell.Column).End(xlUp)
    If lastCell.Row < 2 Then
        ' Vacío: sembrar defaults solo si se enviaron
        If Not IsMissing(defaults) Then
            firstData.Resize(UBound(defaults) - LBound(defaults) + 1, 1).value = _
                Application.WorksheetFunction.Transpose(defaults)
        End If
        Set dataRng = WS.Range(firstData, WS.Cells(WS.Rows.Count, hdrCell.Column).End(xlUp))
    Else
        ' Ya hay datos: respetar existentes
        Set dataRng = WS.Range(firstData, lastCell)
    End If

    ' Crear/actualizar nombres en minúsculas y mayúsculas
    AddOrUpdateName header, dataRng
    AddOrUpdateName UCase(header), dataRng
End Sub
