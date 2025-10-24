Attribute VB_Name = "modFormSupport"
Option Explicit

' Carga de catálogos para el formulario frmRegistroESV
' Nota: se asume que existen Named Ranges en una hoja de "Catálogos".
' Las rutinas usan frm As Object y On Error Resume Next para no romper si un control aún no existe.

Public Sub LoadAllCatalogs(frm As Object)
    On Error Resume Next
    LoadIncidentCatalogs frm
    LoadPersonaCatalogs frm
    LoadVehiculoCatalogs frm
    LoadFactoresCatalogs frm
    On Error GoTo 0
End Sub

Public Sub LoadIncidentCatalogs(frm As Object)
    On Error Resume Next
    LoadCatalogToCombo frm.cmbPais, RangeByName("CAT_PAIS")
    LoadCatalogToCombo frm.cmbProvincia, RangeByName("CAT_PROVINCIA")
    LoadCatalogToCombo frm.cmbLocalidad, RangeByName("CAT_LOCALIDAD_ZONA")

    LoadCatalogToCombo frm.cmbUOIncidente, RangeByName("CAT_UO_INCIDENTE")
    LoadCatalogToCombo frm.cmbUOAccidentado, RangeByName("CAT_UO_ACCIDENTADO")

    LoadCatalogToCombo frm.cmbDenuncia, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbAlcohol, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbSustancias, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbEntrevistas, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbConsecuencias, RangeByName("CAT_SI_NO_NA")

    LoadCatalogToCombo frm.cmbClaseEvento, RangeByName("CAT_CLASE_EVENTO")
    LoadCatalogToCombo frm.cmbTipoColision, RangeByName("CAT_TIPO_COLISION")
    LoadCatalogToCombo frm.cmbNivelSeveridad, RangeByName("CAT_NIVEL_SEVERIDAD")
    LoadCatalogToCombo frm.cmbClasificacion, RangeByName("CAT_CLASIFICACION_ESV")
    On Error GoTo 0
End Sub

Public Sub LoadPersonaCatalogs(frm As Object)
    On Error Resume Next
    LoadCatalogToCombo frm.cmbPTipo, RangeByName("CAT_TIPO_PERSONA")
    LoadCatalogToCombo frm.cmbPRol, RangeByName("CAT_ROL_PERSONA")
    LoadCatalogToCombo frm.cmbPAntiguedad, RangeByName("CAT_ANTIGUEDAD")
    LoadCatalogToCombo frm.cmbPTarea, RangeByName("CAT_TAREA_OPERATIVA")
    LoadCatalogToCombo frm.cmbPTurno, RangeByName("CAT_TURNO")
    LoadCatalogToCombo frm.cmbPTipoDanio, RangeByName("CAT_TIPO_DANIO")
    LoadCatalogToCombo frm.cmbPAtencion, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbPInItinere, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbPAfectacion, RangeByName("CAT_TIPO_AFECTACION")
    LoadCatalogToCombo frm.cmbPParte, RangeByName("CAT_PARTE_AFECTADA")
    On Error GoTo 0
End Sub

Public Sub LoadVehiculoCatalogs(frm As Object)
    On Error Resume Next
    LoadCatalogToCombo frm.cmbVTipo, RangeByName("CAT_TIPO_VEHICULO")
    LoadCatalogToCombo frm.cmbVDueno, RangeByName("CAT_DUENIO_VEHICULO")
    LoadCatalogToCombo frm.cmbVUso, RangeByName("CAT_USO_VEHICULO")

    LoadCatalogToCombo frm.cmbVPoseePatente, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbVTarea, RangeByName("CAT_TAREA_VEHICULO")
    LoadCatalogToCombo frm.cmbVTipoDanio, RangeByName("CAT_TIPO_DANIO_VEHICULO")

    LoadCatalogToCombo frm.cmbVCinturon, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbVCuchetas, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbVAirbags, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbVFlotas, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbVToken, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbVMarcaDisp, RangeByName("CAT_MARCA_DISPOSITIVO")
    LoadCatalogToCombo frm.cmbVFatiga, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbVCamTras, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbVLimitador, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbVCamDel, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbVPtoCiego, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbV360, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbVEspejoPC, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbVAlarmaMA, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbVSisFrenos, RangeByName("CAT_SISTEMA_FRENOS")
    LoadCatalogToCombo frm.cmbVTPMS, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbVProtLateral, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbVProtTrasera, RangeByName("CAT_SI_NO_NA")

    LoadCatalogToCombo frm.cmbVAA, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbVCalef, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbVHandsfree, RangeByName("CAT_HANDSFREE_ESTADO")
    LoadCatalogToCombo frm.cmbVKitAlcoh, RangeByName("CAT_KIT_ALCOHOLEMIA")
    LoadCatalogToCombo frm.cmbVKitEmerg, RangeByName("CAT_KIT_EMERGENCIA")
    LoadCatalogToCombo frm.cmbVEPPs, RangeByName("CAT_EPPS")
    On Error GoTo 0
End Sub

Public Sub LoadFactoresCatalogs(frm As Object)
    On Error Resume Next
    LoadCatalogToCombo frm.cmbFTipoSuperficie, RangeByName("CAT_TIPO_SUPERFICIE")
    LoadCatalogToCombo frm.cmbFBanquina, RangeByName("CAT_SI_NO_NA")
    LoadCatalogToCombo frm.cmbFTipoRuta, RangeByName("CAT_TIPO_RUTA")
    LoadCatalogToCombo frm.cmbFDensidad, RangeByName("CAT_DENSIDAD_TRAFICO")
    LoadCatalogToCombo frm.cmbFCondicion, RangeByName("CAT_CONDICION_RUTA")
    LoadCatalogToCombo frm.cmbFIluminacion, RangeByName("CAT_ILUMINACION")
    LoadCatalogToCombo frm.cmbFSenalizacion, RangeByName("CAT_SENALIZACION")
    LoadCatalogToCombo frm.cmbFGeometria, RangeByName("CAT_GEOMETRIA")
    LoadCatalogToCombo frm.cmbFClima, RangeByName("CAT_CLIMA")
    On Error GoTo 0
End Sub
