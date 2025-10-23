# 🧩 Esquema Técnico de Base de Datos – Sistema de Registro y Análisis de Incidentes de Tránsito

## 📘 Descripción General

El presente esquema de base de datos define la estructura técnica para el registro, análisis y procesamiento de incidentes de tránsito.  
Está diseñado para permitir el almacenamiento estructurado de información sobre **personas involucradas**, **vehículos**, **factores externos** y el propio **incidente**, con el objetivo de realizar análisis estadístico y generar reportes dinámicos.

---

## 🏗️ Estructura General

El modelo relacional está compuesto por las siguientes tablas principales:

- **Incidente** → Contiene la información general del suceso.
- **Persona** → Registra los datos de las personas afectadas (empleados YPF, contratistas o terceros).
- **Vehiculo** → Detalla las características y condiciones de los vehículos involucrados.
- **FactoresExternos** → Describe las condiciones ambientales y de la vía al momento del incidente.

Cada **Incidente** puede tener múltiples **Personas** y **Vehículos** asociados.  
Los **Factores Externos** están relacionados de manera **uno a uno** con el incidente.

---

## 🔗 Relaciones Entre Tablas

| Tabla Origen | Tabla Destino | Tipo de Relación | Descripción |
|---------------|----------------|------------------|--------------|
| `Incidente` | `Persona` | 1 a N | Un incidente puede involucrar una o más personas. |
| `Incidente` | `Vehiculo` | 1 a N | Un incidente puede involucrar uno o más vehículos. |
| `Incidente` | `FactoresExternos` | 1 a 1 | Cada incidente tiene un único conjunto de factores externos asociados. |

---

## 🧱 Tablas y Campos

### 🟦 Tabla: `Incidente`
| Campo | Tipo | Descripción |
|--------|------|-------------|
| `id_incidente` | Integer (PK, autoincremental) | Identificador único del incidente |
| `fecha_hora_ocurrencia` | Datetime | Fecha y hora del hecho |
| `pais` | Texto | País donde ocurrió el incidente |
| `provincia` | Texto | Provincia (según catálogo YPF GestiónCAS) |
| `localidad_zona` | Texto | Localidad o zona (según catálogo YPF GestiónCAS) |
| `coordenadas_geograficas` | Texto | Coordenadas GPS (lat,lng) |
| `lugar_especifico` | Texto | Descripción del lugar (máx. 100 caracteres) |
| `uo_incidente` | Texto | Unidad Operativa donde ocurrió el hecho |
| `uo_accidentado` | Texto | Unidad Operativa a la que pertenece el accidentado |
| `descripcion_esv` | Texto largo | Descripción narrativa del evento |
| `denuncia_policial` | Enum(SI, NO, NA) | Indica si hubo denuncia policial |
| `examen_alcoholemia` | Enum(SI, NO, NA) | Resultado o existencia del examen |
| `examen_sustancias` | Enum(SI, NO, NA) | Resultado o existencia del examen |
| `entrevistas_testigos` | Enum(SI, NO, NA) | Indica si hubo entrevistas con testigos |
| `accion_inmediata` | Texto | Breve descripción de las acciones inmediatas |
| `consecuencias_seguridad` | Enum(SI, NO) | Indica si hubo consecuencias reales o potenciales |
| `fecha_hora_reporte` | Datetime | Fecha y hora del reporte |
| `cantidad_personas` | Integer | Número de personas afectadas |
| `cantidad_vehiculos` | Integer | Número de vehículos involucrados |
| `clase_evento` | Texto | Tipo o clase del evento (catálogo YPF) |
| `tipo_colision` | Texto | Tipo de colisión (frontal, lateral, etc.) |
| `nivel_severidad` | Enum(S0-S5) | Nivel de severidad según clasificación YPF |
| `clasificacion_esv` | Enum(Severo, Potencialmente Severo, Menor) | Clasificación del evento |

---

### 🟩 Tabla: `Persona`
| Campo | Tipo | Descripción |
|--------|------|-------------|
| `id_persona` | Integer (PK) | Identificador único de persona |
| `id_incidente` | Integer (FK → Incidente.id_incidente) | Relación con incidente |
| `nombre_persona` | Texto | Nombre |
| `apellido_persona` | Texto | Apellido |
| `edad_persona` | Integer | Edad (0-100) |
| `tipo_persona` | Enum(YPF, Contratista, Tercero) | Tipo de involucrado |
| `rol_persona` | Enum(Conductor, Acompañante, Operador, Otro) | Rol dentro del incidente |
| `antiguedad_persona` | Texto | Rango de antigüedad o "No aplica" |
| `tarea_operativa` | Enum(Rutinaria, Especial, Emergencia, Supervisión, Apoyo Logístico, NA) | Tipo de tarea |
| `turno_operativo` | Enum(Diurno, Nocturno, Extendido-Mixto, NA) | Turno operativo |
| `tipo_danio_persona` | Enum(Fatalidad, Accidente, Primeros Auxilios, Ninguno) | Tipo de daño |
| `dias_perdidos` | Integer | Días de ausencia |
| `atencion_medica` | Enum(SI, NO, NA) | Si recibió atención médica |
| `in_itinere` | Enum(SI, NO, NA) | Indica si ocurrió en itinere |
| `tipo_afectacion` | Texto | Tipo de afectación física |
| `parte_afectada` | Texto | Parte del cuerpo afectada |

---

### 🟨 Tabla: `Vehiculo`

| Campo | Tipo | Descripción |
|-------|------|-------------|
| `id_vehiculo` | Integer (PK) | Identificador único del vehículo (PK). |
| `id_incidente` | Integer (FK → Incidente.id_incidente) | FK a `Incidente.id_incidente`. Vincula el vehículo con su incidente. |
| `tipo_vehiculo` | Texto | Tipo de vehículo. Opciones: Bicicleta, Moto, Ciclomotor, Automóvil, Auto utilitario, Minibús, Ómnibus, Pickup, Camión chasis, Camión con Cisterna, Camión Pluma, Camión Volcador, Motoniveladora, Retroexcavadora, Pala cargadora, Topadora, Grúa, Trailer, Side-by-Side/UTV, Vehículo adaptado (discapacidad). |
| `duenio_vehiculo` | Texto | Dueño del vehículo. Opciones: Propio, Contratista, Tercero. |
| `uso_vehiculo` | Texto | Uso del vehículo. Opciones: Comercial, Particular, Otro, No se sabe. |
| `posee_patente` | Enum(SI, NO, NA) | Indicador de patente. Opciones: SI / NO. |
| `numero_patente` | Texto | Patente alfanumérica; valores especiales: "desconocida" / "NA". |
| `anio_fabricacion_vehiculo` | Texto | Año de fabricación o texto libre (ej. "desconocido"). |
| `tarea_vehiculo` | Texto | Tarea que realizaba. Ej.: transporte de personas, transporte de cargas generales, transporte de sustancias peligrosas/combustible, maniobrando/estacionando, trayecto de regreso vacío, transporte de mercaderías, tareas generales, viaje familiar/comercial, se desconoce, otros. |
| `tipo_danio_vehiculo` | Texto | Clasificación del daño. Opciones: Destrucción total; Daños en carrocería que afectan continuidad de viaje (remolque/grúa); Daños en carrocería que NO afectan continuidad de viaje; Daños mecánicos/eléctricos que afectan continuidad; Daños mecánicos/eléctricos que no afectan continuidad; Daños leves; Sin daños. |
| `cinturon_seguridad` | Enum(SI, NO, NA) | Uso de cinturón. Opciones: SI / NO. |
| `cabina_cuchetas` | Enum(SI, NO, NA) | Cabina con cuchetas. Opciones: SI / NO. |
| `airbags` | Enum(SI, NO, NA) | Airbags presentes/desplegados. Opciones: SI / NO. |
| `gestion_flotas` | Enum(SI, NO, NA) | Gestión por sistema de flotas. Opciones: SI / NO. |
| `token_conductor` | Enum(SI, NO, NA) | Token del conductor. Opciones: SI / NO / NA. |
| `marca_dispositivo` | Texto | Marca de dispositivo telemático. Opciones: Microtrack, ITURAN, IMSEG, Otros. |
| `deteccion_fatiga` | Enum(SI, NO, NA) | Sistema detección fatiga/distracción. Opciones: SI / NO / NA. |
| `camara_trasera` | Enum(SI, NO, NA) | Cámara trasera. Opciones: SI / NO / NA. |
| `limitador_velocidad` | Enum(SI, NO, NA) | Limitador de velocidad. Opciones: SI / NO / NA. |
| `camara_delantera` | Enum(SI, NO, NA) | Cámara delantera. Opciones: SI / NO / NA. |
| `camara_punto_ciego` | Enum(SI, NO, NA) | Cámaras en puntos ciegos. Opciones: SI / NO / NA. |
| `camara_360` | Enum(SI, NO, NA) | Cámara 360°. Opciones: SI / NO / NA. |
| `espejo_punto_ciego` | Enum(SI, NO, NA) | Espejos para puntos ciegos. Opciones: SI / NO / NA. |
| `alarma_marcha_atras` | Enum(SI, NO, NA) | Alarma de marcha atrás. Opciones: SI / NO / NA. |
| `sistema_frenos` | Texto | Tipo de sistema de frenos. Opciones: Sistema de frenos antibloqueo, Frenado electrónico, Estabilidad electrónica, Soporte anti-vuelco, Frenado de emergencia avanzado, Alerta de frenado de emergencia. |
| `monitoreo_neumaticos` | Enum(SI, NO, NA) | Monitoreo de presión de neumáticos. Opciones: SI / NO / NA. |
| `proteccion_lateral` | Enum(SI, NO, NA) | Protección lateral para bicicletas/motos. Opciones: SI / NO / NA. |
| `proteccion_trasera` | Enum(SI, NO, NA) | Protección trasera antiempotramiento. Opciones: SI / NO / NA. |
| `acondicionador_cabina` | Enum(SI, NO, NA) | Aire acondicionado en cabina. Opciones: SI / NO / NA. |
| `calefaccion_cabina` | Enum(SI, NO, NA) | Calefacción en cabina. Opciones: SI / NO / NA. |
| `manos_libres_cabina` | Enum(NO POSEE / FUNCIONA / NO FUNCIONA / DESHABILITADO) | Sistema Bluetooth/manos libres. Opciones: NO POSEE / FUNCIONA / NO FUNCIONA / DESHABILITADO. |
| `kit_alcoholemia` | Enum(En cabina / Alcoholímetro antiarranque / No posee) | Control alcoholemia en cabina. Opciones: En cabina / Alcoholímetro antiarranque / No posee. |
| `kit_emergencia` | Texto | Kit de emergencia presente. Opciones múltiples: Primeros auxilios, Chaleco reflectivo, Conos/triángulos, Matafuegos, Kit para derrames, No Tiene, Incompleto. |
| `epps_vehiculo` | Texto | EPPs disponibles para conductor/acompañante. Opciones múltiples: Botines seguridad, Casco y máscara/anteojos, Ropa ignífuga (transp. sustancias peligrosas), Guantes de descarga, Guantes auxilio mecánico. |
| `observaciones_vehiculo` | Texto | Campo libre para observaciones adicionales (daños, nota técnica, referencia a evidencias/ fotos). |
| `creado_por` | Texto | Usuario que registró el vehículo. |
| `creado_en` | Datetime | Marca temporal de creación. |
| `actualizado_por` | Texto | Último usuario que modificó el registro. |
| `actualizado_en` | Datetime | Marca temporal de última modificación. |


---

### 🟧 Tabla: `FactoresExternos`
| Campo | Tipo | Descripción |
|--------|------|-------------|
| `id_factores` | Integer (PK) | Identificador único |
| `id_incidente` | Integer (FK → Incidente.id_incidente) | Relación con incidente |
| `tipo_superficie` | Texto | Tipo de superficie (asfalto, ripio, etc.) |
| `posee_banquina` | Enum(SI, NO, NA) | Indica si posee banquina |
| `tipo_ruta` | Texto | Tipo de vía |
| `densidad_trafico` | Enum(Alta, Media, Baja) | Tráfico al momento del hecho |
| `condicion_ruta` | Texto | Estado general del camino |
| `iluminacion_ruta` | Texto | Condición de luz natural o artificial |
| `senalizacion_ruta` | Texto | Estado de la señalización |
| `geometria_ruta` | Texto | Curvatura o pendiente |
| `condiciones_climaticas` | Texto | Condiciones meteorológicas |
| `rango_temperaturas` | Texto | Rango de temperatura ambiental |
