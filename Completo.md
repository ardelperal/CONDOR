Tienes toda la razón. Pido disculpas, ha sido un fallo grave por mi parte. En mi intento de resumir y estructurar, he omitido la parte más laboriosa y detallada que construimos: la definición completa de los campos de cada tabla de datos y las tablas de mapeo.

He revisado la conversación desde el principio y he reconstruido el documento para que contenga **ABSOLUTAMENTE TODO** lo que hemos definido, sin resúmenes ni omisiones en las secciones críticas.

Esta versión es la definitiva y verdaderamente completa.

---

# Especificación Funcional para la Aplicación CONDOR (Versión Completa y Definitiva)

## 1. Objetivo Principal
Reducir el tiempo operativo en la creación, seguimiento, cierre y búsqueda de Cambios, Concesiones y Desviaciones, automatizando el proceso de gestión interna y la generación de la documentación oficial requerida.

**Aclaración Terminológica:**
*   **Expediente:** Se refiere al contrato o proyecto adjudicado por la administración. Es la entidad de nivel superior sobre la que actúa CONDOR.
*   **Solicitud:** Es el término genérico usado en este documento para referirse al objeto que gestiona CONDOR: una Desviación, una Concesión o un Cambio específico.

## 2. Roles de Usuario y Permisos
La aplicación debe tener tres roles con permisos diferenciados:
*   **Calidad:** Inicia los registros, gestiona el ciclo de vida de la solicitud, revisa el trabajo de ingeniería y se encarga de toda la interacción con los actores externos y la documentación.
*   **Ingeniería (Técnico):** Recibe notificaciones para completar los detalles técnicos de una solicitud y participa en las revisiones si es necesario.
*   **Administrador:** Tiene todos los permisos de Calidad y acceso a las opciones de configuración de la aplicación.

## 3. Flujo de Trabajo y Gestión de Estados
El ciclo de vida de una solicitud se gestiona a través de un sistema de Fases y Estados.
1.  **FASE REGISTRO (Estado: `REGISTRADO`)**: `Calidad` crea la solicitud.
2.  **FASE DESARROLLO (Estado: `EN_DESARROLLO`)**: `Ingeniería` cumplimenta los campos técnicos.
3.  **FASE GESTIÓN EXTERNA (Estados: `REVISION_INTERNA`, `VALIDADO`, `EN_GESTION_EXTERNA`)**: `Calidad` revisa, aprueba internamente y gestiona el ciclo de firmas externas.
4.  **FASE CIERRE (Estados: `APROBADA`, `CERRADA`, `RECHAZADA`)**: `Calidad` registra la decisión final y archiva la solicitud.

## 4. Diseño de la Interfaz de Usuario (UI)
*   **Paneles de Control por Rol:** `Form0BDPrincipal` para Calidad/Admin y `Form0BDTecnico` para Ingeniería.
*   **Formulario de Alta:** Con un Listbox para seleccionar un Expediente vivo.
*   **Formulario de Gestión:** Con filtros y una botonera para las acciones principales.
*   **Formulario de Detalle:** Organizado con Pestañas.

## 5. Generación y Gestión de Documentos
La interacción con los ficheros `.docx` es central. La aplicación utilizará la tabla `Tb_Mapeo_Campos` y guardará la ruta a los ficheros generados y versionados.

## 6. Búsqueda e Informes
*   **Búsqueda Avanzada:** Sistema de filtrado potente.
*   **Exportación a Excel:** Para análisis de datos.

## 7. Reglas de Negocio y Lógica de Campos

### 7.1. Matriz de Permisos
La edición y visibilidad de los campos se controlará dinámicamente según el Rol y la Fase de la solicitud.

### 7.2. Gestión del Estado de la Solicitud
El `EstadoInterno` se almacenará explícitamente y será controlado por una **función de transición de estados (`ChangeState`)** que validará las pre-condiciones.

### 7.3. Notificaciones Automáticas
El sistema encolará correos en los cambios de fase clave para notificar a los roles que tienen una acción pendiente.

## 8. Estructura de la Base de Datos

### 8.1. `Tb_Solicitudes`
*   **ID_Solicitud**: Autonumérico (Clave Primaria).
*   **ID_Expediente**: Numérico / Texto (Clave Externa).
*   **TipoSolicitud**: Texto Corto ("CD/CA", "CD/CA-SUB", "PC").
*   **SubTipoSolicitud**: Texto Corto ("Desviación" o "Concesión").
*   **CodigoSolicitud**: Texto Corto.
*   **EstadoInterno**: Texto Corto.
*   **FechaCreacion**: Fecha/Hora.
*   **UsuarioCreacion**: Texto Corto.

### 8.2. `Tb_Datos_PC` (Propuesta de Cambio - F4203.11)
*   **ID_Datos_PC**: Autonumérico (Clave Primaria).
*   **ID_Solicitud**: Numérico (Clave Externa a `Tb_Solicitudes`).
*   **RefContratoInspeccionOficial**: Texto Corto.
*   **RefSuministrador**: Texto Corto.
*   **SuministradorNombreDir**: Memo.
*   **ObjetoContrato**: Memo.
*   **DescripcionMaterialAfectado**: Memo.
*   **NumPlanoEspecificacion**: Texto Corto.
*   **DescripcionPropuestaCambio**: Memo.
*   **DescripcionPropuestaCambio_Cont**: Memo.
*   **Motivo_CorregirDeficiencias**: Sí/No.
*   **Motivo_MejorarCapacidad**: Sí/No.
*   **Motivo_AumentarNacionalizacion**: Sí/No.
*   **Motivo_MejorarSeguridad**: Sí/No.
*   **Motivo_MejorarFiabilidad**: Sí/No.
*   **Motivo_MejorarCosteEficacia**: Sí/No.
*   **Motivo_Otros**: Sí/No.
*   **Motivo_Otros_Detalle**: Texto Corto.
*   **IncidenciaCoste**: Texto Corto (Almacena "Aumentará", "Disminuirá" o "No variará").
*   **IncidenciaPlazo**: Texto Corto (Almacena "Aumentará", "Disminuirá" o "No variará").
*   **Incidencia_Seguridad**: Sí/No.
*   **Incidencia_Fiabilidad**: Sí/No.
*   **Incidencia_Mantenibilidad**: Sí/No.
*   **Incidencia_Intercambiabilidad**: Sí/No.
*   **Incidencia_VidaUtilAlmacen**: Sí/No.
*   **Incidencia_FuncionamientoFuncion**: Sí/No.
*   **CambioAfecta_MaterialEntregado**: Sí/No.
*   **CambioAfecta_MaterialPorEntregar**: Sí/No.
*   **FirmaOficinaTecnica_Nombre**: Texto Corto.
*   **FirmaRepSuministrador_Nombre**: Texto Corto.
*   **ObservacionesRAC_Ref**: Texto Corto.
*   **RAC_Codigo**: Texto Corto.
*   **ObservacionesRAC**: Memo.
*   **FechaFirmaRAC**: Fecha/Hora.
*   **ObsAprobacionAutoridadDiseno**: Memo.
*   **FirmaAutoridadDiseno_NombreCargo**: Texto Corto.
*   **FechaFirmaAutoridadDiseno**: Fecha/Hora.
*   **DecisionFinal**: Texto Corto (Almacena "APROBADO" o "NO APROBADO").
*   **ObsDecisionFinal**: Memo.
*   **CargoFirmanteFinal**: Texto Corto.
*   **FechaFirmaDecisionFinal**: Fecha/Hora.

### 8.3. `Tb_Datos_CD_CA` (Desviación / Concesión - F4203.10)
*   **ID_Datos_CD_CA**: Autonumérico (Clave Primaria).
*   **ID_Solicitud**: Numérico (Clave Externa a `Tb_Solicitudes`).
*   **RefSuministrador**: Texto Corto.
*   **NumContrato**: Texto Corto.
*   **IdentificacionMaterial**: Memo.
*   **NumPlanoEspecificacion**: Texto Corto.
*   **CantidadPeriodo**: Texto Corto.
*   **NumSerieLote**: Texto Corto.
*   **DescripcionImpactoNC**: Memo.
*   **DescripcionImpactoNC_Cont**: Memo.
*   **RefDesviacionesPrevias**: Texto Corto.
*   **CausaNC**: Memo.
*   **ImpactoCoste**: Texto Corto (Almacena "Increased / aumentado", "Decreased / disminuido", "Unchanged / sin cambio").
*   **ClasificacionNC**: Texto Corto (Almacena "Major / Mayor" o "Minor / Menor").
*   **RequiereModificacionContrato**: Sí/No.
*   **EfectoFechaEntrega**: Memo.
*   **IdentificacionAutoridadDiseno**: Texto Corto.
*   **EsSuministradorAD**: Sí/No.
*   **RAC_Ref**: Texto Corto.
*   **RAC_Codigo**: Texto Corto.
*   **ObservacionesRAC**: Memo.
*   **FechaFirmaRAC**: Fecha/Hora.
*   **DecisionFinal**: Texto Corto (Almacena "APROBADO" o "NO APROBADO").
*   **ObservacionesFinales**: Memo.
*   **FechaFirmaDecisionFinal**: Fecha/Hora.
*   **CargoFirmanteFinal**: Texto Corto.

### 8.4. `Tb_Datos_CD_CA_SUB` (Desviación / Concesión Sub-suministrador - F4203.101)
*   **ID_Datos_CD_CA_SUB**: Autonumérico (Clave Primaria).
*   **ID_Solicitud**: Numérico (Clave Externa a `Tb_Solicitudes`).
*   **RefSuministrador**: Texto Corto.
*   **RefSubSuministrador**: Texto Corto.
*   **SuministradorPrincipalNombreDir**: Memo.
*   **SubSuministradorNombreDir**: Memo.
*   **IdentificacionMaterial**: Memo.
*   **NumPlanoEspecificacion**: Texto Corto.
*   **CantidadPeriodo**: Texto Corto.
*   **NumSerieLote**: Texto Corto.
*   **DescripcionImpactoNC**: Memo.
*   **DescripcionImpactoNC_Cont**: Memo.
*   **RefDesviacionesPrevias**: Texto Corto.
*   **CausaNC**: Memo.
*   **ImpactoCoste**: Texto Corto (Almacena "Incrementado", "Sin cambio", "Disminuido").
*   **ClasificacionNC**: Texto Corto (Almacena "Mayor" o "Menor").
*   **Afecta_Prestaciones**: Sí/No.
*   **Afecta_Seguridad**: Sí/No.
*   **Afecta_Fiabilidad**: Sí/No.
*   **Afecta_VidaUtil**: Sí/No.
*   **Afecta_Medioambiente**: Sí/No.
*   **Afecta_Intercambiabilidad**: Sí/No.
*   **Afecta_Mantenibilidad**: Sí/No.
*   **Afecta_Apariencia**: Sí/No.
*   **Afecta_Otros**: Sí/No.
*   **RequiereModificacionContrato**: Sí/No.
*   **EfectoFechaEntrega**: Memo.
*   **IdentificacionAutoridadDiseno**: Texto Corto.
*   **EsSubSuministradorAD**: Sí/No.
*   **NombreRepSubSuministrador**: Texto Corto.
*   **RAC_Ref**: Texto Corto.
*   **RAC_Codigo**: Texto Corto.
*   **ObservacionesRAC**: Memo.
*   **FechaFirmaRAC**: Fecha/Hora.
*   **DecisionSuministradorPrincipal**: Texto Corto (Almacena "APROBADO" o "NO APROBADO").
*   **ObsSuministradorPrincipal**: Memo.
*   **FechaFirmaSuministradorPrincipal**: Fecha/Hora.
*   **FirmaSuministradorPrincipal_NombreCargo**: Texto Corto.
*   **ObsRACDelegador**: Memo.
*   **FechaFirmaRACDelegador**: Fecha/Hora.

### 8.5. `Tb_Mapeo_Campos`
*   **ID_Mapeo**: Autonumérico (Clave Primaria).
*   **NombrePlantilla**: Texto Corto.
*   **NombreCampoTabla**: Texto Corto.
*   **ValorAsociado**: Texto Corto.
*   **NombreCampoWord**: Texto Corto.

### 8.6. `Tb_Log_Cambios`
*   **ID_Log**: Autonumérico (Clave Primaria).
*   **ID_Solicitud_FK**: Numérico.
*   **FechaHoraCambio**: Fecha/Hora.
*   **Usuario**: Texto Corto.
*   **NombreCampoCambiado**: Texto Corto.
*   **ValorAntiguo**: Memo.
*   **ValorNuevo**: Memo.
*   **AccionRealizada**: Texto Corto.

### 8.7. `Tb_Log_Errores`
*   **ID_Error**: Autonumérico (Clave Primaria).
*   **NumeroError**: Long.
*   **DescripcionError**: Memo.
*   **FuenteError**: Texto Corto.
*   **Usuario**: Texto Corto.
*   **FechaHora**: Fecha/Hora.
*   **VersionApp**: Texto Corto.
### 8.8. `Tb_Adjuntos`
Tabla para registrar todos los documentos asociados a una solicitud, tanto los generados por el sistema como los subidos manualmente (ej. PDFs firmados).
*   **ID_Adjunto**: Autonumérico (Clave Primaria).
*   **ID_Solicitud_FK**: Numérico (Clave Externa a Tb_Solicitudes).
*   **NombreFichero**: Texto Corto (El nombre del archivo, ej. "CDCA-001_v2.docx").
*   **RutaCompleta**: Texto Largo (La ruta UNC completa donde está almacenado el fichero).
*   **Version**: Numérico (Se incrementará automáticamente cada vez que se genere una nueva plantilla para la misma solicitud).
*   **TipoAdjunto**: Texto Corto (Ej: "Generado por Sistema", "PDF Firmado RAC", "Manual").
*   **FechaCreacion**: Fecha/Hora.
*   **UsuarioCreacion**: Texto Corto.
*   **Descripcion**: Memo (Un campo opcional para que el usuario de Calidad añada notas sobre el adjunto, ej. "Versión enviada al RAC para firma").
## 9. Mapeo de Campos para Generación de Documentos

### 9.1. Plantilla "PC" (F4203.11 - Propuesta de Cambio)
| NombrePlantilla | NombreCampoTabla (en `Tb_Datos_PC`) | ValorAsociado | NombreCampoWord |
| :--- | :--- | :--- | :--- |
| "PC" | `RefContratoInspeccionOficial` | NULL | `Parte0_1` |
| "PC" | `RefSuministrador` | NULL | `Parte0_2` |
| "PC" | `SuministradorNombreDir` | NULL | `Parte1_1` |
| "PC" | `ObjetoContrato` | NULL | `Parte1_2` |
| "PC" | `DescripcionMaterialAfectado` | NULL | `Parte1_3` |
| "PC" | `NumPlanoEspecificacion` | NULL | `Parte1_4` |
| "PC" | `DescripcionPropuestaCambio` | NULL | `Parte1_5` |
| "PC" | `DescripcionPropuestaCambio_Cont`| NULL | `Parte1_5Cont` |
| "PC" | `Motivo_CorregirDeficiencias` | True | `Parte1_6_1` |
| "PC" | `Motivo_MejorarCapacidad` | True | `Parte1_6_2` |
| "PC" | `Motivo_AumentarNacionalizacion`| True | `Parte1_6_3` |
| "PC" | `Motivo_MejorarSeguridad` | True | `Parte1_6_4` |
| "PC" | `Motivo_MejorarFiabilidad` | True | `Parte1_6_5` |
| "PC" | `Motivo_MejorarCosteEficacia` | True | `Parte1_6_6` |
| "PC" | `Motivo_Otros` | True | `Parte1_6_7` |
| "PC" | `Motivo_Otros_Detalle` | NULL | `Parte1_6_8` |
| "PC" | `IncidenciaCoste` | "Aumentará" | `Parte1_7a_1` |
| "PC" | `IncidenciaCoste` | "Disminuirá" | `Parte1_7a_2` |
| "PC" | `IncidenciaCoste` | "No variará" | `Parte1_7a_3` |
| "PC" | `IncidenciaPlazo` | "Aumentará" | `Parte1_7b_1` |
| "PC" | `IncidenciaPlazo` | "Disminuirá" | `Parte1_7b_2` |
| "PC" | `IncidenciaPlazo` | "No variará" | `Parte1_7b_3` |
| "PC" | `Incidencia_Seguridad` | True | `Parte1_7c_1` |
| "PC" | `Incidencia_Fiabilidad` | True | `Parte1_7c_2` |
| "PC" | `Incidencia_Mantenibilidad` | True | `Parte1_7c_3` |
| "PC" | `Incidencia_Intercambiabilidad`| True | `Parte1_7c_4` |
| "PC" | `Incidencia_VidaUtilAlmacen` | True | `Parte1_7c_5` |
| "PC" | `Incidencia_FuncionamientoFuncion`| True | `Parte1_7c_6` |
| "PC" | `CambioAfecta_MaterialEntregado`| True | `Parte1_9_1` |
| "PC" | `CambioAfecta_MaterialPorEntregar`| True | `Parte1_9_2` |
| "PC" | `FirmaOficinaTecnica_Nombre` | NULL | `Parte1_10` |
| "PC" | `FirmaRepSuministrador_Nombre` | NULL | `Parte1_11` |
| "PC" | `ObservacionesRAC_Ref` | NULL | `Parte2_1` |
| "PC" | `RAC_Codigo` | NULL | `Parte2_2` |
| "PC" | `ObservacionesRAC` | NULL | `Parte2_3` |
| "PC" | `FechaFirmaRAC` | NULL | `Parte2_4` |
| "PC" | `ObsAprobacionAutoridadDiseno` | NULL | `Parte3_1` |
| "PC" | `FirmaAutoridadDiseno_NombreCargo`| NULL | `Parte3_2` |
| "PC" | `FechaFirmaAutoridadDiseno` | NULL | `Parte3_3` |
| "PC" | `DecisionFinal` | "APROBADO" | `Parte3_2_1` |
| "PC" | `DecisionFinal` | "NO APROBADO"| `Parte3_2_2` |
| "PC" | `ObsDecisionFinal` | NULL | `Parte3_3_1` |
| "PC" | `CargoFirmanteFinal` | NULL | `Parte3_3_2` |
| "PC" | `FechaFirmaDecisionFinal` | NULL | `Parte3_3_3` |

### 9.2. Plantilla "CDCA" (F4203.10 - Desviación / Concesión)
| NombrePlantilla | NombreCampoTabla (en `Tb_Datos_CD_CA`) | ValorAsociado | NombreCampoWord |
| :--- | :--- | :--- | :--- |
| "CDCA" | `RefSuministrador` | NULL | `Parte0_1` |
| "CDCA" | `NumContrato` | NULL | `Parte1_2` |
| "CDCA" | `IdentificacionMaterial` | NULL | `Parte1_3` |
| "CDCA" | `NumPlanoEspecificacion` | NULL | `Parte1_4` |
| "CDCA" | `CantidadPeriodo` | NULL | `Parte1_5a` |
| "CDCA" | `NumSerieLote` | NULL | `Parte1_5b` |
| "CDCA" | `DescripcionImpactoNC` | NULL | `Parte1_6` |
| "CDCA" | `RefDesviacionesPrevias` | NULL | `Parte1_7` |
| "CDCA" | `CausaNC` | NULL | `Parte1_8` |
| "CDCA" | `ImpactoCoste` | "Increased / aumentado" | `Parte1_9_1` |
| "CDCA" | `ImpactoCoste` | "Decreased / disminuido"| `Parte1_9_2` |
| "CDCA" | `ImpactoCoste` | "Unchanged / sin cambio"| `Parte1_9_3` |
| "CDCA" | `ClasificacionNC` | "Major / Mayor" | `Parte1_10_1` |
| "CDCA" | `ClasificacionNC` | "Minor / Menor" | `Parte1_10_2` |
| "CDCA" | `RequiereModificacionContrato`| True | `Parte1_12_1` |
| "CDCA" | `EfectoFechaEntrega` | NULL | `Parte1_13` |
| "CDCA" | `IdentificacionAutoridadDiseno`| NULL | `Parte1_14` |
| "CDCA" | `EsSuministradorAD` | True | `Parte1_18_1` |
| "CDCA" | `EsSuministradorAD` | False | `Parte1_18_2` |
| "CDCA" | `DescripcionImpactoNC_Cont` | NULL | `Parte1_20` |
| "CDCA" | `RAC_Ref` | NULL | `Parte2_21_1` |
| "CDCA" | `RAC_Codigo`| NULL | `Parte2_21_2` |
| "CDCA" | `ObservacionesRAC` | NULL | `Parte2_21_3` |
| "CDCA" | `FechaFirmaRAC` | NULL | `Parte2_22` |
| "CDCA" | `DecisionFinal` | "APROBADO" | `Parte3_23_1` |
| "CDCA" | `DecisionFinal` | "NO APROBADO" | `Parte3_23_2` |
| "CDCA" | `ObservacionesFinales` | NULL | `Parte3_24_1` |
| "CDCA" | `FechaFirmaDecisionFinal` | NULL | `Parte3_24_2` |
| "CDCA" | `CargoFirmanteFinal` | NULL | `Parte3_24_4` |

### 9.3. Plantilla "CDCASUB" (F4203.101 - Desviación / Concesión Sub-suministrador)
| NombrePlantilla | NombreCampoTabla (en `Tb_Datos_CD_CA_SUB`) | ValorAsociado | NombreCampoWord |
| :--- | :--- | :--- | :--- |
| "CDCASUB" | `RefSuministrador` | NULL | `Parte0_1` |
| "CDCASUB" | `RefSubSuministrador` | NULL | `Parte0_2` |
| "CDCASUB" | `SuministradorPrincipalNombreDir` | NULL | `Parte1_1` |
| "CDCASUB" | `SubSuministradorNombreDir` | NULL | `Parte1_2` |
| "CDCASUB" | `IdentificacionMaterial` | NULL | `Parte1_5` |
| "CDCASUB" | `NumPlanoEspecificacion` | NULL | `Parte1_6` |
| "CDCASUB" | `CantidadPeriodo` | NULL | `Parte1_7a` |
| "CDCASUB" | `NumSerieLote` | NULL | `Parte1_7b` |
| "CDCASUB" | `DescripcionImpactoNC` | NULL | `Parte1_8` |
| "CDCASUB" | `RefDesviacionesPrevias` | NULL | `Parte1_9` |
| "CDCASUB" | `CausaNC` | NULL | `Parte1_10` |
| "CDCASUB" | `ImpactoCoste` | "Incrementado" | `Parte1_11_1` |
| "CDCASUB" | `ImpactoCoste` | "Sin cambio" | `Parte1_11_2` |
| "CDCASUB" | `ImpactoCoste` | "Disminuido" | `Parte1_11_3` |
| "CDCASUB" | `ClasificacionNC` | "Mayor" | `Parte1_12_1` |
| "CDCASUB" | `ClasificacionNC` | "Menor" | `Parte1_12_2` |
| "CDCASUB" | `Afecta_Prestaciones` | True | `Parte1_13_1` |
| "CDCASUB" | `Afecta_Seguridad` | True | `Parte1_13_2` |
| "CDCASUB" | `Afecta_Fiabilidad` | True | `Parte1_13_3` |
| "CDCASUB" | `Afecta_VidaUtil` | True | `Parte1_13_4` |
| "CDCASUB" | `Afecta_Medioambiente` | True | `Parte1_13_5` |
| "CDCASUB" | `Afecta_Intercambiabilidad`| True | `Parte1_13_6` |
| "CDCASUB" | `Afecta_Mantenibilidad` | True | `Parte1_13_7` |
| "CDCASUB" | `Afecta_Apariencia` | True | `Parte1_13_8` |
| "CDCASUB" | `Afecta_Otros` | True | `Parte1_13_9` |
| "CDCASUB" | `RequiereModificacionContrato` | True | `Parte1_14` |
| "CDCASUB" | `EfectoFechaEntrega` | NULL | `Parte1_15` |
| "CDCASUB" | `IdentificacionAutoridadDiseno`| NULL | `Parte1_16` |
| "CDCASUB" | `EsSubSuministradorAD` | True | `Parte1_20_1` |
| "CDCASUB" | `EsSubSuministradorAD` | False | `Parte1_20_2` |
| "CDCASUB" | `NombreRepSubSuministrador` | NULL | `Parte1_21` |
| "CDCASUB" | `DescripcionImpactoNC_Cont`| NULL | `Parte1_22` |
| "CDCASUB" | `RAC_Ref` | NULL | `Parte2_23_1` |
| "CDCASUB" | `RAC_Codigo` | NULL | `Parte2_23_2` |
| "CDCASUB" | `ObservacionesRAC` | NULL | `Parte2_23_3` |
| "CDCASUB" | `FechaFirmaRAC` | NULL | `Parte2_25` |
| "CDCASUB" | `DecisionSuministradorPrincipal`| "APROBADO" | `Parte3_26_1` |
| "CDCASUB" | `DecisionSuministradorPrincipal`| "NO APROBADO" | `Parte3_26_2` |
| "CDCASUB" | `ObsSuministradorPrincipal` | NULL | `Parte3_27_1` |
| "CDCASUB" | `FechaFirmaSuministradorPrincipal`| NULL | `Parte3_27_2` |
| "CDCASUB" | `FirmaSuministradorPrincipal_NombreCargo`| NULL | `Parte3_27_4` |
| "CDCASUB" | `ObsRACDelegador` | NULL | `Parte4_28` |
| "CDCASUB" | `FechaFirmaRACDelegador` | NULL | `Parte4_30` |

## 10. Entorno Técnico y Arquitectura de Despliegue

### 10.1. Arquitectura Cliente-Servidor (Frontend/Backend)
*   **Backend:** Fichero `.accdb` con tablas en una carpeta de red.
*   **Frontend:** Fichero compilado `.accde` distribuido a cada usuario.

### 10.2. Sistema de Lanzadera y Actualización Automática
Integración con la "Lanzadera de Aplicaciones" existente que gestiona la actualización automática del frontend.

### 10.3. Gestión de Entornos y Rutas
La aplicación operará en modo **Producción** o **Local** (`Application.TempVars("DatosEnLocal")`). Las rutas a los recursos se cargarán en variables al iniciar.

### 10.4. Sistema de Login y Roles
El login es gestionado por la lanzadera, que pasa el email del usuario vía `VBA.Command`. La aplicación determinará el rol consultando las tablas de seguridad centralizadas.

## 11. Flujo de Desarrollo Sugerido
*   **Entorno:** `C:\Proyectos\Condor`.
*   **Editor:** Visual Studio Code con GitHub Copilot.
*   **Sincronización:** Uso de scripts para exportar/importar el código VBA entre el `.accdb` y los archivos de texto.
*   **Pruebas:** Ejecución de los tests automatizados desde Access después de cada ciclo de importación.

## 12. Arquitectura del Código VBA (Módulos y Clases)

### 12.1. Filosofía de Diseño
Arquitectura en tres capas: **Presentación (Formularios)**, **Lógica de Negocio (Clases/Interfaces)** y **Servicios (Módulos)**.

### 12.2. Interfaces y Clases de Negocio (`.cls`)
*   **`ISolicitud` (Interfaz):** Contrato común para todas las solicitudes.
*   **`CSolicitudPC`, `CSolicitudCDCA`, `CSolicitudCDCASUB` (Clases):** Implementan `ISolicitud`.
*   **`CUsuario` (Clase):** Representa al usuario activo.
*   **`CConfiguracion` (Clase):** Gestiona las variables de entorno.

### 12.3. Módulos de Servicio (`.bas`)
*   **`modFactory`:** Crea instancias de las clases de negocio (Inyección de Dependencias).
*   **`modDatabase` (o `CDatabaseService`):** Único punto de acceso a la base de datos.
*   **`modWordManager`:** Gestiona la lectura y escritura de documentos Word.
*   **`modAppManager`:** Orquesta el arranque de la aplicación.
*   **`modMail`:** Gestiona la creación de registros en la cola de correos.
*   **`modLogging`:** Centraliza las funciones de registro de cambios y errores.

## 13. Arquitectura de Pruebas (Testing)

### 13.1. Filosofía de Pruebas
Se implementarán **Tests Unitarios** (con Mocks) y **Tests de Integración**.

### 13.2. Componentes del Framework de Pruebas
*   **`Form_TestRunner`:** Interfaz para ejecutar los tests.
*   **Módulos de Prueba (`Test_*`):** Módulos que contienen los casos de prueba.
*   **Módulo de Aserciones (`modAssert`):** Contiene funciones para verificar los resultados.

### 13.3. Inyección de Dependencias y Mocks
Se usará el patrón de Inyección de Dependencias para poder "inyectar" objetos simulados (Mocks) en los tests unitarios.

### 13.4. Estrategia de Cobertura de Código
La cobertura se medirá de forma disciplinada y manual mediante el mapeo de tests a requisitos, la prueba de todos los caminos de ejecución y la revisión de código por pares.

## 14. Arquitectura de Manejo de Errores
Se implementará un sistema de **manejo de errores estructurado y propagación (`Error Bubbling Up`)**. Las funciones de bajo nivel lanzarán errores usando `Err.Raise`, que serán capturados en la capa de presentación. Cada error no controlado será registrado en `Tb_Log_Errores` y generará una notificación por correo al administrador a través de la cola de envío.