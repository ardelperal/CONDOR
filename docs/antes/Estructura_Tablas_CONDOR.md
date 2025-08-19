# Estructura de Tablas del Sistema CONDOR

## Resumen

Este documento describe la estructura completa de las tablas creadas en la base de datos `CONDOR_datos.accdb` del sistema CONDOR, siguiendo la especificación funcional del proyecto.

## Tablas Principales

### 1. Tb_Solicitudes
**Descripción**: Tabla central que almacena todas las solicitudes del sistema.

| Campo | Tipo | Descripción |
|-------|------|-------------|
| ID | AUTOINCREMENT | Identificador único de la solicitud |
| NumeroExpediente | TEXT(50) | Número del expediente asociado |
| TipoSolicitud | TEXT(10) | Tipo: PC, CD_CA, CD_CA_SUB |
| EstadoInterno | TEXT(20) | Estado interno de la solicitud |
| EstadoRAC | TEXT(20) | Estado en el sistema RAC |
| FechaCreacion | DATETIME | Fecha de creación de la solicitud |
| FechaUltimaModificacion | DATETIME | Última modificación |
| Usuario | TEXT(50) | Usuario que creó la solicitud |
| Observaciones | MEMO | Observaciones generales |
| Activo | YESNO | Indica si el registro está activo |

### 2. TbExpedientes
**Descripción**: Cache local de expedientes obtenidos del sistema RAC.

| Campo | Tipo | Descripción |
|-------|------|-------------|
| ID | AUTOINCREMENT | Identificador único |
| NumeroExpediente | TEXT(50) | Número del expediente |
| TituloExpediente | TEXT(255) | Título del expediente |
| EstadoExpediente | TEXT(50) | Estado actual del expediente |
| FechaCreacionExpediente | DATETIME | Fecha de creación en RAC |
| FechaUltimaActualizacion | DATETIME | Última actualización desde RAC |
| DatosCompletos | MEMO | Datos completos en formato JSON |
| Activo | YESNO | Indica si el registro está activo |

### 3. TbConfiguracion
**Descripción**: Configuración de entornos y parámetros del sistema.

| Campo | Tipo | Descripción |
|-------|------|-------------|
| ID | AUTOINCREMENT | Identificador único |
| Entorno | TEXT(50) | Nombre del entorno (Desarrollo, Producción) |
| RutaBackend | TEXT(255) | Ruta a la base de datos backend |
| Descripcion | TEXT(255) | Descripción del entorno |
| FechaCreacion | DATETIME | Fecha de creación |
| Activo | YESNO | Indica si el entorno está activo |

## Tablas de Datos Específicas

### 4. TbDatos_PC
**Descripción**: Datos específicos para Propuestas de Cambio.

| Campo | Tipo | Descripción |
|-------|------|-------------|
| ID | AUTOINCREMENT | Identificador único |
| SolicitudID | LONG | Referencia a Tb_Solicitudes |
| NumeroExpediente | TEXT(50) | Número del expediente |
| TipoSolicitud | TEXT(10) | Tipo de solicitud (PC) |
| DescripcionCambio | MEMO | Descripción del cambio propuesto |
| JustificacionCambio | MEMO | Justificación del cambio |
| ImpactoSeguridad | MEMO | Análisis de impacto en seguridad |
| ImpactoCalidad | MEMO | Análisis de impacto en calidad |
| FechaCreacion | DATETIME | Fecha de creación |
| FechaUltimaModificacion | DATETIME | Última modificación |
| Estado | TEXT(20) | Estado actual |
| Activo | YESNO | Indica si el registro está activo |

### 5. TbDatos_CD_CA
**Descripción**: Datos específicos para Concesiones y Desviaciones.

| Campo | Tipo | Descripción |
|-------|------|-------------|
| ID | AUTOINCREMENT | Identificador único |
| SolicitudID | LONG | Referencia a Tb_Solicitudes |
| NumeroExpediente | TEXT(50) | Número del expediente |
| TipoSolicitud | TEXT(10) | Tipo de solicitud (CD_CA) |
| DescripcionDesviacion | MEMO | Descripción de la desviación |
| JustificacionDesviacion | MEMO | Justificación de la desviación |
| ImpactoSeguridad | MEMO | Análisis de impacto en seguridad |
| ImpactoCalidad | MEMO | Análisis de impacto en calidad |
| MedidasCorrectivas | MEMO | Medidas correctivas propuestas |
| FechaCreacion | DATETIME | Fecha de creación |
| FechaUltimaModificacion | DATETIME | Última modificación |
| Estado | TEXT(20) | Estado actual |
| Activo | YESNO | Indica si el registro está activo |

### 6. TbDatos_CD_CA_SUB
**Descripción**: Datos específicos para Concesiones y Desviaciones de Sub-suministrador.

| Campo | Tipo | Descripción |
|-------|------|-------------|
| ID | AUTOINCREMENT | Identificador único |
| SolicitudID | LONG | Referencia a Tb_Solicitudes |
| NumeroExpediente | TEXT(50) | Número del expediente |
| TipoSolicitud | TEXT(10) | Tipo de solicitud (CD_CA_SUB) |
| NombreSubsuministrador | TEXT(255) | Nombre del sub-suministrador |
| DescripcionDesviacion | MEMO | Descripción de la desviación |
| JustificacionDesviacion | MEMO | Justificación de la desviación |
| ImpactoSeguridad | MEMO | Análisis de impacto en seguridad |
| ImpactoCalidad | MEMO | Análisis de impacto en calidad |
| MedidasCorrectivas | MEMO | Medidas correctivas propuestas |
| FechaCreacion | DATETIME | Fecha de creación |
| FechaUltimaModificacion | DATETIME | Última modificación |
| Estado | TEXT(20) | Estado actual |
| Activo | YESNO | Indica si el registro está activo |

## Tablas de Soporte

### 7. TbMapeo_Campos
**Descripción**: Mapeo entre campos de las tablas de datos y marcadores en plantillas Word para generación de documentos.

| Campo | Tipo | Descripción |
|-------|------|-------------|
| ID | AUTOINCREMENT | Identificador único |
| NombrePlantilla | TEXT(50) | Tipo de plantilla (PC, CDCA, CDCASUB) |
| NombreCampoTabla | TEXT(100) | Nombre del campo en la tabla de datos |
| ValorAsociado | TEXT(255) | Valor específico para campos con opciones múltiples |
| NombreCampoWord | TEXT(100) | Nombre del marcador en la plantilla Word |
| FechaCreacion | DATETIME | Fecha de creación del registro |
| Activo | YESNO | Indica si el mapeo está activo |

**Nota**: Esta tabla contiene 116 registros que mapean los campos de las plantillas:
- **PC**: Propuesta de Cambio (F4203.11)
- **CDCA**: Desviación/Concesión (F4203.10) 
- **CDCASUB**: Desviación/Concesión Sub-suministrador (F4203.101)

### 8. TbLog_Cambios
**Descripción**: Auditoría de cambios en el sistema.

| Campo | Tipo | Descripción |
|-------|------|-------------|
| ID | AUTOINCREMENT | Identificador único |
| TablaAfectada | TEXT(50) | Nombre de la tabla modificada |
| RegistroID | LONG | ID del registro modificado |
| CampoModificado | TEXT(50) | Campo que fue modificado |
| ValorAnterior | MEMO | Valor anterior del campo |
| ValorNuevo | MEMO | Nuevo valor del campo |
| Usuario | TEXT(50) | Usuario que realizó el cambio |
| FechaCambio | DATETIME | Fecha y hora del cambio |
| TipoOperacion | TEXT(20) | Tipo: INSERT, UPDATE, DELETE |
| Observaciones | MEMO | Observaciones adicionales |

### 9. TbLog_Errores
**Descripción**: Registro de errores del sistema.

| Campo | Tipo | Descripción |
|-------|------|-------------|
| ID | AUTOINCREMENT | Identificador único |
| FechaError | DATETIME | Fecha y hora del error |
| Usuario | TEXT(50) | Usuario que experimentó el error |
| Modulo | TEXT(50) | Módulo donde ocurrió el error |
| Procedimiento | TEXT(100) | Procedimiento específico |
| NumeroError | LONG | Número de error de VBA |
| DescripcionError | MEMO | Descripción del error |
| ParametrosEntrada | MEMO | Parámetros que causaron el error |
| EstadoSistema | MEMO | Estado del sistema al momento del error |
| Resuelto | YESNO | Indica si el error fue resuelto |

### 10. TbAdjuntos
**Descripción**: Gestión de archivos adjuntos.

| Campo | Tipo | Descripción |
|-------|------|-------------|
| ID | AUTOINCREMENT | Identificador único |
| SolicitudID | LONG | Referencia a Tb_Solicitudes |
| NombreArchivo | TEXT(255) | Nombre original del archivo |
| RutaArchivo | TEXT(255) | Ruta donde se almacena el archivo |
| TipoArchivo | TEXT(50) | Tipo/extensión del archivo |
| TamanoArchivo | LONG | Tamaño del archivo en bytes |
| FechaSubida | DATETIME | Fecha de subida del archivo |
| Usuario | TEXT(50) | Usuario que subió el archivo |
| Descripcion | TEXT(255) | Descripción del archivo |
| Activo | YESNO | Indica si el archivo está activo |

## Tablas de Usuarios

### 11. TbUsuariosAplicaciones
**Descripción**: Gestión de usuarios del sistema.

| Campo | Tipo | Descripción |
|-------|------|-------------|
| ID | AUTOINCREMENT | Identificador único |
| Usuario | TEXT(50) | Nombre de usuario |
| NombreCompleto | TEXT(100) | Nombre completo del usuario |
| Email | TEXT(100) | Correo electrónico |
| Departamento | TEXT(50) | Departamento al que pertenece |
| Rol | TEXT(30) | Rol del usuario en el sistema |
| FechaCreacion | DATETIME | Fecha de creación del usuario |
| FechaUltimoAcceso | DATETIME | Último acceso al sistema |
| Activo | YESNO | Indica si el usuario está activo |

### 12. TbUsuariosAplicacionesPermisos
**Descripción**: Permisos específicos de usuarios por módulo.

| Campo | Tipo | Descripción |
|-------|------|-------------|
| ID | AUTOINCREMENT | Identificador único |
| UsuarioID | LONG | Referencia a TbUsuariosAplicaciones |
| Modulo | TEXT(50) | Módulo del sistema |
| Permiso | TEXT(50) | Tipo de permiso |
| Lectura | YESNO | Permiso de lectura |
| Escritura | YESNO | Permiso de escritura |
| Eliminacion | YESNO | Permiso de eliminación |
| Administracion | YESNO | Permiso de administración |
| FechaAsignacion | DATETIME | Fecha de asignación del permiso |
| Activo | YESNO | Indica si el permiso está activo |

## Tablas de Workflow

### 13. TbEstados
**Descripción**: Estados disponibles en el sistema de workflow.

| Campo | Tipo | Descripción |
|-------|------|-------------|
| ID | AUTOINCREMENT | Identificador único |
| CodigoEstado | TEXT(20) | Código único del estado |
| NombreEstado | TEXT(50) | Nombre descriptivo del estado |
| DescripcionEstado | TEXT(255) | Descripción detallada |
| TipoSolicitud | TEXT(10) | Tipo de solicitud aplicable |
| EsEstadoInicial | YESNO | Indica si es estado inicial |
| EsEstadoFinal | YESNO | Indica si es estado final |
| RequiereAprobacion | YESNO | Indica si requiere aprobación |
| FechaCreacion | DATETIME | Fecha de creación |
| Activo | YESNO | Indica si el estado está activo |

### 14. TbTransiciones
**Descripción**: Transiciones permitidas entre estados.

| Campo | Tipo | Descripción |
|-------|------|-------------|
| ID | AUTOINCREMENT | Identificador único |
| EstadoOrigenID | LONG | Referencia al estado origen |
| EstadoDestinoID | LONG | Referencia al estado destino |
| TipoSolicitud | TEXT(10) | Tipo de solicitud aplicable |
| CondicionTransicion | TEXT(255) | Condición para la transición |
| RequiereAprobacion | YESNO | Indica si requiere aprobación |
| RolRequerido | TEXT(30) | Rol necesario para la transición |
| FechaCreacion | DATETIME | Fecha de creación |
| Activo | YESNO | Indica si la transición está activa |

### 15. TbHistorialEstados
**Descripción**: Historial de cambios de estado de las solicitudes.

| Campo | Tipo | Descripción |
|-------|------|-------------|
| ID | AUTOINCREMENT | Identificador único |
| SolicitudID | LONG | Referencia a Tb_Solicitudes |
| EstadoAnterior | TEXT(20) | Estado anterior |
| EstadoNuevo | TEXT(20) | Nuevo estado |
| FechaCambio | DATETIME | Fecha del cambio de estado |
| Usuario | TEXT(50) | Usuario que realizó el cambio |
| Comentarios | MEMO | Comentarios del cambio |
| TipoTransicion | TEXT(50) | Tipo de transición realizada |
| Activo | YESNO | Indica si el registro está activo |

## Relaciones entre Tablas

### Relaciones Principales:
1. **Tb_Solicitudes** → **TbDatos_PC/TbDatos_CD_CA/TbDatos_CD_CA_SUB** (1:1)
2. **Tb_Solicitudes** → **TbAdjuntos** (1:N)
3. **Tb_Solicitudes** → **TbHistorialEstados** (1:N)
4. **TbUsuariosAplicaciones** → **TbUsuariosAplicacionesPermisos** (1:N)
5. **TbEstados** → **TbTransiciones** (1:N como origen y destino)

## Índices Recomendados

Para optimizar el rendimiento, se recomienda crear índices en:
- `Tb_Solicitudes.NumeroExpediente`
- `Tb_Solicitudes.TipoSolicitud`
- `Tb_Solicitudes.EstadoInterno`
- `TbExpedientes.NumeroExpediente`
- `TbHistorialEstados.SolicitudID`
- `TbAdjuntos.SolicitudID`

---

**Fecha de creación**: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
**Versión**: 1.0
**Sistema**: CONDOR - Control de Documentos y Registros