# Especificaciones Técnicas de Integración - Sistema CONDOR

## Introducción

Este documento detalla las especificaciones técnicas para la integración del Sistema CONDOR con los sistemas externos, basado en la arquitectura definida en la especificación funcional del proyecto.

---

## 1. Integración con ExpedienteService

### 1.1 Descripción General

El ExpedienteService es el servicio principal para obtener información de expedientes desde el sistema externo de gestión de contratos. CONDOR utiliza este servicio para:

- Obtener datos básicos del expediente al crear solicitudes
- Validar la existencia y estado de expedientes
- Sincronizar información de responsables y contratistas

### 1.2 Arquitectura de Integración

```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────────┐
│   CONDOR.accdb  │    │  ExpedienteService │    │ Sistema Expedientes │
│   (Frontend)    │◄──►│   (Capa Servicio) │◄──►│   (Base de Datos)   │
└─────────────────┘    └──────────────────┘    └─────────────────────┘
```

### 1.3 Interfaz IExpedienteService

#### Definición de la Interfaz
```vba
' Archivo: IExpedienteService.cls
' Descripción: Interfaz para el servicio de expedientes

Option Compare Database
Option Explicit

' Método principal para obtener datos de expediente
Public Function ObtenerExpediente(numeroExpediente As String) As T_Expediente
End Function

' Método para validar existencia de expediente
Public Function ExisteExpediente(numeroExpediente As String) As Boolean
End Function

' Método para obtener lista de expedientes por responsable
Public Function ObtenerExpedientesPorResponsable(emailResponsable As String) As Collection
End Function
```

#### Implementación CExpedienteService
```vba
' Archivo: CExpedienteService.cls
' Descripción: Implementación concreta del servicio de expedientes

Option Compare Database
Option Explicit
Implements IExpedienteService

Private Function IExpedienteService_ObtenerExpediente(numeroExpediente As String) As T_Expediente
    ' Implementación que consulta la base de datos de expedientes
    ' Retorna objeto T_Expediente con todos los datos necesarios
End Function

Private Function IExpedienteService_ExisteExpediente(numeroExpediente As String) As Boolean
    ' Validación rápida de existencia sin cargar todos los datos
End Function

Private Function IExpedienteService_ObtenerExpedientesPorResponsable(emailResponsable As String) As Collection
    ' Retorna colección de expedientes asignados al responsable
End Function
```

### 1.4 Estructura de Datos T_Expediente

```vba
' Archivo: T_Expediente.cls
' Descripción: Tipo de datos para expedientes

Option Compare Database
Option Explicit

' Propiedades principales
Public NumeroExpediente As String
Public Nemotecnico As String
Public TituloExpediente As String
Public EstadoExpediente As String
Public ResponsableCalidad As String
Public EmailResponsable As String
Public JefeProyecto As String
Public ContratistaPrincipal As String
Public ValorContrato As Currency
Public FechaInicio As Date
Public FechaFinPrevista As Date
Public UltimaActualizacion As Date

' Método para validar completitud de datos
Public Function EsValido() As Boolean
    EsValido = (Len(NumeroExpediente) > 0 And _
                Len(ResponsableCalidad) > 0 And _
                Len(EmailResponsable) > 0)
End Function
```

### 1.5 Configuración de Conexión

#### Parámetros de Configuración (TbConfiguracion)
```sql
-- Configuración de conexión a ExpedienteService
INSERT INTO TbConfiguracion (Clave, Valor, Descripcion, Categoria) VALUES
('EXPEDIENTE_SERVICE_URL', 'http://servidor-expedientes:8080/api', 'URL base del servicio de expedientes', 'Integración'),
('EXPEDIENTE_SERVICE_TIMEOUT', '30', 'Timeout en segundos para llamadas al servicio', 'Integración'),
('EXPEDIENTE_SERVICE_RETRY', '3', 'Número de reintentos en caso de fallo', 'Integración'),
('EXPEDIENTE_SERVICE_AUTH', 'Bearer', 'Tipo de autenticación', 'Integración');
```

### 1.6 Manejo de Errores

```vba
' Códigos de error específicos para ExpedienteService
Public Enum ExpedienteServiceError
    ESE_EXPEDIENTE_NO_ENCONTRADO = 1001
    ESE_SERVICIO_NO_DISPONIBLE = 1002
    ESE_TIMEOUT_CONEXION = 1003
    ESE_DATOS_INCOMPLETOS = 1004
    ESE_ACCESO_DENEGADO = 1005
End Enum

' Función para manejo centralizado de errores
Public Function ManejarErrorExpedienteService(errorCode As Long, errorDesc As String) As String
    Select Case errorCode
        Case ESE_EXPEDIENTE_NO_ENCONTRADO
            ManejarErrorExpedienteService = "El expediente especificado no existe en el sistema"
        Case ESE_SERVICIO_NO_DISPONIBLE
            ManejarErrorExpedienteService = "El servicio de expedientes no está disponible"
        Case ESE_TIMEOUT_CONEXION
            ManejarErrorExpedienteService = "Timeout al conectar con el servicio de expedientes"
        Case Else
            ManejarErrorExpedienteService = "Error desconocido: " & errorDesc
    End Select
End Function
```

---

## 2. Integración con Sistema RAC

### 2.1 Descripción General

El Sistema RAC (Registro y Control) es el sistema externo donde se registran oficialmente las solicitudes aprobadas. La integración permite:

- Envío automático de solicitudes aprobadas
- Sincronización de estados entre CONDOR y RAC
- Obtención de números de registro oficiales

### 2.2 Flujo de Integración

```
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│ CONDOR          │    │ Servicio RAC    │    │ Sistema RAC     │
│ Estado: Aprobado│───►│ Envío Automático│───►│ Registro Oficial│
└─────────────────┘    └─────────────────┘    └─────────────────┘
         ▲                       │                       │
         │                       ▼                       │
         └─────────────── Confirmación ◄─────────────────┘
                        + Número RAC
```

### 2.3 Interfaz de Integración RAC

```vba
' Archivo: IRACService.cls
' Descripción: Interfaz para integración con Sistema RAC

Option Compare Database
Option Explicit

' Enviar solicitud al Sistema RAC
Public Function EnviarSolicitudRAC(solicitudId As String) As T_RespuestaRAC
End Function

' Consultar estado en RAC
Public Function ConsultarEstadoRAC(numeroRAC As String) As String
End Function

' Obtener número de registro RAC
Public Function ObtenerNumeroRAC(solicitudId As String) As String
End Function
```

### 2.4 Estructura T_RespuestaRAC

```vba
' Archivo: T_RespuestaRAC.cls
' Descripción: Respuesta del Sistema RAC

Option Compare Database
Option Explicit

Public Exitoso As Boolean
Public NumeroRAC As String
Public FechaRegistro As Date
Public EstadoRAC As String
Public MensajeError As String
Public CodigoError As Long

Public Function EsExitoso() As Boolean
    EsExitoso = (Exitoso And Len(NumeroRAC) > 0)
End Function
```

### 2.5 Configuración RAC

```sql
-- Configuración para integración RAC
INSERT INTO TbConfiguracion (Clave, Valor, Descripcion, Categoria) VALUES
('RAC_SERVICE_URL', 'https://rac.gobierno.es/api/v2', 'URL del servicio RAC', 'RAC'),
('RAC_API_KEY', '[ENCRYPTED]', 'Clave API para RAC', 'RAC'),
('RAC_TIMEOUT', '60', 'Timeout para envíos a RAC (segundos)', 'RAC'),
('RAC_AUTO_SEND', 'true', 'Envío automático al aprobar solicitudes', 'RAC'),
('RAC_RETRY_ATTEMPTS', '5', 'Intentos de reenvío en caso de fallo', 'RAC');
```

### 2.6 Estados de Sincronización

| Estado CONDOR | Estado RAC | Acción Requerida |
|---------------|------------|------------------|
| Borrador | - | Ninguna |
| En Revisión | - | Ninguna |
| Aprobado | Pendiente | Envío automático |
| Enviado | En Proceso | Monitoreo |
| Cerrado | Registrado | Sincronización completa |
| Cancelado | Cancelado | Notificación RAC |

---

## 3. Sistema de Lanzadera

### 3.1 Descripción General

El Sistema de Lanzadera gestiona el despliegue y actualización automática de CONDOR. Componentes principales:

- **Lanzadera_Datos.accdb**: Base de datos de control de versiones y usuarios
- **condor_cli.vbs**: Herramienta de línea de comandos para operaciones
- **Sistema de actualización automática**

### 3.2 Arquitectura de Despliegue

```
┌─────────────────────┐    ┌─────────────────────┐    ┌─────────────────────┐
│ Servidor Central    │    │ Lanzadera_Datos     │    │ Cliente Local       │
│ - CONDOR.accde      │◄──►│ - Control Versiones │◄──►│ - CONDOR_Local.accde│
│ - CONDOR_datos.accdb│    │ - Gestión Usuarios  │    │ - Actualización Auto│
└─────────────────────┘    └─────────────────────┘    └─────────────────────┘
```

### 3.3 Estructura Lanzadera_Datos.accdb

#### Tabla TbVersiones
```sql
CREATE TABLE TbVersiones (
    ID AUTOINCREMENT PRIMARY KEY,
    NumeroVersion VARCHAR(20) NOT NULL,
    FechaPublicacion DATETIME NOT NULL,
    RutaArchivo VARCHAR(255) NOT NULL,
    Descripcion MEMO,
    EsObligatoria YESNO DEFAULT FALSE,
    EstadoVersion VARCHAR(20) DEFAULT 'Activa',
    CreadoPor VARCHAR(100),
    FechaCreacion DATETIME DEFAULT NOW()
);
```

#### Tabla TbUsuariosAplicaciones
```sql
CREATE TABLE TbUsuariosAplicaciones (
    ID AUTOINCREMENT PRIMARY KEY,
    NombreUsuario VARCHAR(100) NOT NULL,
    EmailUsuario VARCHAR(150) NOT NULL,
    RolUsuario VARCHAR(50) NOT NULL, -- 'Calidad', 'Ingenieria', 'Administrador'
    EstadoUsuario VARCHAR(20) DEFAULT 'Activo',
    VersionInstalada VARCHAR(20),
    UltimoAcceso DATETIME,
    FechaRegistro DATETIME DEFAULT NOW()
);
```

### 3.4 Proceso de Actualización Automática

#### Algoritmo de Verificación
```vba
' Archivo: ModuloActualizacion.bas
' Función principal de verificación de actualizaciones

Public Function VerificarActualizacion() As Boolean
    Dim versionLocal As String
    Dim versionServidor As String
    Dim esObligatoria As Boolean
    
    ' Obtener versión local
    versionLocal = ObtenerVersionLocal()
    
    ' Consultar versión en servidor
    versionServidor = ConsultarVersionServidor()
    
    ' Comparar versiones
    If CompararVersiones(versionLocal, versionServidor) < 0 Then
        esObligatoria = EsActualizacionObligatoria(versionServidor)
        
        If esObligatoria Then
            ' Forzar actualización
            EjecutarActualizacion versionServidor, True
        Else
            ' Preguntar al usuario
            If MsgBox("Nueva versión disponible. ¿Desea actualizar?", vbYesNo) = vbYes Then
                EjecutarActualizacion versionServidor, False
            End If
        End If
        
        VerificarActualizacion = True
    Else
        VerificarActualizacion = False
    End If
End Function
```

### 3.5 Herramienta condor_cli.vbs

#### Funcionalidades Principales
```vbscript
' Archivo: condor_cli.vbs
' Herramienta de línea de comandos para CONDOR

' Comandos disponibles:
' condor_cli.vbs compile [modulo] - Compilar módulos VBA
' condor_cli.vbs deploy [version] - Desplegar nueva versión
' condor_cli.vbs backup - Crear respaldo de base de datos
' condor_cli.vbs test [suite] - Ejecutar pruebas
' condor_cli.vbs users [accion] - Gestionar usuarios

Sub Main()
    Dim args
    args = WScript.Arguments
    
    If args.Count = 0 Then
        MostrarAyuda
        Exit Sub
    End If
    
    Select Case LCase(args(0))
        Case "compile"
            EjecutarCompilacion args
        Case "deploy"
            EjecutarDespliegue args
        Case "backup"
            EjecutarBackup
        Case "test"
            EjecutarPruebas args
        Case "users"
            GestionarUsuarios args
        Case Else
            WScript.Echo "Comando no reconocido: " & args(0)
            MostrarAyuda
    End Select
End Sub
```

### 3.6 Gestión de Entornos

#### Configuración por Entorno
```vba
' Detección automática de entorno
Public Function DetectarEntorno() As String
    Dim rutaAplicacion As String
    rutaAplicacion = CurrentProject.Path
    
    If InStr(rutaAplicacion, "\\servidor\condor") > 0 Then
        DetectarEntorno = "REMOTO"
    Else
        DetectarEntorno = "LOCAL"
    End If
End Function

' Configuración específica por entorno
Public Sub ConfigurarEntorno()
    Dim entorno As String
    entorno = DetectarEntorno()
    
    Select Case entorno
        Case "LOCAL"
            ' Configuración para desarrollo local
            SetConfigValue "DEV_MODE", "true"
            SetConfigValue "LOG_LEVEL", "DEBUG"
            SetConfigValue "AUTO_UPDATE", "false"
            
        Case "REMOTO"
            ' Configuración para producción
            SetConfigValue "DEV_MODE", "false"
            SetConfigValue "LOG_LEVEL", "INFO"
            SetConfigValue "AUTO_UPDATE", "true"
    End Select
End Sub
```

---

## 4. Protocolos de Comunicación

### 4.1 Formato de Mensajes

Todos los servicios utilizan formato JSON para intercambio de datos:

```json
{
  "version": "1.0",
  "timestamp": "2024-12-20T10:30:00Z",
  "source": "CONDOR",
  "target": "ExpedienteService",
  "operation": "ObtenerExpediente",
  "data": {
    "numeroExpediente": "EXP-2024-INF-001"
  },
  "metadata": {
    "usuario": "maria.garcia@empresa.com",
    "sessionId": "sess_123456789"
  }
}
```

### 4.2 Códigos de Respuesta Estándar

| Código | Descripción | Acción Recomendada |
|--------|-------------|-------------------|
| 200 | Éxito | Procesar respuesta |
| 400 | Solicitud inválida | Validar parámetros |
| 401 | No autorizado | Renovar autenticación |
| 404 | Recurso no encontrado | Verificar identificador |
| 500 | Error interno | Reintentar más tarde |
| 503 | Servicio no disponible | Activar modo offline |

### 4.3 Autenticación y Seguridad

#### Tokens de Acceso
```vba
' Gestión de tokens de autenticación
Public Function ObtenerTokenAcceso(servicio As String) As String
    Dim token As String
    
    ' Verificar si el token existe y es válido
    token = GetConfigValue(servicio & "_TOKEN")
    
    If Len(token) = 0 Or TokenExpirado(token) Then
        ' Renovar token
        token = RenovarToken(servicio)
        SetConfigValue servicio & "_TOKEN", token
    End If
    
    ObtenerTokenAcceso = token
End Function
```

---

## 5. Monitoreo y Logging

### 5.1 Registro de Integraciones

```sql
-- Tabla para logging de integraciones
CREATE TABLE TbLog_Integraciones (
    ID AUTOINCREMENT PRIMARY KEY,
    FechaHora DATETIME DEFAULT NOW(),
    Servicio VARCHAR(50) NOT NULL,
    Operacion VARCHAR(100) NOT NULL,
    Parametros MEMO,
    Resultado VARCHAR(20), -- 'EXITO', 'ERROR', 'TIMEOUT'
    TiempoRespuesta INTEGER, -- en milisegundos
    MensajeError MEMO,
    Usuario VARCHAR(100),
    DireccionIP VARCHAR(15)
);
```

### 5.2 Métricas de Rendimiento

```vba
' Función para registrar métricas de integración
Public Sub RegistrarMetricaIntegracion(servicio As String, operacion As String, _
                                      tiempoMs As Long, resultado As String, _
                                      Optional mensajeError As String = "")
    
    Dim sql As String
    sql = "INSERT INTO TbLog_Integraciones " & _
          "(Servicio, Operacion, TiempoRespuesta, Resultado, MensajeError, Usuario) " & _
          "VALUES ('" & servicio & "', '" & operacion & "', " & tiempoMs & ", " & _
          "'" & resultado & "', '" & mensajeError & "', '" & CurrentUser() & "')"
    
    CurrentDb.Execute sql
End Sub
```

---

## 6. Configuración de Desarrollo

### 6.1 Variables de Entorno

```vba
' Configuración para entorno de desarrollo
Public Sub ConfigurarDesarrollo()
    ' URLs de servicios de desarrollo
    SetConfigValue "EXPEDIENTE_SERVICE_URL", "http://localhost:8080/api"
    SetConfigValue "RAC_SERVICE_URL", "http://test-rac.local/api"
    
    ' Configuración de logging extendido
    SetConfigValue "LOG_LEVEL", "DEBUG"
    SetConfigValue "LOG_INTEGRACIONES", "true"
    
    ' Desactivar actualizaciones automáticas
    SetConfigValue "AUTO_UPDATE", "false"
    
    ' Timeouts extendidos para debugging
    SetConfigValue "EXPEDIENTE_SERVICE_TIMEOUT", "300"
    SetConfigValue "RAC_TIMEOUT", "300"
End Sub
```

### 6.2 Modo Simulación

```vba
' Implementación de servicios mock para desarrollo
Public Function CrearExpedienteServiceMock() As IExpedienteService
    Dim mockService As New CExpedienteServiceMock
    Set CrearExpedienteServiceMock = mockService
End Function

' Clase mock para ExpedienteService
' Archivo: CExpedienteServiceMock.cls
Option Compare Database
Option Explicit
Implements IExpedienteService

Private Function IExpedienteService_ObtenerExpediente(numeroExpediente As String) As T_Expediente
    ' Retorna datos simulados para desarrollo
    Dim expediente As New T_Expediente
    
    expediente.NumeroExpediente = numeroExpediente
    expediente.TituloExpediente = "Expediente de Prueba - " & numeroExpediente
    expediente.ResponsableCalidad = "Usuario Prueba"
    expediente.EmailResponsable = "prueba@test.com"
    
    Set IExpedienteService_ObtenerExpediente = expediente
End Function
```

---

## 7. Troubleshooting

### 7.1 Problemas Comunes

#### ExpedienteService No Responde
```vba
' Diagnóstico de conectividad
Public Function DiagnosticarExpedienteService() As String
    Dim resultado As String
    
    ' Verificar conectividad de red
    If Not PingServidor(GetConfigValue("EXPEDIENTE_SERVICE_URL")) Then
        resultado = "Error: No hay conectividad con el servidor de expedientes"
    ElseIf Not ValidarCredenciales() Then
        resultado = "Error: Credenciales de acceso inválidas"
    Else
        resultado = "Conectividad OK - Verificar logs del servicio"
    End If
    
    DiagnosticarExpedienteService = resultado
End Function
```

#### Fallos en Actualización Automática
```vba
' Recuperación de fallos de actualización
Public Sub RecuperarActualizacionFallida()
    ' Restaurar versión anterior
    Dim versionAnterior As String
    versionAnterior = GetConfigValue("VERSION_ANTERIOR")
    
    If Len(versionAnterior) > 0 Then
        ' Copiar archivos de respaldo
        RestaurarBackup versionAnterior
        
        ' Actualizar registro de versión
        SetConfigValue "VERSION_ACTUAL", versionAnterior
        
        ' Notificar al administrador
        NotificarFalloActualizacion
    End If
End Sub
```

### 7.2 Logs de Diagnóstico

```sql
-- Consulta para análisis de errores de integración
SELECT 
    Servicio,
    COUNT(*) as TotalLlamadas,
    SUM(CASE WHEN Resultado = 'ERROR' THEN 1 ELSE 0 END) as Errores,
    AVG(TiempoRespuesta) as TiempoPromedio,
    MAX(FechaHora) as UltimaLlamada
FROM TbLog_Integraciones 
WHERE FechaHora >= DateAdd('d', -7, Date())
GROUP BY Servicio
ORDER BY Errores DESC;
```

---

*Documento generado según la Especificación Funcional y Arquitectura CONDOR*  
*Versión: 1.0*  
*Fecha: Diciembre 2024*