Option Compare Database
Option Explicit
'******************************************************************************
' CLASE: T_Configuracion
' DESCRIPCIÓN: Tipo de datos para almacenar configuración del sistema
' AUTOR: Sistema CONDOR
' FECHA: 2024
' NOTA: Propiedades definidas basándose en TbConfiguracion y CConfig
'******************************************************************************


'******************************************************************************
' PROPIEDADES PÚBLICAS
' Basadas en la estructura de TbConfiguracion y uso en CConfig
'******************************************************************************

' Identificador único de la configuración
Public ID As Long

' Clave de configuración
Public Clave As String

' Valor de configuración
Public Valor As String

' Descripción de la configuración
Public Descripcion As String

' Categoría de la configuración
Public Categoria As String

' Entorno al que aplica (Desarrollo, Producción)
Public Entorno As String

' Ruta a la base de datos backend
Public RutaBackend As String

' Ruta de datos
Public DataPath As String

' Ruta de expedientes
Public ExpedientesPath As String

' Fecha de creación
Public fechaCreacion As Date

' Indica si está activo
Public Activo As Boolean

'******************************************************************************
' PROPIEDADES ESPECÍFICAS DE INTEGRACIÓN
' Basadas en Especificaciones_Integracion_CONDOR.md
'******************************************************************************

' URL del servicio de expedientes
Public ExpedienteServiceURL As String

' Timeout para servicio de expedientes
Public ExpedienteServiceTimeout As Long

' Número de reintentos para servicio de expedientes
Public ExpedienteServiceRetry As Long

' Tipo de autenticación para servicio de expedientes
Public ExpedienteServiceAuth As String

' URL del servicio RAC
Public RACServiceURL As String

' Clave API para RAC
Public RACAPIKey As String

' Timeout para RAC
Public RACTimeout As Long

' Envío automático a RAC
Public RACAutoSend As Boolean

' Intentos de reenvío RAC
Public RACRetryAttempts As Long

' Nivel de logging
Public LogLevel As String

' Logging de integraciones activado
Public LogIntegraciones As Boolean

' Actualización automática activada
Public AutoUpdate As Boolean

'******************************************************************************
' MÉTODOS PÚBLICOS
'******************************************************************************

' Valida si la configuración tiene datos mínimos requeridos
Public Function EsValida() As Boolean
    EsValida = (Len(Clave) > 0 And _
                Len(Valor) > 0 And _
                Activo = True)
End Function

' Determina si es una configuración de entorno
Public Function EsConfiguracionEntorno() As Boolean
    EsConfiguracionEntorno = (Len(Entorno) > 0 And _
                             Len(RutaBackend) > 0)
End Function

' Determina si es una configuración de integración
Public Function EsConfiguracionIntegracion() As Boolean
    EsConfiguracionIntegracion = (Categoria = "Integración" Or _
                                 Categoria = "RAC" Or _
                                 InStr(Clave, "SERVICE_") > 0)
End Function













