VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "T_Configuracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' CLASE: T_Configuracion
' DESCRIPCI├ôN: Tipo de datos para almacenar configuraci├│n del sistema
' AUTOR: Sistema CONDOR
' FECHA: 2024
' NOTA: Propiedades definidas bas├índose en TbConfiguracion y CConfig
'******************************************************************************

Option Compare Database
Option Explicit

'******************************************************************************
' PROPIEDADES P├ÜBLICAS
' Basadas en la estructura de TbConfiguracion y uso en CConfig
'******************************************************************************

' Identificador ├║nico de la configuraci├│n
Public ID As Long

' Clave de configuraci├│n
Public Clave As String

' Valor de configuraci├│n
Public Valor As String

' Descripci├│n de la configuraci├│n
Public Descripcion As String

' Categor├¡a de la configuraci├│n
Public Categoria As String

' Entorno al que aplica (Desarrollo, Producci├│n)
Public Entorno As String

' Ruta a la base de datos backend
Public RutaBackend As String

' Ruta de datos
Public DataPath As String

' Ruta de expedientes
Public ExpedientesPath As String

' Fecha de creaci├│n
Public FechaCreacion As Date

' Indica si est├í activo
Public Activo As Boolean

'******************************************************************************
' PROPIEDADES ESPEC├ìFICAS DE INTEGRACI├ôN
' Basadas en Especificaciones_Integracion_CONDOR.md
'******************************************************************************

' URL del servicio de expedientes
Public ExpedienteServiceURL As String

' Timeout para servicio de expedientes
Public ExpedienteServiceTimeout As Long

' N├║mero de reintentos para servicio de expedientes
Public ExpedienteServiceRetry As Long

' Tipo de autenticaci├│n para servicio de expedientes
Public ExpedienteServiceAuth As String

' URL del servicio RAC
Public RACServiceURL As String

' Clave API para RAC
Public RACAPIKey As String

' Timeout para RAC
Public RACTimeout As Long

' Env├¡o autom├ítico a RAC
Public RACAutoSend As Boolean

' Intentos de reenv├¡o RAC
Public RACRetryAttempts As Long

' Nivel de logging
Public LogLevel As String

' Logging de integraciones activado
Public LogIntegraciones As Boolean

' Actualizaci├│n autom├ítica activada
Public AutoUpdate As Boolean

'******************************************************************************
' M├ëTODOS P├ÜBLICOS
'******************************************************************************

' Valida si la configuraci├│n tiene datos m├¡nimos requeridos
Public Function EsValida() As Boolean
    EsValida = (Len(Clave) > 0 And _
                Len(Valor) > 0 And _
                Activo = True)
End Function

' Determina si es una configuraci├│n de entorno
Public Function EsConfiguracionEntorno() As Boolean
    EsConfiguracionEntorno = (Len(Entorno) > 0 And _
                             Len(RutaBackend) > 0)
End Function

' Determina si es una configuraci├│n de integraci├│n
Public Function EsConfiguracionIntegracion() As Boolean
    EsConfiguracionIntegracion = (Categoria = "Integraci├│n" Or _
                                 Categoria = "RAC" Or _
                                 InStr(Clave, "SERVICE_") > 0)
End Function
