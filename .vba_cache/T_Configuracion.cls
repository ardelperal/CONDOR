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
' VARIABLES MIEMBRO PRIVADAS
'******************************************************************************
Private m_ID As Long
Private m_Clave As String
Private m_Valor As String
Private m_Descripcion As String
Private m_Categoria As String
Private m_Entorno As String
Private m_RutaBackend As String
Private m_DataPath As String
Private m_ExpedientesPath As String
Private m_fechaCreacion As Date
Private m_Activo As Boolean
Private m_ExpedienteServiceURL As String
Private m_LogLevel As String

'******************************************************************************
' PROPIEDADES PÚBLICAS (GET/LET)
'******************************************************************************

Public Property Get ID() As Long
    ID = m_ID
End Property
Public Property Let ID(ByVal value As Long)
    m_ID = value
End Property

Public Property Get Clave() As String
    Clave = m_Clave
End Property
Public Property Let Clave(ByVal value As String)
    m_Clave = value
End Property

Public Property Get Valor() As String
    Valor = m_Valor
End Property
Public Property Let Valor(ByVal value As String)
    m_Valor = value
End Property

Public Property Get Descripcion() As String
    Descripcion = m_Descripcion
End Property
Public Property Let Descripcion(ByVal value As String)
    m_Descripcion = value
End Property

Public Property Get Categoria() As String
    Categoria = m_Categoria
End Property
Public Property Let Categoria(ByVal value As String)
    m_Categoria = value
End Property

Public Property Get Entorno() As String
    Entorno = m_Entorno
End Property
Public Property Let Entorno(ByVal value As String)
    m_Entorno = value
End Property

Public Property Get RutaBackend() As String
    RutaBackend = m_RutaBackend
End Property
Public Property Let RutaBackend(ByVal value As String)
    m_RutaBackend = value
End Property

Public Property Get DataPath() As String
    DataPath = m_DataPath
End Property
Public Property Let DataPath(ByVal value As String)
    m_DataPath = value
End Property

Public Property Get ExpedientesPath() As String
    ExpedientesPath = m_ExpedientesPath
End Property
Public Property Let ExpedientesPath(ByVal value As String)
    m_ExpedientesPath = value
End Property

Public Property Get fechaCreacion() As Date
    fechaCreacion = m_fechaCreacion
End Property
Public Property Let fechaCreacion(ByVal value As Date)
    m_fechaCreacion = value
End Property

Public Property Get Activo() As Boolean
    Activo = m_Activo
End Property
Public Property Let Activo(ByVal value As Boolean)
    m_Activo = value
End Property

Public Property Get ExpedienteServiceURL() As String
    ExpedienteServiceURL = m_ExpedienteServiceURL
End Property
Public Property Let ExpedienteServiceURL(ByVal value As String)
    m_ExpedienteServiceURL = value
End Property



Public Property Get LogLevel() As String
    LogLevel = m_LogLevel
End Property
Public Property Let LogLevel(ByVal value As String)
    m_LogLevel = value
End Property



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
                                 InStr(Clave, "SERVICE_") > 0)
End Function













