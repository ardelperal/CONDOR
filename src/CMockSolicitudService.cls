VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockSolicitudService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
' CONDOR - Sistema de Gestión de Expedientes
' Mock del Servicio de Solicitudes para Testing
' Autor: CONDOR-Expert
' Fecha: 2025
'*******************************************************************************

Option Explicit

Implements ISolicitudService

' Propiedades para verificar llamadas
Private m_CreateSolicitud_WasCalled As Boolean
Private m_SaveSolicitud_WasCalled As Boolean

' Propiedades para capturar parámetros
Private m_LastExpediente As EExpediente
Private m_LastTipoSolicitud As String
Private m_LastSavedSolicitud As ESolicitud

' Propiedades para configurar valores de retorno
Private m_CreateSolicitudReturnValue As ESolicitud
Private m_SaveSolicitudReturnValue As Long

' Contadores de llamadas
Private m_CreateSolicitudCallCount As Long
Private m_SaveSolicitudCallCount As Long

' Constructor
Private Sub Class_Initialize()
    Reset
End Sub

' Propiedades públicas para verificación en tests
Public Property Get CreateSolicitud_WasCalled() As Boolean
    CreateSolicitud_WasCalled = m_CreateSolicitud_WasCalled
End Property

Public Property Get SaveSolicitud_WasCalled() As Boolean
    SaveSolicitud_WasCalled = m_SaveSolicitud_WasCalled
End Property

Public Property Get LastExpediente() As EExpediente
    Set LastExpediente = m_LastExpediente
End Property

Public Property Get LastTipoSolicitud() As String
    LastTipoSolicitud = m_LastTipoSolicitud
End Property

Public Property Get LastSavedSolicitud() As ESolicitud
    Set LastSavedSolicitud = m_LastSavedSolicitud
End Property

Public Property Get CreateSolicitudCallCount() As Long
    CreateSolicitudCallCount = m_CreateSolicitudCallCount
End Property

Public Property Get SaveSolicitudCallCount() As Long
    SaveSolicitudCallCount = m_SaveSolicitudCallCount
End Property

' Métodos para configurar valores de retorno
Public Sub ConfigureCreateSolicitud(ByRef solicitud As ESolicitud)
    Set m_CreateSolicitudReturnValue = solicitud
End Sub

Public Sub ConfigureSaveSolicitud(ByVal returnId As Long)
    m_SaveSolicitudReturnValue = returnId
End Sub

' Método para resetear el estado del mock
Public Sub Reset()
    m_CreateSolicitud_WasCalled = False
    m_SaveSolicitud_WasCalled = False
    Set m_LastExpediente = Nothing
    m_LastTipoSolicitud = ""
    Set m_LastSavedSolicitud = Nothing
    Set m_CreateSolicitudReturnValue = Nothing
    m_SaveSolicitudReturnValue = 0
    m_CreateSolicitudCallCount = 0
    m_SaveSolicitudCallCount = 0
End Sub

' Implementación de ISolicitudService
Private Function ISolicitudService_CreateSolicitud(ByVal paraExpediente As EExpediente) As ESolicitud
    m_CreateSolicitud_WasCalled = True
    Set m_LastExpediente = paraExpediente
    m_LastTipoSolicitud = "PC" ' Tipo por defecto
    m_CreateSolicitudCallCount = m_CreateSolicitudCallCount + 1
    
    Set ISolicitudService_CreateSolicitud = m_CreateSolicitudReturnValue
End Function

Private Function ISolicitudService_SaveSolicitud(ByVal solicitud As ESolicitud) As Long
    m_SaveSolicitud_WasCalled = True
    Set m_LastSavedSolicitud = solicitud
    m_SaveSolicitudCallCount = m_SaveSolicitudCallCount + 1
    
    ISolicitudService_SaveSolicitud = m_SaveSolicitudReturnValue
End Function

' Métodos públicos de conveniencia
Public Function CreateSolicitud(ByVal paraExpediente As EExpediente) As ESolicitud
    Set CreateSolicitud = ISolicitudService_CreateSolicitud(paraExpediente)
End Function

Public Function SaveSolicitud(ByVal solicitud As ESolicitud) As Long
    SaveSolicitud = ISolicitudService_SaveSolicitud(solicitud)
End Function