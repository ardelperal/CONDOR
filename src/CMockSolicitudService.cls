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
Private m_LastExpedienteId As String
Private m_LastTipoSolicitud As String
Private m_LastSavedSolicitud As T_Solicitud

' Propiedades para configurar valores de retorno
Private m_CreateSolicitudReturnValue As T_Solicitud
Private m_SaveSolicitudReturnValue As Boolean

' Contadores de llamadas
Private m_CreateSolicitudCallCount As Long
Private m_SaveSolicitudCallCount As Long

' Propiedades públicas para verificación en tests
Public Property Get CreateSolicitud_WasCalled() As Boolean
    CreateSolicitud_WasCalled = m_CreateSolicitud_WasCalled
End Property

Public Property Get SaveSolicitud_WasCalled() As Boolean
    SaveSolicitud_WasCalled = m_SaveSolicitud_WasCalled
End Property

Public Property Get LastExpedienteId() As String
    LastExpedienteId = m_LastExpedienteId
End Property

Public Property Get LastTipoSolicitud() As String
    LastTipoSolicitud = m_LastTipoSolicitud
End Property

Public Property Get LastSavedSolicitud() As T_Solicitud
    Set LastSavedSolicitud = m_LastSavedSolicitud
End Property

Public Property Get CreateSolicitudCallCount() As Long
    CreateSolicitudCallCount = m_CreateSolicitudCallCount
End Property

Public Property Get SaveSolicitudCallCount() As Long
    SaveSolicitudCallCount = m_SaveSolicitudCallCount
End Property

' Métodos para configurar valores de retorno
Public Sub SetCreateSolicitudReturnValue(ByRef solicitud As T_Solicitud)
    Set m_CreateSolicitudReturnValue = solicitud
End Sub

Public Sub SetSaveSolicitudReturnValue(ByVal success As Boolean)
    m_SaveSolicitudReturnValue = success
End Sub

' Método para resetear el estado del mock
Public Sub Reset()
    m_CreateSolicitud_WasCalled = False
    m_SaveSolicitud_WasCalled = False
    m_LastExpedienteId = ""
    m_LastTipoSolicitud = ""
    Set m_LastSavedSolicitud = Nothing
    Set m_CreateSolicitudReturnValue = Nothing
    m_SaveSolicitudReturnValue = False
    m_CreateSolicitudCallCount = 0
    m_SaveSolicitudCallCount = 0
End Sub

' Implementación de ISolicitudService
Private Function ISolicitudService_CreateSolicitud(ByVal idExpediente As String, ByVal tipo As String) As T_Solicitud
    m_CreateSolicitud_WasCalled = True
    m_LastExpedienteId = idExpediente
    m_LastTipoSolicitud = tipo
    m_CreateSolicitudCallCount = m_CreateSolicitudCallCount + 1
    
    Set ISolicitudService_CreateSolicitud = m_CreateSolicitudReturnValue
End Function

Private Function ISolicitudService_SaveSolicitud(ByVal solicitud As T_Solicitud) As Boolean
    m_SaveSolicitud_WasCalled = True
    Set m_LastSavedSolicitud = solicitud
    m_SaveSolicitudCallCount = m_SaveSolicitudCallCount + 1
    
    ISolicitudService_SaveSolicitud = m_SaveSolicitudReturnValue
End Function

' Métodos públicos de conveniencia
Public Function CreateSolicitud(ByVal idExpediente As String, ByVal tipo As String) As T_Solicitud
    Set CreateSolicitud = ISolicitudService_CreateSolicitud(idExpediente, tipo)
End Function

Public Function SaveSolicitud(ByVal solicitud As T_Solicitud) As Boolean
    SaveSolicitud = ISolicitudService_SaveSolicitud(solicitud)
End Function