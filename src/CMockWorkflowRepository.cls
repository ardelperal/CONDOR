VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockWorkflowRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

#If DEV_MODE Then

'==============================================================================
' Clase: CMockWorkflowRepository
' Propósito: Mock del repositorio de workflow para pruebas unitarias
'            Implementa IWorkflowRepository devolviendo valores por defecto seguros
' Autor: CONDOR-Expert
' Fecha: 2025-01-21
'==============================================================================

Implements IWorkflowRepository

'==============================================================================
' VARIABLES PRIVADAS PARA CONFIGURAR VALORES DE RETORNO
'==============================================================================

Private m_IsValidTransitionReturnValue As Boolean
Private m_GetAvailableStatesReturnValue As Collection
Private m_GetNextStatesReturnValue As Collection
Private m_GetInitialStateReturnValue As String
Private m_IsStateFinalReturnValue As Boolean
Private m_RecordStateChangeReturnValue As Boolean
Private m_GetStateHistoryReturnValue As Collection
Private m_HasTransitionPermissionReturnValue As Boolean
Private m_RequiresApprovalReturnValue As Boolean
Private m_GetTransitionRequiredRoleReturnValue As String

'==============================================================================
' EVENTOS DE CLASE
'==============================================================================

Private Sub Class_Initialize()
    ' Inicializar valores por defecto seguros
    m_IsValidTransitionReturnValue = True
    Set m_GetAvailableStatesReturnValue = New Collection
    Set m_GetNextStatesReturnValue = New Collection
    m_GetInitialStateReturnValue = "BORRADOR"
    m_IsStateFinalReturnValue = False
    m_RecordStateChangeReturnValue = True
    Set m_GetStateHistoryReturnValue = New Collection
    m_HasTransitionPermissionReturnValue = True
    m_RequiresApprovalReturnValue = False
    m_GetTransitionRequiredRoleReturnValue = ""
End Sub

Private Sub Class_Terminate()
    Set m_GetAvailableStatesReturnValue = Nothing
    Set m_GetNextStatesReturnValue = Nothing
    Set m_GetStateHistoryReturnValue = Nothing
End Sub

'==============================================================================
' MÉTODOS PARA CONFIGURAR VALORES DE RETORNO
'==============================================================================

Public Sub SetIsValidTransitionReturnValue(ByVal value As Boolean)
    m_IsValidTransitionReturnValue = value
End Sub

Public Sub SetGetAvailableStatesReturnValue(ByVal value As Collection)
    Set m_GetAvailableStatesReturnValue = value
End Sub

Public Sub SetGetNextStatesReturnValue(ByVal value As Collection)
    Set m_GetNextStatesReturnValue = value
End Sub

Public Sub SetGetInitialStateReturnValue(ByVal value As String)
    m_GetInitialStateReturnValue = value
End Sub

Public Sub SetIsStateFinalReturnValue(ByVal value As Boolean)
    m_IsStateFinalReturnValue = value
End Sub

Public Sub SetRecordStateChangeReturnValue(ByVal value As Boolean)
    m_RecordStateChangeReturnValue = value
End Sub

Public Sub SetGetStateHistoryReturnValue(ByVal value As Collection)
    Set m_GetStateHistoryReturnValue = value
End Sub

Public Sub SetHasTransitionPermissionReturnValue(ByVal value As Boolean)
    m_HasTransitionPermissionReturnValue = value
End Sub

Public Sub SetRequiresApprovalReturnValue(ByVal value As Boolean)
    m_RequiresApprovalReturnValue = value
End Sub

Public Sub SetGetTransitionRequiredRoleReturnValue(ByVal value As String)
    m_GetTransitionRequiredRoleReturnValue = value
End Sub

'==============================================================================
' MÉTODO PARA REINICIAR EL MOCK
'==============================================================================

Public Sub Reset()
    ' Reiniciar valores de retorno a valores por defecto seguros
    m_IsValidTransitionReturnValue = True
    Set m_GetAvailableStatesReturnValue = New Collection
    Set m_GetNextStatesReturnValue = New Collection
    m_GetInitialStateReturnValue = "BORRADOR"
    m_IsStateFinalReturnValue = False
    m_RecordStateChangeReturnValue = True
    Set m_GetStateHistoryReturnValue = New Collection
    m_HasTransitionPermissionReturnValue = True
    m_RequiresApprovalReturnValue = False
    m_GetTransitionRequiredRoleReturnValue = ""
End Sub

'==============================================================================
' IMPLEMENTACIÓN DE IWorkflowRepository
'==============================================================================

Private Function IWorkflowRepository_IsValidTransition(ByVal tipoSolicitud As String, ByVal estadoOrigen As String, ByVal estadoDestino As String) As Boolean
    IWorkflowRepository_IsValidTransition = m_IsValidTransitionReturnValue
End Function

Private Function IWorkflowRepository_GetAvailableStates(ByVal tipoSolicitud As String) As Collection
    Set IWorkflowRepository_GetAvailableStates = m_GetAvailableStatesReturnValue
End Function

Private Function IWorkflowRepository_GetNextStates(ByVal estadoActual As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Collection
    Set IWorkflowRepository_GetNextStates = m_GetNextStatesReturnValue
End Function

Private Function IWorkflowRepository_GetInitialState(ByVal tipoSolicitud As String) As String
    IWorkflowRepository_GetInitialState = m_GetInitialStateReturnValue
End Function

Private Function IWorkflowRepository_IsStateFinal(ByVal estadoCodigo As String, ByVal tipoSolicitud As String) As Boolean
    IWorkflowRepository_IsStateFinal = m_IsStateFinalReturnValue
End Function

Private Function IWorkflowRepository_RecordStateChange(ByVal SolicitudID As Long, ByVal estadoAnterior As String, ByVal estadoNuevo As String, ByVal usuario As String, ByVal comentarios As String) As Boolean
    IWorkflowRepository_RecordStateChange = m_RecordStateChangeReturnValue
End Function

Private Function IWorkflowRepository_GetStateHistory(ByVal SolicitudID As Long) As Collection
    Set IWorkflowRepository_GetStateHistory = m_GetStateHistoryReturnValue
End Function

Private Function IWorkflowRepository_HasTransitionPermission(ByVal usuarioRol As String, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As Boolean
    IWorkflowRepository_HasTransitionPermission = m_HasTransitionPermissionReturnValue
End Function

Private Function IWorkflowRepository_RequiresApproval(ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As Boolean
    IWorkflowRepository_RequiresApproval = m_RequiresApprovalReturnValue
End Function

Private Function IWorkflowRepository_GetTransitionRequiredRole(ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As String
    IWorkflowRepository_GetTransitionRequiredRole = m_GetTransitionRequiredRoleReturnValue
End Function

#End If