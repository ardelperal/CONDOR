VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockWorkflowService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


#If DEV_MODE Then

'==============================================================================
' Clase: CMockWorkflowService
' Propósito: Mock del servicio de workflow para pruebas unitarias
'            Implementa IWorkflowService permitiendo configurar comportamientos
' Autor: CONDOR-Expert
' Fecha: 2025-01-21
'==============================================================================


Implements IWorkflowService

'==============================================================================
' VARIABLES PRIVADAS PARA VERIFICACIÓN DE LLAMADAS
'==============================================================================

Private m_ValidateTransition_WasCalled As Boolean
Private m_GetAvailableStates_WasCalled As Boolean
Private m_GetNextStates_WasCalled As Boolean
Private m_GetInitialState_WasCalled As Boolean
Private m_IsStateFinal_WasCalled As Boolean
Private m_RecordStateChange_WasCalled As Boolean
Private m_GetStateHistory_WasCalled As Boolean
Private m_HasTransitionPermission_WasCalled As Boolean
Private m_RequiresApproval_WasCalled As Boolean

'==============================================================================
' VARIABLES PRIVADAS PARA CAPTURAR PARÁMETROS
'==============================================================================

Private m_LastSolicitudID As Long
Private m_LastEstadoOrigen As String
Private m_LastEstadoDestino As String
Private m_LastTipoSolicitud As String
Private m_LastUsuarioRol As String
Private m_LastEstadoAnterior As String
Private m_LastEstadoNuevo As String
Private m_LastUsuario As String
Private m_LastComentarios As String
Private m_LastEstadoActual As String
Private m_LastEstadoCodigo As String

'==============================================================================
' VARIABLES PRIVADAS PARA CONFIGURAR VALORES DE RETORNO
'==============================================================================

Private m_ValidateTransitionReturnValue As Boolean
Private m_GetAvailableStatesReturnValue As Scripting.Dictionary
Private m_GetNextStatesReturnValue As Scripting.Dictionary
Private m_GetInitialStateReturnValue As String
Private m_IsStateFinalReturnValue As Boolean
Private m_RecordStateChangeReturnValue As Boolean
Private m_GetStateHistoryReturnValue As Scripting.Dictionary
Private m_HasTransitionPermissionReturnValue As Boolean
Private m_RequiresApprovalReturnValue As Boolean

'==============================================================================
' VARIABLES PRIVADAS PARA CONTAR LLAMADAS
'==============================================================================

Private m_ValidateTransitionCallCount As Long
Private m_GetAvailableStatesCallCount As Long
Private m_GetNextStatesCallCount As Long
Private m_GetInitialStateCallCount As Long
Private m_IsStateFinalCallCount As Long
Private m_RecordStateChangeCallCount As Long
Private m_GetStateHistoryCallCount As Long
Private m_HasTransitionPermissionCallCount As Long
Private m_RequiresApprovalCallCount As Long

'==============================================================================
' EVENTOS DE CLASE
'==============================================================================

Private Sub Class_Initialize()
    Reset
End Sub

Private Sub Class_Terminate()
    Set m_GetAvailableStatesReturnValue = Nothing
    Set m_GetNextStatesReturnValue = Nothing
    Set m_GetStateHistoryReturnValue = Nothing
End Sub

'==============================================================================
' PROPIEDADES PARA VERIFICAR LLAMADAS
'==============================================================================

Public Property Get ValidateTransition_WasCalled() As Boolean
    ValidateTransition_WasCalled = m_ValidateTransition_WasCalled
End Property

Public Property Get GetAvailableStates_WasCalled() As Boolean
    GetAvailableStates_WasCalled = m_GetAvailableStates_WasCalled
End Property

Public Property Get GetNextStates_WasCalled() As Boolean
    GetNextStates_WasCalled = m_GetNextStates_WasCalled
End Property

Public Property Get GetInitialState_WasCalled() As Boolean
    GetInitialState_WasCalled = m_GetInitialState_WasCalled
End Property

Public Property Get IsStateFinal_WasCalled() As Boolean
    IsStateFinal_WasCalled = m_IsStateFinal_WasCalled
End Property

Public Property Get RecordStateChange_WasCalled() As Boolean
    RecordStateChange_WasCalled = m_RecordStateChange_WasCalled
End Property

Public Property Get GetStateHistory_WasCalled() As Boolean
    GetStateHistory_WasCalled = m_GetStateHistory_WasCalled
End Property

Public Property Get HasTransitionPermission_WasCalled() As Boolean
    HasTransitionPermission_WasCalled = m_HasTransitionPermission_WasCalled
End Property

Public Property Get RequiresApproval_WasCalled() As Boolean
    RequiresApproval_WasCalled = m_RequiresApproval_WasCalled
End Property

'==============================================================================
' PROPIEDADES PARA CAPTURAR PARÁMETROS
'==============================================================================

Public Property Get LastSolicitudID() As Long
    LastSolicitudID = m_LastSolicitudID
End Property

Public Property Get LastEstadoOrigen() As String
    LastEstadoOrigen = m_LastEstadoOrigen
End Property

Public Property Get LastEstadoDestino() As String
    LastEstadoDestino = m_LastEstadoDestino
End Property

Public Property Get LastTipoSolicitud() As String
    LastTipoSolicitud = m_LastTipoSolicitud
End Property

Public Property Get LastUsuarioRol() As String
    LastUsuarioRol = m_LastUsuarioRol
End Property

Public Property Get LastEstadoAnterior() As String
    LastEstadoAnterior = m_LastEstadoAnterior
End Property

Public Property Get LastEstadoNuevo() As String
    LastEstadoNuevo = m_LastEstadoNuevo
End Property

Public Property Get LastUsuario() As String
    LastUsuario = m_LastUsuario
End Property

Public Property Get LastComentarios() As String
    LastComentarios = m_LastComentarios
End Property

Public Property Get LastEstadoActual() As String
    LastEstadoActual = m_LastEstadoActual
End Property

Public Property Get LastEstadoCodigo() As String
    LastEstadoCodigo = m_LastEstadoCodigo
End Property

'==============================================================================
' PROPIEDADES PARA CONTAR LLAMADAS
'==============================================================================

Public Property Get ValidateTransitionCallCount() As Long
    ValidateTransitionCallCount = m_ValidateTransitionCallCount
End Property

Public Property Get GetAvailableStatesCallCount() As Long
    GetAvailableStatesCallCount = m_GetAvailableStatesCallCount
End Property

Public Property Get GetNextStatesCallCount() As Long
    GetNextStatesCallCount = m_GetNextStatesCallCount
End Property

Public Property Get GetInitialStateCallCount() As Long
    GetInitialStateCallCount = m_GetInitialStateCallCount
End Property

Public Property Get IsStateFinalCallCount() As Long
    IsStateFinalCallCount = m_IsStateFinalCallCount
End Property

Public Property Get RecordStateChangeCallCount() As Long
    RecordStateChangeCallCount = m_RecordStateChangeCallCount
End Property

Public Property Get GetStateHistoryCallCount() As Long
    GetStateHistoryCallCount = m_GetStateHistoryCallCount
End Property

Public Property Get HasTransitionPermissionCallCount() As Long
    HasTransitionPermissionCallCount = m_HasTransitionPermissionCallCount
End Property

Public Property Get RequiresApprovalCallCount() As Long
    RequiresApprovalCallCount = m_RequiresApprovalCallCount
End Property

'==============================================================================
' MÉTODOS PARA CONFIGURAR VALORES DE RETORNO
'==============================================================================

Public Sub ConfigureValidateTransition(ByVal value As Boolean)
    m_ValidateTransitionReturnValue = value
End Sub

Public Sub ConfigureGetAvailableStates(ByVal value As Scripting.Dictionary)
    Set m_GetAvailableStatesReturnValue = value
End Sub

Public Sub ConfigureGetNextStates(ByVal value As Scripting.Dictionary)
    Set m_GetNextStatesReturnValue = value
End Sub

Public Sub ConfigureGetInitialState(ByVal value As String)
    m_GetInitialStateReturnValue = value
End Sub

Public Sub ConfigureIsStateFinal(ByVal value As Boolean)
    m_IsStateFinalReturnValue = value
End Sub

Public Sub ConfigureRecordStateChange(ByVal value As Boolean)
    m_RecordStateChangeReturnValue = value
End Sub

Public Sub ConfigureGetStateHistory(ByVal value As Scripting.Dictionary)
    Set m_GetStateHistoryReturnValue = value
End Sub

Public Sub ConfigureHasTransitionPermission(ByVal value As Boolean)
    m_HasTransitionPermissionReturnValue = value
End Sub

Public Sub ConfigureRequiresApproval(ByVal value As Boolean)
    m_RequiresApprovalReturnValue = value
End Sub

'==============================================================================
' MÉTODO PARA REINICIAR EL MOCK
'==============================================================================

Public Sub Reset()
    ' Reiniciar flags de llamadas
    m_ValidateTransition_WasCalled = False
    m_GetAvailableStates_WasCalled = False
    m_GetNextStates_WasCalled = False
    m_GetInitialState_WasCalled = False
    m_IsStateFinal_WasCalled = False
    m_RecordStateChange_WasCalled = False
    m_GetStateHistory_WasCalled = False
    m_HasTransitionPermission_WasCalled = False
    m_RequiresApproval_WasCalled = False
    
    ' Reiniciar parámetros capturados
    m_LastSolicitudID = 0
    m_LastEstadoOrigen = ""
    m_LastEstadoDestino = ""
    m_LastTipoSolicitud = ""
    m_LastUsuarioRol = ""
    m_LastEstadoAnterior = ""
    m_LastEstadoNuevo = ""
    m_LastUsuario = ""
    m_LastComentarios = ""
    m_LastEstadoActual = ""
    m_LastEstadoCodigo = ""
    
    ' Reiniciar contadores
    m_ValidateTransitionCallCount = 0
    m_GetAvailableStatesCallCount = 0
    m_GetNextStatesCallCount = 0
    m_GetInitialStateCallCount = 0
    m_IsStateFinalCallCount = 0
    m_RecordStateChangeCallCount = 0
    m_GetStateHistoryCallCount = 0
    m_HasTransitionPermissionCallCount = 0
    m_RequiresApprovalCallCount = 0
    
    ' Reiniciar valores de retorno a valores por defecto
    m_ValidateTransitionReturnValue = True
    Set m_GetAvailableStatesReturnValue = New Scripting.Dictionary
    m_GetAvailableStatesReturnValue.CompareMode = TextCompare
    Set m_GetNextStatesReturnValue = New Scripting.Dictionary
    m_GetNextStatesReturnValue.CompareMode = TextCompare
    m_GetInitialStateReturnValue = "BORRADOR"
    m_IsStateFinalReturnValue = False
    m_RecordStateChangeReturnValue = True
    Set m_GetStateHistoryReturnValue = New Scripting.Dictionary
    m_GetStateHistoryReturnValue.CompareMode = TextCompare
    m_HasTransitionPermissionReturnValue = True
    m_RequiresApprovalReturnValue = False
End Sub

'==============================================================================
' IMPLEMENTACIÓN DE IWorkflowService
'==============================================================================

Private Function IWorkflowService_ValidateTransition(ByVal SolicitudID As Long, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Boolean
    m_ValidateTransition_WasCalled = True
    m_ValidateTransitionCallCount = m_ValidateTransitionCallCount + 1
    
    ' Capturar parámetros
    m_LastSolicitudID = SolicitudID
    m_LastEstadoOrigen = estadoOrigen
    m_LastEstadoDestino = estadoDestino
    m_LastTipoSolicitud = tipoSolicitud
    m_LastUsuarioRol = usuarioRol
    
    IWorkflowService_ValidateTransition = m_ValidateTransitionReturnValue
End Function

Private Function IWorkflowService_GetAvailableStates(ByVal tipoSolicitud As String) As Scripting.Dictionary
    m_GetAvailableStates_WasCalled = True
    m_GetAvailableStatesCallCount = m_GetAvailableStatesCallCount + 1
    
    ' Capturar parámetros
    m_LastTipoSolicitud = tipoSolicitud
    
    Set IWorkflowService_GetAvailableStates = m_GetAvailableStatesReturnValue
End Function

Private Function IWorkflowService_GetNextStates(ByVal estadoActual As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Scripting.Dictionary
    m_GetNextStates_WasCalled = True
    m_GetNextStatesCallCount = m_GetNextStatesCallCount + 1
    
    ' Capturar parámetros
    m_LastEstadoActual = estadoActual
    m_LastTipoSolicitud = tipoSolicitud
    m_LastUsuarioRol = usuarioRol
    
    Set IWorkflowService_GetNextStates = m_GetNextStatesReturnValue
End Function

Private Function IWorkflowService_GetInitialState(ByVal tipoSolicitud As String) As String
    m_GetInitialState_WasCalled = True
    m_GetInitialStateCallCount = m_GetInitialStateCallCount + 1
    
    ' Capturar parámetros
    m_LastTipoSolicitud = tipoSolicitud
    
    IWorkflowService_GetInitialState = m_GetInitialStateReturnValue
End Function

Private Function IWorkflowService_IsStateFinal(ByVal estadoCodigo As String, ByVal tipoSolicitud As String) As Boolean
    m_IsStateFinal_WasCalled = True
    m_IsStateFinalCallCount = m_IsStateFinalCallCount + 1
    
    ' Capturar parámetros
    m_LastEstadoCodigo = estadoCodigo
    m_LastTipoSolicitud = tipoSolicitud
    
    IWorkflowService_IsStateFinal = m_IsStateFinalReturnValue
End Function

Private Function IWorkflowService_RecordStateChange(ByVal SolicitudID As Long, ByVal estadoAnterior As String, ByVal estadoNuevo As String, ByVal usuario As String, ByVal comentarios As String) As Boolean
    m_RecordStateChange_WasCalled = True
    m_RecordStateChangeCallCount = m_RecordStateChangeCallCount + 1
    
    ' Capturar parámetros
    m_LastSolicitudID = SolicitudID
    m_LastEstadoAnterior = estadoAnterior
    m_LastEstadoNuevo = estadoNuevo
    m_LastUsuario = usuario
    m_LastComentarios = comentarios
    
    IWorkflowService_RecordStateChange = m_RecordStateChangeReturnValue
End Function

Private Function IWorkflowService_GetStateHistory(ByVal SolicitudID As Long) As Scripting.Dictionary
    m_GetStateHistory_WasCalled = True
    m_GetStateHistoryCallCount = m_GetStateHistoryCallCount + 1
    
    ' Capturar parámetros
    m_LastSolicitudID = SolicitudID
    
    Set IWorkflowService_GetStateHistory = m_GetStateHistoryReturnValue
End Function

Private Function IWorkflowService_HasTransitionPermission(ByVal usuarioRol As String, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As Boolean
    m_HasTransitionPermission_WasCalled = True
    m_HasTransitionPermissionCallCount = m_HasTransitionPermissionCallCount + 1
    
    ' Capturar parámetros
    m_LastUsuarioRol = usuarioRol
    m_LastEstadoOrigen = estadoOrigen
    m_LastEstadoDestino = estadoDestino
    m_LastTipoSolicitud = tipoSolicitud
    
    IWorkflowService_HasTransitionPermission = m_HasTransitionPermissionReturnValue
End Function

Private Function IWorkflowService_RequiresApproval(ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As Boolean
    m_RequiresApproval_WasCalled = True
    m_RequiresApprovalCallCount = m_RequiresApprovalCallCount + 1
    
    ' Capturar parámetros
    m_LastEstadoOrigen = estadoOrigen
    m_LastEstadoDestino = estadoDestino
    m_LastTipoSolicitud = tipoSolicitud
    
    IWorkflowService_RequiresApproval = m_RequiresApprovalReturnValue
End Function

'==============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA (DELEGACIÓN A INTERFAZ)
'==============================================================================

Public Function ValidateTransition(ByVal SolicitudID As Long, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Boolean
    ValidateTransition = IWorkflowService_ValidateTransition(SolicitudID, estadoOrigen, estadoDestino, tipoSolicitud, usuarioRol)
End Function

Public Function GetAvailableStates(ByVal tipoSolicitud As String) As Collection
    Set GetAvailableStates = IWorkflowService_GetAvailableStates(tipoSolicitud)
End Function

Public Function GetNextStates(ByVal estadoActual As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Collection
    Set GetNextStates = IWorkflowService_GetNextStates(estadoActual, tipoSolicitud, usuarioRol)
End Function

Public Function GetInitialState(ByVal tipoSolicitud As String) As String
    GetInitialState = IWorkflowService_GetInitialState(tipoSolicitud)
End Function

Public Function IsStateFinal(ByVal estadoCodigo As String, ByVal tipoSolicitud As String) As Boolean
    IsStateFinal = IWorkflowService_IsStateFinal(estadoCodigo, tipoSolicitud)
End Function

Public Function RecordStateChange(ByVal SolicitudID As Long, ByVal estadoAnterior As String, ByVal estadoNuevo As String, ByVal usuario As String, ByVal comentarios As String) As Boolean
    RecordStateChange = IWorkflowService_RecordStateChange(SolicitudID, estadoAnterior, estadoNuevo, usuario, comentarios)
End Function

Public Function GetStateHistory(ByVal SolicitudID As Long) As Collection
    Set GetStateHistory = IWorkflowService_GetStateHistory(SolicitudID)
End Function

Public Function HasTransitionPermission(ByVal usuarioRol As String, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As Boolean
    HasTransitionPermission = IWorkflowService_HasTransitionPermission(usuarioRol, estadoOrigen, estadoDestino, tipoSolicitud)
End Function

Public Function RequiresApproval(ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As Boolean
    RequiresApproval = IWorkflowService_RequiresApproval(estadoOrigen, estadoDestino, tipoSolicitud)
End Function

#End If