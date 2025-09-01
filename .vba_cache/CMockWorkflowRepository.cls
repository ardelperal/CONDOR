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

Implements IWorkflowRepository

' Inicialización de la clase
Private Sub Class_Initialize()
    Call Reset
End Sub

' Miembros privados para el mock
Private m_isValidTransitionReturnValue As Boolean
Private m_isValidTransitionWasCalled As Boolean
Private m_lastEstadoOrigen As String
Private m_lastEstadoDestino As String

' Miembros privados para el mock de GetAvailableStates
Private m_getAvailableStatesReturnValue As Scripting.Dictionary
Private m_getAvailableStatesWasCalled As Boolean
Private m_lastTipoSolicitudForStates As String

' Configuración del valor de retorno para IsValidTransition
Public Sub ConfigureIsValidTransition(ByVal returnValue As Boolean)
    m_isValidTransitionReturnValue = returnValue
End Sub

' Propiedades "espía" para verificar llamadas en los tests
Public Property Get IsValidTransition_WasCalled() As Boolean
    IsValidTransition_WasCalled = m_isValidTransitionWasCalled
End Property

Public Property Get LastEstadoOrigen() As String
    LastEstadoOrigen = m_lastEstadoOrigen
End Property

Public Property Get LastEstadoDestino() As String
    LastEstadoDestino = m_lastEstadoDestino
End Property

' Configuración del valor de retorno para GetAvailableStates
Public Sub ConfigureGetAvailableStates(ByVal returnValue As Scripting.Dictionary)
    Set m_getAvailableStatesReturnValue = returnValue
End Sub

' Propiedades "espía" para GetAvailableStates
Public Property Get GetAvailableStates_WasCalled() As Boolean
    GetAvailableStates_WasCalled = m_getAvailableStatesWasCalled
End Property

Public Property Get LastTipoSolicitudForStates() As String
    LastTipoSolicitudForStates = m_lastTipoSolicitudForStates
End Property

' Limpia el estado del mock para asegurar el aislamiento de los tests
Public Sub Reset()
    ' Inicializar valores por defecto
    m_isValidTransitionReturnValue = False
    Set m_getAvailableStatesReturnValue = New Scripting.Dictionary
    m_getAvailableStatesReturnValue.CompareMode = TextCompare
    Set m_getNextStatesReturnValue = New Scripting.Dictionary
    m_getNextStatesReturnValue.CompareMode = TextCompare
    m_getInitialStateReturnValue = ""
    m_isStateFinalReturnValue = False
    m_recordStateChangeReturnValue = True
    Set m_getStateHistoryReturnValue = New Scripting.Dictionary
    m_getStateHistoryReturnValue.CompareMode = TextCompare
    m_hasTransitionPermissionReturnValue = False
    m_requiresApprovalReturnValue = False
    m_getTransitionRequiredRoleReturnValue = ""
    
    ' Inicializar indicadores de llamadas
    m_isValidTransitionWasCalled = False
    m_getAvailableStatesWasCalled = False
    m_getNextStatesWasCalled = False
    m_getInitialStateWasCalled = False
    m_isStateFinalWasCalled = False
    m_recordStateChangeWasCalled = False
    m_getStateHistoryWasCalled = False
    m_hasTransitionPermissionWasCalled = False
    m_requiresApprovalWasCalled = False
    m_getTransitionRequiredRoleWasCalled = False
    
    ' Limpiar argumentos capturados
    m_lastEstadoOrigen = ""
    m_lastEstadoDestino = ""
    m_lastTipoSolicitudForStates = ""
    m_lastEstadoActual = ""
    m_lastUsuarioRol = ""
    m_lastEstadoCodigo = ""
    m_lastSolicitudID = 0
    m_lastEstadoAnterior = ""
    m_lastEstadoNuevo = ""
    m_lastUsuario = ""
    m_lastComentarios = ""
End Sub

' Implementación del método de la interfaz
Private Function IWorkflowRepository_IsValidTransition(ByVal tipoSolicitud As String, ByVal estadoOrigen As String, ByVal estadoDestino As String) As Boolean
    ' Registrar que el método fue llamado y capturar los argumentos
    m_isValidTransitionWasCalled = True
    m_lastEstadoOrigen = estadoOrigen
    m_lastEstadoDestino = estadoDestino
    
    ' Devolver el valor preconfigurado
    IWorkflowRepository_IsValidTransition = m_isValidTransitionReturnValue
End Function

' Implementación del método de la interfaz GetAvailableStates
Private Function IWorkflowRepository_GetAvailableStates(ByVal tipoSolicitud As String) As Scripting.Dictionary
    ' Registrar que el método fue llamado y capturar los argumentos
    m_getAvailableStatesWasCalled = True
    m_lastTipoSolicitudForStates = tipoSolicitud
    
    ' Devolver la colección preconfigurada
    Set IWorkflowRepository_GetAvailableStates = m_getAvailableStatesReturnValue
End Function

' Variables para GetNextStates
Private m_getNextStatesReturnValue As Scripting.Dictionary
Private m_getNextStatesWasCalled As Boolean
Private m_lastEstadoActual As String
Private m_lastUsuarioRol As String

' Configuración para GetNextStates
Public Sub ConfigureGetNextStates(ByVal returnValue As Scripting.Dictionary)
    Set m_getNextStatesReturnValue = returnValue
End Sub

Public Property Get GetNextStates_WasCalled() As Boolean
    GetNextStates_WasCalled = m_getNextStatesWasCalled
End Property

Public Property Get LastEstadoActual() As String
    LastEstadoActual = m_lastEstadoActual
End Property

Public Property Get LastUsuarioRol() As String
    LastUsuarioRol = m_lastUsuarioRol
End Property

' Implementación de GetNextStates
Private Function IWorkflowRepository_GetNextStates(ByVal estadoActual As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Scripting.Dictionary
    m_getNextStatesWasCalled = True
    m_lastEstadoActual = estadoActual
    m_lastUsuarioRol = usuarioRol
    Set IWorkflowRepository_GetNextStates = m_getNextStatesReturnValue
End Function

' Variables para GetInitialState
Private m_getInitialStateReturnValue As String
Private m_getInitialStateWasCalled As Boolean

' Configuración para GetInitialState
Public Sub ConfigureGetInitialState(ByVal returnValue As String)
    m_getInitialStateReturnValue = returnValue
End Sub

Public Property Get GetInitialState_WasCalled() As Boolean
    GetInitialState_WasCalled = m_getInitialStateWasCalled
End Property

' Implementación de GetInitialState
Private Function IWorkflowRepository_GetInitialState(ByVal tipoSolicitud As String) As String
    m_getInitialStateWasCalled = True
    IWorkflowRepository_GetInitialState = m_getInitialStateReturnValue
End Function

' Variables para IsStateFinal
Private m_isStateFinalReturnValue As Boolean
Private m_isStateFinalWasCalled As Boolean
Private m_lastEstadoCodigo As String

' Configuración para IsStateFinal
Public Sub ConfigureIsStateFinal(ByVal returnValue As Boolean)
    m_isStateFinalReturnValue = returnValue
End Sub

Public Property Get IsStateFinal_WasCalled() As Boolean
    IsStateFinal_WasCalled = m_isStateFinalWasCalled
End Property

Public Property Get LastEstadoCodigo() As String
    LastEstadoCodigo = m_lastEstadoCodigo
End Property

' Implementación de IsStateFinal
Private Function IWorkflowRepository_IsStateFinal(ByVal estadoCodigo As String, ByVal tipoSolicitud As String) As Boolean
    m_isStateFinalWasCalled = True
    m_lastEstadoCodigo = estadoCodigo
    IWorkflowRepository_IsStateFinal = m_isStateFinalReturnValue
End Function

' Variables para RecordStateChange
Private m_recordStateChangeReturnValue As Boolean
Private m_recordStateChangeWasCalled As Boolean
Private m_lastSolicitudID As Long
Private m_lastEstadoAnterior As String
Private m_lastEstadoNuevo As String
Private m_lastUsuario As String
Private m_lastComentarios As String

' Configuración para RecordStateChange
Public Sub ConfigureRecordStateChange(ByVal returnValue As Boolean)
    m_recordStateChangeReturnValue = returnValue
End Sub

Public Property Get RecordStateChange_WasCalled() As Boolean
    RecordStateChange_WasCalled = m_recordStateChangeWasCalled
End Property

Public Property Get LastSolicitudID() As Long
    LastSolicitudID = m_lastSolicitudID
End Property

Public Property Get LastEstadoAnterior() As String
    LastEstadoAnterior = m_lastEstadoAnterior
End Property

Public Property Get LastEstadoNuevo() As String
    LastEstadoNuevo = m_lastEstadoNuevo
End Property

Public Property Get LastUsuario() As String
    LastUsuario = m_lastUsuario
End Property

Public Property Get LastComentarios() As String
    LastComentarios = m_lastComentarios
End Property

' Implementación de RecordStateChange
Private Function IWorkflowRepository_RecordStateChange(ByVal SolicitudID As Long, ByVal estadoAnterior As String, ByVal estadoNuevo As String, ByVal usuario As String, ByVal comentarios As String) As Boolean
    m_recordStateChangeWasCalled = True
    m_lastSolicitudID = SolicitudID
    m_lastEstadoAnterior = estadoAnterior
    m_lastEstadoNuevo = estadoNuevo
    m_lastUsuario = usuario
    m_lastComentarios = comentarios
    IWorkflowRepository_RecordStateChange = m_recordStateChangeReturnValue
End Function

' Variables para GetStateHistory
Private m_getStateHistoryReturnValue As Scripting.Dictionary
Private m_getStateHistoryWasCalled As Boolean

' Configuración para GetStateHistory
Public Sub ConfigureGetStateHistory(ByVal returnValue As Scripting.Dictionary)
    Set m_getStateHistoryReturnValue = returnValue
End Sub

Public Property Get GetStateHistory_WasCalled() As Boolean
    GetStateHistory_WasCalled = m_getStateHistoryWasCalled
End Property

' Implementación de GetStateHistory
Private Function IWorkflowRepository_GetStateHistory(ByVal SolicitudID As Long) As Scripting.Dictionary
    m_getStateHistoryWasCalled = True
    m_lastSolicitudID = SolicitudID
    Set IWorkflowRepository_GetStateHistory = m_getStateHistoryReturnValue
End Function

' Variables para HasTransitionPermission
Private m_hasTransitionPermissionReturnValue As Boolean
Private m_hasTransitionPermissionWasCalled As Boolean

' Configuración para HasTransitionPermission
Public Sub ConfigureHasTransitionPermission(ByVal returnValue As Boolean)
    m_hasTransitionPermissionReturnValue = returnValue
End Sub

Public Property Get HasTransitionPermission_WasCalled() As Boolean
    HasTransitionPermission_WasCalled = m_hasTransitionPermissionWasCalled
End Property

' Implementación de HasTransitionPermission
Private Function IWorkflowRepository_HasTransitionPermission(ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Boolean
    m_hasTransitionPermissionWasCalled = True
    m_lastEstadoOrigen = estadoOrigen
    m_lastEstadoDestino = estadoDestino
    m_lastUsuarioRol = usuarioRol
    IWorkflowRepository_HasTransitionPermission = m_hasTransitionPermissionReturnValue
End Function

' Variables para RequiresApproval
Private m_requiresApprovalReturnValue As Boolean
Private m_requiresApprovalWasCalled As Boolean

' Configuración para RequiresApproval
Public Sub ConfigureRequiresApproval(ByVal returnValue As Boolean)
    m_requiresApprovalReturnValue = returnValue
End Sub

Public Property Get RequiresApproval_WasCalled() As Boolean
    RequiresApproval_WasCalled = m_requiresApprovalWasCalled
End Property

' Implementación de RequiresApproval
Private Function IWorkflowRepository_RequiresApproval(ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As Boolean
    m_requiresApprovalWasCalled = True
    m_lastEstadoOrigen = estadoOrigen
    m_lastEstadoDestino = estadoDestino
    IWorkflowRepository_RequiresApproval = m_requiresApprovalReturnValue
End Function

' Variables para GetTransitionRequiredRole
Private m_getTransitionRequiredRoleReturnValue As String
Private m_getTransitionRequiredRoleWasCalled As Boolean

' Configuración para GetTransitionRequiredRole
Public Sub ConfigureGetTransitionRequiredRole(ByVal returnValue As String)
    m_getTransitionRequiredRoleReturnValue = returnValue
End Sub

Public Property Get GetTransitionRequiredRole_WasCalled() As Boolean
    GetTransitionRequiredRole_WasCalled = m_getTransitionRequiredRoleWasCalled
End Property

' Implementación de GetTransitionRequiredRole
Private Function IWorkflowRepository_GetTransitionRequiredRole(ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As String
    m_getTransitionRequiredRoleWasCalled = True
    m_lastEstadoOrigen = estadoOrigen
    m_lastEstadoDestino = estadoDestino
    IWorkflowRepository_GetTransitionRequiredRole = m_getTransitionRequiredRoleReturnValue
End Function