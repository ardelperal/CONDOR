VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWorkflowService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'******************************************************************************
' CLASE: CWorkflowService
' DESCRIPCIÓN: Implementación concreta del servicio de workflow y estados
' AUTOR: Sistema CONDOR
' FECHA: 2024
' NOTAS: Lección 1 - Implementa la interfaz IWorkflowService
'        Lección 3 - Sigue la arquitectura de 3 capas (Capa de Servicios)
'        Lección 6 - Utiliza módulos de acceso a datos como especificación
'******************************************************************************


' Lección 1: Implementar la interfaz
Implements IWorkflowService

'******************************************************************************
' VARIABLES PRIVADAS PARA DEPENDENCIAS INYECTADAS
'******************************************************************************

Private m_Config As IConfig
Private m_OperationLogger As IOperationLogger
Private m_WorkflowRepository As IWorkflowRepository

'******************************************************************************
' CONSTANTES PRIVADAS
'******************************************************************************

Private Const ESTADO_BORRADOR As String = "Borrador"
Private Const ESTADO_EN_PROCESO As String = "EnProceso"
Private Const ESTADO_APROBADO As String = "Aprobado"
Private Const ESTADO_RECHAZADO As String = "Rechazado"
Private Const ESTADO_CANCELADO As String = "Cancelado"

Private Const ROL_USUARIO As String = "Usuario"
Private Const ROL_APROBADOR As String = "Aprobador"
Private Const ROL_ADMINISTRADOR As String = "Administrador"

'******************************************************************************
' MÉTODO DE INICIALIZACIÓN CON DEPENDENCIAS
'******************************************************************************

' Inicializa el servicio con sus dependencias
Public Sub Initialize(ByVal config As IConfig, ByVal operationLogger As IOperationLogger, ByVal workflowRepository As IWorkflowRepository)
    Set m_Config = config
    Set m_OperationLogger = operationLogger
    Set m_WorkflowRepository = workflowRepository
End Sub

'******************************************************************************
' IMPLEMENTACIÓN DE MÉTODOS DE VALIDACIÓN DE TRANSICIONES
'******************************************************************************

Private Function IWorkflowService_ValidateTransition(ByVal SolicitudID As Long, _
                                                    ByVal estadoOrigen As String, _
                                                    ByVal estadoDestino As String, _
                                                    ByVal tipoSolicitud As String, _
                                                    ByVal usuarioRol As String) As Boolean
    
    ' Logging de auditoría: Validación de transición de estado
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "WORKFLOW_TRANSITION_VALIDATION", "Validando transición solicitud ID: " & SolicitudID & ", de " & estadoOrigen & " a " & estadoDestino & " por usuario rol: " & usuarioRol, "CWorkflowService.ValidateTransition"
    End If
    
    Dim Resultado As Boolean
    Resultado = False
    
    ' Validar parámetros de entrada
    If Len(Trim(estadoOrigen)) = 0 Or Len(Trim(estadoDestino)) = 0 Or _
       Len(Trim(tipoSolicitud)) = 0 Or Len(Trim(usuarioRol)) = 0 Then
        IWorkflowService_ValidateTransition = False
        Exit Function
    End If
    
    ' Verificar si la transición existe usando el repositorio
    If Not m_WorkflowRepository.IsValidTransition(tipoSolicitud, estadoOrigen, estadoDestino) Then
        IWorkflowService_ValidateTransition = False
        Exit Function
    End If
    
    ' Verificar permisos del usuario para esta transición
    If Not IWorkflowService_HasTransitionPermission(usuarioRol, estadoOrigen, estadoDestino, tipoSolicitud) Then
        IWorkflowService_ValidateTransition = False
        Exit Function
    End If
    
    ' Si llegamos aquí, la transición es válida
    IWorkflowService_ValidateTransition = True
    
    ' Logging de auditoría: Transición validada exitosamente
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "WORKFLOW_TRANSITION_VALIDATED", "Transición validada exitosamente para solicitud ID: " & SolicitudID & ", de " & estadoOrigen & " a " & estadoDestino, "CWorkflowService.ValidateTransition"
    End If
End Function

'******************************************************************************
' IMPLEMENTACIÓN DE MÉTODOS DE OBTENCIÓN DE ESTADOS
'******************************************************************************

Private Function IWorkflowService_GetAvailableStates(ByVal tipoSolicitud As String) As Scripting.Dictionary
    ' Delegar al repositorio
    Set IWorkflowService_GetAvailableStates = m_WorkflowRepository.GetAvailableStates(tipoSolicitud)
End Function

Private Function IWorkflowService_GetNextStates(ByVal estadoActual As String, _
                                               ByVal tipoSolicitud As String, _
                                               ByVal usuarioRol As String) As Scripting.Dictionary
    ' Delegar al repositorio
    Set IWorkflowService_GetNextStates = m_WorkflowRepository.GetNextStates(estadoActual, tipoSolicitud, usuarioRol)
End Function

Private Function IWorkflowService_GetInitialState(ByVal tipoSolicitud As String) As String
    ' Delegar al repositorio
    IWorkflowService_GetInitialState = m_WorkflowRepository.GetInitialState(tipoSolicitud)
End Function

Private Function IWorkflowService_IsStateFinal(ByVal estadoCodigo As String, _
                                              ByVal tipoSolicitud As String) As Boolean
    ' Delegar al repositorio
    IWorkflowService_IsStateFinal = m_WorkflowRepository.IsStateFinal(estadoCodigo, tipoSolicitud)
End Function

'******************************************************************************
' IMPLEMENTACIÓN DE MÉTODOS DE HISTORIAL DE ESTADOS
'******************************************************************************

Private Function IWorkflowService_RecordStateChange(ByVal SolicitudID As Long, _
                                                   ByVal estadoAnterior As String, _
                                                   ByVal estadoNuevo As String, _
                                                   ByVal usuario As String, _
                                                   ByVal comentarios As String) As Boolean
    
    ' Logging de auditoría: Registro de cambio de estado
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "WORKFLOW_STATE_CHANGE", "Registrando cambio de estado para solicitud ID: " & SolicitudID & ", de " & estadoAnterior & " a " & estadoNuevo & " por usuario: " & usuario, "CWorkflowService.RecordStateChange"
    End If
    
    ' Delegar al repositorio
    Dim resultado As Boolean
    resultado = m_WorkflowRepository.RecordStateChange(SolicitudID, estadoAnterior, estadoNuevo, usuario, comentarios)
    
    If resultado Then
        ' Logging de auditoría: Cambio de estado registrado exitosamente
        If Not m_OperationLogger Is Nothing Then
            m_OperationLogger.LogOperation "WORKFLOW_STATE_RECORDED", "Cambio de estado registrado exitosamente para solicitud ID: " & SolicitudID & ", de " & estadoAnterior & " a " & estadoNuevo, "CWorkflowService.RecordStateChange"
        End If
    Else
        ' Logging de auditoría: Error al registrar cambio de estado
        If Not m_OperationLogger Is Nothing Then
            m_OperationLogger.LogOperation "WORKFLOW_STATE_RECORD_ERROR", "Error al registrar cambio de estado para solicitud ID: " & SolicitudID, "CWorkflowService.RecordStateChange"
        End If
    End If
    
    IWorkflowService_RecordStateChange = resultado
End Function

Private Function IWorkflowService_GetStateHistory(ByVal SolicitudID As Long) As Scripting.Dictionary
    ' Delegar al repositorio
    Set IWorkflowService_GetStateHistory = m_WorkflowRepository.GetStateHistory(SolicitudID)
End Function

'******************************************************************************
' IMPLEMENTACIÓN DE MÉTODOS DE VALIDACIÓN DE PERMISOS
'******************************************************************************

Private Function IWorkflowService_HasTransitionPermission(ByVal usuarioRol As String, _
                                                         ByVal estadoOrigen As String, _
                                                         ByVal estadoDestino As String, _
                                                         ByVal tipoSolicitud As String) As Boolean
    
    ' Obtener rol requerido del repositorio
    Dim rolRequerido As String
    rolRequerido = m_WorkflowRepository.GetTransitionRequiredRole(estadoOrigen, estadoDestino, tipoSolicitud)
    
    ' Si no hay rol requerido específico, cualquier usuario puede hacer la transición
    If Len(Trim(rolRequerido)) = 0 Then
        IWorkflowService_HasTransitionPermission = True
    ' Si el usuario es administrador, puede hacer cualquier transición
    ElseIf usuarioRol = ROL_ADMINISTRADOR Then
        IWorkflowService_HasTransitionPermission = True
    ' Verificar si el rol del usuario coincide con el requerido
    Else
        IWorkflowService_HasTransitionPermission = (usuarioRol = rolRequerido)
    End If
End Function

Private Function IWorkflowService_RequiresApproval(ByVal estadoOrigen As String, _
                                                  ByVal estadoDestino As String, _
                                                  ByVal tipoSolicitud As String) As Boolean
    ' Delegar al repositorio
    IWorkflowService_RequiresApproval = m_WorkflowRepository.RequiresApproval(estadoOrigen, estadoDestino, tipoSolicitud)
End Function

'******************************************************************************
' MÃ‰TODOS AUXILIARES PRIVADOS
'******************************************************************************

Private Function TransitionExists(ByVal estadoOrigen As String, _
                                 ByVal estadoDestino As String, _
                                 ByVal tipoSolicitud As String) As Boolean
    ' Delegar al repositorio
    TransitionExists = m_WorkflowRepository.IsValidTransition(tipoSolicitud, estadoOrigen, estadoDestino)
End Function

'******************************************************************************
' MÉTODOS PÚBLICOS DE CONVENIENCIA
'******************************************************************************

' Método público de conveniencia que delega a la implementación de la interfaz
Public Function ValidateTransition(ByVal SolicitudID As Long, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Boolean
    ValidateTransition = IWorkflowService_ValidateTransition(SolicitudID, estadoOrigen, estadoDestino, tipoSolicitud, usuarioRol)
End Function

' Método público de conveniencia que delega a la implementación de la interfaz
Public Function GetNextStates(ByVal estadoActual As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Scripting.Dictionary
    Set GetNextStates = IWorkflowService_GetNextStates(estadoActual, tipoSolicitud, usuarioRol)
End Function
















