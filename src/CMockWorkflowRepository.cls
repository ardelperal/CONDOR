Attribute VB_Name = "CMockWorkflowRepository"
Option Compare Database
Option Explicit
Implements IWorkflowRepository

' Implementar todos los métodos de IWorkflowRepository devolviendo valores por defecto
' (ej. False para Booleans, Nothing para Objects, 0 para Longs, "" para Strings)

Private Function IWorkflowRepository_GetInitialState(ByVal tipoSolicitud As String) As String
    IWorkflowRepository_GetInitialState = "BORRADOR"
End Function

Private Function IWorkflowRepository_IsValidTransition(ByVal tipoSolicitud As String, ByVal estadoOrigen As String, ByVal estadoDestino As String) As Boolean
    IWorkflowRepository_IsValidTransition = True ' Simula éxito por defecto
End Function

Private Function IWorkflowRepository_GetAvailableStates(ByVal tipoSolicitud As String) As Collection
    Set IWorkflowRepository_GetAvailableStates = New Collection
End Function

Private Function IWorkflowRepository_GetNextStates(ByVal estadoActual As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Collection
    Set IWorkflowRepository_GetNextStates = New Collection
End Function

Private Function IWorkflowRepository_IsStateFinal(ByVal estadoCodigo As String, ByVal tipoSolicitud As String) As Boolean
    IWorkflowRepository_IsStateFinal = False
End Function

Private Function IWorkflowRepository_RecordStateChange(ByVal SolicitudID As Long, ByVal estadoAnterior As String, ByVal estadoNuevo As String, ByVal usuario As String, ByVal comentarios As String) As Boolean
    IWorkflowRepository_RecordStateChange = True
End Function

Private Function IWorkflowRepository_GetStateHistory(ByVal SolicitudID As Long) As Collection
    Set IWorkflowRepository_GetStateHistory = New Collection
End Function

Private Function IWorkflowRepository_HasTransitionPermission(ByVal usuarioRol As String, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As Boolean
    IWorkflowRepository_HasTransitionPermission = True
End Function

Private Function IWorkflowRepository_RequiresApproval(ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As Boolean
    IWorkflowRepository_RequiresApproval = False
End Function

Private Function IWorkflowRepository_GetTransitionRequiredRole(ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As String
    IWorkflowRepository_GetTransitionRequiredRole = ""
End Function