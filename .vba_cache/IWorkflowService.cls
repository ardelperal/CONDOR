VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWorkflowService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ===================================================================
' Interfaz: IWorkflowService
' Descripción: Contrato para el servicio de workflow y estados.
' ===================================================================

' Métodos de validación de transiciones
Public Function ValidateTransition(ByVal SolicitudID As Long, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Boolean
End Function

' Métodos de obtención de estados
Public Function GetAvailableStates(ByVal tipoSolicitud As String) As Collection
End Function

Public Function GetNextStates(ByVal estadoActual As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Collection
End Function

Public Function GetInitialState(ByVal tipoSolicitud As String) As String
End Function

Public Function IsStateFinal(ByVal estadoCodigo As String, ByVal tipoSolicitud As String) As Boolean
End Function

' Métodos de historial de estados
Public Function RecordStateChange(ByVal SolicitudID As Long, ByVal estadoAnterior As String, ByVal estadoNuevo As String, ByVal usuario As String, ByVal comentarios As String) As Boolean
End Function

Public Function GetStateHistory(ByVal SolicitudID As Long) As Collection
End Function

' Métodos de validación de permisos
Public Function HasTransitionPermission(ByVal usuarioRol As String, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As Boolean
End Function

Public Function RequiresApproval(ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As Boolean
End Function

