Option Compare Database
Option Explicit
' ===================================================================
' Interfaz: IWorkflowService
' DescripciÃ³n: Contrato para el servicio de workflow y estados.
' ===================================================================

Public Function ValidateTransition(ByVal SolicitudID As Long, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Boolean
End Function

Public Function GetNextStates(ByVal estadoActual As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Collection
End Function
