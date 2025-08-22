Option Compare Database
Option Explicit
' ============================================================================
' Interface: ISolicitudService
' Descripcion: Contrato para la logica de negocio de Solicitudes.
' Version: 3.0
' ============================================================================

Public Function CreateSolicitud(ByVal idExpediente As String, ByVal tipo As String, ByVal subTipo As String) As T_Solicitud
End Function

Public Function GetSolicitud(ByVal idSolicitud As Long) As T_Solicitud
End Function

Public Function SaveSolicitud(solicitud As T_Solicitud) As Boolean
End Function

Public Function ChangeState(ByVal idSolicitud As Long, ByVal newState As String) As Boolean
End Function











