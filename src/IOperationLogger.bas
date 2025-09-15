Option Compare Database
Option Explicit

' Contrato para el servicio de registro de operaciones.
' Versión 2.0 - Refactorizado para usar el objeto de entidad.

Public Sub LogOperation(ByVal logEntry As EOperationLog)
End Sub

Public Sub LogSolicitudOperation(ByVal logEntry As EOperationLog, ByVal solicitud As ESolicitud, ByVal userId As String)
End Sub