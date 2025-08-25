Option Compare Database
Option Explicit

' Interfaz para el servicio de logging de operaciones
' Define el contrato para registrar acciones importantes del usuario y del sistema.

Public Sub LogOperation(ByVal operationType As String, ByVal entityId As String, ByVal details As String)
End Sub

' Puedes aÃ±adir mÃ©todos mÃ¡s especÃ­ficos si lo deseas, por ejemplo:
' Public Sub LogSolicitudCreation(ByVal solicitudId As Long, ByVal userId As String)
' Public Sub LogStateChange(ByVal solicitudId As Long, ByVal oldState As String, ByVal newState As String, ByVal userId As String)
