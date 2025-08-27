VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IOperationLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =====================================================
' Interfaz: IOperationLogger
' Propósito: Define el contrato para el logging de operaciones
' Autor: CONDOR-Expert
' Fecha: 2025-01-15
' =====================================================


' Registra una operación general del sistema
' @param operationType: Tipo de operación realizada
' @param entityId: ID de la entidad afectada
' @param details: Detalles adicionales de la operación
Public Sub LogOperation(ByVal operationType As String, ByVal entityId As String, ByVal details As String)
End Sub

' Registra la creación de una solicitud
' @param solicitudId: ID de la solicitud creada
' @param userId: ID del usuario que creó la solicitud
Public Sub LogSolicitudCreation(ByVal SolicitudID As Long, ByVal userId As String)
End Sub

' Registra un cambio de estado
' @param solicitudId: ID de la solicitud
' @param oldState: Estado anterior
' @param newState: Estado nuevo
' @param userId: ID del usuario que realizó el cambio
Public Sub LogStateChange(ByVal SolicitudID As Long, ByVal oldState As String, ByVal newState As String, ByVal userId As String)
End Sub

' Registra un error del sistema
' @param errNumber: Número del error
' @param errDescription: Descripción del error
' @param moduleName: Nombre del módulo donde ocurrió el error (opcional)
' @param isCritical: Indica si el error es crítico (opcional, por defecto False)
Public Sub LogError(ByVal errNumber As Long, ByVal errDescription As String, Optional ByVal moduleName As String, Optional ByVal isCritical As Boolean = False)
End Sub


