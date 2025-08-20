Option Compare Database
Option Explicit
' Interfaz: ILoggingService
' Descripción: Define el contrato para servicios de logging del sistema CONDOR
' Arquitectura: Capa de Servicios - Interface

' Registra un error en el sistema de logging
' @param errNumber: Número del error
' @param errDescription: Descripción del error
' @param moduleName: Nombre del módulo donde ocurrió el error (opcional)
' @param context: Contexto adicional del error (opcional)
Public Sub LogError(ByVal errNumber As Long, ByVal errDescription As String, Optional ByVal moduleName As String, Optional ByVal context As String)
End Sub

Public Sub LogInfo(ByVal message As String, Optional ByVal moduleName As String)
End Sub






