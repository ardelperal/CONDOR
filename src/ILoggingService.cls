Option Compare Database
Option Explicit
' Interfaz: ILoggingService
' DescripciÃ³n: Define el contrato para servicios de logging del sistema CONDOR
' Arquitectura: Capa de Servicios - Interface

' Registra un error en el sistema de logging
' @param errNumber: NÃºmero del error
' @param errDescription: DescripciÃ³n del error
' @param moduleName: Nombre del mÃ³dulo donde ocurriÃ³ el error (opcional)
' @param context: Contexto adicional del error (opcional)
Public Sub LogError(ByVal errNumber As Long, ByVal errDescription As String, Optional ByVal moduleName As String, Optional ByVal context As String)
End Sub

Public Sub LogInfo(ByVal message As String, Optional ByVal moduleName As String)
End Sub

Public Sub LogWarning(ByVal message As String, Optional ByVal moduleName As String)
End Sub






