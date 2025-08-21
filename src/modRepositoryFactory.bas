Option Compare Database
Option Explicit
'******************************************************************************
' MÓDULO: modRepositoryFactory
' DESCRIPCIÓN: Factory para la inyección de dependencias del repositorio de solicitudes
' AUTOR: Sistema CONDOR
' FECHA: 2024
'******************************************************************************


'******************************************************************************
' FACTORY METHODS
'******************************************************************************

'******************************************************************************
' FUNCIÓN: CreateSolicitudRepository
' DESCRIPCIÓN: Crea una instancia del repositorio de solicitudes según el modo
' RETORNA: ISolicitudRepository - Instancia del repositorio (Mock o Real)
'******************************************************************************
Public Function CreateSolicitudRepository() As ISolicitudRepository
    On Error GoTo ErrorHandler
    
    ' Obtener la instancia del logger de operaciones
    Dim operationLogger As IOperationLogger
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' TODO: Implementar lógica para alternar entre Mock y Real según configuración
    ' Por ahora usamos el repositorio real con inyección de dependencias
    Dim repositoryInstance As New CSolicitudRepository
    
    ' Inyectar la dependencia del logger
    repositoryInstance.Initialize operationLogger
    
    Set CreateSolicitudRepository = repositoryInstance
    
    Exit Function
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "modRepositoryFactory.CreateSolicitudRepository")
    Set CreateSolicitudRepository = Nothing
End Function






