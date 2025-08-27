Attribute VB_Name = "modSolicitudServiceFactory"
Option Compare Database
Option Explicit

'******************************************************************************
' MÓDULO: modSolicitudServiceFactory
' DESCRIPCIÓN: Factory para la inyección de dependencias del servicio de solicitudes
' AUTOR: Sistema CONDOR
' FECHA: 2025-01-14
'******************************************************************************

'******************************************************************************
' FACTORY METHODS
'******************************************************************************

'******************************************************************************
' FUNCIÓN: CreateSolicitudService
' DESCRIPCIÓN: Crea una instancia del servicio de solicitudes con todas sus dependencias
' RETORNA: ISolicitudService - Instancia del servicio completamente inicializada
'******************************************************************************
Public Function CreateSolicitudService(ByVal errorHandler As IErrorHandlerService) As ISolicitudService
    On Error GoTo ErrorHandler
    
    ' Obtener todas las dependencias necesarias
    Dim solicitudRepository As ISolicitudRepository
    Set solicitudRepository = modRepositoryFactory.CreateSolicitudRepository(errorHandler)
    
    Dim operationLogger As IOperationLogger
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger(errorHandler)
    
    ' Crear la instancia del servicio
    Dim serviceInstance As New CSolicitudService
    
    ' Inyectar las dependencias requeridas por Initialize
    serviceInstance.Initialize solicitudRepository, operationLogger
    
    Set CreateSolicitudService = serviceInstance
    
    Exit Function
    
ErrorHandler:
    errorHandler.LogError Err.Number, Err.Description, "modSolicitudServiceFactory.CreateSolicitudService"
    Set CreateSolicitudService = Nothing
End Function

