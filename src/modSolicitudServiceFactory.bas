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
Public Function CreateSolicitudService() As ISolicitudService
    On Error GoTo ErrorHandler
    
    ' 1. Obtener dependencias de sus propias factorías
    Dim repo As ISolicitudRepository
    Dim logger As IOperationLogger
    Dim errorHandler As IErrorHandlerService
    Dim config As IConfig
    
    ' Crear errorHandler primero para poder usarlo en caso de errores
    Set config = modConfig.CreateConfigService()
    Dim fileSystem As IFileSystem
    Set fileSystem = modFileSystemFactory.CreateFileSystem()
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    ' Crear las demás dependencias
    Set repo = modRepositoryFactory.CreateSolicitudRepository(config, errorHandler)
    Set logger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' 2. Crear e inicializar el servicio
    Dim serviceInstance As New CSolicitudService
    serviceInstance.Initialize repo, logger, errorHandler
    
    ' 3. Devolver la instancia
    Set CreateSolicitudService = serviceInstance
    
    Exit Function
    
ErrorHandler:
    If Not errorHandler Is Nothing Then
        errorHandler.LogError Err.Number, Err.Description, "modSolicitudServiceFactory.CreateSolicitudService"
    End If
    Set CreateSolicitudService = Nothing
End Function

