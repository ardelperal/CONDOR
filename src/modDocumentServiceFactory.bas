Attribute VB_Name = "modDocumentServiceFactory"
Option Compare Database
Option Explicit


' =====================================================
' MÓDULO: modDocumentServiceFactory
' DESCRIPCIÓN: Factory para la creación del servicio de documentos
' AUTOR: Sistema CONDOR
' FECHA: 2025-08-22
' =====================================================

Public Function CreateDocumentService(ByVal errorHandler As IErrorHandlerService) As IDocumentService
    On Error GoTo ErrorHandler
    
    ' Obtener las dependencias necesarias
    Dim configService As IConfig
    Set configService = modConfig.CreateConfigService(errorHandler)
    
    ' Crear las dependencias usando los factories correspondientes
    Set wordManager = modWordManagerFactory.CreateWordManager(errorHandler, configService)
    Set mapeoRepository = modRepositoryFactory.CreateMapeoRepository(errorHandler, configService)
    Set solicitudRepository = modRepositoryFactory.CreateSolicitudRepository(errorHandler, configService)
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger(errorHandler)
    
    ' Crear una instancia de la clase concreta
    Dim documentServiceInstance As New CDocumentService
    
    ' Inicializar la instancia concreta con las dependencias (incluyendo errorHandler)
    documentServiceInstance.Initialize configService, solicitudRepository, operationLogger, wordManager, mapeoRepository, errorHandler
    
    ' Devolver la instancia inicializada como el tipo de la interfaz
    Set CreateDocumentService = documentServiceInstance
    
    Exit Function
    
ErrorHandler:
    errorHandler.LogError Err.Number, Err.Description, "modDocumentServiceFactory.CreateDocumentService"
    Set CreateDocumentService = Nothing
End Function


