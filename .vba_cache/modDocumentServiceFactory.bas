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
    
    Dim solicitudRepository As ISolicitudRepository
    Set solicitudRepository = modRepositoryFactory.CreateSolicitudRepository(errorHandler)
    
    Dim operationLogger As IOperationLogger
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger(errorHandler)
    
    Dim wordManager As IWordManager
    Set wordManager = modWordManagerFactory.CreateWordManager(errorHandler)
    
    Dim mapeoRepository As IMapeoRepository
    Set mapeoRepository = modRepositoryFactory.CreateMapeoRepository(errorHandler)
    
    ' Crear una instancia de la clase concreta
    Dim documentServiceInstance As New CDocumentService
    
    ' Inicializar la instancia concreta con las dependencias
    documentServiceInstance.Initialize configService, solicitudRepository, operationLogger, wordManager, mapeoRepository
    
    ' Devolver la instancia inicializada como el tipo de la interfaz
    Set CreateDocumentService = documentServiceInstance
    
    Exit Function
    
ErrorHandler:
    errorHandler.LogError Err.Number, Err.Description, "modDocumentServiceFactory.CreateDocumentService"
    Set CreateDocumentService = Nothing
End Function


