Attribute VB_Name = "modDocumentServiceFactory"
Option Compare Database
Option Explicit

' =====================================================
' MÓDULO: modDocumentServiceFactory
' DESCRIPCIÓN: Factory para la creación del servicio de documentos
' AUTOR: Sistema CONDOR
' FECHA: 2025-08-22
' =====================================================

Public Function CreateDocumentService() As IDocumentService
    On Error GoTo ErrorHandler
    
    ' Obtener las dependencias necesarias
    Dim configService As IConfig
    Set configService = modConfig.CreateConfigService()
    
    Dim solicitudRepository As ISolicitudRepository
    Set solicitudRepository = modRepositoryFactory.CreateSolicitudRepository()
    
    Dim operationLogger As IOperationLogger
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger()
    
    Dim wordManager As IWordManager
    Set wordManager = modWordManagerFactory.CreateWordManager()
    
    Dim mapeoRepository As IMapeoRepository
    Set mapeoRepository = modRepositoryFactory.CreateMapeoRepository()
    
    ' Crear una instancia de la clase concreta
    Dim documentServiceInstance As New CDocumentService
    
    ' Inicializar la instancia concreta con las dependencias
    documentServiceInstance.Initialize configService, solicitudRepository, operationLogger, wordManager, mapeoRepository
    
    ' Devolver la instancia inicializada como el tipo de la interfaz
    Set CreateDocumentService = documentServiceInstance
    
    Exit Function
    
ErrorHandler:
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "modDocumentServiceFactory.CreateDocumentService"
    Set CreateDocumentService = Nothing
End Function
