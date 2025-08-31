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
    
    Dim fileSystem As IFileSystem
    Set fileSystem = modFileSystemFactory.CreateFileSystem()
    
    Dim configService As IConfig
    Set configService = modConfigFactory.CreateConfigService()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(configService, fileSystem)
    
    Dim wordManager As IWordManager
    Set wordManager = modWordManagerFactory.CreateWordManager()
    
    Dim mapeoRepository As IMapeoRepository
    Set mapeoRepository = modRepositoryFactory.CreateMapeoRepository()
    
    Dim solicitudRepository As ISolicitudRepository
    Set solicitudRepository = modRepositoryFactory.CreateSolicitudRepository()
    
    Dim operationLogger As IOperationLogger
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger()
    
    Dim documentServiceInstance As New CDocumentService
    
    documentServiceInstance.Initialize configService, solicitudRepository, operationLogger, wordManager, mapeoRepository, errorHandler
    
    Set CreateDocumentService = documentServiceInstance
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error en modDocumentServiceFactory.CreateDocumentService: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


