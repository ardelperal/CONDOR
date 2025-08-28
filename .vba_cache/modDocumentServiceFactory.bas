Attribute VB_Name = "modDocumentServiceFactory"
Option Compare Database
Option Explicit


' =====================================================
' MÓDULO: modDocumentServiceFactory
' DESCRIPCIÓN: Factory para la creación del servicio de documentos
' AUTOR: Sistema CONDOR
' FECHA: 2025-08-22
' =====================================================

Public Function CreateDocumentService(Optional ByVal testConfig As IConfig = Nothing) As IDocumentService
    On Error GoTo ErrorHandler
    
    ' Crear las dependencias necesarias usando sus respectivas factorías
    Dim fileSystem As IFileSystem
    Set fileSystem = modFileSystemFactory.CreateFileSystem()
    
    Dim configService As IConfig
    If Not testConfig Is Nothing Then
        Set configService = testConfig
    Else
        Set configService = modConfig.CreateConfigService()
    End If
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(configService, fileSystem)
    
    ' Crear las dependencias usando los factories correspondientes
    Dim wordManager As IWordManager
    Set wordManager = modWordManagerFactory.CreateWordManager()
    
    Dim mapeoRepository As IMapeoRepository
    Set mapeoRepository = modRepositoryFactory.CreateMapeoRepository(configService, errorHandler)
    
    Dim solicitudRepository As ISolicitudRepository
    Set solicitudRepository = modRepositoryFactory.CreateSolicitudRepository(configService, errorHandler)
    
    Dim operationLogger As IOperationLogger
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' Crear una instancia de la clase concreta
    Dim documentServiceInstance As New CDocumentService
    
    ' Inicializar la instancia concreta con las dependencias (incluyendo errorHandler)
    documentServiceInstance.Initialize configService, solicitudRepository, operationLogger, wordManager, mapeoRepository, errorHandler
    
    ' Devolver la instancia inicializada como el tipo de la interfaz
    Set CreateDocumentService = documentServiceInstance
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error en modDocumentServiceFactory.CreateDocumentService: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


