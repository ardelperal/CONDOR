Attribute VB_Name = "modDocumentServiceFactory"
Option Compare Database
Option Explicit

' =====================================================
' MODULO: modDocumentServiceFactory
' DESCRIPCION: Factory para la creación del servicio de documentos
' AUTOR: Sistema CONDOR
' FECHA: 2025-08-22
' =====================================================

Public Function CreateDocumentService() As IDocumentService
    On Error GoTo ErrorHandler
    
    ' Obtener las dependencias necesarias
    Dim configService As IConfig
    Set configService = modConfigFactory.CreateConfigService()
    
    Dim solicitudRepository As ISolicitudRepository
    Set solicitudRepository = modRepositoryFactory.CreateSolicitudRepository()
    
    Dim operationLogger As IOperationLogger
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' Crear una instancia de la clase concreta
    Dim documentServiceInstance As New CDocumentService
    
    ' Inicializar la instancia concreta con todas las dependencias
    documentServiceInstance.Initialize configService, solicitudRepository, operationLogger
    
    ' Devolver la instancia inicializada como el tipo de la interfaz
    Set CreateDocumentService = documentServiceInstance
    
    Exit Function
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "modDocumentServiceFactory.CreateDocumentService")
    Set CreateDocumentService = Nothing
End Function
