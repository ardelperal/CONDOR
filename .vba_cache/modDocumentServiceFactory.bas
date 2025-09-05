Attribute VB_Name = "modDocumentServiceFactory"
Option Compare Database
Option Explicit


Public Function CreateDocumentService() As IDocumentService
    On Error GoTo errorHandler
    
    Dim config As IConfig
    Set config = modTestContext.GetTestConfig()
    
    Dim serviceImpl As New CDocumentService
    
    ' Crear TODAS las dependencias
    Dim wordMgr As IWordManager
    Set wordMgr = modWordManagerFactory.CreateWordManager()
    
    Dim errHandler As IErrorHandlerService
    Set errHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    Dim solicitudSrv As ISolicitudService
    Set solicitudSrv = modSolicitudServiceFactory.CreateSolicitudService()
    
    Dim mapeoRepo As IMapeoRepository
    Set mapeoRepo = modRepositoryFactory.CreateMapeoRepository()
    
    ' Inyectar dependencias en el orden correcto
    serviceImpl.Initialize wordMgr, errHandler, solicitudSrv, mapeoRepo
    
    Set CreateDocumentService = serviceImpl
    Exit Function
    
errorHandler:
    Debug.Print "Error crítico en modDocumentServiceFactory: " & Err.Description
    Set CreateDocumentService = Nothing
End Function




