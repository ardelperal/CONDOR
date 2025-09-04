Attribute VB_Name = "modDocumentServiceFactory"
Option Compare Database
Option Explicit


Public Function CreateDocumentService(Optional ByVal config As IConfig = Nothing) As IDocumentService
    On Error GoTo errorHandler
    
    ' Determinar configuración final
    Dim finalConfig As IConfig
    If config Is Nothing Then
        Set finalConfig = modConfigFactory.CreateConfigService()
    Else
        Set finalConfig = config
    End If
    
    Dim serviceImpl As New CDocumentService
    
    ' Crear TODAS las dependencias propagando la configuración
    Dim wordMgr As IWordManager
    Set wordMgr = modWordManagerFactory.CreateWordManager(finalConfig)
    
    Dim errHandler As IErrorHandlerService
    Set errHandler = modErrorHandlerFactory.CreateErrorHandlerService(finalConfig)
    
    Dim solicitudSrv As ISolicitudService
    Set solicitudSrv = modSolicitudServiceFactory.CreateSolicitudService(finalConfig)
    
    Dim mapeoRepo As IMapeoRepository
    Set mapeoRepo = modRepositoryFactory.CreateMapeoRepository(finalConfig)
    
    ' Inyectar dependencias en el orden correcto
    serviceImpl.Initialize wordMgr, errHandler, solicitudSrv, mapeoRepo
    
    Set CreateDocumentService = serviceImpl
    Exit Function
    
errorHandler:
    Debug.Print "Error crítico en modDocumentServiceFactory: " & Err.Description
    Set CreateDocumentService = Nothing
End Function




