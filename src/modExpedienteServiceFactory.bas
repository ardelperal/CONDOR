Attribute VB_Name = "modExpedienteServiceFactory"
Option Compare Database
Option Explicit

Public Function CreateExpedienteService() As IExpedienteService
    On Error GoTo ErrorHandler
    
    Dim serviceImpl As New CExpedienteService
    
    Dim repo As IExpedienteRepository
    Set repo = modRepositoryFactory.CreateExpedienteRepository()
    
    Dim logger As IOperationLogger
    Set logger = modOperationLoggerFactory.CreateOperationLogger()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    ' La dependencia de IConfig ahora es manejada por el repositorio, no por el servicio.
    serviceImpl.Initialize repo, logger, errorHandler
    
    Set CreateExpedienteService = serviceImpl
    Exit Function
    
ErrorHandler:
    Debug.Print "Error crítico en modExpedienteServiceFactory: " & Err.Description
    Set CreateExpedienteService = Nothing
End Function

