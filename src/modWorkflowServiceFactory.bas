Attribute VB_Name = "modWorkflowServiceFactory"
Option Compare Database
Option Explicit


Public Function CreateWorkflowService() As IWorkflowService
    On Error GoTo errorHandler
    
    Dim serviceImpl As New CWorkflowService
    
    ' Crear dependencias llamando a sus factorías
    Dim repo As IWorkflowRepository
    Set repo = modRepositoryFactory.CreateWorkflowRepository()
    
    Dim logger As IOperationLogger
    Set logger = modOperationLoggerFactory.CreateOperationLogger()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    ' Inyectar dependencias
    serviceImpl.Initialize repo, logger, errorHandler
    
    Set CreateWorkflowService = serviceImpl
    Exit Function
    
errorHandler:
    Debug.Print "Error crítico en modWorkflowServiceFactory: " & Err.Description
    Set CreateWorkflowService = Nothing
End Function

