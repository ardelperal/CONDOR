Attribute VB_Name = "modWorkflowServiceFactory"
Option Compare Database
Option Explicit


Public Function CreateWorkflowService() As IWorkflowService
    On Error GoTo errorHandler

    Dim config As IConfig
    Set config = modTestContext.GetTestConfig()

    ' 2. Crear las dependencias
    Dim repo As IWorkflowRepository
    Set repo = modRepositoryFactory.CreateWorkflowRepository()

    Dim logger As IOperationLogger
    Set logger = modOperationLoggerFactory.CreateOperationLogger()

    Dim errorHandlerSvc As IErrorHandlerService
    Set errorHandlerSvc = modErrorHandlerFactory.CreateErrorHandlerService()

    ' 3. Crear e inicializar la implementación del servicio
    Dim serviceImpl As New CWorkflowService
    serviceImpl.Initialize repo, logger, errorHandlerSvc

    Set CreateWorkflowService = serviceImpl
    Exit Function

errorHandler:
    Debug.Print "Error crítico en modWorkflowServiceFactory: " & Err.Description
    Set CreateWorkflowService = Nothing
End Function

