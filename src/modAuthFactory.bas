Attribute VB_Name = "modAuthFactory"
Option Compare Database
Option Explicit


Public Function CreateAuthService() As IAuthService
    On Error GoTo errorHandler
    
    Dim configSvc As IConfig
    Set configSvc = modTestContext.GetTestConfig()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    Dim operationLogger As IOperationLogger
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger()
    
    Dim authRepo As IAuthRepository
    Set authRepo = modRepositoryFactory.CreateAuthRepository() ' CORRECTO
    
    Dim authSvc As New CAuthService
    authSvc.Initialize configSvc, operationLogger, authRepo, errorHandler
    
    Set CreateAuthService = authSvc
    
    Exit Function
    
errorHandler:
    Debug.Print "Error en modAuthFactory.CreateAuthService: " & Err.Description
    Set CreateAuthService = Nothing
End Function

