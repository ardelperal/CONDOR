Option Compare Database
Option Explicit

' MÃ³dulo: modOperationLoggerFactory
' DescripciÃ³n: Factory para la creaciÃ³n de servicios de logging de operaciones.

Private g_MockLogger As IOperationLogger ' Para inyectar un mock en tests

Public Function CreateOperationLogger() As IOperationLogger
    On Error GoTo ErrorHandler
    
    If Not g_MockLogger Is Nothing Then
        Set CreateOperationLogger = g_MockLogger ' Devolver el mock si estÃ¡ configurado
    Else
        Dim loggerInstance As COperationLogger
        Dim repositoryInstance As COperationRepository
        Dim configService As IConfig
        
        Set loggerInstance = New COperationLogger
        Set repositoryInstance = New COperationRepository
        
        ' Obtener instancia de configuraciÃ³n usando el nuevo factory
        Set configService = modConfig.CreateConfigService()
        
        ' Inicializar el repositorio con la configuraciÃ³n
        repositoryInstance.Initialize configService
        
        ' Inyectar las dependencias en el logger
        loggerInstance.Initialize configService, repositoryInstance
        
        Set CreateOperationLogger = loggerInstance
    End If
    
    Exit Function
    
ErrorHandler:
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "modOperationLoggerFactory.CreateOperationLogger"
    Set CreateOperationLogger = Nothing
End Function

' MÃ©todo para configurar el mock logger en tests
Public Sub SetMockLogger(ByVal mockLogger As IOperationLogger)
    On Error GoTo ErrorHandler
    Set g_MockLogger = mockLogger
    Exit Sub
ErrorHandler:
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "modOperationLoggerFactory.SetMockLogger"
End Sub

' MÃ©todo para resetear el mock logger
Public Sub ResetMockLogger()
    On Error GoTo ErrorHandler
    Set g_MockLogger = Nothing
    Exit Sub
ErrorHandler:
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "modOperationLoggerFactory.ResetMockLogger"
End Sub
