Attribute VB_Name = "modOperationLoggerFactory"
Option Compare Database
Option Explicit


' Módulo: modOperationLoggerFactory
' Descripción: Factory para la creación de servicios de logging de operaciones.

Private g_MockLogger As IOperationLogger ' Para inyectar un mock en tests

Public Function CreateOperationLogger() As IOperationLogger
    On Error GoTo ErrorHandler
    
    ' Crear dependencias internamente
    Dim errorHandler As IErrorHandlerService
    Dim configService As IConfig
    Dim fileSystem As IFileSystem
    
    Set configService = modConfig.CreateConfigService()
    Set fileSystem = modFileSystemFactory.CreateFileSystem()
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(configService, fileSystem)
    
    ' Crear instancia real del logger
    Dim loggerInstance As COperationLogger
    Dim repositoryInstance As COperationRepository
    
    Set loggerInstance = New COperationLogger
    Set repositoryInstance = New COperationRepository
    
    ' Inicializar el repositorio con la configuración y errorHandler
    repositoryInstance.Initialize configService, errorHandler
    
    ' Inyectar las dependencias en el logger
    loggerInstance.Initialize configService, repositoryInstance, errorHandler
    
    Set CreateOperationLogger = loggerInstance
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error en modOperationLoggerFactory.CreateOperationLogger: " & Err.Description
    Err.Raise Err.Number, "modOperationLoggerFactory.CreateOperationLogger", Err.Description
End Function

' Método para configurar el mock logger en tests
Public Sub SetMockLogger(ByVal mockLogger As IOperationLogger)
    On Error GoTo ErrorHandler
    Set g_MockLogger = mockLogger
    Exit Sub
ErrorHandler:
    Debug.Print "Error en modOperationLoggerFactory.SetMockLogger: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' Método para resetear el mock logger
Public Sub ResetMockLogger()
    On Error GoTo ErrorHandler
    Set g_MockLogger = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print "Error en modOperationLoggerFactory.ResetMockLogger: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


