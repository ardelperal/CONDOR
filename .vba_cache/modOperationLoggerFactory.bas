Attribute VB_Name = "modOperationLoggerFactory"
Option Compare Database
Option Explicit



' Módulo: modOperationLoggerFactory
' Descripción: Factory para la creación de servicios de logging de operaciones.

Public Function CreateOperationLogger() As IOperationLogger
    On Error GoTo errorHandler
    
    Dim config As IConfig
    Set config = modTestContext.GetTestConfig()
    
    ' Crear dependencias
    Dim errorHandler As IErrorHandlerService
    Dim fileSystem As IFileSystem
    
    Set fileSystem = modFileSystemFactory.CreateFileSystem()
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    ' Crear instancia real del logger
    Dim loggerInstance As COperationLogger
    Dim repositoryInstance As COperationRepository
    
    Set loggerInstance = New COperationLogger
    Set repositoryInstance = New COperationRepository
    
    ' Inicializar el repositorio con la configuración y errorHandler
    repositoryInstance.Initialize config, errorHandler
    
    ' Inyectar las dependencias en el logger
    loggerInstance.Initialize config, repositoryInstance, errorHandler
    
    Set CreateOperationLogger = loggerInstance
    
    Exit Function
    
errorHandler:
    Debug.Print "Error en modOperationLoggerFactory.CreateOperationLogger: " & Err.Description
    Err.Raise Err.Number, "modOperationLoggerFactory.CreateOperationLogger", Err.Description
End Function




