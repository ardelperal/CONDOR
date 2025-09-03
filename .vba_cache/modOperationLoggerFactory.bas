Attribute VB_Name = "modOperationLoggerFactory"
Option Compare Database
Option Explicit



' Módulo: modOperationLoggerFactory
' Descripción: Factory para la creación de servicios de logging de operaciones.

Public Function CreateOperationLogger(Optional ByVal config As IConfig = Nothing) As IOperationLogger
    On Error GoTo errorHandler
    
    ' Determinar configuración final
    Dim finalConfig As IConfig
    If config Is Nothing Then
        Set finalConfig = modConfigFactory.CreateConfigService()
    Else
        Set finalConfig = config
    End If
    
    ' Crear dependencias propagando la configuración
    Dim errorHandler As IErrorHandlerService
    Dim fileSystem As IFileSystem
    
    Set fileSystem = modFileSystemFactory.CreateFileSystem(finalConfig)
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(finalConfig)
    
    ' Crear instancia real del logger
    Dim loggerInstance As COperationLogger
    Dim repositoryInstance As COperationRepository
    
    Set loggerInstance = New COperationLogger
    Set repositoryInstance = New COperationRepository
    
    ' Inicializar el repositorio con la configuración y errorHandler
    repositoryInstance.Initialize finalConfig, errorHandler
    
    ' Inyectar las dependencias en el logger
    loggerInstance.Initialize finalConfig, repositoryInstance, errorHandler
    
    Set CreateOperationLogger = loggerInstance
    
    Exit Function
    
errorHandler:
    Debug.Print "Error en modOperationLoggerFactory.CreateOperationLogger: " & Err.Description
    Err.Raise Err.Number, "modOperationLoggerFactory.CreateOperationLogger", Err.Description
End Function




