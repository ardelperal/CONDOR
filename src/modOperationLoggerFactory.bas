Attribute VB_Name = "modOperationLoggerFactory"
Option Compare Database
Option Explicit



' Módulo: modOperationLoggerFactory
' Descripción: Factory para la creación de servicios de logging de operaciones.

Public Function CreateOperationLogger(Optional ByVal config As IConfig = Nothing) As IOperationLogger
    On Error GoTo errorHandler
    
    Dim effectiveConfig As IConfig
    If config Is Nothing Then
        ' Si no se pasa una configuración, usar la global por defecto
        Set effectiveConfig = modTestContext.GetTestConfig()
    Else
        ' Si se pasa una configuración (desde un test), usarla
        Set effectiveConfig = config
    End If
    
    ' Crear dependencias
    Dim errorHandler As IErrorHandlerService
    Dim fileSystem As IFileSystem
    
    Set fileSystem = modFileSystemFactory.CreateFileSystem(effectiveConfig)
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(effectiveConfig)
    
    ' Crear instancia real del logger
    Dim loggerInstance As COperationLogger
    Dim repositoryInstance As COperationRepository
    
    Set loggerInstance = New COperationLogger
    Set repositoryInstance = New COperationRepository
    
    ' Inicializar el repositorio con la configuración y errorHandler
    repositoryInstance.Initialize effectiveConfig, errorHandler
    
    ' Inyectar las dependencias en el logger
    loggerInstance.Initialize effectiveConfig, repositoryInstance, errorHandler
    
    Set CreateOperationLogger = loggerInstance
    
    Exit Function
    
errorHandler:
    Debug.Print "Error en modOperationLoggerFactory.CreateOperationLogger: " & Err.Description
    Err.Raise Err.Number, "modOperationLoggerFactory.CreateOperationLogger", Err.Description
End Function




