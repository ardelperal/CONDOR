Option Compare Database
Option Explicit

' Módulo: modConfigFactory
' Descripción: Factory para la creación de servicios de configuración.

Public Function CreateConfigService() As IConfig
    On Error GoTo ErrorHandler
    
    ' Obtener la instancia del logger de operaciones
    Dim operationLogger As IOperationLogger
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' Crear la instancia de configuración
    Dim configInstance As New CConfig
    
    ' Inicializar con la dependencia del logger
    configInstance.Initialize operationLogger
    
    Set CreateConfigService = configInstance
    
    Exit Function
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "modConfigFactory.CreateConfigService")
    Set CreateConfigService = Nothing
End Function
