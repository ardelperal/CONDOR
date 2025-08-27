Attribute VB_Name = "modOperationLoggerFactory"
Option Compare Database
Option Explicit


' Módulo: modOperationLoggerFactory
' Descripción: Factory para la creación de servicios de logging de operaciones.

Private g_MockLogger As IOperationLogger ' Para inyectar un mock en tests

Public Function CreateOperationLogger(ByVal errorHandler As IErrorHandlerService) As IOperationLogger
    On Error GoTo ErrorHandler
    
    ' Obtener la instancia de configuración
    Dim configService As IConfig
    Set configService = modConfig.CreateConfigService(errorHandler)
    
    ' Decidir si usar mock o clase concreta basado en DEV_MODE
    If CBool(configService.GetValue("DEV_MODE")) Then
        ' Modo desarrollo - usar mock
        Dim mockLogger As New CMockOperationLogger
        Set CreateOperationLogger = mockLogger
    Else
        ' Modo producción - usar clase concreta
        Dim loggerInstance As COperationLogger
        Dim repositoryInstance As COperationRepository
        
        Set loggerInstance = New COperationLogger
        Set repositoryInstance = New COperationRepository
        
        ' Inicializar el repositorio con la configuración y errorHandler
        repositoryInstance.Initialize configService, errorHandler
        
        ' Inyectar las dependencias en el logger
        loggerInstance.Initialize configService, repositoryInstance, errorHandler
        
        Set CreateOperationLogger = loggerInstance
    End If
    
    Exit Function
    
ErrorHandler:
    ' Usar el manejador de errores inyectado
    errorHandler.LogError Err.Number, Err.Description, "modOperationLoggerFactory.CreateOperationLogger", True ' Mark as critical
    Err.Raise Err.Number, Err.Source, Err.Description
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


