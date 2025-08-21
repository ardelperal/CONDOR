Option Compare Database
Option Explicit

' Módulo: modOperationLoggerFactory
' Descripción: Factory para la creación de servicios de logging de operaciones.

Private g_MockLogger As IOperationLogger ' Para inyectar un mock en tests

Public Function CreateOperationLogger() As IOperationLogger
    On Error GoTo ErrorHandler
    
    If Not g_MockLogger Is Nothing Then
        Set CreateOperationLogger = g_MockLogger ' Devolver el mock si está configurado
    Else
        Dim loggerInstance As COperationLogger
        Set loggerInstance = New COperationLogger
        
        ' Crear la dependencia de configuración
        Dim configService As IConfig
        Set configService = modConfigFactory.CreateConfigService() ' Asumiendo que existe un factory para config
        
        ' Inyectar la dependencia
        loggerInstance.Initialize configService
        
        Set CreateOperationLogger = loggerInstance
    End If
    
    Exit Function
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "modOperationLoggerFactory.CreateOperationLogger")
    Set CreateOperationLogger = Nothing
End Function

' Método para configurar el mock logger en tests
Public Sub SetMockLogger(ByVal mockLogger As IOperationLogger)
    On Error GoTo ErrorHandler
    Set g_MockLogger = mockLogger
    Exit Sub
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "modOperationLoggerFactory.SetMockLogger")
End Sub

' Método para resetear el mock logger
Public Sub ResetMockLogger()
    On Error GoTo ErrorHandler
    Set g_MockLogger = Nothing
    Exit Sub
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "modOperationLoggerFactory.ResetMockLogger")
End Sub
