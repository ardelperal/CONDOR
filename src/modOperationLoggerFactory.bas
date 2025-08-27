Attribute VB_Name = "modOperationLoggerFactory"
Option Compare Database
Option Explicit


' Módulo: modOperationLoggerFactory
' Descripción: Factory para la creación de servicios de logging de operaciones.

Private g_MockLogger As IOperationLogger ' Para inyectar un mock en tests

Public Function CreateOperationLogger(ByVal errorHandler As IErrorHandlerService) As IOperationLogger
    On Error GoTo ErrorHandler
    
    If Not g_MockLogger Is Nothing Then
        Set CreateOperationLogger = g_MockLogger ' Devolver el mock si está configurado
    Else
        Dim loggerInstance As COperationLogger
        Dim repositoryInstance As COperationRepository
        Dim configService As IConfig
        
        Set loggerInstance = New COperationLogger
        Set repositoryInstance = New COperationRepository
        
        ' Obtener instancia de configuración usando el nuevo factory
        Set configService = modConfig.CreateConfigService(errorHandler) ' Pass errorHandler
        
        ' Inicializar el repositorio con la configuración
        repositoryInstance.Initialize configService
        
        ' Inyectar las dependencias en el logger
        loggerInstance.Initialize configService, repositoryInstance
        
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


