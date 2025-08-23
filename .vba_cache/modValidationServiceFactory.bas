Option Compare Database
Option Explicit

' =====================================================
' MODULO: modValidationServiceFactory
' DESCRIPCION: Factory especializada para la creación del servicio de validación
' AUTOR: Sistema CONDOR
' FECHA: 2024
' =====================================================

' Función factory para crear y configurar el servicio de validación
Public Function CreateValidationService() As IValidationService
    On Error GoTo ErrorHandler
    
    ' Obtener la instancia del logger de operaciones
    Dim operationLogger As IOperationLogger
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' Crear una instancia de la clase concreta
    Dim validationServiceInstance As New CValidationService
    
    ' Inicializar la instancia concreta con la dependencia
    validationServiceInstance.Initialize operationLogger
    
    ' Devolver la instancia inicializada como el tipo de la interfaz
    Set CreateValidationService = validationServiceInstance
    
    Exit Function
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "modValidationServiceFactory.CreateValidationService")
    Set CreateValidationService = Nothing
End Function