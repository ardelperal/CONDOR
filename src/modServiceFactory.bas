Option Compare Database
Option Explicit

' =====================================================
' MODULO: modServiceFactory
' DESCRIPCION: Factory para la creación de servicios principales
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
    Call modErrorHandler.LogError(Err.Number, Err.Description, "modServiceFactory.CreateValidationService")
    Set CreateValidationService = Nothing
End Function

' Función factory para crear y configurar el servicio de documentos
Public Function CreateDocumentService() As IDocumentService
    On Error GoTo ErrorHandler
    
    ' Obtener las dependencias necesarias
    Dim validationService As IValidationService
    Set validationService = CreateValidationService()
    
    Dim solicitudRepository As ISolicitudRepository
    Set solicitudRepository = modRepositoryFactory.CreateSolicitudRepository()
    
    Dim operationLogger As IOperationLogger
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' Crear una instancia de la clase concreta
    Dim documentServiceInstance As New CDocumentService
    
    ' Inicializar la instancia concreta con todas las dependencias
    documentServiceInstance.Initialize validationService, solicitudRepository, operationLogger
    
    ' Devolver la instancia inicializada como el tipo de la interfaz
    Set CreateDocumentService = documentServiceInstance
    
    Exit Function
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "modServiceFactory.CreateDocumentService")
    Set CreateDocumentService = Nothing
End Function

' Función factory para crear y configurar el servicio de notificaciones
Public Function CreateNotificationService() As INotificationService
    On Error GoTo ErrorHandler
    
    ' Obtener la instancia del logger de operaciones
    Dim operationLogger As IOperationLogger
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' Crear una instancia de la clase concreta
    Dim notificationServiceInstance As New CNotificationService
    
    ' Inicializar la instancia concreta con la dependencia
    notificationServiceInstance.Initialize operationLogger
    
    ' Devolver la instancia inicializada como el tipo de la interfaz
    Set CreateNotificationService = notificationServiceInstance
    
    Exit Function
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "modServiceFactory.CreateNotificationService")
    Set CreateNotificationService = Nothing
End Function

' Función factory para crear y configurar el servicio de expedientes
Public Function CreateExpedienteService() As IExpedienteService
    On Error GoTo ErrorHandler
    
    ' Obtener la instancia del logger de operaciones
    Dim operationLogger As IOperationLogger
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' Crear una instancia de la clase concreta
    Dim expedienteServiceInstance As New CExpedienteService
    
    ' Inicializar la instancia concreta con la dependencia
    expedienteServiceInstance.Initialize operationLogger
    
    ' Devolver la instancia inicializada como el tipo de la interfaz
    Set CreateExpedienteService = expedienteServiceInstance
    
    Exit Function
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "modServiceFactory.CreateExpedienteService")
    Set CreateExpedienteService = Nothing
End Function