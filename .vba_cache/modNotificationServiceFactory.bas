Attribute VB_Name = "modNotificationServiceFactory"
Option Compare Database
Option Explicit



' =====================================================
' MÓDULO: modNotificationServiceFactory
' DESCRIPCIÓN: Factory especializada para la creación del servicio de notificaciones
' AUTOR: Sistema CONDOR
' FECHA: 2024
' =====================================================

' Función factory para crear y configurar el servicio de notificaciones
Public Function CreateNotificationService() As INotificationService
    On Error GoTo errorHandler
    
    ' Crear las dependencias necesarias usando sus respectivas factorías
    Dim config As IConfig
    Set config = modConfigFactory.CreateConfigService()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    Dim operationLogger As IOperationLogger
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' Crear el repositorio usando el factory
    Dim notificationRepository As INotificationRepository
    Set notificationRepository = modRepositoryFactory.CreateNotificationRepository(config, errorHandler)
    
    ' Crear una instancia de la clase concreta
    Dim notificationServiceInstance As New CNotificationService
    
    ' Inicializar la instancia concreta con todas las dependencias
    Call notificationServiceInstance.Initialize(config, operationLogger, notificationRepository, errorHandler)
    
    ' Devolver la instancia inicializada como el tipo de la interfaz
    Set CreateNotificationService = notificationServiceInstance
    
    Exit Function
    
errorHandler:
    Debug.Print "Error en modNotificationServiceFactory.CreateNotificationService: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Function



