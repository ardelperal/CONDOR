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
Public Function CreateNotificationService(ByVal config As IConfig, ByVal operationLogger As IOperationLogger, ByVal errorHandler As IErrorHandlerService) As INotificationService
    On Error GoTo ErrorHandler
    
    ' Crear el repositorio usando el factory
    Dim notificationRepository As INotificationRepository
    Set notificationRepository = modRepositoryFactory.CreateNotificationRepository(errorHandler, config)
    
    ' Crear una instancia de la clase concreta
    Dim notificationServiceInstance As New CNotificationService
    
    ' Inicializar la instancia concreta con todas las dependencias
    Call notificationServiceInstance.Initialize(config, operationLogger, notificationRepository, errorHandler)
    
    ' Devolver la instancia inicializada como el tipo de la interfaz
    Set CreateNotificationService = notificationServiceInstance
    
    Exit Function
    
ErrorHandler:
    errorHandler.LogError Err.Number, Err.Description, "modNotificationServiceFactory.CreateNotificationService"
    Set CreateNotificationService = Nothing
End Function

