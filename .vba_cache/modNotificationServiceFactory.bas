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
Public Function CreateNotificationService(ByVal errorHandler As IErrorHandlerService) As INotificationService
    On Error GoTo ErrorHandler
    
    ' Obtener las dependencias requeridas
    Dim config As IConfig
    Dim operationLogger As IOperationLogger
    Dim notificationRepository As INotificationRepository
    
    Set config = modConfig.CreateConfigService(errorHandler)
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger(errorHandler)
    
    ' Crear el repositorio de notificaciones
    Dim repositoryInstance As New CNotificationRepository
    repositoryInstance.Initialize config
    Set notificationRepository = repositoryInstance
    
    ' Crear una instancia de la clase concreta
    Dim notificationServiceInstance As New CNotificationService
    
    ' Inicializar la instancia concreta con todas las dependencias
    notificationServiceInstance.Initialize config, operationLogger, notificationRepository
    
    ' Devolver la instancia inicializada como el tipo de la interfaz
    Set CreateNotificationService = notificationServiceInstance
    
    Exit Function
    
ErrorHandler:
    errorHandler.LogError Err.Number, Err.Description, "modNotificationServiceFactory.CreateNotificationService"
    Set CreateNotificationService = Nothing
End Function

