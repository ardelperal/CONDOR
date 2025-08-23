Option Compare Database
Option Explicit

' =====================================================
' MODULO: modNotificationServiceFactory
' DESCRIPCION: Factory especializada para la creación del servicio de notificaciones
' AUTOR: Sistema CONDOR
' FECHA: 2024
' =====================================================

' Función factory para crear y configurar el servicio de notificaciones
Public Function CreateNotificationService() As INotificationService
    On Error GoTo ErrorHandler
    
    ' Obtener las dependencias requeridas
    Dim config As IConfig
    Dim operationLogger As IOperationLogger
    
    Set config = modConfigFactory.CreateConfig()
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' Crear una instancia de la clase concreta
    Dim notificationServiceInstance As New CNotificationService
    
    ' Inicializar la instancia concreta con ambas dependencias
    notificationServiceInstance.Initialize config, operationLogger
    
    ' Devolver la instancia inicializada como el tipo de la interfaz
    Set CreateNotificationService = notificationServiceInstance
    
    Exit Function
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "modNotificationServiceFactory.CreateNotificationService")
    Set CreateNotificationService = Nothing
End Function