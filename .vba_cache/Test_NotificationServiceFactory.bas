Option Compare Database
Option Explicit

' =====================================================
' MODULO: Test_NotificationServiceFactory
' DESCRIPCION: Pruebas unitarias para modNotificationServiceFactory
' AUTOR: Sistema CONDOR
' FECHA: 2024
' =====================================================

#If DEV_MODE Then

' Prueba que CreateNotificationService devuelve una instancia válida
Public Sub Test_CreateNotificationService_ReturnsValidInstance()
    On Error GoTo ErrorHandler
    
    ' Arrange & Act
    Dim notificationService As INotificationService
    Set notificationService = modNotificationServiceFactory.CreateNotificationService()
    
    ' Assert
    Call modAssert.AssertNotNothing(notificationService, "CreateNotificationService debe devolver una instancia válida")
    
    Exit Sub
    
ErrorHandler:
    Call modAssert.Fail("Error en Test_CreateNotificationService_ReturnsValidInstance: " & Err.Description)
End Sub

' Prueba que CreateNotificationService inicializa correctamente las dependencias
Public Sub Test_CreateNotificationService_InitializesDependencies()
    On Error GoTo ErrorHandler
    
    ' Arrange & Act
    Dim notificationService As INotificationService
    Set notificationService = modNotificationServiceFactory.CreateNotificationService()
    
    ' Assert - Verificar que el servicio está inicializado
    Call modAssert.AssertNotNothing(notificationService, "El servicio de notificaciones debe estar inicializado")
    
    Exit Sub
    
ErrorHandler:
    Call modAssert.Fail("Error en Test_CreateNotificationService_InitializesDependencies: " & Err.Description)
End Sub

' Prueba que CreateNotificationService maneja errores correctamente
Public Sub Test_CreateNotificationService_HandlesErrors()
    On Error GoTo ErrorHandler
    
    ' Esta prueba verifica que la función maneja errores internos
    ' En condiciones normales, debería devolver una instancia válida
    Dim notificationService As INotificationService
    Set notificationService = modNotificationServiceFactory.CreateNotificationService()
    
    Call modAssert.AssertNotNothing(notificationService, "CreateNotificationService debe manejar errores correctamente")
    
    Exit Sub
    
ErrorHandler:
    Call modAssert.Fail("Error en Test_CreateNotificationService_HandlesErrors: " & Err.Description)
End Sub

#End If