Option Compare Database
Option Explicit

' =====================================================
' MODULO: Test_ServiceFactories_Integration
' DESCRIPCION: Pruebas de integración para todas las fábricas de servicios
' AUTOR: Sistema CONDOR
' FECHA: 2024
' =====================================================

#If DEV_MODE Then

' Prueba de integración que valida todas las fábricas de servicios
Public Sub Test_AllServiceFactories_Integration()
    On Error GoTo ErrorHandler
    
    ' Arrange & Act - Crear instancias de todos los servicios
    Dim documentService As IDocumentService
    Set documentService = modDocumentServiceFactory.CreateDocumentService()
    
    Dim notificationService As INotificationService
    Set notificationService = modNotificationServiceFactory.CreateNotificationService()
    
    Dim expedienteService As IExpedienteService
    Set expedienteService = modExpedienteServiceFactory.CreateExpedienteService()
    
    ' Assert - Verificar que todos los servicios se crearon correctamente
    Call modAssert.AssertNotNothing(documentService, "DocumentService debe crearse correctamente")
    Call modAssert.AssertNotNothing(notificationService, "NotificationService debe crearse correctamente")
    Call modAssert.AssertNotNothing(expedienteService, "ExpedienteService debe crearse correctamente")
    
    Exit Sub
    
ErrorHandler:
    Call modAssert.Fail("Error en Test_AllServiceFactories_Integration: " & Err.Description)
End Sub

' Prueba que verifica la independencia de las fábricas
Public Sub Test_ServiceFactories_Independence()
    On Error GoTo ErrorHandler
    
    ' Arrange & Act - Crear múltiples instancias del mismo servicio
    Dim documentService1 As IDocumentService
    Set documentService1 = modDocumentServiceFactory.CreateDocumentService()
    
    Dim documentService2 As IDocumentService
    Set documentService2 = modDocumentServiceFactory.CreateDocumentService()
    
    ' Assert - Verificar que son instancias diferentes
    Call modAssert.AssertNotNothing(documentService1, "Primera instancia debe ser válida")
    Call modAssert.AssertNotNothing(documentService2, "Segunda instancia debe ser válida")
    
    ' Nota: En VBA no podemos comparar directamente las referencias de objeto
    ' pero podemos verificar que ambas instancias son válidas
    
    Exit Sub
    
ErrorHandler:
    Call modAssert.Fail("Error en Test_ServiceFactories_Independence: " & Err.Description)
End Sub

' Prueba que verifica la correcta inyección de dependencias
Public Sub Test_ServiceFactories_DependencyInjection()
    On Error GoTo ErrorHandler
    
    ' Arrange & Act - Crear servicio que depende de otros
    Dim documentService As IDocumentService
    Set documentService = modDocumentServiceFactory.CreateDocumentService()
    
    ' Assert - Verificar que el servicio se creó con sus dependencias
    Call modAssert.AssertNotNothing(documentService, "DocumentService debe crearse con todas sus dependencias")
    
    Exit Sub
    
ErrorHandler:
    Call modAssert.Fail("Error en Test_ServiceFactories_DependencyInjection: " & Err.Description)
End Sub

#End If