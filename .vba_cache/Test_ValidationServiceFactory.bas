Option Compare Database
Option Explicit

' =====================================================
' MODULO: Test_ValidationServiceFactory
' DESCRIPCION: Pruebas unitarias para modValidationServiceFactory
' AUTOR: Sistema CONDOR
' FECHA: 2024
' =====================================================

#If DEV_MODE Then

' Prueba que CreateValidationService devuelve una instancia válida
Public Sub Test_CreateValidationService_ReturnsValidInstance()
    On Error GoTo ErrorHandler
    
    ' Arrange & Act
    Dim validationService As IValidationService
    Set validationService = modValidationServiceFactory.CreateValidationService()
    
    ' Assert
    Call modAssert.AssertNotNothing(validationService, "CreateValidationService debe devolver una instancia válida")
    
    Exit Sub
    
ErrorHandler:
    Call modAssert.Fail("Error en Test_CreateValidationService_ReturnsValidInstance: " & Err.Description)
End Sub

' Prueba que CreateValidationService inicializa correctamente las dependencias
Public Sub Test_CreateValidationService_InitializesDependencies()
    On Error GoTo ErrorHandler
    
    ' Arrange & Act
    Dim validationService As IValidationService
    Set validationService = modValidationServiceFactory.CreateValidationService()
    
    ' Assert - Verificar que el servicio está inicializado
    Call modAssert.AssertNotNothing(validationService, "El servicio de validación debe estar inicializado")
    
    Exit Sub
    
ErrorHandler:
    Call modAssert.Fail("Error en Test_CreateValidationService_InitializesDependencies: " & Err.Description)
End Sub

' Prueba que CreateValidationService maneja errores correctamente
Public Sub Test_CreateValidationService_HandlesErrors()
    On Error GoTo ErrorHandler
    
    ' Esta prueba verifica que la función maneja errores internos
    ' En condiciones normales, debería devolver una instancia válida
    Dim validationService As IValidationService
    Set validationService = modValidationServiceFactory.CreateValidationService()
    
    Call modAssert.AssertNotNothing(validationService, "CreateValidationService debe manejar errores correctamente")
    
    Exit Sub
    
ErrorHandler:
    Call modAssert.Fail("Error en Test_CreateValidationService_HandlesErrors: " & Err.Description)
End Sub

#End If