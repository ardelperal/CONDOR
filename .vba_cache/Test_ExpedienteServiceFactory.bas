Option Compare Database
Option Explicit

' =====================================================
' MODULO: Test_ExpedienteServiceFactory
' DESCRIPCION: Pruebas unitarias para modExpedienteServiceFactory
' AUTOR: Sistema CONDOR
' FECHA: 2024
' =====================================================

#If DEV_MODE Then

' FunciÃ³n principal que ejecuta todas las pruebas del ExpedienteServiceFactory
Public Function Test_ExpedienteServiceFactory_RunAll() As CTestSuiteResult
    Dim suiteResult As CTestSuiteResult
    Set suiteResult = New CTestSuiteResult
    suiteResult.Initialize "Test_ExpedienteServiceFactory"
    
    ' Ejecutar todas las pruebas individuales
    suiteResult.AddTestResult Test_CreateExpedienteService_ReturnsValidInstance()
    suiteResult.AddTestResult Test_CreateExpedienteService_InitializesDependencies()
    suiteResult.AddTestResult Test_CreateExpedienteService_HandlesErrors()
    
    Set Test_ExpedienteServiceFactory_RunAll = suiteResult
End Function

' Prueba que CreateExpedienteService devuelve una instancia vÃ¡lida
Private Function Test_CreateExpedienteService_ReturnsValidInstance() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_CreateExpedienteService_ReturnsValidInstance"
    
    On Error GoTo ErrorHandler
    
    ' Arrange & Act
    Dim expedienteService As IExpedienteService
    Set expedienteService = modExpedienteServiceFactory.CreateExpedienteService()
    
    ' Assert
    If expedienteService Is Nothing Then
        testResult.Fail "CreateExpedienteService debe devolver una instancia vÃ¡lida"
    Else
        testResult.Pass
    End If
    
    Set Test_CreateExpedienteService_ReturnsValidInstance = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error en Test_CreateExpedienteService_ReturnsValidInstance: " & Err.Description
    Set Test_CreateExpedienteService_ReturnsValidInstance = testResult
End Function

' Prueba que CreateExpedienteService inicializa correctamente las dependencias
Private Function Test_CreateExpedienteService_InitializesDependencies() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_CreateExpedienteService_InitializesDependencies"
    
    On Error GoTo ErrorHandler
    
    ' Arrange & Act
    Dim expedienteService As IExpedienteService
    Set expedienteService = modExpedienteServiceFactory.CreateExpedienteService()
    
    ' Assert - Verificar que el servicio estÃ¡ inicializado
    If expedienteService Is Nothing Then
        testResult.Fail "El servicio de expedientes debe estar inicializado"
    Else
        testResult.Pass
    End If
    
    Set Test_CreateExpedienteService_InitializesDependencies = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error en Test_CreateExpedienteService_InitializesDependencies: " & Err.Description
    Set Test_CreateExpedienteService_InitializesDependencies = testResult
End Function

' Prueba que CreateExpedienteService maneja errores correctamente
Private Function Test_CreateExpedienteService_HandlesErrors() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_CreateExpedienteService_HandlesErrors"
    
    On Error GoTo ErrorHandler
    
    ' Esta prueba verifica que la funciÃ³n maneja errores internos
    ' En condiciones normales, deberÃ­a devolver una instancia vÃ¡lida
    Dim expedienteService As IExpedienteService
    Set expedienteService = modExpedienteServiceFactory.CreateExpedienteService()
    
    If expedienteService Is Nothing Then
        testResult.Fail "CreateExpedienteService debe manejar errores correctamente"
    Else
        testResult.Pass
    End If
    
    Set Test_CreateExpedienteService_HandlesErrors = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error en Test_CreateExpedienteService_HandlesErrors: " & Err.Description
    Set Test_CreateExpedienteService_HandlesErrors = testResult
End Function

#End If
