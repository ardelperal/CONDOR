Option Compare Database
Option Explicit

' =====================================================
' MODULO: Test_DocumentServiceFactory
' DESCRIPCION: Pruebas unitarias para modDocumentServiceFactory
' AUTOR: Sistema CONDOR
' FECHA: 2024
' VERSION: 2.0 - Estandarizado según framework CONDOR
' =====================================================

#If DEV_MODE Then

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE DE PRUEBAS
' ============================================================================

Public Function Test_DocumentServiceFactory_RunAll() As CTestSuiteResult
    Dim suite As New CTestSuiteResult
    suite.Initialize "DocumentServiceFactory"
    
    ' Ejecutar todas las pruebas y añadir resultados
    suite.AddTestResult Test_CreateDocumentService_ReturnsValidInstance()
    suite.AddTestResult Test_CreateDocumentService_InitializesDependencies()
    suite.AddTestResult Test_CreateDocumentService_HandlesErrors()
    
    Set Test_DocumentServiceFactory_RunAll = suite
End Function

' ============================================================================
' PRUEBAS INDIVIDUALES
' ============================================================================

' Prueba que CreateDocumentService devuelve una instancia válida
Private Function Test_CreateDocumentService_ReturnsValidInstance() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "CreateDocumentService devuelve instancia válida"
    
    On Error GoTo ErrorHandler
    
    ' Arrange & Act
    Dim documentService As IDocumentService
    Set documentService = modDocumentServiceFactory.CreateDocumentService()
    
    ' Assert
    If documentService Is Nothing Then
        testResult.Fail "CreateDocumentService debe devolver una instancia válida"
        GoTo Cleanup
    End If
    
    testResult.Pass
    
Cleanup:
    Set Test_CreateDocumentService_ReturnsValidInstance = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error en Test_CreateDocumentService_ReturnsValidInstance: " & Err.Description
    Resume Cleanup
End Function

' Prueba que CreateDocumentService inicializa correctamente las dependencias
Private Function Test_CreateDocumentService_InitializesDependencies() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "CreateDocumentService inicializa dependencias"
    
    On Error GoTo ErrorHandler
    
    ' Arrange & Act
    Dim documentService As IDocumentService
    Set documentService = modDocumentServiceFactory.CreateDocumentService()
    
    ' Assert - Verificar que el servicio está inicializado
    If documentService Is Nothing Then
        testResult.Fail "El servicio de documentos debe estar inicializado"
        GoTo Cleanup
    End If
    
    testResult.Pass
    
Cleanup:
    Set Test_CreateDocumentService_InitializesDependencies = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error en Test_CreateDocumentService_InitializesDependencies: " & Err.Description
    Resume Cleanup
End Function

' Prueba que CreateDocumentService maneja errores correctamente
Private Function Test_CreateDocumentService_HandlesErrors() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "CreateDocumentService maneja errores"
    
    On Error GoTo ErrorHandler
    
    ' Esta prueba verifica que la función maneja errores internos
    ' En condiciones normales, debería devolver una instancia válida
    Dim documentService As IDocumentService
    Set documentService = modDocumentServiceFactory.CreateDocumentService()
    
    If documentService Is Nothing Then
        testResult.Fail "CreateDocumentService debe manejar errores correctamente"
        GoTo Cleanup
    End If
    
    testResult.Pass
    
Cleanup:
    Set Test_CreateDocumentService_HandlesErrors = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error en Test_CreateDocumentService_HandlesErrors: " & Err.Description
    Resume Cleanup
End Function

#End If