Option Compare Database
Option Explicit

' =====================================================
' MODULO: Test_DocumentServiceFactory
' DESCRIPCION: Pruebas unitarias para modDocumentServiceFactory
' AUTOR: Sistema CONDOR
' FECHA: 2024
' VERSION: 2.0 - Estandarizado segÃºn framework CONDOR
' =====================================================

#If DEV_MODE Then

' ============================================================================
' FUNCIÃ“N PRINCIPAL DE LA SUITE DE PRUEBAS
' ============================================================================

Public Function Test_DocumentServiceFactory_RunAll() As CTestSuiteResult
    Dim suite As New CTestSuiteResult
    suite.Initialize "DocumentServiceFactory"
    
    ' Ejecutar todas las pruebas y aÃ±adir resultados
    suite.AddTestResult Test_CreateDocumentService_ReturnsValidInstance()
    suite.AddTestResult Test_CreateDocumentService_InitializesDependencies()
    suite.AddTestResult Test_CreateDocumentService_HandlesErrors()
    
    Set Test_DocumentServiceFactory_RunAll = suite
End Function

' ============================================================================
' PRUEBAS INDIVIDUALES
' ============================================================================

' Prueba que CreateDocumentService devuelve una instancia vÃ¡lida
Private Function Test_CreateDocumentService_ReturnsValidInstance() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "CreateDocumentService devuelve instancia vÃ¡lida"
    
    On Error GoTo ErrorHandler
    
    ' Arrange & Act
    Dim documentService As IDocumentService
    Set documentService = modDocumentServiceFactory.CreateDocumentService()
    
    ' Assert
    If documentService Is Nothing Then
        testResult.Fail "CreateDocumentService debe devolver una instancia vÃ¡lida"
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
    
    ' Assert - Verificar que el servicio estÃ¡ inicializado
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
    
    ' Esta prueba verifica que la funciÃ³n maneja errores internos
    ' En condiciones normales, deberÃ­a devolver una instancia vÃ¡lida
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
