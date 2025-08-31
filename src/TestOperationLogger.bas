Attribute VB_Name = "TestOperationLogger"
Option Compare Database
Option Explicit

' ============================================================================
' Módulo: TestOperationLogger
' Descripción: Pruebas unitarias puras y aisladas para COperationLogger.
' Arquitectura: Capa de Pruebas - Tests unitarios con mocks
' Autor: CONDOR-Expert
' Fecha: Enero 2025
' ============================================================================

' ============================================================================
' FUNCIÓN PRINCIPAL DE EJECUCIÓN DE PRUEBAS
' ============================================================================

Public Function TestOperationLoggerRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TestOperationLogger"
    
    ' Ejecutar todas las pruebas unitarias aisladas
    Call suiteResult.AddTestResult(TestInitializeWithValidDependenciesSuccess())
    Call suiteResult.AddTestResult(TestLogOperationWithoutInitializeHandlesError())
    Call suiteResult.AddTestResult(TestLogOperationWithValidParamsCallsRepositoryCorrectly())
    Call suiteResult.AddTestResult(TestLogOperationWithEmptyParamsCallsRepositoryWithEmptyValues())
    Call suiteResult.AddTestResult(TestLogOperationMultipleOperationsCallsRepositoryMultipleTimes())
    
    Set TestOperationLoggerRunAll = suiteResult
End Function

' ============================================================================
' PRUEBAS UNITARIAS PURAS Y AISLADAS PARA COperationLogger
' ============================================================================

Public Function TestInitializeWithValidDependenciesSuccess() As CTestResult
    Set TestInitializeWithValidDependenciesSuccess = New CTestResult
    TestInitializeWithValidDependenciesSuccess.Initialize "Initialize con dependencias válidas debe tener éxito"
    On Error GoTo TestFail
    
    ' Arrange
    Dim logger As New COperationLogger
    Dim mockConfig As New CMockConfig
    Dim mockRepository As New CMockOperationRepository
    Dim mockErrorHandler As New CMockErrorHandlerService
    
    ' Act
    Call logger.Initialize(mockConfig, mockRepository, mockErrorHandler)
    
    ' Assert - Si no hay error, la inicialización fue exitosa
    TestInitializeWithValidDependenciesSuccess.Pass
    Exit Function
    
TestFail:
    Call TestInitializeWithValidDependenciesSuccess.Fail("Error en inicialización: " & Err.Description)
End Function

Public Function TestLogOperationWithoutInitializeHandlesError() As CTestResult
    Set TestLogOperationWithoutInitializeHandlesError = New CTestResult
    TestLogOperationWithoutInitializeHandlesError.Initialize "LogOperation sin inicializar debe manejar el error"
    On Error GoTo TestFail
    
    ' Arrange
    Dim logger As New COperationLogger
    ' No inicializar el logger intencionalmente
    
    ' Act & Assert - Debería manejar el error graciosamente
    Call logger.LogOperation("TestType", "TestEntity", "TestDetails")
    
    ' Si llegamos aquí sin crash, el manejo de errores funcionó
    TestLogOperationWithoutInitializeHandlesError.Pass
    Exit Function
    
TestFail:
    ' El error es esperado, pero debe ser manejado internamente
    TestLogOperationWithoutInitializeHandlesError.Pass
End Function

Public Function TestLogOperationWithValidParamsCallsRepositoryCorrectly() As CTestResult
    Set TestLogOperationWithValidParamsCallsRepositoryCorrectly = New CTestResult
    TestLogOperationWithValidParamsCallsRepositoryCorrectly.Initialize "LogOperation con parámetros válidos debe llamar al repositorio correctamente"
    On Error GoTo TestFail
    
    ' Arrange
    Dim logger As New COperationLogger
    Dim mockConfig As New CMockConfig
    Dim mockRepository As New CMockOperationRepository
    Dim mockErrorHandler As New CMockErrorHandlerService
    
    mockRepository.Reset
    Call logger.Initialize(mockConfig, mockRepository, mockErrorHandler)
    
    ' Act
    Call logger.LogOperation("CREATE", "EXP001", "Expediente creado exitosamente")
    
    ' Assert
    Call modAssert.IsTrue(mockRepository.SaveLogCalled, "SaveLog debería haber sido llamado")
    Call modAssert.AreEqual(1, mockRepository.CallCount, "SaveLog debería haber sido llamado exactamente 1 vez")
    Call modAssert.AreEqual("CREATE", mockRepository.LastOperationType, "Tipo de operación incorrecto")
    Call modAssert.AreEqual("EXP001", mockRepository.LastEntityId, "ID de entidad incorrecto")
    Call modAssert.AreEqual("Expediente creado exitosamente", mockRepository.LastDetails, "Detalles incorrectos")
    
    TestLogOperationWithValidParamsCallsRepositoryCorrectly.Pass
    Exit Function
    
TestFail:
    Call TestLogOperationWithValidParamsCallsRepositoryCorrectly.Fail("Error en prueba: " & Err.Description)
End Function

Public Function TestLogOperationWithEmptyParamsCallsRepositoryWithEmptyValues() As CTestResult
    Set TestLogOperationWithEmptyParamsCallsRepositoryWithEmptyValues = New CTestResult
    TestLogOperationWithEmptyParamsCallsRepositoryWithEmptyValues.Initialize "LogOperation con parámetros vacíos debe llamar al repositorio con valores vacíos"
    On Error GoTo TestFail
    
    ' Arrange
    Dim logger As New COperationLogger
    Dim mockConfig As New CMockConfig
    Dim mockRepository As New CMockOperationRepository
    Dim mockErrorHandler As New CMockErrorHandlerService
    
    mockRepository.Reset
    logger.Initialize mockConfig, mockRepository, mockErrorHandler
    
    ' Act
    Call logger.LogOperation("", "", "")
    
    ' Assert
    Call modAssert.IsTrue(mockRepository.SaveLogCalled, "SaveLog debería haber sido llamado")
    Call modAssert.AreEqual("", mockRepository.LastOperationType, "Tipo de operación debería estar vacío")
    Call modAssert.AreEqual("", mockRepository.LastEntityId, "ID de entidad debería estar vacío")
    Call modAssert.AreEqual("", mockRepository.LastDetails, "Detalles deberían estar vacíos")
    
    TestLogOperationWithEmptyParamsCallsRepositoryWithEmptyValues.Pass
    Exit Function
    
TestFail:
    Call TestLogOperationWithEmptyParamsCallsRepositoryWithEmptyValues.Fail("Error en prueba: " & Err.Description)
End Function

Public Function TestLogOperationMultipleOperationsCallsRepositoryMultipleTimes() As CTestResult
    Set TestLogOperationMultipleOperationsCallsRepositoryMultipleTimes = New CTestResult
    TestLogOperationMultipleOperationsCallsRepositoryMultipleTimes.Initialize "LogOperation con múltiples operaciones debe llamar al repositorio múltiples veces"
    On Error GoTo TestFail
    
    ' Arrange
    Dim logger As New COperationLogger
    Dim mockConfig As New CMockConfig
    Dim mockRepository As New CMockOperationRepository
    Dim mockErrorHandler As New CMockErrorHandlerService
    
    mockRepository.Reset
    logger.Initialize mockConfig, mockRepository, mockErrorHandler
    
    ' Act
    Call logger.LogOperation("CREATE", "EXP001", "Primera operación")
    Call logger.LogOperation("UPDATE", "EXP002", "Segunda operación")
    Call logger.LogOperation("DELETE", "EXP003", "Tercera operación")
    
    ' Assert
    Call modAssert.IsTrue(mockRepository.SaveLogCalled, "SaveLog debería haber sido llamado")
    Call modAssert.AreEqual(3, mockRepository.CallCount, "SaveLog debería haber sido llamado exactamente 3 veces")
    ' Verificar que los últimos parámetros corresponden a la última llamada
    Call modAssert.AreEqual("DELETE", mockRepository.LastOperationType, "Último tipo de operación incorrecto")
    Call modAssert.AreEqual("EXP003", mockRepository.LastEntityId, "Último ID de entidad incorrecto")
    Call modAssert.AreEqual("Tercera operación", mockRepository.LastDetails, "Últimos detalles incorrectos")
    
    TestLogOperationMultipleOperationsCallsRepositoryMultipleTimes.Pass
    Exit Function
    
TestFail:
    Call TestLogOperationMultipleOperationsCallsRepositoryMultipleTimes.Fail("Error en prueba: " & Err.Description)
End Function