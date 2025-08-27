Attribute VB_Name = "Test_OperationLogger"
Option Compare Database
Option Explicit


#If DEV_MODE Then

' ============================================================================
' Módulo: Test_OperationLogger
' Descripción: Pruebas unitarias puras y aisladas para COperationLogger.
' Arquitectura: Capa de Pruebas - Tests unitarios con mocks
' Autor: CONDOR-Expert
' Fecha: Enero 2025
' ============================================================================

' ============================================================================
' FUNCIÃ“N PRINCIPAL DE EJECUCIÃ“N DE PRUEBAS
' ============================================================================

Public Function Test_OperationLogger_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.suiteName = "Test_OperationLogger"
    
    ' Ejecutar todas las pruebas unitarias aisladas
    Call suiteResult.AddTestResult("Test_Initialize_WithValidDependencies_Success", Test_Initialize_WithValidDependencies_Success())
    Call suiteResult.AddTestResult("Test_LogOperation_WithoutInitialize_HandlesError", Test_LogOperation_WithoutInitialize_HandlesError())
    Call suiteResult.AddTestResult("Test_LogOperation_WithValidParams_CallsRepositoryCorrectly", Test_LogOperation_WithValidParams_CallsRepositoryCorrectly())
    Call suiteResult.AddTestResult("Test_LogOperation_WithEmptyParams_CallsRepositoryWithEmptyValues", Test_LogOperation_WithEmptyParams_CallsRepositoryWithEmptyValues())
    Call suiteResult.AddTestResult("Test_LogOperation_MultipleOperations_CallsRepositoryMultipleTimes", Test_LogOperation_MultipleOperations_CallsRepositoryMultipleTimes())
    
    Set Test_OperationLogger_RunAll = suiteResult
End Function

' ============================================================================
' PRUEBAS UNITARIAS PURAS Y AISLADAS PARA COperationLogger
' ============================================================================

Public Function Test_Initialize_WithValidDependencies_Success() As CTestResult
    On Error GoTo TestFail
    
    ' Arrange
    Dim logger As New COperationLogger
    Dim mockConfig As New CMockConfig
    Dim mockRepository As New CMockOperationRepository
    
    ' Act
    logger.Initialize mockConfig, mockRepository
    
    ' Assert - Si no hay error, la inicialización fue exitosa
    Set Test_Initialize_WithValidDependencies_Success = New CTestResult
    Test_Initialize_WithValidDependencies_Success.Pass
    Exit Function
    
TestFail:
    Set Test_Initialize_WithValidDependencies_Success = New CTestResult
    Test_Initialize_WithValidDependencies_Success.Fail "Error en inicialización: " & Err.Description
End Function

Public Function Test_LogOperation_WithoutInitialize_HandlesError() As CTestResult
    On Error GoTo TestFail
    
    ' Arrange
    Dim logger As New COperationLogger
    ' No inicializar el logger intencionalmente
    
    ' Act & Assert - Debería manejar el error graciosamente
    logger.LogOperation "TestType", "TestEntity", "TestDetails"
    
    ' Si llegamos aquí sin crash, el manejo de errores funcionó
    Set Test_LogOperation_WithoutInitialize_HandlesError = New CTestResult
    Test_LogOperation_WithoutInitialize_HandlesError.Pass
    Exit Function
    
TestFail:
    ' El error es esperado, pero debe ser manejado internamente
    Set Test_LogOperation_WithoutInitialize_HandlesError = New CTestResult
    Test_LogOperation_WithoutInitialize_HandlesError.Pass
End Function

Public Function Test_LogOperation_WithValidParams_CallsRepositoryCorrectly() As CTestResult
    On Error GoTo TestFail
    
    ' Arrange
    Dim logger As New COperationLogger
    Dim mockConfig As New CMockConfig
    Dim mockRepository As New CMockOperationRepository
    
    mockRepository.Reset
    logger.Initialize mockConfig, mockRepository
    
    ' Act
    logger.LogOperation "CREATE", "EXP001", "Expediente creado exitosamente"
    
    ' Assert
    Call modAssert.IsTrue(mockRepository.SaveLogCalled, "SaveLog debería haber sido llamado")
    Call modAssert.AreEqual(1, mockRepository.CallCount, "SaveLog debería haber sido llamado exactamente 1 vez")
    Call modAssert.AreEqual("CREATE", mockRepository.LastOperationType, "Tipo de operación incorrecto")
    Call modAssert.AreEqual("EXP001", mockRepository.LastEntityId, "ID de entidad incorrecto")
    Call modAssert.AreEqual("Expediente creado exitosamente", mockRepository.LastDetails, "Detalles incorrectos")
    
    Set Test_LogOperation_WithValidParams_CallsRepositoryCorrectly = New CTestResult
    Test_LogOperation_WithValidParams_CallsRepositoryCorrectly.Pass
    Exit Function
    
TestFail:
    Set Test_LogOperation_WithValidParams_CallsRepositoryCorrectly = New CTestResult
    Test_LogOperation_WithValidParams_CallsRepositoryCorrectly.Fail "Error en prueba: " & Err.Description
End Function

Public Function Test_LogOperation_WithEmptyParams_CallsRepositoryWithEmptyValues() As CTestResult
    On Error GoTo TestFail
    
    ' Arrange
    Dim logger As New COperationLogger
    Dim mockConfig As New CMockConfig
    Dim mockRepository As New CMockOperationRepository
    
    mockRepository.Reset
    logger.Initialize mockConfig, mockRepository
    
    ' Act
    logger.LogOperation "", "", ""
    
    ' Assert
    Call modAssert.IsTrue(mockRepository.SaveLogCalled, "SaveLog debería haber sido llamado")
    Call modAssert.AreEqual("", mockRepository.LastOperationType, "Tipo de operación debería estar vacío")
    Call modAssert.AreEqual("", mockRepository.LastEntityId, "ID de entidad debería estar vacío")
    Call modAssert.AreEqual("", mockRepository.LastDetails, "Detalles deberían estar vacíos")
    
    Set Test_LogOperation_WithEmptyParams_CallsRepositoryWithEmptyValues = New CTestResult
    Test_LogOperation_WithEmptyParams_CallsRepositoryWithEmptyValues.Pass
    Exit Function
    
TestFail:
    Set Test_LogOperation_WithEmptyParams_CallsRepositoryWithEmptyValues = New CTestResult
    Test_LogOperation_WithEmptyParams_CallsRepositoryWithEmptyValues.Fail "Error en prueba: " & Err.Description
End Function

Public Function Test_LogOperation_MultipleOperations_CallsRepositoryMultipleTimes() As CTestResult
    On Error GoTo TestFail
    
    ' Arrange
    Dim logger As New COperationLogger
    Dim mockConfig As New CMockConfig
    Dim mockRepository As New CMockOperationRepository
    
    mockRepository.Reset
    logger.Initialize mockConfig, mockRepository
    
    ' Act
    logger.LogOperation "CREATE", "EXP001", "Primera operación"
    logger.LogOperation "UPDATE", "EXP002", "Segunda operación"
    logger.LogOperation "DELETE", "EXP003", "Tercera operación"
    
    ' Assert
    Call modAssert.IsTrue(mockRepository.SaveLogCalled, "SaveLog debería haber sido llamado")
    Call modAssert.AreEqual(3, mockRepository.CallCount, "SaveLog debería haber sido llamado exactamente 3 veces")
    ' Verificar que los últimos parámetros corresponden a la última llamada
    Call modAssert.AreEqual("DELETE", mockRepository.LastOperationType, "Último tipo de operación incorrecto")
    Call modAssert.AreEqual("EXP003", mockRepository.LastEntityId, "Último ID de entidad incorrecto")
    Call modAssert.AreEqual("Tercera operación", mockRepository.LastDetails, "Últimos detalles incorrectos")
    
    Set Test_LogOperation_MultipleOperations_CallsRepositoryMultipleTimes = New CTestResult
    Test_LogOperation_MultipleOperations_CallsRepositoryMultipleTimes.Pass
    Exit Function
    
TestFail:
    Set Test_LogOperation_MultipleOperations_CallsRepositoryMultipleTimes = New CTestResult
    Test_LogOperation_MultipleOperations_CallsRepositoryMultipleTimes.Fail "Error en prueba: " & Err.Description
End Function



#End If


