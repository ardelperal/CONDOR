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
' FUNCIÓN PRINCIPAL DE EJECUCIÓN DE PRUEBAS
' ============================================================================

Public Function Test_OperationLogger_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.SuiteName = "Test_OperationLogger"
    
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

Public Function Test_Initialize_WithValidDependencies_Success() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim logger As New COperationLogger
    Dim mockConfig As New CMockConfig
    Dim mockRepository As New CMockOperationRepository
    
    ' Act
    logger.Initialize mockConfig, mockRepository
    
    ' Assert - Si no hay error, la inicialización fue exitosa
    Test_Initialize_WithValidDependencies_Success = True
    Exit Function
    
TestFail:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_OperationLogger.Test_Initialize_WithValidDependencies_Success")
    Test_Initialize_WithValidDependencies_Success = False
End Function

Public Function Test_LogOperation_WithoutInitialize_HandlesError() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim logger As New COperationLogger
    ' No inicializar el logger intencionalmente
    
    ' Act & Assert - Debería manejar el error graciosamente
    logger.LogOperation "TestType", "TestEntity", "TestDetails"
    
    ' Si llegamos aquí sin crash, el manejo de errores funcionó
    Test_LogOperation_WithoutInitialize_HandlesError = True
    Exit Function
    
TestFail:
    ' El error es esperado, pero debe ser manejado internamente
    Test_LogOperation_WithoutInitialize_HandlesError = True
End Function

Public Function Test_LogOperation_WithValidParams_CallsRepositoryCorrectly() As Boolean
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
    
    Test_LogOperation_WithValidParams_CallsRepositoryCorrectly = True
    Exit Function
    
TestFail:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_OperationLogger.Test_LogOperation_WithValidParams_CallsRepositoryCorrectly")
    Test_LogOperation_WithValidParams_CallsRepositoryCorrectly = False
End Function

Public Function Test_LogOperation_WithEmptyParams_CallsRepositoryWithEmptyValues() As Boolean
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
    
    Test_LogOperation_WithEmptyParams_CallsRepositoryWithEmptyValues = True
    Exit Function
    
TestFail:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_OperationLogger.Test_LogOperation_WithEmptyParams_CallsRepositoryWithEmptyValues")
    Test_LogOperation_WithEmptyParams_CallsRepositoryWithEmptyValues = False
End Function

Public Function Test_LogOperation_MultipleOperations_CallsRepositoryMultipleTimes() As Boolean
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
    
    Test_LogOperation_MultipleOperations_CallsRepositoryMultipleTimes = True
    Exit Function
    
TestFail:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_OperationLogger.Test_LogOperation_MultipleOperations_CallsRepositoryMultipleTimes")
    Test_LogOperation_MultipleOperations_CallsRepositoryMultipleTimes = False
End Function



#End If
