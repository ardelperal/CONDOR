Option Compare Database
Option Explicit

' ============================================================================
' Módulo: Test_OperationLogger
' Descripción: Pruebas unitarias aisladas para el sistema de logging de operaciones.
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
    Call suiteResult.AddTestResult("Test_MockOperationLogger_LogOperation_RecordsEntry", Test_MockOperationLogger_LogOperation_RecordsEntry())
    Call suiteResult.AddTestResult("Test_MockOperationLogger_ClearLog_ResetsCount", Test_MockOperationLogger_ClearLog_ResetsCount())
    Call suiteResult.AddTestResult("Test_OperationLoggerFactory_ReturnsMockWhenSet", Test_OperationLoggerFactory_ReturnsMockWhenSet())
    Call suiteResult.AddTestResult("Test_OperationLoggerFactory_ReturnsRealWhenReset", Test_OperationLoggerFactory_ReturnsRealWhenReset())
    
    Set Test_OperationLogger_RunAll = suiteResult
End Function

' ============================================================================
' PRUEBAS UNITARIAS AISLADAS PARA CMockOperationLogger
' ============================================================================

Public Function Test_MockOperationLogger_LogOperation_RecordsEntry() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockLogger As New CMockOperationLogger
    Call modOperationLoggerFactory.SetMockLogger(mockLogger) ' Inyectar el mock
    
    Dim logger As IOperationLogger
    Set logger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' Act
    logger.LogOperation "TestType", "TestEntityID", "TestDetails"
    
    ' Assert
    Call modAssert.AreEqual(1, mockLogger.LoggedOperations.Count, "Debería haber 1 operación loggeada.")
    
    Dim loggedEntry As Collection
    Set loggedEntry = mockLogger.LoggedOperations(1)
    Call modAssert.AreEqual("TestType", loggedEntry("OperationType"), "Tipo de operación incorrecto.")
    Call modAssert.AreEqual("TestEntityID", loggedEntry("EntityId"), "ID de entidad incorrecto.")
    Call modAssert.AreEqual("TestDetails", loggedEntry("Details"), "Detalles incorrectos.")
    
    Test_MockOperationLogger_LogOperation_RecordsEntry = True
    Exit Function
    
TestFail:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_OperationLogger.Test_MockOperationLogger_LogOperation_RecordsEntry")
    Test_MockOperationLogger_LogOperation_RecordsEntry = False
End Function

Public Function Test_MockOperationLogger_ClearLog_ResetsCount() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockLogger As New CMockOperationLogger
    Call modOperationLoggerFactory.SetMockLogger(mockLogger)
    
    Dim logger As IOperationLogger
    Set logger = modOperationLoggerFactory.CreateOperationLogger()
    
    logger.LogOperation "Type1", "ID1", "Details1"
    logger.LogOperation "Type2", "ID2", "Details2"
    Call modAssert.AreEqual(2, mockLogger.LoggedOperations.Count, "Debería haber 2 operaciones antes de limpiar.")
    
    ' Act
    mockLogger.ClearLog()
    
    ' Assert
    Call modAssert.AreEqual(0, mockLogger.LoggedOperations.Count, "La cuenta de operaciones debería ser 0 después de limpiar.")
    
    Test_MockOperationLogger_ClearLog_ResetsCount = True
    Exit Function
    
TestFail:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_OperationLogger.Test_MockOperationLogger_ClearLog_ResetsCount")
    Test_MockOperationLogger_ClearLog_ResetsCount = False
End Function

' ============================================================================
' PRUEBAS UNITARIAS AISLADAS PARA modOperationLoggerFactory
' ============================================================================

Public Function Test_OperationLoggerFactory_ReturnsMockWhenSet() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockLogger As New CMockOperationLogger
    Call modOperationLoggerFactory.SetMockLogger(mockLogger)
    
    ' Act
    Dim logger As IOperationLogger
    Set logger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' Assert
    Call modAssert.IsTrue(TypeOf logger Is CMockOperationLogger, "El factory debería devolver una instancia de CMockOperationLogger.")
    
    Test_OperationLoggerFactory_ReturnsMockWhenSet = True
    Exit Function
    
TestFail:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_OperationLogger.Test_OperationLoggerFactory_ReturnsMockWhenSet")
    Test_OperationLoggerFactory_ReturnsMockWhenSet = False
End Function

Public Function Test_OperationLoggerFactory_ReturnsRealWhenReset() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockLogger As New CMockOperationLogger
    Call modOperationLoggerFactory.SetMockLogger(mockLogger)
    Call modOperationLoggerFactory.ResetMockLogger()
    
    ' Act
    Dim logger As IOperationLogger
    Set logger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' Assert
    Call modAssert.IsTrue(TypeOf logger Is COperationLogger, "El factory debería devolver una instancia de COperationLogger después de resetear el mock.")
    
    Test_OperationLoggerFactory_ReturnsRealWhenReset = True
    Exit Function
    
TestFail:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_OperationLogger.Test_OperationLoggerFactory_ReturnsRealWhenReset")
    Test_OperationLoggerFactory_ReturnsRealWhenReset = False
End Function
