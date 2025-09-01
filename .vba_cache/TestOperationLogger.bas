Attribute VB_Name = "TestOperationLogger"
Option Compare Database
Option Explicit

' ============================================================================
' Módulo: TestOperationLogger
' Descripción: Pruebas unitarias para COperationLogger.
' ESTÁNDAR: Oro - Prueba la implementación real con mocks.
' ============================================================================

Public Function TestOperationLoggerRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TestOperationLogger"
    
    suiteResult.AddResult Test_LogOperation_DelegatesToRepository()
    
    Set TestOperationLoggerRunAll = suiteResult
End Function

Private Function Test_LogOperation_DelegatesToRepository() As CTestResult
    Set Test_LogOperation_DelegatesToRepository = New CTestResult
    Test_LogOperation_DelegatesToRepository.Initialize "LogOperation debe delegar la llamada a SaveLog del repositorio"
    
    ' --- Declaraciones ---
    Dim loggerImpl As COperationLogger
    Dim mockRepo As CMockOperationRepository
    Dim mockConfig As CMockConfig
    Dim mockError As CMockErrorHandlerService
    Dim logger As IOperationLogger
    
    On Error GoTo TestFail
    
    ' --- ARRANGE ---
    ' 1. Crear las dependencias mockeadas
    Set mockRepo = New CMockOperationRepository
    Set mockConfig = New CMockConfig
    Set mockError = New CMockErrorHandlerService
    
    ' 2. Instanciar la CLASE REAL que vamos a probar
    Set loggerImpl = New COperationLogger
    
    ' 3. Inyectar los mocks en la clase real
    loggerImpl.Initialize mockConfig, mockRepo, mockError
    
    ' 4. Asignar a la variable de interfaz para probar contra el contrato
    Set logger = loggerImpl
    
    ' --- ACT ---
    logger.LogOperation "TEST_OP", "ID_123", "Detalles de prueba"
    
    ' --- ASSERT ---
    ' Verificar que el método del MOCK fue llamado por la clase real
    modAssert.AssertTrue mockRepo.SaveLogCalled, "El método SaveLog del repositorio debería haber sido llamado."
    modAssert.AssertEquals 1, mockRepo.CallCount, "SaveLog debería haber sido llamado exactamente una vez."
    modAssert.AssertEquals "TEST_OP", mockRepo.LastOperationType, "El tipo de operación no se delegó correctamente."
    modAssert.AssertEquals "ID_123", mockRepo.LastEntityId, "El ID de entidad no se delegó correctamente."
    
    Test_LogOperation_DelegatesToRepository.Pass
    GoTo Cleanup
    
TestFail:
    Test_LogOperation_DelegatesToRepository.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    ' Limpieza de recursos
    Set loggerImpl = Nothing
    Set mockRepo = Nothing
    Set mockConfig = Nothing
    Set mockError = Nothing
    Set logger = Nothing
End Function