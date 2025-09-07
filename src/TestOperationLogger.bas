Attribute VB_Name = "TestOperationLogger"
Option Compare Database
Option Explicit

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
    Set mockRepo = New CMockOperationRepository
    Set mockConfig = New CMockConfig
    Set mockError = New CMockErrorHandlerService
    
    ' ¡¡CONFIGURACIÓN DEL MOCK!!
    ' Aquí le decimos al mock qué debe devolver cuando se le pregunte por "USUARIO_ACTUAL"
    mockConfig.SetSetting "USUARIO_ACTUAL", "test.user@condor.com"
    
    Set loggerImpl = New COperationLogger
    loggerImpl.Initialize mockConfig, mockRepo, mockError
    Set logger = loggerImpl
    
    ' --- ACT ---
    logger.LogOperation "TEST_OP", 123, "Detalles de prueba"
    
    ' --- ASSERT ---
    modAssert.AssertTrue mockRepo.SaveLogCalled, "El método SaveLog del repositorio debería haber sido llamado."
    modAssert.AssertEquals 1, mockRepo.CallCount, "SaveLog debería haber sido llamado exactamente una vez."
    modAssert.AssertEquals "TEST_OP", mockRepo.LastOperationType, "El tipo de operación no se delegó correctamente."
    
    Test_LogOperation_DelegatesToRepository.Pass
    GoTo Cleanup
    
TestFail:
    Test_LogOperation_DelegatesToRepository.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set loggerImpl = Nothing
    Set mockRepo = Nothing
    Set mockConfig = Nothing
    Set mockError = Nothing
    Set logger = Nothing
End Function

