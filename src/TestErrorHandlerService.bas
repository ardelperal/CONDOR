Attribute VB_Name = "TestErrorHandlerService"
Option Compare Database
Option Explicit

' =====================================================
' TestErrorHandlerService.bas
' Módulo de pruebas para CErrorHandlerService
' =====================================================

Public Function TestErrorHandlerServiceRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TestErrorHandlerService"
    
    suiteResult.AddResult TestLogError_WritesToFile_Success()
    
    Set TestErrorHandlerServiceRunAll = suiteResult
End Function

Private Function TestLogError_WritesToFile_Success() As CTestResult
    Set TestLogError_WritesToFile_Success = New CTestResult
    TestLogError_WritesToFile_Success.Initialize "LogError debe escribir en el fichero de log correctamente"
    
    Dim errorHandlerService As IErrorHandlerService
    Dim mockConfig As CMockConfig
    Dim mockFileSystem As CMockFileSystem
    Dim mockTextFile As CMockTextFile
    
    On Error GoTo TestFail
    
    ' Arrange: Instanciar la CLASE REAL a probar y sus DEPENDENCIAS MOCKEADAS
    Set errorHandlerService = New CErrorHandlerService
    Set mockConfig = New CMockConfig
    Set mockFileSystem = New CMockFileSystem
    
    ' Configurar los mocks
    mockConfig.SetSetting "LOG_FILE_PATH", "C:\temp\test.log"
    Set mockTextFile = mockFileSystem.GetMockTextFile() ' Asumimos que CMockFileSystem tiene este método auxiliar
    
    ' Act: Inicializar el servicio REAL con los MOCKS
    errorHandlerService.Initialize mockConfig, mockFileSystem
    errorHandlerService.LogError 1001, "Error de prueba", "TestModule.TestFunction"
    
    ' Assert: Verificar que los MOCKS fueron llamados correctamente
    Dim writtenContent As String
    writtenContent = mockTextFile.LastWrittenLine
    
    modAssert.AssertTrue mockFileSystem.WasOpenTextFileCalled, "OpenTextFile debe haber sido llamado"
    modAssert.AssertTrue mockTextFile.WasWriteLineCalled, "WriteLine debe haber sido llamado"
    modAssert.AssertTrue mockTextFile.WasCloseCalled, "Close debe haber sido llamado"
    modAssert.AssertTrue InStr(writtenContent, "1001") > 0, "El número de error debe estar en el log"
    modAssert.AssertTrue InStr(writtenContent, "Error de prueba") > 0, "La descripción del error debe estar en el log"
    modAssert.AssertTrue InStr(writtenContent, "TestModule.TestFunction") > 0, "La fuente del error debe estar en el log"
    
    TestLogError_WritesToFile_Success.Pass
    GoTo Cleanup
    
TestFail:
    TestLogError_WritesToFile_Success.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set errorHandlerService = Nothing
    Set mockConfig = Nothing
    Set mockFileSystem = Nothing
    Set mockTextFile = Nothing
End Function