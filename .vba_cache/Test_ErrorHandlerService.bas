Attribute VB_Name = "Test_ErrorHandlerService"
Option Compare Database
Option Explicit

' =====================================================
' Test_ErrorHandlerService.bas
' Módulo de pruebas para CErrorHandlerService
' Implementa pruebas unitarias aisladas usando mocks
' =====================================================

' Función principal para ejecutar todas las pruebas del ErrorHandlerService
Public Function Test_ErrorHandlerService_RunAll() As CTestSuiteResult
    On Error GoTo ErrorHandler
    
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "Test_ErrorHandlerService"
    
    ' Ejecutar todas las pruebas
    Call suiteResult.AddTestResult(Test_LogError_WritesToFile_Success())
    Call suiteResult.AddTestResult(Test_LogInfo_WritesToFile_Success())
    Call suiteResult.AddTestResult(Test_LogWarning_WritesToFile_Success())
    Call suiteResult.AddTestResult(Test_Initialize_WithValidConfig_Success())
    
    Set Test_ErrorHandlerService_RunAll = suiteResult
    Exit Function
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_ErrorHandlerService.Test_ErrorHandlerService_RunAll")
    Set Test_ErrorHandlerService_RunAll = Nothing
End Function

' Prueba principal: LogError escribe correctamente usando mock
Public Function Test_LogError_WritesToFile_Success() As CTestResult
    On Error GoTo ErrorHandler
    
    Dim testResult As New CTestResult
    testResult.Initialize "Test_LogError_WritesToFile_Success"
    
    ' Variables para la prueba
    Dim errorHandlerService As New CErrorHandlerService
    Dim mockConfig As New CMockConfig
    Dim mockFSO As New CMockFileSystemObject
    Dim logFilePath As String
    Dim writtenContent As String
    
    ' Arrange: Configurar el mock y el servicio
    logFilePath = "C:\Temp\test_error_log.txt"
    mockConfig.SetValue "LOG_FILE_PATH", logFilePath
    
    ' Inicializar el servicio con el mock FileSystemObject
    errorHandlerService.Initialize mockConfig, mockFSO
    
    ' Act: Llamar al método LogError
    errorHandlerService.LogError 1001, "Error de prueba", "Test_Module.Test_Function"
    
    ' Assert: Verificar que el mock capturó el contenido esperado
    writtenContent = mockFSO.GetLastWrittenContent()
    
    Call modAssert.IsTrue(mockFSO.WasOpenTextFileCalled(), "OpenTextFile debe haber sido llamado")
    Call modAssert.IsTrue(mockFSO.WasWriteLineCalled(), "WriteLine debe haber sido llamado")
    Call modAssert.IsTrue(mockFSO.WasCloseCalled(), "Close debe haber sido llamado")
    Call modAssert.IsTrue(InStr(writtenContent, "1001") > 0, "El número de error debe estar en el log")
    Call modAssert.IsTrue(InStr(writtenContent, "Error de prueba") > 0, "La descripción del error debe estar en el log")
    Call modAssert.IsTrue(InStr(writtenContent, "Test_Module.Test_Function") > 0, "La fuente del error debe estar en el log")
    Call modAssert.IsTrue(InStr(writtenContent, "ERROR") > 0, "El nivel ERROR debe estar en el log")
    
    testResult.SetResult True, "LogError escribió correctamente usando mock"
    
    Set Test_LogError_WritesToFile_Success = testResult
    Exit Function
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_ErrorHandlerService.Test_LogError_WritesToFile_Success")
    testResult.SetResult False, "Error durante la prueba: " & Err.Description
End Function

' Prueba: LogInfo escribe correctamente usando mock
Public Function Test_LogInfo_WritesToFile_Success() As CTestResult
    On Error GoTo ErrorHandler
    
    Dim testResult As New CTestResult
    testResult.Initialize "Test_LogInfo_WritesToFile_Success"
    
    ' Variables para la prueba
    Dim errorHandlerService As New CErrorHandlerService
    Dim mockConfig As New CMockConfig
    Dim mockFSO As New CMockFileSystemObject
    Dim logFilePath As String
    Dim writtenContent As String
    
    ' Arrange
    logFilePath = "C:\Temp\test_info_log.txt"
    mockConfig.SetValue "LOG_FILE_PATH", logFilePath
    
    errorHandlerService.Initialize mockConfig, mockFSO
    
    ' Act
    errorHandlerService.LogInfo "Información de prueba", "Test_Module.Test_Function"
    
    ' Assert
    writtenContent = mockFSO.GetLastWrittenContent()
    
    Call modAssert.IsTrue(mockFSO.WasOpenTextFileCalled(), "OpenTextFile debe haber sido llamado")
    Call modAssert.IsTrue(mockFSO.WasWriteLineCalled(), "WriteLine debe haber sido llamado")
    Call modAssert.IsTrue(mockFSO.WasCloseCalled(), "Close debe haber sido llamado")
    Call modAssert.IsTrue(InStr(writtenContent, "Información de prueba") > 0, "El mensaje de info debe estar en el log")
    Call modAssert.IsTrue(InStr(writtenContent, "INFO") > 0, "El nivel INFO debe estar en el log")
    
    testResult.SetResult True, "LogInfo escribió correctamente usando mock"
    
    Set Test_LogInfo_WritesToFile_Success = testResult
    Exit Function
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_ErrorHandlerService.Test_LogInfo_WritesToFile_Success")
    testResult.SetResult False, "Error durante la prueba: " & Err.Description
End Function

' Prueba: LogWarning escribe correctamente usando mock
Public Function Test_LogWarning_WritesToFile_Success() As CTestResult
    On Error GoTo ErrorHandler
    
    Dim testResult As New CTestResult
    testResult.Initialize "Test_LogWarning_WritesToFile_Success"
    
    ' Variables para la prueba
    Dim errorHandlerService As New CErrorHandlerService
    Dim mockConfig As New CMockConfig
    Dim mockFSO As New CMockFileSystemObject
    Dim logFilePath As String
    Dim writtenContent As String
    
    ' Arrange
    logFilePath = "C:\Temp\test_warning_log.txt"
    mockConfig.SetValue "LOG_FILE_PATH", logFilePath
    
    errorHandlerService.Initialize mockConfig, mockFSO
    
    ' Act
    errorHandlerService.LogWarning "Advertencia de prueba", "Test_Module.Test_Function"
    
    ' Assert
    writtenContent = mockFSO.GetLastWrittenContent()
    
    Call modAssert.IsTrue(mockFSO.WasOpenTextFileCalled(), "OpenTextFile debe haber sido llamado")
    Call modAssert.IsTrue(mockFSO.WasWriteLineCalled(), "WriteLine debe haber sido llamado")
    Call modAssert.IsTrue(mockFSO.WasCloseCalled(), "Close debe haber sido llamado")
    Call modAssert.IsTrue(InStr(writtenContent, "Advertencia de prueba") > 0, "El mensaje de warning debe estar en el log")
    Call modAssert.IsTrue(InStr(writtenContent, "WARNING") > 0, "El nivel WARNING debe estar en el log")
    
    testResult.SetResult True, "LogWarning escribió correctamente usando mock"
    
    Set Test_LogWarning_WritesToFile_Success = testResult
    Exit Function
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_ErrorHandlerService.Test_LogWarning_WritesToFile_Success")
    testResult.SetResult False, "Error durante la prueba: " & Err.Description
End Function

' Prueba: Initialize funciona correctamente con configuración válida y mock
Public Function Test_Initialize_WithValidConfig_Success() As CTestResult
    On Error GoTo ErrorHandler
    
    Dim testResult As New CTestResult
    testResult.Initialize "Test_Initialize_WithValidConfig_Success"
    
    ' Variables para la prueba
    Dim errorHandlerService As New CErrorHandlerService
    Dim mockConfig As New CMockConfig
    Dim mockFSO As New CMockFileSystemObject
    
    ' Arrange
    mockConfig.SetValue "LOG_FILE_PATH", "C:\Temp\test_init_log.txt"
    
    ' Act
    errorHandlerService.Initialize mockConfig, mockFSO
    
    ' Assert: Si no hay error, la inicialización fue exitosa
    testResult.SetResult True, "Initialize completado exitosamente con configuración válida y mock"
    
    Set Test_Initialize_WithValidConfig_Success = testResult
    Exit Function
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_ErrorHandlerService.Test_Initialize_WithValidConfig_Success")
    testResult.SetResult False, "Error durante la inicialización: " & Err.Description
End Function